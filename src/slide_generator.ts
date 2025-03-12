// Copyright 2016 Google Inc.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

import Debug from 'debug';
import extractSlides from './parser/extract_slides.js';
import {SlideDefinition, ImageDefinition} from './slides.js';
import matchLayout from './layout/match_layout.js';
import {URL} from 'url';
import {google, Auth, slides_v1 as SlidesV1} from 'googleapis';
import uploadLocalImage from './images/upload.js';
import probeImage from './images/probe.js';
import maybeGenerateImage from './images/generate.js';
import assert from 'assert';

const debug = Debug('md2gslides');
import fs from 'fs';
import path from 'path';
import cliProgress from 'cli-progress';

const USER_HOME =
  process.env.HOME || process.env.HOMEPATH || process.env.USERPROFILE || "";

const STORED_API_KEY_PATH = path.join(
  USER_HOME,
  '.md2googleslides',
  'fileio_key.json'
);

// wait a given number of milliseconds
async function sleep(ms: number) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}


/**
 * Generates slides from Markdown or HTML. Requires an authorized
 * oauth2 client.
 *
 * @example
 *
 *   var SlideGenerator = require('md2slides');
 *   var fs = require('fs');
 *
 *   var oauth2Client = ...; // See Google API client for details
 *   var generator = SlideGenerator.newPresentation(oauth2Client);
 *   var markdown = fs.readFileSync('mydeck.md');
 *   generator.generateFromMarkdown(markdown).then(function(id) {
 *     console.log("Presentation ID: " + id);
 *   });
 *
 * @see https://github.com/google/google-api-nodejs-client
 */
export default class SlideGenerator {
  private slides: SlideDefinition[] = [];
  private api: SlidesV1.Slides;
  private presentation: SlidesV1.Schema$Presentation;
  private allowUpload = false;
  private fileIO_key?: string;
  /**
   * @param {Object} api Authorized API client instance
   * @param {Object} presentation Initial presentation data
   * @private
   */
  public constructor(
    api: SlidesV1.Slides,
    presentation: SlidesV1.Schema$Presentation
  ) {
    this.api = api;
    this.presentation = presentation;

    // Load and parse api key from fileio_api.json file. (must 
    // be stored in ~/.md2googleslides_)
    let data; // needs to be scoped outside of try-catch
    try {
      data = fs.readFileSync(STORED_API_KEY_PATH, 'utf8');
      this.fileIO_key = JSON.parse(data).api_key;
    } catch (err) {
      console.log('Error loading api key data:', err);
    }
    if (!!this.fileIO_key && typeof this.fileIO_key !== "string") {
      this.fileIO_key = undefined;
      console.log('No valid api key was found. Image uploading will be limited');
    }
  }

  /**
   * Returns a generator that writes to a new blank presentation.
   *
   * @param {OAuth2Client} oauth2Client User credentials
   * @param {string} title Title of presentation
   * @returns {Promise.<SlideGenerator>}
   */
  public static async newPresentation(
    oauth2Client: Auth.OAuth2Client,
    title: string,
    parentId: string = "",
  ): Promise<SlideGenerator> {
    const api = google.slides({version: 'v1', auth: oauth2Client});
    const res = await api.presentations.create({
      requestBody: {
        title: title,
      },
    });
    const presentation = res.data;
    return new SlideGenerator(api, presentation);
  }

  /**
   * Returns a generator that copies an existing presentation.
   *
   * @param {OAuth2Client} oauth2Client User credentials
   * @param {string} title Title of presentation
   * @param {string} presentationId ID of presentation to copy
   * @returns {Promise.<SlideGenerator>}
   */
  public static async copyPresentation(
    oauth2Client: Auth.OAuth2Client,
    title: string,
    presentationId: string,
    parentId: string = "",
  ): Promise<SlideGenerator> {
    const drive = google.drive({version: 'v3', auth: oauth2Client});
    
    const res = await drive.files.copy({
      fileId: presentationId,
      requestBody: { name: title, parents: [parentId] }
    });

    assert(res.data.id);
    return SlideGenerator.forPresentation(oauth2Client, res.data.id);
  }

  /**
   * Returns a generator that writes to an existing presentation.
   *
   * @param {OAuth2Client} oauth2Client User credentials
   * @param {string} presentationId ID of presentation to use
   * @returns {Promise.<SlideGenerator>}
   */
  public static async forPresentation(
    oauth2Client: Auth.OAuth2Client,
    presentationId: string
  ): Promise<SlideGenerator> {
    const api = google.slides({version: 'v1', auth: oauth2Client});
    const res = await api.presentations
      .get({presentationId: presentationId})
      .catch(e => { 
        e.errors[0].message = "could not find presentation with ID="+presentationId;
        throw e;
      });
    const presentation = res.data;
    return new SlideGenerator(api, presentation);
  }

  /**
   * Generate slides from markdown
   *
   * @param {String} markdown Markdown to import
   * @param css
   * @param useFileio
   * @returns {Promise.<String>} ID of generated slide
   */
  public async generateFromMarkdown(
    markdown: string,
    {css, useFileio}: {css: string; useFileio: boolean}
  ): Promise<string> {
    assert(this.presentation?.presentationId);
    this.slides = extractSlides(markdown, css);
    this.allowUpload = useFileio;
    await this.generateImages();
    await this.probeImageSizes();
    await this.uploadLocalImages();
    await this.updatePresentation(this.createSlides());
    await this.reloadPresentation();
    await this.updatePresentation(this.populateSlides());
    return this.presentation.presentationId;
  }

  /**
   * Removes any existing slides from the presentation.
   *
   * @returns {Promise.<*>}
   */
  public async erase(): Promise<void> {
    debug('Erasing previous slides');
    assert(this.presentation?.presentationId);
    if (!this.presentation.slides) {
      return Promise.resolve();
    }

    const requests = this.presentation.slides.map(slide => ({
      deleteObject: {
        objectId: slide.objectId,
      },
    }));
    const batch = {requests};
    await this.api.presentations.batchUpdate({
      presentationId: this.presentation.presentationId,
      requestBody: batch,
    });
  }

  protected async processImages<T>(
    fn: (img: ImageDefinition) => Promise<T>,
    upload: boolean = false
  ): Promise<void> {
    const promises = [];
    const images: ImageDefinition[] = [];
    
    // collect all the background images and body images
    this.slides.forEach(slide => {
      if (slide.backgroundImage) {
        images.push(slide.backgroundImage);
      }
      slide.bodies.forEach(body =>
        body.images.forEach(image => images.push(image)));
    });

    // process each image, throttling if it's an upload
    if(upload && (images.length > 0)) {
      console.log("Uploading images for this slide deck to file.io");
      const bar = new cliProgress.SingleBar({
        format: 'Sending {value}/{total} image files to file.io',
        hideCursor: false
      });
      bar.start(images.length, 0);
      // make sure to avoid going over 8 requests/sec
      // 1) space out requests every 150ms
      // 2) if we've uploaded 7 images, wait an extra 500
      for(const [i, image] of images.entries()) {
        bar.increment();
        await sleep(150);
        if(i%6 == 0) { await sleep(250); }
        promises.push(fn(image));
      }
      bar.stop();
    } else {
      images.forEach(image => promises.push(fn(image)));
    }

    await Promise.all(promises);
    //if(upload) { console.log(JSON.stringify(images.map(i =>i.url), null, 2)); }
  }
  protected async generateImages(): Promise<void> {
    return this.processImages(maybeGenerateImage);
  }

  protected async uploadLocalImages(): Promise<void> {
    const urlCache: { [key: string]: string } = {};
    const uploadImageifLocal = async (
      image: ImageDefinition
    ): Promise<void> => {
      assert(image.url);
      const parsedUrl = new URL(image.url);

      // if it's not a file, just terminate
      if (parsedUrl.protocol !== 'file:') {
        return Promise.reject(new Error(`The url ${parsedUrl}was not a valid file`));
      }
      // if we've already uploaded it, use the cached link
      else if(urlCache[parsedUrl.pathname]) { 
        image.url = urlCache[parsedUrl.pathname];
        return Promise.resolve();
      }
      // reject the promise if we're not allowed to upload
      else if (!this.allowUpload || !this.fileIO_key) {
        return Promise.reject(new Error('Local images require --use-fileio option with a valid key!'));
      }
      else {
        image.url = await uploadLocalImage(parsedUrl.pathname, this.fileIO_key);
        urlCache[parsedUrl.pathname] = image.url;
        return Promise.resolve();
      }
    };
    return this.processImages(uploadImageifLocal, true);
  }

  /**
   * Fetches the image sizes for each image in the presentation. Allows
   * for more accurate layout of images.
   *
   * Image sizes are stored as data attributes on the image elements.
   *
   * @returns {Promise.<*>}
   * @private
   */
  protected async probeImageSizes(): Promise<void> {
    return this.processImages(probeImage);
  }

  /**
   * 1st pass at generation -- creates slides using the apporpriate
   * layout based on the content.
   *
   * Note this only returns the batch requests, but does not execute it.
   *
   * @returns {{requests: Array}}
   */
  protected createSlides(): SlidesV1.Schema$BatchUpdatePresentationRequest {
    debug('Creating slides');
    const batch = {
      requests: [],
    };
    for (const slide of this.slides) {
      const layout = matchLayout(this.presentation, slide);
      layout.appendCreateSlideRequest(batch.requests);
    }
    return batch;
  }

  /**
   * 2nd pass at generation -- fills in placeholders and adds any other
   * elements to the slides.
   *
   * Note this only returns the batch requests, but does not execute it.
   *
   * @returns {{requests: Array}}
   */
  protected populateSlides(): SlidesV1.Schema$BatchUpdatePresentationRequest {
    debug('Populating slides');
    const batch = {
      requests: [],
    };
    for (const slide of this.slides) {
      const layout = matchLayout(this.presentation, slide);
      layout.appendContentRequests(batch.requests);
    }
    return batch;
  }

  /**
   * Updates the remote presentation.
   *
   * @param batch Batch of operations to execute
   * @returns {Promise.<*>}
   */
  protected async updatePresentation(
    batch: SlidesV1.Schema$BatchUpdatePresentationRequest
  ): Promise<void> {
    debug('Updating presentation: %O', batch);
    assert(this.presentation?.presentationId);
    if (!batch.requests || batch.requests.length === 0) {
      return Promise.resolve();
    }

    /* 
      IMAGE THROTTLING
      If the slide deck includes images, we risk hitting file.io's 8 req/sec 
      limit. To deal with this, we process the requests *in-order*, splitting
      them into chunks. Each chunk is an array of requests: 
       - falling in the same order they were before
       - createImage requests happen <= MAX_IMAGES_PER_CHUNK times in each chunk
    */
    const MAX_IMAGES_PER_CHUNK = 6;  
    let requestChunks:SlidesV1.Schema$Request[][] = [];
    let currentChunk:SlidesV1.Schema$Request[] = [];
    let createImageRequestCount = 0;

    batch.requests.forEach( (req) => {
      if(req["createImage"]) { createImageRequestCount++; }
      if(createImageRequestCount > MAX_IMAGES_PER_CHUNK) {
        requestChunks.push(currentChunk);
        createImageRequestCount = 0;
        currentChunk = [req];
      } else {
        currentChunk.push(req);
      }
    });
    requestChunks.push(currentChunk);

    /* 
      Throttle the processing of these chunks, so that Google Slides
      never sends more than one chunk every 2 seconds. This keeps us
      from blowing file.io's limit in any 1 second window
    */
    const DELAY_BTW_REQUESTS = 2000; 
    const bar = new cliProgress.SingleBar({
      format: 'Sending {value}/{total} request batches to google',
      hideCursor: false
    });
    bar.start(requestChunks.length, 0);
    for await (const [i, chunk] of requestChunks.entries()) {
      bar.increment();
      batch.requests = chunk;
      let response = await this.api.presentations.batchUpdate({
        presentationId: this.presentation.presentationId,
        requestBody: batch,
      });
      debug('API response: %O', response.data);
      await sleep(DELAY_BTW_REQUESTS);
    }
    bar.stop();
  }

  /**
   * Refreshes the local copy of the presentation.
   *
   * @returns {Promise.<*>}
   */
  protected async reloadPresentation(): Promise<void> {
    assert(this.presentation?.presentationId);
    const res = await this.api.presentations.get({
      presentationId: this.presentation.presentationId,
    });
    this.presentation = res.data;
  }
}
