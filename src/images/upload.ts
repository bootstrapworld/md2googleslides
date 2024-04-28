// Copyright 2019 Google Inc.
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
import fs from 'fs';
import * as path from 'path';
import { blob } from 'node:stream/consumers';
import { FormData } from 'formdata-polyfill/esm.min.js'

const debug = Debug('md2gslides');

/**
 * Uploads a local file to temporary storage so it is HTTP/S accessible.
 *
 * Currently uses https://file.io for free emphemeral file hosting.
 *
 * @param {string} filePath -- Local path to image to upload
 * @returns {Promise<string>} URL to hosted image
 */
async function uploadLocalImage(filePath: string, key?: string): Promise<string> {
  debug('Registering file %s', filePath);
  const stream = fs.createReadStream(filePath);

  try {
    const data = new FormData();
    data.append('file', await blob(stream), path.basename(filePath));
    data.append('expires', '5m');
    data.append('autoDelete', 'true');

    const request : RequestInfo = new Request('https://file.io', {
      method: 'POST',
      body: data,
    });

    // add the authorization key, if one is defined
    if(key) {
      request.headers.append("Authorization", key);
    }

    // Make a POST request using fetch
    return fetch('https://file.io', request)
      .then(response => response.json())
      .then(responseJSON => {
        if(!responseJSON.success){ throw responseJSON; }
        debug('Temporary link: %s', responseJSON.link);
        return responseJSON.link;
      })
      .catch(error => {
        debug('Unable to upload file: %O', error);
        if(error.status == 492 || error.code == 'TOO_MANY_REQUESTS') {
          console.error(`\n\n❌ Too many image requests/sec with file.io! 
   Someone else probably is upoading images with the same API key right now.
   Please wait a few seconds and try again.\n\n`);
        }
        console.error('Error uploading file:', error);
        throw error;
      });
  } finally {
    stream.destroy();
  }
}

export default uploadLocalImage;