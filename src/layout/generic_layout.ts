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
import {v1 as uuid} from 'uuid';
// @ts-ignore
import Layout from 'layout';
import {slides_v1 as SlidesV1} from 'googleapis';
import {
  ImageDefinition,
  SlideDefinition,
  TableDefinition,
  TextDefinition,
  VideoDefinition,
} from '../slides.js';
import {
  findLayoutIdByName,
  findPlaceholder,
  findSpeakerNotesObjectId,
  calculateFontSize,
} from './presentation_helpers.js';
import assert from 'assert';

const debug = Debug('md2gslides');

// from https://stackoverflow.com/a/66543738/12026982
const EMUperPixel = 9525;

interface BoundingBox {
  height: number;
  width: number;
  x: number;
  y: number;
}

/**
 * Performs most of the work of converting a slide into API requests.
 *
 */
export default class GenericLayout {
  public name: string;
  public presentation: SlidesV1.Schema$Presentation;
  private slide: SlideDefinition;

  public constructor(
    name: string,
    presentation: SlidesV1.Schema$Presentation,
    slide: SlideDefinition
  ) {
    this.name = name;
    this.presentation = presentation;
    this.slide = slide;

    const addPageEltsToMap = (page:SlidesV1.Schema$Page) => {
      page?.pageElements?.forEach(e => this.pageElementMap.set(e.objectId, e));
    }
    presentation.slides?.forEach(addPageEltsToMap);
    presentation.masters?.forEach(addPageEltsToMap);
    presentation.layouts?.forEach(addPageEltsToMap);
  }

  private pageElementMap = new Map();

  // Added by Emmanuel Schanzer 1/4/24
  // Given a PageElement, produce the parent element so we can extract the inherited properties
  // if there's no parent, return undefined.
  private findParentObject(
    pageElement: SlidesV1.Schema$PageElement,
  ): SlidesV1.Schema$PageElement | undefined {

    // if it's a placeholder, lookup the parent shape whose style we need to match
    const parentObjectId = pageElement?.shape?.placeholder?.parentObjectId;
    return this.pageElementMap.get(parentObjectId);
  }

  public appendCreateSlideRequest(
    requests: SlidesV1.Schema$Request[]
  ): SlidesV1.Schema$Request[] {
    const layoutId = findLayoutIdByName(this.presentation, this.name);
    if (!layoutId) {
      throw new Error(`Unable to find layout ${this.name}`);
    }
    this.slide.objectId = uuid();

    debug('Creating slide %s with layout %s', this.slide.objectId, this.name);
    requests.push({
      createSlide: {
        slideLayoutReference: {
          layoutId: layoutId,
        },
        objectId: this.slide.objectId,
      },
    });
    return requests;
  }

  // called once, for each slide!
  public appendContentRequests(
    requests: SlidesV1.Schema$Request[]
  ): SlidesV1.Schema$Request[] {
    this.appendFillPlaceholderTextRequest(
      this.slide.title, 
      'TITLE', 
      requests, 
      "horizontal"
    );
    this.appendFillPlaceholderTextRequest(
      this.slide.title,
      'CENTERED_TITLE',
      requests
    );
    this.appendFillPlaceholderTextRequest(
      this.slide.subtitle,
      'SUBTITLE',
      requests
    );

    if (this.slide.backgroundImage) {
      this.appendSetBackgroundImageRequest(
        this.slide.backgroundImage,
        requests
      );
    }

    if (this.slide.tables.length) {
      this.slide.tables.forEach(t => this.appendCreateTableRequest(t, requests));
    }

    if (this.slide.bodies) {
      assert(this.slide.objectId);
      const bodyElements = findPlaceholder(
        this.presentation,
        this.slide.objectId,
        'BODY',
      );
      const bodyCount = Math.min(
        bodyElements?.length ?? 0,
        this.slide.bodies.length
      );

      // grab all the image placeholders for this slide
      const imagePlaceholders = findPlaceholder(
        this.presentation,
        this.slide.objectId,
        'PICTURE'
      ) || [];

      for (let i = 0; i < bodyCount; ++i) {
        const placeholder = bodyElements![i];
        const body = this.slide.bodies[i];
        this.appendFillPlaceholderTextRequest(body.text, placeholder, requests, "vertical");

        if (body.images && body.images.length) {
          // send all the images, and just the first placeholder
          this.appendCreateImageRequests(body.images, imagePlaceholders, requests);
        }
        if (body.videos && body.videos.length) {
          this.appendCreateVideoRequests(body.videos, placeholder, requests);
        }
      }

      // if there were any, remove all image placeholders from the slide
      imagePlaceholders.forEach(p => 
        requests.push({'deleteObject': {'objectId': p!['objectId']}}));

    }

    if (this.slide.notes) {
      assert(this.slide.objectId);
      const objectId = findSpeakerNotesObjectId(
        this.presentation,
        this.slide.objectId
      );
      this.appendInsertTextRequests(
        this.slide.notes,
        {objectId: objectId},
        requests
      );
    }

    return requests;
  }

  protected appendFillPlaceholderTextRequest(
    value: TextDefinition | undefined,
    placeholder: string | SlidesV1.Schema$PageElement,
    requests: SlidesV1.Schema$Request[],
    constraints?: string
  ): void {
    if (!value) {
      debug('No text for placeholder %s');
      return;
    }
    if (typeof placeholder === 'string') {
      assert(this.slide.objectId);
      const pageElements = findPlaceholder(
        this.presentation,
        this.slide.objectId,
        placeholder
      );
      if (!pageElements) {
        debug('Skipping undefined placeholder %s', placeholder);
        return;
      }
      placeholder = pageElements[0];
    }

    // compute an array of ancestor elts, from oldest-to-youngest
    let ancestors = [];
    let parentObject = this.findParentObject(placeholder);
    while (parentObject) {
      ancestors.unshift(parentObject); // add to beginning
      parentObject = this.findParentObject(parentObject);
    }

    this.appendInsertTextRequests(
      value,
      {objectId: placeholder.objectId},
      requests,
      constraints, // passing constraints turns on auto-fitting
      ancestors
    );
  }

  protected appendInsertTextRequests(
    text: TextDefinition,
    locationProps:
      | Partial<SlidesV1.Schema$UpdateTextStyleRequest>
      | Partial<SlidesV1.Schema$CreateParagraphBulletsRequest>,
    requests: SlidesV1.Schema$Request[],
    constraints?: string,
    ancestors?: SlidesV1.Schema$PageElement[],
  ): void {

    // make sure there's text to insert
    if(text.rawText.trimLeft().length == 0) {
      return; // empty requests are invalid, so don't bother
    }

    // Insert the raw text first
    const request = {
      insertText: Object.assign(
        {
          text: text.rawText
        },
        locationProps
      ),
    };
    requests.push(request);


    // Apply any text styles present.
    // Most of the work for generating the text runs
    // is performed when parsing markdown.
    text.textRuns.reverse();
    for (const textRun of text.textRuns) {

      const request: SlidesV1.Schema$Request = {
        updateTextStyle: Object.assign(
          {
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: textRun.start,
              endIndex: textRun.end,
            },
            style: {
              bold: textRun.bold,
              italic: textRun.italic,
              foregroundColor: textRun.foregroundColor,
              backgroundColor: textRun.backgroundColor,
              strikethrough: textRun.strikethrough,
              underline: textRun.underline,
              smallCaps: textRun.smallCaps,
              fontFamily: textRun.fontFamily,
              // if we're autofitting, call out to calculateFontSize
              // to render a canvas element and estimate size
              fontSize: (ancestors?.length && constraints && !textRun.fontSize)? {
                magnitude: calculateFontSize(ancestors, text, constraints),
                unit: 'PT'
              } : textRun.fontSize,
              link: textRun.link,
              baselineOffset: textRun.baselineOffset,
            },
          },
          locationProps
        ),
      };
      assert(request.updateTextStyle?.style);
      request.updateTextStyle.fields = this.computeShallowFieldMask(
        request.updateTextStyle.style
      );

      if (request.updateTextStyle.fields.length) {
        requests.push(request); // Only push if at least one style set
      }
    }

    // Convert paragraphs to lists.
    // Note that leading tabs for nested lists in the raw text are removed.
    // In this case, we're assuming that lists are supplied in order of
    // appearance and they're non-overlapping.
    // Processing in the reverse order avoids having to readjust indices.
    for (const listMarker of text.listMarkers.reverse()) {
      const request = {
        createParagraphBullets: Object.assign(
          {
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: listMarker.start,
              endIndex: listMarker.end,
            },
            bulletPreset:
              listMarker.type === 'ordered'
                ? 'NUMBERED_DIGIT_ALPHA_ROMAN'
                : 'BULLET_DISC_CIRCLE_SQUARE',
          },
          locationProps
        ),
      };
      requests.push(request);
    }
  }

  protected appendSetBackgroundImageRequest(
    image: ImageDefinition,
    requests: SlidesV1.Schema$Request[]
  ): void {
    debug(
      'Slide #%d: setting background image to %s',
      this.slide.index,
      image.url
    );

    requests.push({
      updatePageProperties: {
        objectId: this.slide.objectId,
        fields: 'pageBackgroundFill.stretchedPictureFill.contentUrl',
        pageProperties: {
          pageBackgroundFill: {
            stretchedPictureFill: {
              contentUrl: image.url,
            },
          },
        },
      },
    });
  }

  protected appendCreateImageRequests(
      images: ImageDefinition[],
      placeholders: SlidesV1.Schema$PageElement[],
      requests: SlidesV1.Schema$Request[]
    ): void {
      const that = this;

      // Convert pixels to EMUs. If there's a placeholder, resize to fit.
      function transformImageSize(
        image: ImageDefinition, 
        placeholder: SlidesV1.Schema$PageElement
      ): void {
        // TODO - Fix weird cast
        const layer = (Layout as (s: string) => Layout.PackingSmith)('left-right'); // TODO - Configurable?

        debug('Slide #%d: adding inline image %s', that.slide.index, image.url);
        layer.addItem({
          width: image.width + image.padding * 2,
          height: image.height + image.padding * 2,
          meta: image,
        });

        // bounding box of the entire slide
        const {width: slideWidth, height: slideHeight} = (that.presentation.pageSize)!;

        // if there's a placeholder, use it as a bounding box
        // otherwise: if the image is smaller than 50% of the slide
        // use actual size. Otherwise 
        let box;
        if(placeholder) {
          box = that.getBodyBoundingBox(placeholder);
        } else {
          const maxWidth  = slideWidth!.magnitude!  * 0.5;
          const maxHeight = slideHeight!.magnitude! * 0.5;
          const imgWidth  = Math.min(image.width  * EMUperPixel, maxWidth);
          const imgHeight = Math.min(image.height * EMUperPixel, maxHeight);
          box = {
            width:  imgWidth,
            height: imgHeight,
            x: maxWidth - imgWidth / 2, 
            y: maxHeight - imgHeight / 2, 
          };
        }

        const computedLayout = layer.export();

        // assume we're just converting Pixels->EMU, but scale to fit
        // box will be the dimensions of the placeholder (if defined),
        // or the entire slide
        let scaleRatio = EMUperPixel;  
        scaleRatio = Math.min(
          box.width  / computedLayout.width,
          box.height / computedLayout.height
        );

        const scaledWidth  = computedLayout.width  * scaleRatio;
        const scaledHeight = computedLayout.height * scaleRatio;

        const baseTranslateX = box.x + (box.width  - scaledWidth)  / 2;
        const baseTranslateY = box.y + (box.height - scaledHeight) / 2;

        if(computedLayout.items.length > 1) {
          console.error('IMPOSSIBLE - multiple images in transformAndReplacePlaceholder');
        }
        const item = computedLayout.items[0];

        const itemOffsetX = item.meta.offsetX ? item.meta.offsetX : 0;
        const itemOffsetY = item.meta.offsetY ? item.meta.offsetY : 0;
        const itemPadding = item.meta.padding ? item.meta.padding : 0;
        const width  = item.meta.width  * scaleRatio;
        const height = item.meta.height * scaleRatio;
        const translateX =
          baseTranslateX + (item.x + itemPadding + itemOffsetX) * scaleRatio;
        const translateY =
          baseTranslateY + (item.y + itemPadding + itemOffsetY) * scaleRatio;
        const imageId = uuid();

        // add the image at about the same size/position as the placeholder
        const createImageReq = {
          createImage: {
            objectId: imageId,
            elementProperties: {
              pageObjectId: that.slide.objectId,
              size: {
                height: { magnitude: height, unit: 'EMU'},
                width:  { magnitude: width,  unit: 'EMU'},
              },
              transform: {
                scaleX: 1, scaleY: 1,
                translateX: translateX, translateY: translateY,
                shearX: 0, shearY: 0,
                unit: 'EMU',
              },
            },
            url: item.meta.url,
          },
        };
        requests.push(createImageReq);

        // add the altText
        const altTextReq = {
          updatePageElementAltText: {
            "objectId": imageId,
            "title": "",
            "description": image.altText,
          }
        };
        requests.push(altTextReq);
      }

      images.forEach((image, i) => {
        debug('Slide #%d: adding inline image %s', this.slide.index, image.url);
        const placeholder = placeholders[i] || undefined;
        transformImageSize(image, placeholder)
      });
    }

  protected appendCreateVideoRequests(
    videos: VideoDefinition[],
    placeholder: SlidesV1.Schema$PageElement | undefined,
    requests: SlidesV1.Schema$Request[]
  ): void {
    if (videos.length > 1) {
      throw new Error('Multiple videos per slide are not supported.');
    }
    const video = videos[0];

    debug('Slide #%d: adding video %s', this.slide.index, video.id);

    const box = this.getBodyBoundingBox(placeholder);

    const scaleRatio = Math.min(
      box.width / video.width,
      box.height / video.height
    );

    const scaledWidth = video.width * scaleRatio;
    const scaledHeight = video.height * scaleRatio;

    const translateX = box.x + (box.width - scaledWidth) / 2;
    const translateY = box.y + (box.height - scaledHeight) / 2;

    const objectId = uuid();
    requests.push({
      createVideo: {
        source: 'YOUTUBE',
        objectId: objectId,
        id: video.id,
        elementProperties: {
          pageObjectId: this.slide.objectId,
          size: {
            height: {
              magnitude: scaledHeight,
              unit: 'EMU',
            },
            width: {
              magnitude: scaledWidth,
              unit: 'EMU',
            },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: translateX,
            translateY: translateY,
            shearX: 0,
            shearY: 0,
            unit: 'EMU',
          },
        },
      },
    });
    requests.push({
      updateVideoProperties: {
        objectId: objectId,
        fields: 'autoPlay',
        videoProperties: {
          autoPlay: video.autoPlay,
        },
      },
    });
  }

  protected appendCreateTableRequest(
    table: TableDefinition,
    requests: SlidesV1.Schema$Request[]
  ): void {
    const tableId = uuid();
    let hasHeader = true;

    // check to see if the first cell of the first row asks us to delete
    // necessary to handle racket workaround for the markdown limitation
    // that all tables have headers
    if(table.cells[0][0].rawText == "DELETE THIS ROW") {
      hasHeader = false;
    }
    // If it's a table with no header...
    if(!hasHeader) {
      table.cells.shift();               // delete the row
      table.rows = table.rows - 1;       // update row count
    } 

    requests.push({
      createTable: {
        objectId: tableId,
        elementProperties: {
          pageObjectId: this.slide.objectId,
          // Use default size/transform for tables
        },
        rows: table.rows,
        columns: table.columns,
      },
    });

    for (const r in table.cells) {
      const row = table.cells[r];
      for (const c in row) {
        this.appendInsertTextRequests(
          row[c],
          {
            objectId: tableId,
            cellLocation: {
              rowIndex: parseInt(r),
              columnIndex: parseInt(c),
            },
          },
          requests
        );
      }
    }

    // if there's a header, update the formatting
    // from cell (0,0) through table.columns, give it a half-gray BG
    requests.push({
      updateTableCellProperties: {
        objectId: tableId,
        tableRange: {
          location: {
            rowIndex: 0,
            columnIndex: 0
          },
          rowSpan: 1,
          columnSpan: table.columns
        },
        tableCellProperties: {
          tableCellBackgroundFill: {
            solidFill: {
              color: {
                rgbColor: {
                  red: 0.75,
                  green: 0.75,
                  blue: 0.75
                }
              }
            }
          }
        },
        fields: "tableCellBackgroundFill.solidFill.color"
      }
    });
  }

  protected calculateBoundingBox(
    element: SlidesV1.Schema$PageElement
  ): BoundingBox {
    assert(element);
    assert(element.size?.height?.magnitude);
    assert(element.size?.width?.magnitude);
    const height = element.size.height.magnitude;
    const width = element.size.width.magnitude;
    const scaleX = element.transform?.scaleX ?? 1;
    const scaleY = element.transform?.scaleY ?? 1;
    const shearX = element.transform?.shearX ?? 0;
    const shearY = element.transform?.shearY ?? 0;

    return {
      width: scaleX * width + shearX * height,
      height: scaleY * height + shearY * width,
      x: element.transform?.translateX ?? 0,
      y: element.transform?.translateY ?? 0,
    };
  }

  protected getBodyBoundingBox(
    placeholder: SlidesV1.Schema$PageElement | undefined
  ): BoundingBox {
    if (placeholder) {
      return this.calculateBoundingBox(placeholder);
    }
    assert(this.presentation.pageSize?.width?.magnitude);
    assert(this.presentation.pageSize?.height?.magnitude);
    return {
      width: this.presentation.pageSize.width.magnitude,
      height: this.presentation.pageSize.height.magnitude,
      x: 0,
      y: 0,
    };
  }

  protected computeShallowFieldMask<T extends object>(object: T): string {
    const fields = [];
    for (const field of Object.keys(object)) {
      if (object[field as keyof T] !== undefined) {
        fields.push(field);
      }
    }
    return fields.join(',');
  }
}
