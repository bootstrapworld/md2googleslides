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

import {slides_v1 as SlidesV1} from 'googleapis';
import { createCanvas } from 'canvas';
import assert from 'assert';

export interface Dimensions {
  width: number;
  height: number;
}

/**
 * Locates a page by ID
 *
 * @param presentation
 * @param {string} pageId Object ID of page to find
 * @returns {Object} Page or null if not found
 */
export function findPage(
  presentation: SlidesV1.Schema$Presentation,
  pageId: string
): SlidesV1.Schema$Page | undefined {
  if (!presentation.slides) {
    return undefined;
  }
  return presentation.slides.find((p): boolean => p.objectId === pageId);
}

export function pageSize(
  presentation: SlidesV1.Schema$Presentation
): Dimensions {
  assert(presentation.pageSize?.width?.magnitude);
  assert(presentation.pageSize?.height?.magnitude);
  return {
    width: presentation.pageSize.width.magnitude,
    height: presentation.pageSize.height.magnitude,
  };
}

/**
 * Locates a layout.
 *
 * @param presentation
 * @param {string} name
 * @returns {string} layout ID or null if not found
 */
export function findLayoutIdByName(
  presentation: SlidesV1.Schema$Presentation,
  name: string
): string | undefined {
  if (!presentation.layouts) {
    return undefined;
  }
  const layout = presentation.layouts.find(
    (l): boolean => l.layoutProperties?.name === name
  );
  if (!layout) {
    return undefined;
  }
  return layout.objectId ?? undefined;
}

/**
 * Find a named placeholder on the page.
 *
 * @param presentation
 * @param {string} pageId Object ID of page to find element on
 * @param name Placeholder name.
 * @returns {Array} Array of placeholders
 */
export function findPlaceholder(
  presentation: SlidesV1.Schema$Presentation,
  pageId: string,
  name: string
): SlidesV1.Schema$PageElement[] | undefined {
  const page = findPage(presentation, pageId);
  if (!page) {
    throw new Error(`Can't find page ${pageId}`);
  }

  const placeholders = [];
  if (!page.pageElements) {
    return undefined;
  }

  // Check for textboxes (name == element.shape.placeholder.type)
  // But also check for image placeholders (name == 'PICTURE')
  for (const element of page.pageElements) {
    if (
      element.shape &&
      element.shape.placeholder &&
      name === element.shape.placeholder.type
    ) {
      placeholders.push(element);
    }
    if (element.image && 
      element.image.placeholder && 
      name == element.image.placeholder.type) {
      placeholders.push(element);
    }
  }

  if (placeholders.length) {
    return placeholders;
  }

  return undefined;
}

export function findSpeakerNotesObjectId(
  presentation: SlidesV1.Schema$Presentation,
  pageId: string
): string | undefined {
  const page = findPage(presentation, pageId);
  if (page) {
    return (
      page.slideProperties?.notesPage?.notesProperties?.speakerNotesObjectId ??
      undefined
    );
  }
  return undefined;
}

// Added by Emmanuel Schanzer 1/4/24
// Given a PageElement, produce the parent element so we can extract the inherited properties
// if there's no parent, return undefined.
export function findParentObject(
  presentation: SlidesV1.Schema$Presentation,
  pageElement: SlidesV1.Schema$PageElement,
): SlidesV1.Schema$PageElement | undefined {

  // if it's a placeholder, lookup the parent shape whose style we need to match
  const parentObjectId = pageElement?.shape?.placeholder?.parentObjectId;
  if(parentObjectId) {
    // flatten all the pageElts across all slides, masters, and layouts, then search
    return presentation.slides?.map(_ => _.pageElements)
      .concat(presentation.masters?.map(_ => _.pageElements))
      .concat(presentation.layouts?.map(_ => _.pageElements))
      .flat()
      .find(elt => elt?.objectId == parentObjectId);    
  } else {
    return;
  }
}


// Added by Emmanuel Schanzer 2/5/23
// based on code from https://stackoverflow.com/questions/75228506/google-slides-autofit-text-alternative-calculate-based-on-dimensions-of-elem/75278719#75278719
// this code had a few bugs, but more importantly we try to measure a sample of the actual text
// instead of a single 'W' character

// ASSUMED DEFAULTS (all values are in PT unless otherwise specified)
const DEFAULT_FONT_WEIGHT  = 'normal';
const DEFAULT_FONT_SIZE    = 16;

// An English Metric Unit (EMU) is defined as 1/360,000 of a centimeter and thus there are 914,400 EMUs per inch, and 12,700 EMUs per point.
export const convertEMUtoPT = (emu: number): number => emu / 12700;
// convert pixles to PT, there is 0.75pt to a px
export const convertPXtoPT = (px: number): number => px * 0.75;
// convert PT to PX, there is 0.75pt to a px
export const convertPTtoPX = (px: number): number => px / 0.75;
export const convertEMUtoPX = (emu: number): number => convertPTtoPX(convertEMUtoPT(emu));

// this is a very simple example of what i have, obviously you'll need error handling if those values don't exist
// The below will return the dimensions in EMU, to convert to PT divide the EMU value by `12700`
export function getElementSizePT(element: SlidesV1.Schema$PageElement) {
    assert(element);
    assert(element.size?.height?.magnitude);
    assert(element.size?.width?.magnitude);
    assert(element.transform?.scaleX);
    assert(element.transform?.scaleY);
    const width = element.size.width.magnitude * element.transform?.scaleX;
    const height = element?.size?.height?.magnitude * element.transform?.scaleY;
    return { width: Math.round(convertEMUtoPT(width)), height: Math.round(convertEMUtoPT(height)) };
}
/**
 * @name findByKey
 * @description This was introduced as the fontWeight key for example could be on a mixture of elements, and we
 * want to find them whereever they may be on the element so we can average out the values
 * @function
 * @param obj - any object to search
 * @param kee - representing the needle to search
 * @returns any - returns the value by the key if found
 */
export const findByKey = (obj: any, kee: string): any | undefined => {
  if (kee in obj) {
    return obj[kee];
  }
  for (const n of Object.values(obj).filter(Boolean).filter(v => typeof v === 'object')) {
      const found = findByKey(n, kee);
      if (typeof found !== 'undefined') {
        return found;
      }
  }
};

/**
 * @name splitter
 * @description Based on the maximum allowed characters on a single line, we split the lines
 * based on this value so we can calculate multi line text wrapping and adjust the font size
 * continually within a while loop
 * @function
 * @param str - the input string
 * @param l - the length of each "line" of text
 * @returns string[] - an array of strings representing each new line of text
 */

export function splitter(str: string, l: number): string[] {
  const strs = [];
  while (str.length > l) {
      let pos = str.substring(0, l).lastIndexOf(' ');
      pos = pos <= 0 ? l : pos;
      strs.push(str.substring(0, pos));
      let i = str.indexOf(' ', pos) + 1;
      if (i < pos || i > pos + l)
          i = pos;
      str = str.substring(i);
  }
  strs.push(str);
  return strs;
}

// An amalgamation of text and paragraph style objects, whose fields do not overlap
// used when computing the measurements of a textRun, which are impacted by fields from both
// irrelevant fields like color are included for completeness
const DEFAULT_STYLE = {
  backgroundColor: {},
  foregroundColor: {},
  bold: false,
  italic: false,
  fontFamily: "Arial",
  fontSize: { magnitude: 16, unit: 'PT' },
  link: undefined,
  baselineOffset: 0,
  smallCaps: false,
  strikethrough: false,
  underline: false,
  weightedFontFamily: { fontFamily: "Arial", weight: 400 },
  lineSpacing: 115, // % of line height (115 = 14pt lines are separated by 14pt*115%)
  alignment: 'START',
  indentStart: { magnitude: 0, unit: 'PT' },
  indentEnd:   { magnitude: 0, unit: 'PT' },
  spaceAbove:  { magnitude: 0, unit: 'PT' },
  spaceBelow: { magnitude: 16, unit: 'PT' },
  indentFirstLine: { magnitude: 0, unit: 'PT' },
  direction: 'LEFT_TO_RIGHT',
  spacingMode: 'NEVER_COLLAPSE'
}

export function calculateFontSize(
  ancestors: SlidesV1.Schema$PageElement[],  // oldest-to-youngest
  text: string, 
  constraints: string): number {

  // ancestor 0 is the elt we're trying to fit
  const element = ancestors[0];
  const sizePT = getElementSizePT(element);

  //console.log(`fitting "${text}" into ${sizePT.width}pt x ${sizePT.height}pt. Constraints are ${constraints}`);

  // create a canvas with the same size as the element, this most likely does not matter as we're only measuring a fake
  // representation of the text with ctx.measureText
  const canvas = createCanvas(10000, 10000);
  const ctx = canvas.getContext('2d');
  // try to extract all the font-sizes
  const fontSizes = element.shape?.text?.textElements?.map(textElement => textElement.textRun?.style?.fontSize?.magnitude).filter((a): a is number => Number.isInteger(a)) ?? [];
  
  // get the first textElement of each ancestor (use '.at' instead of '[n]' to salvage the ? operator)
  // then merge the pmarkers from oldest to youngest, so children can override parents
  const pmarkers = ancestors.map(a => a.shape?.text?.textElements?.at(0)?.paragraphMarker?.style);
  const computedStyle = Object.assign({}, DEFAULT_STYLE, ...pmarkers);

  //console.log('inheritedProps', computedStyle)

  // try to extract all the font-weights
  const fontWeights = element.shape?.text?.textElements?.map(textElement => textElement.textRun?.style?.weightedFontFamily?.weight).filter((a): a is number => Number.isInteger(a)) ?? [];
  // fallback to arial if not found, if there's more than one fontFamily used in a single element, we just pick the first one, no way i can think of
  // to be smart here and not really necessary to create multiple strings with different fonts and calculate those, this seems to work fine
  const fontFamily = findByKey(element, 'fontFamily') ?? 'Arial';
  // calulate the average as there can be different fonts with different weights within a single text element
  const averageFontWeight = fontWeights.reduce((a, n) => a + n, 0) / fontWeights.length;
  const averageFontSize = fontSizes.reduce((a, n) => a + n, 0) / fontSizes.length;
  // if the average font-weight is not a number, usae the default
  const fontWeight = isNaN(averageFontWeight) ? DEFAULT_FONT_WEIGHT : averageFontWeight;
  // use the average fontSize if available, else start at an arbitrary default
  let fontSize = isNaN(averageFontSize) ? computedStyle.fontSize.magnitude : averageFontSize;
  // if the input value is an empty string, don't bother with any calculations
  if (text.length === 0) { return fontSize; }

  const MIN_FONT_SIZE = 12;

  const WIDTH_ADJUSTMENT = 1.15; // smaller adj = larger font

  const estCharsPerLine = (): number => {
    const numChars = Math.min(text.length, 100);
    const avgCharWidth = ctx.measureText(text.substring(0, numChars)).width / numChars;
    return convertPTtoPX(sizePT.width) / (avgCharWidth * WIDTH_ADJUSTMENT);
  }

  // convenience function: given a property, get its magnitude or produce zero
  function amountPT(prop:string){ return computedStyle[prop].magnitude || 0; }

  // used for the while loop, to continually resize the shape and multiline text, until it fits within the bounds
  // of the element
  const isOutsideBounds = (): boolean => {
    ctx.font = `${fontWeight} ${fontSize}pt ${fontFamily}`;
    // based on the maximum amount of chars available in the horizontal axis for this font size
    // we split onto a new line to get the width/height correctly
    const lines = splitter(text, estCharsPerLine());

    // if the only constraint is horizontal, we're outside bounds if there's more than one line
    if ((constraints == "horizontal") && (lines.length > 1)) {
      return true;
    }

    // compute the horizontal and vertical whitespace in PT
    // Horizontal: <indentStart> + <indentEnd>
    // Vertical: <space below and after all n lines> + <spacing between each line (n-1)>
    const horizontalWhiteSpace = amountPT("indentStart") + amountPT("indentEnd");
    const spaceAroundLines = lines.length * (amountPT("spaceBelow") + amountPT("spaceAbove"))
    const spaceBetweenLines = (lines.length-1) * (computedStyle.lineSpacing/100 * fontSize)
    const verticalWhiteSpace = spaceAroundLines + spaceBetweenLines;

    // get the measurements of the current multiline string
    const metrics = ctx.measureText(lines.join('\n'));
    // the emAcent/Decent values do exist, it's the types that are wrong from canvas
    // @ts-expect-error
    const emAcent = metrics.emHeightAscent as number;
    // @ts-expect-error
    const emDecent = metrics.emHeightDescent as number;
    // get dimensions in PT, and compare to element size
    const height = convertPXtoPT(emAcent + emDecent) + verticalWhiteSpace;
    const width  = convertPXtoPT(metrics.width) + horizontalWhiteSpace;
    //console.log(`at ${fontSize}pt, text is ${width} x ${height}`)
    return width > sizePT.width || height > sizePT.height;
  };
  // continually loop over until the size of the text element is within bounds,
  // decreasing by 0.25pt increments until it fits within the width
  while (isOutsideBounds() && (fontSize > MIN_FONT_SIZE)) { fontSize = fontSize - 0.25; }
  
  return fontSize;
}
