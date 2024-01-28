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
import { TextDefinition } from '../slides.js'

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

// Code below added by Emmanuel Schanzer 
// based on code from https://stackoverflow.com/questions/75228506/google-slides-autofit-text-alternative-calculate-based-on-dimensions-of-elem/75278719#75278719
// - enhanced to use smarter calculation of average chars (use an avg'd sample instead of 'W')
// - enhanced to use computed style objects (takes into acct indenting, lineSpacing, and more)
// - enhanced to use a minimum font size

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
      strs.push(...str.substring(0, pos).split('\n'));
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
  spaceBelow:  { magnitude: 0, unit: 'PT' },
  indentFirstLine: { magnitude: 0, unit: 'PT' },
  direction: 'LEFT_TO_RIGHT',
  spacingMode: 'NEVER_COLLAPSE'
}

const MIN_SIZE = 14; // Anything smaller than 14pt is not readable on a projector

const DEFAULT_PADDING = 0.2 * 72;  // 72pt per inch, assume 0.1in padding on all sizes

// in practice, characters seem to be roughly 1.15x wider in GSlides than in canvas elt
const WTF_CHAR_WIDTH_HACK = 1.15; 

// NOTE(Emmanuel): probably unneeded if we ever fix the regular markdown parser
const cachedFontCalculations = new Map();
  
export function calculateFontSize(
  ancestors: SlidesV1.Schema$PageElement[],  // oldest-to-youngest
  text: TextDefinition,
  constraints: string): number {

  // the youngest (last) ancestor is the actual element we're trying to fit
  const element = ancestors[ancestors.length-1];

  // check to see if we've already done this work
  // NOTE(Emmanuel): probably unneeded if we ever fix the regular markdown parser
  const key = text.rawText + element.objectId;
  if(cachedFontCalculations.has(key)) { 
    return cachedFontCalculations.get(key); 
  }

  // Dreate a canvas with the same size as the element. This probably doesn't matter,
  // as we're only measuring a fake representation of the text with ctx.measureText
  const sizePT = getElementSizePT(element);
  // adjust the size to account for space lost to padding
  sizePT.width -= DEFAULT_PADDING;
  sizePT.height -= DEFAULT_PADDING;
  const canvas = createCanvas(sizePT.width, sizePT.height);
  const ctx = canvas.getContext('2d');

  
  // starting with the default, merge style rules from oldest-to-youngest
  const computedStyle = Object.assign(
    {}, DEFAULT_STYLE, // use '.at' instead of '[n]' below to salvage the ? operator
    ...ancestors.map(a => a.shape?.text?.textElements?.at(0)?.paragraphMarker?.style)
  );
  //console.log('inheritedProps', computedStyle)

  // try to extract all the font-sizes
  const fontSizes = element.shape?.text?.textElements?.map(textElement => textElement.textRun?.style?.fontSize?.magnitude).filter((a): a is number => Number.isInteger(a)) ?? [];
  // try to extract all the font-weights
  const fontWeights = element.shape?.text?.textElements?.map(textElement => textElement.textRun?.style?.weightedFontFamily?.weight).filter((a): a is number => Number.isInteger(a)) ?? [];
  // fallback to arial if not found, if there's more than one fontFamily used in a single element, we just pick the first one, no way i can think of
  // to be smart here and not really necessary to create multiple strings with different fonts and calculate those, this seems to work fine
  const fontFamily = findByKey(element, 'fontFamily') ?? computedStyle.fontFamily;
  // calulate the average as there can be different fonts with different weights within a single text element
  const averageFontWeight = fontWeights.reduce((a, n) => a + n, 0) / fontWeights.length;
  const averageFontSize = fontSizes.reduce((a, n) => a + n, 0) / fontSizes.length;
  // if the average font-weight is not a number, use the default
  const fontWeight = isNaN(averageFontWeight) ? computedStyle.weightedFontFamily.weight : averageFontWeight;
  // use the average fontSize if available, else start at an arbitrary default
  let fontSize = isNaN(averageFontSize) ? computedStyle.fontSize.magnitude : averageFontSize;
  // if the input value is an empty string, don't bother with any calculations
  if (text.rawText.length === 0) { return fontSize; }

  const estCharsPerLine = (): number => {
    const numChars = Math.min(text.rawText.length, 100);
    const avgCharWidth = ctx.measureText(text.rawText.substring(0, numChars)).width / numChars;
    return convertPTtoPX(sizePT.width) / (avgCharWidth * WTF_CHAR_WIDTH_HACK);
  }

  // convenience function: given a property, get its magnitude or produce zero
  function amountPT(prop:string){ return computedStyle[prop].magnitude || 0; }

  // used for the while loop, to continually resize the shape and multiline
  // text, until it fits within the bounds of the element
  const isOutsideBounds = (): boolean => {
    ctx.font = `${fontWeight} ${fontSize}pt ${fontFamily}`;
    
    // # of hard AND wrapped lines - computed using the estimated chars per line at this size
    // # of newlines - relevant for spaceAbove/spaceBelow/spaceBetween calculation
    const lines = splitter(text.rawText, estCharsPerLine());
    const newlines = text.rawText.match(/\\n/g)?.length || 1;

    // if the only constraint is horizontal, we're outside bounds if there's more than one line
    if ((constraints == "horizontal") && (lines.length > 1)) {
      return true;
    }

    // compute the horizontal and vertical whitespace in PT
    // Horizontal: <indentStart> + <indentEnd>
    // Vertical: <space below and above all n linebreaks> + <spacing between each line (n-1)>
    const horizontalWhiteSpace = amountPT("indentStart") + amountPT("indentEnd");
    const spaceAroundLines = newlines * (amountPT("spaceBelow") + amountPT("spaceAbove"))
    const spaceBetweenLines = (newlines-1) * (computedStyle.lineSpacing/100 * fontSize)
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

  //console.log(`fitting "${text.rawText}" into ${sizePT.width}pt x ${sizePT.height}pt. Constraints are ${constraints}`);

  // continually loop over until the size of the text element is within bounds,
  // decreasing by 0.25pt increments until it fits within the width
  while (isOutsideBounds() && fontSize > MIN_SIZE) { fontSize = fontSize - 0.25; }
  
  cachedFontCalculations.set(key, fontSize);
  return fontSize;
}
