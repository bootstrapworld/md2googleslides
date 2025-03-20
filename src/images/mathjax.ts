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

import { mathjax } from 'mathjax-full/js/mathjax.js'
import { TeX } from 'mathjax-full/js/input/tex.js'
import { SVG } from 'mathjax-full/js/output/svg.js'
import { AllPackages } from 'mathjax-full/js/input/tex/AllPackages.js'
import { liteAdaptor } from 'mathjax-full/js/adaptors/liteAdaptor.js'
import { RegisterHTMLHandler } from 'mathjax-full/js/handlers/html.js'
import renderSVG from './svg.js';
import {ImageDefinition} from '../slides.js';
import assert from 'assert';

const debug = Debug('md2gslides');


const adaptor = liteAdaptor()
RegisterHTMLHandler(adaptor)

const mathjax_document = mathjax.document('', {
  InputJax: new TeX({ packages: AllPackages }),
  OutputJax: new SVG({ fontCache: 'local' })
})


export function get_mathjax_svg(math: string): string {
  const node = mathjax_document.convert(math)
  return adaptor.innerHTML(node)
}

function formatFor(expression: string): string {
  return expression.trim().startsWith('<math>') ? 'MathML' : 'TeX';
}

function addOrMergeStyles(svg: string, style?: string): string {
  if (!style) {
    return svg;
  }
  const match = svg.match(/(<svg[^>]+)(style="([^"]+)")([^>]+>)/);
  if (match) {
    return (
      svg.slice(0, match[1].length) +
      `style="${style};${match[3]}"` +
      svg.slice(match[1].length + match[2].length)
    );
  } else {
    const i = svg.indexOf('>');
    return svg.slice(0, i) + ` style="${style}"` + svg.slice(i);
  }
}

async function renderMathJax(image: ImageDefinition): Promise<string> {
  debug('Generating math image: %O', image);
  assert(image.source);
  const svg = get_mathjax_svg(image.source);
  image.source = addOrMergeStyles(svg, image.style);
  image.type = 'svg';
  return await renderSVG(image);
}

export default renderMathJax;
