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

import renderSVG from './svg.js';
import {ImageDefinition} from '../slides.js';

type MathJaxApi = typeof import('@mathjax/src').default;
import assert from 'assert';

const debug = Debug('md2gslides');

// TeX extensions that MathJax 3's AllPackages provided and that MathJax 4
// does not already include in input/tex
const TEX_PACKAGES = [
  'action',
  'amscd',
  'bbox',
  'boldsymbol',
  'braket',
  'bussproofs',
  'cancel',
  'cases',
  'centernot',
  'color',
  'colortbl',
  'empheq',
  'enclose',
  'extpfeil',
  'gensymb',
  'html',
  'mathtools',
  'mhchem',
  'physics',
  'setoptions',
  'tagformat',
  'unicode',
  'upgreek',
  'verb',
];

// MathJax 4 loads its components asynchronously, so start it once and share
// the promise. markdown-it-mathjax3 bundles MathJax 3, and both use the global
// `MathJax` object: hide it while MathJax 4 loads, then put it back.
let mathjaxReady: Promise<MathJaxApi> | undefined;

function startMathJax(): Promise<MathJaxApi> {
  mathjaxReady ??= (async () => {
    const globals = globalThis as {MathJax?: unknown};
    const saved = globals.MathJax;
    delete globals.MathJax;
    try {
      const {default: MathJax} = await import('@mathjax/src');
      await MathJax.init({
        loader: {
          load: [
            'input/tex',
            'output/svg',
            'adaptors/liteDOM',
            ...TEX_PACKAGES.map(name => `[tex]/${name}`),
          ],
        },
        tex: {packages: {'[+]': TEX_PACKAGES}},
        svg: {fontCache: 'local'},
      });
      return MathJax;
    } finally {
      if (saved !== undefined) {
        globals.MathJax = saved;
      }
    }
  })();
  return mathjaxReady;
}

export async function get_mathjax_svg(math: string): Promise<string> {
  const MathJax = await startMathJax();
  const node = await MathJax.tex2svgPromise(math, {display: true});
  return MathJax.startup.adaptor.innerHTML(node);
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
  const svg = await get_mathjax_svg(image.source);
  image.source = addOrMergeStyles(svg, image.style);
  image.type = 'svg';
  return await renderSVG(image);
}

export default renderMathJax;
