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

import * as chai from 'chai';
// load the markdown parser first: it bundles MathJax 3, which shares the
// global `MathJax` object with MathJax 4
import '../src/parser/parser';
import {get_mathjax_svg} from '../src/images/mathjax';

const expect = chai.expect;

describe('get_mathjax_svg', () => {
  it('should render TeX to an svg', async () => {
    const svg = await get_mathjax_svg('\\frac{a}{b}+\\sqrt{x^2}');
    expect(svg).to.match(/^<svg/);
  });

  it('should render align environments and \\cancel', async () => {
    const svg = await get_mathjax_svg(
      '\\begin{align} a &= b \\\\ c &= \\cancel{d}\\end{align}',
    );
    expect(svg).to.match(/^<svg/);
    expect(svg).not.to.include('data-mjx-error');
  });

  it('should render chemistry with mhchem', async () => {
    const svg = await get_mathjax_svg('\\ce{H2O + CO2 -> H2CO3}');
    expect(svg).to.match(/^<svg/);
    expect(svg).not.to.include('data-mjx-error');
  });
});
