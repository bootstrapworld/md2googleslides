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

import markdownIt from 'markdown-it';
import Token from 'markdown-it/lib/token';
// @ts-ignore
import attrs from 'markdown-it-attrs';
// @ts-ignore
import lazyHeaders from 'markdown-it-lazy-headers';
// @ts-ignore
import { full as emoji } from 'markdown-it-emoji';
// @ts-ignore
import expandTabs from 'markdown-it-expand-tabs';
// @ts-ignore
import video from 'markdown-it-video';
// @ts-ignore
import mathjax3 from 'markdown-it-mathjax3';

const mdOptions = {
  html: true,
  langPrefix: 'highlight ',
  linkify: false,
  breaks: false,
};

const parser = markdownIt(mdOptions)
  .use(attrs)
  .use(lazyHeaders)
  .use(emoji, {shortcuts: {}})
  .use(expandTabs, {tabWidth: 4})
  .use(mathjax3)
  .use(video, {youtube: {width: 640, height: 390}});

function parseMarkdown(markdown: string): Token[] {
  const parseTree = parser.parse(markdown, {});
  return parseTree;
}

export default parseMarkdown;
