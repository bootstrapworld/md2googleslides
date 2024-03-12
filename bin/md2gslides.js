#!/usr/bin/env node

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

/* eslint-disable no-console, @typescript-eslint/no-var-requires */

import fs   from 'fs';
import path from 'path';
import { ArgumentParser } from 'argparse';
import UserAuthorizer from '../lib/auth.js';
import SlideGenerator from '../lib/slide_generator.js';
import opener  from 'opener';
import readline from 'readline';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const SCOPES = [
  'https://www.googleapis.com/auth/presentations',
  'https://www.googleapis.com/auth/drive',
];

const USER_HOME =
  process.env.HOME || process.env.HOMEPATH || process.env.USERPROFILE;
const STORED_CREDENTIALS_PATH = path.join(
  USER_HOME,
  '.md2googleslides',
  'credentials.json'
);
const STORED_CLIENT_ID_PATH = path.join(
  USER_HOME,
  '.md2googleslides',
  'client_id.json'
);

const parser = new ArgumentParser({
  add_help: true,
  description: 'Markdown to Slides converter',
});

parser.add_argument('-v', '--version', {
  action: 'version',
  version: '1.0.1',
});
parser.add_argument('file', {
  help: 'Path to markdown file to convert, If omitted, reads from stdin',
  nargs: '?',
});
parser.add_argument('-u', '--user', {
  help: 'Email address of user',
  required: false,
  dest: 'user',
  default: 'default',
});
parser.add_argument('-a', '--append', {
  dest: 'id',
  help: 'Appends slides to an existing presentation',
  required: false,
});
parser.add_argument('-e', '--erase', {
  dest: 'erase',
  action: 'store_true',
  help: 'Erase existing slides prior to appending.',
  required: false,
});
parser.add_argument('-n', '--no-browser', {
  action: 'store_true',
  dest: 'headless',
  help: 'Headless mode - do not launch browsers, just shows URLs',
  required: false,
});
parser.add_argument('-s', '--style', {
  help: 'Name of highlight.js theme for code formatting',
  dest: 'style',
  required: false,
  default: 'default',
});
parser.add_argument('-t', '--title', {
  help: 'Title of the presentation',
  dest: 'title',
  required: false,
});
parser.add_argument('-p', '--parent', {
  help: 'Id of the GDrive folder in which so save the presentation',
  dest: 'parentId',
  required: false,
});
parser.add_argument('-c', '--copy', {
  help: 'Id of the presentation to copy and use as a base',
  dest: 'copy',
  required: false,
});
parser.add_argument('--use-fileio', {
  help: 'Acknolwedge local and generated images are uploaded to https://file.io',
  action: 'store_true',
  dest: 'useFileio',
  required: false,
});

const args = parser.parse_args();

function handleError(err) {
  if(err && err.errors && err.errors[0] && err.errors[0].message) {
    let errorString = '❌ Unable to generate slides:\n';
    errorString += err.errors[0].message + '\n';
    const reqIdx = err.errors[0].message.match(/requests\[([0-9]+)\]/)[1];
    // if there's a specific request that went bad, print it
    if(reqIdx && err?.config?.data?.requests) {
      errorString += 'The request that failed was:\n'
      errorString += JSON.stringify(err.config.data.requests[reqIdx], null, 2);
    }
    console.error(errorString);  
  } else {
    throw err;
  }
  
}

function prompt(url) {
  if (args.headless) {
    console.log('Authorize this app by visiting this url: ');
    console.log(url);
  } else {
    console.log('Authorize this app in your browser.');
    opener(url);
  }
  return new Promise((resolve, reject) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Enter the code here: ', code => {
      rl.close();
      code = code.trim();
      if (code.length > 0) {
        resolve(code);
      } else {
        reject(new Error('No code provided'));
      }
    });
  });
}

function authorizeUser() {
  // Google OAuth2 clients always have a secret, even if the client is an installed
  // application/utility such as this.  Of course, in such cases the "secret" is
  // actually publicly known; security depends entirely on the secrecy of refresh
  // tokens, which effectively become bearer tokens.

  // Load and parse client ID and secret from client_id.json file. (Create
  // OAuth client ID from Credentials tab at console.developers.google.com
  // and download the credentials as client_id.json to ~/.md2googleslides
  let data; // needs to be scoped outside of try-catch
  try {
    data = fs.readFileSync(STORED_CLIENT_ID_PATH);
  } catch (err) {
    console.log('Error loading client secret file:', err);
    throw err;
  }
  if (data === undefined) {
    console.log('Error loading client secret data');
    throw 'No client secret found.';
  }
  const creds = JSON.parse(data).installed;

  // Authorize user and get (& store) a valid access token.
  const options = {
    clientId: creds.client_id,
    clientSecret: creds.client_secret,
    filePath: STORED_CREDENTIALS_PATH,
    prompt: prompt,
  };
  const auth = new UserAuthorizer(options);
  return auth.getUserCredentials(args.user, SCOPES);
}

function buildSlideGenerator(oauth2Client) {
  const title = args.title || args.file;
  const presentationId = args.id;
  const copyId = args.copy;
  const parentId = args.parentId;
  if (presentationId) {
    return SlideGenerator.forPresentation(oauth2Client, presentationId);
  } else if (copyId) {
    return SlideGenerator.copyPresentation(oauth2Client, title, copyId, parentId);
  } else {
    return SlideGenerator.newPresentation(oauth2Client, title, parentId);
  }    
}

function eraseIfNeeded(slideGenerator) {
  if (args.erase || !args.id) {
    return slideGenerator.erase().then(() => {
      return slideGenerator;
    });
  } else {
    return Promise.resolve(slideGenerator);
  }
}

function loadCss(theme) {
  let cssPath, css;
  try {
    cssPath = path.join(
      require.resolve('highlight.js'),
      '..',
      '..',
      'styles',
      theme + '.css'
    );
    css = fs.readFileSync(cssPath, {encoding: 'UTF-8'});  
  } catch(e) {
    try {
      cssPath = path.join(theme);
      css = fs.readFileSync(cssPath, {encoding: 'UTF-8'});
    } catch(e) {
      console.log(e)
      throw "Could not load the theme '" + theme + "'. It is not known to highlight.js, or a valid file path" 
    }
  }
  
  return css;
}

function generateSlides(slideGenerator) {
  // try to limit warning spew
  require('events').EventEmitter.defaultMaxListeners = 15;

  let source;
  if (args.file) {
    source = path.resolve(args.file);
    // Set working directory relative to markdown file
    process.chdir(path.dirname(source));
  } else {
    source = 0;
  }
  const input = fs.readFileSync(source, {encoding: 'UTF-8'});
  const css = loadCss(args.style);

  return slideGenerator.generateFromMarkdown(input, {
    css: css,
    useFileio: args.useFileio,
  });
}

function displayResults(id) {
  const url = 'https://docs.google.com/presentation/d/' + id;
  if (args.headless) {
    console.log('View your presentation at: %s', url);
  } else {
    console.log('Opening your presentation (%s)', url);
    opener(url);
  }
}
authorizeUser()
  .then(buildSlideGenerator)
  .then(eraseIfNeeded)
  .then(generateSlides)
  .then(displayResults)
  .catch(handleError);
