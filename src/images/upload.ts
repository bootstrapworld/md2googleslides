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

const debug = Debug('md2gslides');

/**
 * Uploads a local file to temporary storage so it is HTTP/S accessible.
 *
 * Currently uses https://file.io for free emphemeral file hosting.
 *
 * @param {string} filePath -- Local path to image to upload
 * @returns {Promise<string>} URL to hosted image
 */
async function uploadLocalImage(filePath: string, drive: any): Promise<string> {
  debug('Registering file %s', filePath);
  const stream = fs.createReadStream(filePath);
  const filename = filePath.split('/').pop();

  try {
    const fileMetadata = {
      name: filename,
      parents: ['1kfhzbGk2HPD2xkIphI1x1-ObwCzQWSCx']
    };
    const media = { body: stream };

    const response = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id' // Get the file ID after upload
    });

    // If the upload is successful, get the fileId
    const fileId = response.data.id;

    // Set the file at that ID to be world-readable
    await drive.permissions.create({
      fileId: fileId,
      resource: {
        'type': 'anyone',
        'role': 'reader'
      }
    });

    // return the URL to the newly-uploaded, world-readable file
    return `https://drive.usercontent.google.com/uc?id=${fileId}&authuser=0&export=download`

  } catch (e) {
    console.error('Error uploading file:', e);
    throw e;
  } finally {
    stream.destroy();
  }
}

export default uploadLocalImage;