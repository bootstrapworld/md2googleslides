import Debug from 'debug';
import { google, Auth } from 'googleapis';
import path from 'path';
import { mkdirp } from 'mkdirp';
import { Low } from 'lowdb';
import { JSONFile } from 'lowdb/node';
import { Memory } from 'lowdb';

const debug = Debug('md2gslides');

export type UserPrompt = (message: string) => Promise<string>;

export interface AuthOptions {
  clientId: string;
  clientSecret: string;
  prompt: UserPrompt;
  filePath?: string;
}

// Define the shape of your JSON file
interface Schema {
  [key: string]: Auth.Credentials;
}

export default class UserAuthorizer {
  private redirectUrl = 'urn:ietf:wg:oauth:2.0:oob';
  private db: Low<Schema>;
  private clientId: string;
  private clientSecret: string;
  private prompt: UserPrompt;

  public constructor(options: AuthOptions) {
    this.clientId = options.clientId;
    this.clientSecret = options.clientSecret;
    this.prompt = options.prompt;

    // Setup adapter
    let adapter;
    if (options.filePath) {
      const parentDir = path.dirname(options.filePath);
      mkdirp.sync(parentDir);
      adapter = new JSONFile<Schema>(options.filePath);
    } else {
      adapter = new Memory<Schema>();
    }

    // Initialize instance (default data is an empty object)
    this.db = new Low<Schema>(adapter, {});
  }

  public async getUserCredentials(
    user: string,
    scopes: string
  ): Promise<Auth.OAuth2Client> {
    // 1. Read from disk/memory
    await this.db.read();
    this.db.data ||= {};

    const oauth2Client = new google.auth.OAuth2(
      this.clientId,
      this.clientSecret,
      this.redirectUrl
    );

    oauth2Client.on('tokens', async (tokens: Auth.Credentials) => {
      if (tokens.refresh_token) {
        debug('Saving refresh token');
        this.db.data[user] = tokens;
        await this.db.write(); // Must be awaited now
      }
    });

    const tokens = this.db.data[user];
    if (tokens) {
      debug('User previously authorized, refreshing');
      oauth2Client.setCredentials(tokens);
      await oauth2Client.getAccessToken();
      return oauth2Client;
    }

    debug('Challenging for authorization');
    const authUrl = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: scopes,
      login_hint: user,
    });
    const code = await this.prompt(authUrl);
    const tokenResponse = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokenResponse.tokens);

    return oauth2Client;
  }
}