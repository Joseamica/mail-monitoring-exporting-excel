import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import fs from 'fs';
import path from 'path';
import readline from 'readline';
import { config } from '../utils/config';
import { logger } from '../utils/logger';

const SCOPES = ['https://www.googleapis.com/auth/gmail.modify'];

export class GmailAuth {
  private oAuth2Client: OAuth2Client;

  constructor() {
    this.oAuth2Client = new google.auth.OAuth2();
  }

  async initialize(): Promise<void> {
    try {
      // Load client secrets from a local file
      const credentialsPath = path.resolve(config.gmail.credentialsPath);
      const credentials = JSON.parse(fs.readFileSync(credentialsPath, 'utf8'));
      
      const { client_secret, client_id, redirect_uris } = credentials.installed;
      this.oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

      // Check if we have previously stored a token
      const tokenPath = path.resolve(config.gmail.tokenPath);
      if (fs.existsSync(tokenPath)) {
        const token = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
        this.oAuth2Client.setCredentials(token);
        logger.info('Gmail authentication token loaded successfully');
      } else {
        await this.getNewToken();
      }
    } catch (error) {
      logger.error('Error initializing Gmail authentication:', error);
      throw error;
    }
  }

  private async getNewToken(): Promise<void> {
    const authUrl = this.oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
    });

    logger.info(`Authorize this app by visiting this url: ${authUrl}`);
    
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });

    return new Promise((resolve, reject) => {
      rl.question('Enter the code from that page here: ', async (code) => {
        rl.close();
        try {
          // Exchange the authorization code for an access token and a refresh token.
          const { tokens } = await this.oAuth2Client.getToken(code);

          if (!tokens) {
            throw new Error('Failed to receive tokens.');
          }

          // A refresh_token is only issued on the first authorization.
          // If it's missing, the user may need to revoke and re-authorize access.
          if (!tokens.refresh_token) {
            logger.warn('No refresh token was received. If the app fails to re-authenticate later, you may need to revoke its access in your Google account and authorize it again.');
          }

          this.oAuth2Client.setCredentials(tokens);

          // Store the token to disk for later program executions
          const tokenPath = path.resolve(config.gmail.tokenPath);
          fs.writeFileSync(tokenPath, JSON.stringify(tokens));
          logger.info('Token stored successfully to', tokenPath);
          resolve();
        } catch (error) {
          logger.error('Error while trying to retrieve access token', error);
          reject(error);
        }
      });
    });
  }

  async refreshTokenIfNeeded(): Promise<void> {
    try {
      const { credentials } = await this.oAuth2Client.refreshAccessToken();
      this.oAuth2Client.setCredentials(credentials);
      
      // Update the stored token
      const tokenPath = path.resolve(config.gmail.tokenPath);
      fs.writeFileSync(tokenPath, JSON.stringify(credentials));
      logger.info('Access token refreshed successfully');
    } catch (error) {
      logger.error('Error refreshing access token:', error);
      throw error;
    }
  }

  getAuthorizedClient(): OAuth2Client {
    return this.oAuth2Client;
  }

  async getAuthorizedGmail() {
    // Check if token needs refresh
    if (this.oAuth2Client.credentials.expiry_date && 
        this.oAuth2Client.credentials.expiry_date <= Date.now()) {
      await this.refreshTokenIfNeeded();
    }

    return google.gmail({ version: 'v1', auth: this.oAuth2Client });
  }
}

export default GmailAuth;
