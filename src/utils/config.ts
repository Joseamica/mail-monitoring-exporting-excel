import dotenv from 'dotenv';
import { Config } from '../types';

dotenv.config();

export const config: Config = {
  gmail: {
    credentialsPath: process.env.GMAIL_CREDENTIALS_PATH || './credentials.json',
    tokenPath: process.env.GMAIL_TOKEN_PATH || './token.json',
  },
  excel: {
    filePath: process.env.EXCEL_FILE_PATH || './solicitudes.xlsx',
    sheetName: process.env.EXCEL_SHEET_NAME || 'Solicitudes',
  },
  monitoring: {
    pollingIntervalMs: parseInt(process.env.POLLING_INTERVAL_MS || '300000'),
    logLevel: process.env.LOG_LEVEL || 'info',
  },
  errorHandling: {
    retryAttempts: parseInt(process.env.RETRY_ATTEMPTS || '3'),
    retryDelayMs: parseInt(process.env.RETRY_DELAY_MS || '5000'),
  },
};

export default config;
