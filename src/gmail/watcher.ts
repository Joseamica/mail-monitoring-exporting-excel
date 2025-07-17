import { gmail_v1 } from 'googleapis';
import { GmailAuth } from './auth';
import { logger } from '../utils/logger';
import { config } from '../utils/config';
import { ProcessedEmail, GmailAttachment } from '../types';

export class GmailWatcher {
  private gmailAuth: GmailAuth;
  private gmail: gmail_v1.Gmail | null = null;
  private processedMessages: Set<string> = new Set();

  constructor() {
    this.gmailAuth = new GmailAuth();
  }

  async initialize(): Promise<void> {
    try {
      await this.gmailAuth.initialize();
      this.gmail = await this.gmailAuth.getAuthorizedGmail();
      await this.ensureLabelExists('PROCESSED');
      logger.info('Gmail watcher initialized successfully');
    } catch (error) {
      logger.error('Error initializing Gmail watcher:', error);
      throw error;
    }
  }

  private async ensureLabelExists(labelName: string): Promise<string> {
    if (!this.gmail) {
      throw new Error('Gmail not initialized');
    }

    try {
      const res = await this.gmail.users.labels.list({ userId: 'me' });
      const labels = res.data.labels || [];
      const existingLabel = labels.find((label) => label.name === labelName);

      if (existingLabel && existingLabel.id) {
        logger.info(`Label '${labelName}' already exists with ID: ${existingLabel.id}`);
        return existingLabel.id;
      }

      logger.info(`Label '${labelName}' not found, creating it...`);
      const newLabel = await this.gmail.users.labels.create({
        userId: 'me',
        requestBody: {
          name: labelName,
          labelListVisibility: 'labelShow',
          messageListVisibility: 'show',
        },
      });

      if (!newLabel.data.id) {
        throw new Error('Failed to create label.');
      }

      logger.info(`Label '${labelName}' created successfully with ID: ${newLabel.data.id}`);
      return newLabel.data.id;
    } catch (error) {
      logger.error(`Error ensuring label '${labelName}' exists:`, error);
      throw error;
    }
  }

  async checkForNewMessages(): Promise<ProcessedEmail[]> {
    if (!this.gmail) {
      throw new Error('Gmail not initialized');
    }

    try {
      // Search for emails with PDF attachments, excluding already processed ones (regardless of read/unread status)
      const query = 'has:attachment filename:pdf -label:PROCESSED';
      logger.info(`Gmail query: ${query}`);
      
      const response = await this.gmail.users.messages.list({
        userId: 'me',
        q: query,
        maxResults: 10
      });
      
      logger.info(`Gmail API response: Found ${response.data.messages?.length || 0} messages`);
      
      if (!response.data.messages || response.data.messages.length === 0) {
        // Let's also try a broader query to see if there are any unread emails at all
        const broadQuery = 'is:unread';
        const broadResponse = await this.gmail.users.messages.list({
          userId: 'me',
          q: broadQuery,
          maxResults: 5
        });
        logger.info(`Broad query 'is:unread' found ${broadResponse.data.messages?.length || 0} messages`);
        
        // Try query without PROCESSED label exclusion
        const attachmentQuery = 'is:unread has:attachment';
        const attachmentResponse = await this.gmail.users.messages.list({
          userId: 'me',
          q: attachmentQuery,
          maxResults: 5
        });
        logger.info(`Query 'is:unread has:attachment' found ${attachmentResponse.data.messages?.length || 0} messages`);
        
        logger.info('Found 0 unread messages with CARTA attachments');
        return [];
      }

      const messages = response.data.messages || [];
      logger.info(`Found ${messages.length} unread messages with CARTA attachments`);

      const processedEmails: ProcessedEmail[] = [];

      for (const message of messages) {
        if (message.id && !this.processedMessages.has(message.id)) {
          try {
            const processedEmail = await this.getMessageDetails(message.id);
            if (processedEmail) {
              processedEmails.push(processedEmail);
              this.processedMessages.add(message.id);
            }
          } catch (error) {
            logger.error(`Error processing message ${message.id}:`, error);
          }
        }
      }

      return processedEmails;
    } catch (error) {
      logger.error('Error checking for new messages:', error);
      throw error;
    }
  }

  private async getMessageDetails(messageId: string): Promise<ProcessedEmail | null> {
    if (!this.gmail) {
      throw new Error('Gmail not initialized');
    }

    try {
      const response = await this.gmail.users.messages.get({
        userId: 'me',
        id: messageId
      });

      const message = response.data;
      const headers = message.payload?.headers || [];
      
      const subject = headers.find(h => h.name === 'Subject')?.value || '';
      const sender = headers.find(h => h.name === 'From')?.value || '';
      const dateHeader = headers.find(h => h.name === 'Date')?.value || '';
      
      const receivedAt = dateHeader ? new Date(dateHeader) : new Date();
      
      // Extract attachments
      const attachments = await this.extractAttachments(message);
      logger.info(`Message ${messageId} has ${attachments.length} attachments:`, attachments.map(att => `${att.filename} (${att.mimeType})`));
      
      // Filter for all PDF attachments (including CARTA and non-CARTA files)
      const pdfAttachments = attachments.filter(att => {
        const isCartaFile = this.isCartaAttachment(att.filename);
        const isPDF = att.mimeType === 'application/pdf';
        logger.info(`Checking attachment: ${att.filename} - isCarta: ${isCartaFile}, isPDF: ${isPDF}`);
        return isPDF; // Process all PDFs, not just CARTA ones
      });

      if (pdfAttachments.length === 0) {
        logger.warn(`No PDF attachments found in message ${messageId}`);
        return null;
      }
      
      logger.info(`Found ${pdfAttachments.length} PDF attachments in message ${messageId} (including ${attachments.filter(att => this.isCartaAttachment(att.filename) && att.mimeType === 'application/pdf').length} CARTA files)`);

      return {
        messageId,
        subject,
        sender,
        receivedAt,
        attachments: pdfAttachments
      };
    } catch (error) {
      logger.error(`Error getting message details for ${messageId}:`, error);
      return null;
    }
  }

  private isCartaAttachment(filename: string): boolean {
    const normalizedFilename = filename.toLowerCase();
    // Check for variations of CARTA including typos and case variations
    const cartaPatterns = [
      'carta',
      'cartas', 
      'carat',
      'carata',
      'cart',
      'crta',
      'acarta'
    ];
    
    return cartaPatterns.some(pattern => normalizedFilename.includes(pattern));
  }

  private async extractAttachments(message: gmail_v1.Schema$Message): Promise<GmailAttachment[]> {
    const attachments: GmailAttachment[] = [];

    const extractFromParts = (parts: gmail_v1.Schema$MessagePart[]) => {
      for (const part of parts) {
        if (part.filename && part.filename.length > 0 && part.body?.attachmentId) {
          attachments.push({
            filename: part.filename,
            mimeType: part.mimeType || '',
            attachmentId: part.body.attachmentId,
            size: part.body.size || 0
          });
        }
        
        if (part.parts) {
          extractFromParts(part.parts);
        }
      }
    };

    if (message.payload?.parts) {
      extractFromParts(message.payload.parts);
    }

    return attachments;
  }

  async downloadAttachment(messageId: string, attachmentId: string): Promise<Buffer> {
    if (!this.gmail) {
      throw new Error('Gmail not initialized');
    }

    try {
      const response = await this.gmail.users.messages.attachments.get({
        userId: 'me',
        messageId,
        id: attachmentId
      });

      const data = response.data.data;
      if (!data) {
        throw new Error('No attachment data received');
      }

      // Convert from base64url to buffer
      const buffer = Buffer.from(data.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
      logger.info(`Downloaded attachment ${attachmentId} (${buffer.length} bytes)`);
      
      return buffer;
    } catch (error) {
      logger.error(`Error downloading attachment ${attachmentId}:`, error);
      throw error;
    }
  }

  async markAsProcessed(messageId: string): Promise<void> {
    if (!this.gmail) {
      throw new Error('Gmail not initialized');
    }

    try {
      // Get the label ID first
      const labelId = await this.ensureLabelExists('PROCESSED');
      
      // Add a custom label to mark as processed using the label ID
      await this.gmail.users.messages.modify({
        userId: 'me',
        id: messageId,
        requestBody: {
          addLabelIds: [labelId], // Use the actual label ID, not the label name
          removeLabelIds: ['UNREAD']
        }
      });

      logger.info(`Marked message ${messageId} as processed`);
    } catch (error) {
      logger.warn(`Error marking message ${messageId} as processed:`, error);
      // Don't throw here, as this is not critical for the main flow
    }
  }

  async retryOperation<T>(operation: () => Promise<T>, maxRetries: number = config.errorHandling.retryAttempts): Promise<T> {
    for (let i = 0; i < maxRetries; i++) {
      try {
        return await operation();
      } catch (error) {
        logger.warn(`Operation failed, retry ${i + 1}/${maxRetries}:`, error);
        
        if (i === maxRetries - 1) {
          throw error;
        }
        
        // Wait before retrying
        await new Promise(resolve => setTimeout(resolve, config.errorHandling.retryDelayMs));
      }
    }
    
    throw new Error('Max retries exceeded');
  }
}

export default GmailWatcher;
