import { GmailWatcher } from './gmail/watcher';
import { PDFParser } from './pdf/parser';
import { ExcelWriter } from './excel/writer';
import { logger } from './utils/logger';
import { config } from './utils/config';
import { SolicitudData } from './types';

class MonitoreoGmail {
  private gmailWatcher: GmailWatcher;
  private pdfParser: PDFParser;
  private excelWriter: ExcelWriter;
  private isRunning: boolean = false;
  private intervalId: NodeJS.Timeout | null = null;

  constructor() {
    this.gmailWatcher = new GmailWatcher();
    this.pdfParser = new PDFParser();
    this.excelWriter = new ExcelWriter();
  }

  async initialize(): Promise<void> {
    try {
      logger.info('Initializing Monitoreo Gmail system...');
      
      // Initialize all modules
      await this.gmailWatcher.initialize();
      await this.excelWriter.initialize();
      
      logger.info('Monitoreo Gmail system initialized successfully');
    } catch (error) {
      logger.error('Error initializing system:', error);
      throw error;
    }
  }

  async start(): Promise<void> {
    if (this.isRunning) {
      logger.warn('System is already running');
      return;
    }

    try {
      this.isRunning = true;
      logger.info(`Starting monitoring with ${config.monitoring.pollingIntervalMs}ms interval`);
      
      // Run initial check
      await this.checkAndProcessEmails();
      
      // Set up interval polling
      this.intervalId = setInterval(async () => {
        try {
          await this.checkAndProcessEmails();
        } catch (error) {
          logger.error('Error in polling interval:', error);
        }
      }, config.monitoring.pollingIntervalMs);
      
      logger.info('Monitoring started successfully');
    } catch (error) {
      logger.error('Error starting monitoring:', error);
      this.isRunning = false;
      throw error;
    }
  }

  async stop(): Promise<void> {
    if (!this.isRunning) {
      logger.warn('System is not running');
      return;
    }

    logger.info('Stopping monitoring...');
    
    if (this.intervalId) {
      clearInterval(this.intervalId);
      this.intervalId = null;
    }
    
    this.isRunning = false;
    logger.info('Monitoring stopped');
  }

  private async checkAndProcessEmails(): Promise<void> {
    try {
      logger.info('Checking for new emails...');
      
      const emails = await this.gmailWatcher.checkForNewMessages();
      
      if (emails.length === 0) {
        logger.info('No new emails to process');
        return;
      }
      
      logger.info(`Found ${emails.length} emails to process`);
      
      for (const email of emails) {
        try {
          await this.processEmail(email);
        } catch (error) {
          logger.error(`Error processing email ${email.messageId}:`, error);
        }
      }
    } catch (error) {
      logger.error('Error checking for emails:', error);
    }
  }

  private async processEmail(email: any): Promise<void> {
    try {
      logger.info(`Processing email: ${email.subject}`);
      
      // Sort attachments to prioritize non-CARTA PDFs for better company name extraction
      const sortedAttachments = [...email.attachments].sort((a, b) => {
        const aIsCarta = a.filename.toLowerCase().includes('carta');
        const bIsCarta = b.filename.toLowerCase().includes('carta');
        // Non-CARTA PDFs first (false < true), then CARTA PDFs
        return Number(aIsCarta) - Number(bIsCarta);
      });
      
      logger.info(`Processing ${sortedAttachments.length} PDF attachments in prioritized order:`, 
        sortedAttachments.map(att => `${att.filename} (CARTA: ${att.filename.toLowerCase().includes('carta')})`));
      
      // Process each PDF attachment (non-CARTA first for better company name extraction)
      for (const attachment of sortedAttachments) {
        try {
          // Download the PDF attachment
          const pdfBuffer = await this.gmailWatcher.downloadAttachment(
            email.messageId,
            attachment.attachmentId
          );
          
          // Parse the PDF to extract company name only
          const parsedData = await this.pdfParser.parsePDF(pdfBuffer);
          
          // Extract amount and city from email subject
          const cantidad = PDFParser.extractAmountFromSubject(email.subject);
          const ciudad = PDFParser.extractCityFromSubject(email.subject);
          
          // Use company name from PDF as client name, fallback to email extraction
          const cliente = parsedData.domicilioEntrega !== 'No especificada' 
            ? parsedData.domicilioEntrega 
            : this.extractClientName(email.sender, email.subject);
          
          // Prepare data for Excel with new column structure
          const solicitudData: SolicitudData = {
            fecha: this.formatShortDate(email.receivedAt),        // FECHA - formato corto (24-jun)
            operador: cliente,                                    // OPERADOR - nombre del cliente/empresa
            sucursal: ciudad || 'No especificada',              // SUCURSAL - ciudad
            monto: cantidad || 'No especificada',               // MONTO - cantidad
            link: this.generateGmailLink(email.messageId),      // LINK - URL para acceder al correo
            messageId: email.messageId                           // Para identificación interna
          };
          
          // Append to Excel file
          await this.excelWriter.appendSolicitud(solicitudData);
          
          // Mark email as processed
          await this.gmailWatcher.markAsProcessed(email.messageId);
          
          logger.info(`Successfully processed email from ${cliente} for ${ciudad}`);
          
        } catch (error) {
          logger.error(`Error processing attachment ${attachment.filename}:`, error);
        }
      }
    } catch (error) {
      logger.error(`Error processing email ${email.messageId}:`, error);
    }
  }

  private extractClientName(sender: string, subject: string): string {
    // Extract client name from sender email
    const emailMatch = sender.match(/^(.+?)\s*<.*>$/);
    if (emailMatch) {
      return emailMatch[1].trim();
    }
    
    // If no name in sender, try to extract from email address
    const emailAddressMatch = sender.match(/<(.+?)@/);
    if (emailAddressMatch) {
      return emailAddressMatch[1].trim();
    }
    
    // Try to extract from subject as fallback
    const subjectMatch = subject.match(/^(?:solicitud|pedido|orden)\s+(?:de\s+)?(.+?)(?:\s+-|$)/i);
    if (subjectMatch) {
      return subjectMatch[1].trim();
    }
    
    return sender; // Return full sender as fallback
  }

  async getStatus(): Promise<{
    isRunning: boolean;
    lastCheck: Date;
    excelStats: any;
  }> {
    const excelStats = await this.excelWriter.getStatistics();
    
    return {
      isRunning: this.isRunning,
      lastCheck: new Date(),
      excelStats
    };
  }

  // Helper method to format date in short format (24-jun)
  private formatShortDate(date: Date): string {
    const months = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 
                   'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
    const day = date.getDate();
    const month = months[date.getMonth()];
    return `${day}-${month}`;
  }

  // Helper method to generate Gmail link
  private generateGmailLink(messageId: string): string {
    // Generate Gmail web link using the correct format for specific message
    return `https://mail.google.com/mail/u/1/#inbox/${messageId}`;
  }



  // Graceful shutdown handler
  async shutdown(): Promise<void> {
    logger.info('Shutting down Monitoreo Gmail system...');
    
    await this.stop();
    
    // Create backup before shutdown
    try {
      await this.excelWriter.backup();
    } catch (error) {
      logger.error('Error creating backup during shutdown:', error);
    }
    
    logger.info('Shutdown completed');
  }
}

// Main execution
async function main() {
  const monitoreo = new MonitoreoGmail();
  
  // Handle graceful shutdown
  process.on('SIGINT', async () => {
    logger.info('Received SIGINT, shutting down gracefully...');
    await monitoreo.shutdown();
    process.exit(0);
  });
  
  process.on('SIGTERM', async () => {
    logger.info('Received SIGTERM, shutting down gracefully...');
    await monitoreo.shutdown();
    process.exit(0);
  });
  
  try {
    await monitoreo.initialize();
    await monitoreo.start();
    
    logger.info('Monitoreo Gmail is running. Press Ctrl+C to stop.');
    
    // Keep the process running
    process.stdin.resume();
    
  } catch (error) {
    logger.error('Failed to start Monitoreo Gmail:', error);
    process.exit(1);
  }
}

// Run the application
if (require.main === module) {
  main().catch((error) => {
    logger.error('Unhandled error:', error);
    process.exit(1);
  });
}

export default MonitoreoGmail;
