import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { logger } from '../utils/logger';
import { config } from '../utils/config';
import { SolicitudData } from '../types';

export class ExcelWriter {
  private filePath: string;
  private sheetName: string;
  private workbook: ExcelJS.Workbook;
  private worksheet: ExcelJS.Worksheet | null = null;

  constructor() {
    this.filePath = path.resolve(config.excel.filePath);
    this.sheetName = config.excel.sheetName;
    this.workbook = new ExcelJS.Workbook();
  }

  async initialize(): Promise<void> {
    try {
      await this.loadOrCreateWorkbook();
      await this.setupWorksheet();
      logger.info('Excel writer initialized successfully');
    } catch (error) {
      logger.error('Error initializing Excel writer:', error);
      throw error;
    }
  }

  private async loadOrCreateWorkbook(): Promise<void> {
    try {
      if (fs.existsSync(this.filePath)) {
        // Load existing workbook
        await this.workbook.xlsx.readFile(this.filePath);
        logger.info(`Loaded existing Excel file: ${this.filePath}`);
      } else {
        // Create new workbook
        logger.info(`Creating new Excel file: ${this.filePath}`);
        
        // Ensure directory exists
        const dir = path.dirname(this.filePath);
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
        }
      }
    } catch (error) {
      logger.error('Error loading/creating workbook:', error);
      throw error;
    }
  }

  private async setupWorksheet(): Promise<void> {
    try {
      // Get or create worksheet
      this.worksheet = this.workbook.getWorksheet(this.sheetName) || null;
      
      if (!this.worksheet) {
        this.worksheet = this.workbook.addWorksheet(this.sheetName);
        logger.info(`Created new worksheet: ${this.sheetName}`);
      }

      // Setup headers if worksheet is empty
      if (this.worksheet.rowCount === 0) {
        await this.setupHeaders();
      }
    } catch (error) {
      logger.error('Error setting up worksheet:', error);
      throw error;
    }
  }

  private async setupHeaders(): Promise<void> {
    if (!this.worksheet) {
      throw new Error('Worksheet not initialized');
    }

    const headers = [
      'FECHA',
      'OPERADOR', 
      'SUCURSAL',
      'MONTO',
      'LINK'
    ];

    // Add headers to the first row
    this.worksheet.addRow(headers);

    // Style the headers
    const headerRow = this.worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // Auto-fit columns
    this.worksheet.columns.forEach((column) => {
      column.width = 15;
    });

    // Set specific column widths for new columns
    this.worksheet.getColumn(1).width = 12; // FECHA
    this.worksheet.getColumn(2).width = 25; // OPERADOR
    this.worksheet.getColumn(3).width = 15; // SUCURSAL
    this.worksheet.getColumn(4).width = 15; // MONTO
    this.worksheet.getColumn(5).width = 40; // LINK
    this.worksheet.getColumn(6).width = 0;  // MessageId - hidden column
    this.worksheet.getColumn(6).hidden = true;

    await this.saveWorkbook();
    logger.info('Headers added to Excel worksheet');
  }

  async appendSolicitud(solicitud: SolicitudData): Promise<void> {
    if (!this.worksheet) {
      throw new Error('Worksheet not initialized');
    }

    try {
      // Check if this message has already been processed
      const existingRow = this.findExistingRow(solicitud.messageId);
      if (existingRow) {
        logger.warn(`Message ${solicitud.messageId} already exists in Excel file`);
        return;
      }

      // Prepare row data for new column structure (including hidden messageId column)
      const rowData = [
        solicitud.fecha,     // FECHA - already in short format
        solicitud.operador,  // OPERADOR - company name
        solicitud.sucursal,  // SUCURSAL - city
        solicitud.monto,     // MONTO - amount (will be formatted as accounting)
        solicitud.link,      // LINK - Gmail URL
        solicitud.messageId  // MessageId - hidden column for duplicate detection
      ];

      // Add the row
      const newRow = this.worksheet.addRow(rowData);

      // Format MONTO column (column 4) as accounting number
      const montoCell = newRow.getCell(4);
      if (solicitud.monto && solicitud.monto !== 'No especificada') {
        // Try to parse the amount and format as accounting
        const amount = this.parseAmount(solicitud.monto);
        if (amount !== null) {
          montoCell.value = amount;
          montoCell.numFmt = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';
        }
      }

      // Format LINK column (column 5) as hyperlink
      const linkCell = newRow.getCell(5);
      if (solicitud.link) {
        linkCell.value = {
          text: 'Ver correo',
          hyperlink: solicitud.link
        };
        linkCell.font = { color: { argb: 'FF0000FF' }, underline: true };
      }
      
      // Convert data range to table for filtering
      await this.createFilterableTable();

      // Style the new row
      newRow.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      // Date formatting is already handled in the data preparation
      // Currency formatting is already handled above for MONTO column

      await this.saveWorkbook();
      
      logger.info(`Added solicitud to Excel: ${solicitud.operador} - ${solicitud.sucursal}`);
    } catch (error) {
      logger.error('Error appending solicitud to Excel:', error);
      throw error;
    }
  }

  private findExistingRow(messageId: string): ExcelJS.Row | null {
    if (!this.worksheet) {
      return null;
    }

    const messageIdColumn = 6; // MessageId is stored after the main columns (hidden column)
    
    for (let i = 2; i <= this.worksheet.rowCount; i++) {
      const row = this.worksheet.getRow(i);
      const cellValue = row.getCell(messageIdColumn).value;
      
      if (cellValue === messageId) {
        return row;
      }
    }

    return null;
  }

  private parseAmount(amountString: string): number | null {
    if (!amountString || amountString === 'No especificada') {
      return null;
    }
    
    // Remove currency symbols, commas, and extract numeric value
    const cleanAmount = amountString.replace(/[$,]/g, '').trim();
    const numericValue = parseFloat(cleanAmount);
    
    if (isNaN(numericValue)) {
      return null;
    }
    
    return numericValue;
  }

  private async createFilterableTable(): Promise<void> {
    if (!this.worksheet || this.worksheet.rowCount <= 1) {
      return; // Need at least header + 1 data row
    }

    try {
      // Create auto-filter for the data range (A1 to F + last row)
      const filterRange = `A1:F${this.worksheet.rowCount}`;
      
      // Apply auto-filter to make data filterable
      this.worksheet.autoFilter = filterRange;
      
      logger.info(`Created filterable range: ${filterRange}`);
    } catch (error) {
      logger.warn('Could not create filterable range:', error);
    }
  }

  private async saveWorkbook(): Promise<void> {
    try {
      await this.workbook.xlsx.writeFile(this.filePath);
      logger.debug(`Excel file saved: ${this.filePath}`);
    } catch (error) {
      logger.error('Error saving Excel file:', error);
      throw error;
    }
  }

  async getStatistics(): Promise<{
    totalRows: number;
    lastUpdated: Date;
    fileName: string;
  }> {
    return {
      totalRows: this.worksheet ? this.worksheet.rowCount - 1 : 0, // Exclude header
      lastUpdated: new Date(),
      fileName: path.basename(this.filePath)
    };
  }

  async backup(): Promise<string> {
    try {
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupPath = this.filePath.replace('.xlsx', `_backup_${timestamp}.xlsx`);
      
      await this.workbook.xlsx.writeFile(backupPath);
      logger.info(`Backup created: ${backupPath}`);
      
      return backupPath;
    } catch (error) {
      logger.error('Error creating backup:', error);
      throw error;
    }
  }

  // Method to validate Excel file integrity
  async validateFile(): Promise<boolean> {
    try {
      if (!fs.existsSync(this.filePath)) {
        logger.warn('Excel file does not exist');
        return false;
      }

      const tempWorkbook = new ExcelJS.Workbook();
      await tempWorkbook.xlsx.readFile(this.filePath);
      
      const tempWorksheet = tempWorkbook.getWorksheet(this.sheetName);
      if (!tempWorksheet) {
        logger.warn('Expected worksheet not found');
        return false;
      }

      logger.info('Excel file validation passed');
      return true;
    } catch (error) {
      logger.error('Excel file validation failed:', error);
      return false;
    }
  }
}

export default ExcelWriter;
