import pdfParse from 'pdf-parse';
import { logger } from '../utils/logger';
import { SolicitudParsed } from '../types';

export class PDFParser {
  private static readonly REGEX_PATTERNS = {
    // Pattern to extract company name from the top of the PDF (usually in uppercase)
    companyName: /^\s*([A-Z]{3,}(?:\s+[A-Z]{3,})*)\s*$/m,
    // Alternative company name patterns
    companyNameAlt: /\b([A-Z]{4,}(?:\s+[A-Z]{4,})*)\b/g,
    // Common company names we expect (including OPTIMAL)
    knownCompanies: /(CAZAN|OPTIMAL|OFFSHORE|CORPORATIVO|EMPRESA|COMPAÑIA|SERVICIOS)/i,
    // More flexible patterns for CAZAN and OPTIMAL specifically  
    specificCompanies: /\b(CAZAN|OPTIMAL)\b/i,
    // Pattern to match full company names like "CAZAN VISION EMPRESARIAL SA DE CV"
    fullCompanyName: /\b(CAZAN|OPTIMAL)[\s\w]*(?:SA\s+DE\s+CV|S\.A\.|EMPRESARIAL|CORPORATION|CORP|INC|LLC)?/i,
    // Pattern to extract company name from lines that contain key identifiers
    companyLinePattern: /^\s*([A-Z][A-Z\s]+(?:SA\s+DE\s+CV|S\.A\.|EMPRESARIAL|CORPORATION|CORP|INC|LLC)?)\s*$/
  };

  async parsePDF(pdfBuffer: Buffer): Promise<SolicitudParsed> {
    try {
      const data = await pdfParse(pdfBuffer);
      const text = data.text;
      
      logger.info('PDF parsed successfully, extracting data...');
      logger.debug('PDF text content:', text.substring(0, 500) + '...');
      
      const parsedData = this.extractDataFromText(text);
      
      return parsedData;
    } catch (error) {
      logger.error('Error parsing PDF:', error);
      throw new Error(`PDF parsing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  private extractCompanyName(text: string): string {
    logger.debug('Extracting company name from PDF text...');
    
    // Split text into lines and check the first few lines for company name
    const lines = text.split('\n').slice(0, 15); // Check first 15 lines
    
    logger.debug(`Checking first ${lines.length} lines for company name:`);
    lines.forEach((line, index) => {
      logger.debug(`Line ${index + 1}: "${line.trim()}"`);
    });
    
    // First priority: Look for full company names like "CAZAN VISION EMPRESARIAL SA DE CV"
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const fullMatch = line.match(PDFParser.REGEX_PATTERNS.fullCompanyName);
      if (fullMatch) {
        const companyName = fullMatch[0].trim().toUpperCase();
        logger.info(`Found full company name "${companyName}" on line ${i + 1}`);
        return companyName;
      }
    }
    
    // Second priority: Look for company line patterns (complete lines with company structure)
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const lineMatch = line.match(PDFParser.REGEX_PATTERNS.companyLinePattern);
      if (lineMatch && lineMatch[1].length >= 10) { // Ensure meaningful company name
        const companyName = lineMatch[1].trim().toUpperCase();
        logger.info(`Found company line pattern "${companyName}" on line ${i + 1}`);
        return companyName;
      }
    }
    
    // Third priority: Look for CAZAN or OPTIMAL specifically anywhere in first lines
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const specificMatch = line.match(PDFParser.REGEX_PATTERNS.specificCompanies);
      if (specificMatch) {
        const companyName = specificMatch[1].toUpperCase();
        logger.info(`Found specific company name "${companyName}" on line ${i + 1}`);
        return companyName;
      }
    }
    
    // Second priority: Look for other known companies
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const knownMatch = line.match(PDFParser.REGEX_PATTERNS.knownCompanies);
      if (knownMatch) {
        const companyName = knownMatch[1].toUpperCase();
        logger.info(`Found known company name "${companyName}" on line ${i + 1}`);
        return companyName;
      }
    }
    
    // Third priority: Look for uppercase words that could be company names
    for (let i = 0; i < Math.min(5, lines.length); i++) {
      const line = lines[i].trim();
      const companyMatch = line.match(PDFParser.REGEX_PATTERNS.companyName);
      if (companyMatch && companyMatch[1].length >= 4 && companyMatch[1].length <= 25) {
        const companyName = companyMatch[1].trim();
        logger.info(`Found potential company name "${companyName}" on line ${i + 1}`);
        return companyName;
      }
    }
    
    // Fourth priority: Look for any occurrence of CAZAN or OPTIMAL in entire text
    const globalMatch = text.match(PDFParser.REGEX_PATTERNS.specificCompanies);
    if (globalMatch) {
      const companyName = globalMatch[1].toUpperCase();
      logger.info(`Found specific company name "${companyName}" in full text`);
      return companyName;
    }
    
    logger.warn('No company name found in PDF text');
    
    return 'EMPRESA NO IDENTIFICADA';
  }

  static extractAmountFromSubject(subject: string): string {
    // Common patterns for amounts in email subjects
    const amountPatterns = [
      /\$\s*([\d,]+(?:\.\d{2})?)/,
      /([\d,]+(?:\.\d{2})?)\s*(?:pesos|usd|dolares|dólares)/i,
      /(?:monto|cantidad|total|valor)\s*:?\s*\$?\s*([\d,]+(?:\.\d{2})?)/i,
      /([\d,]+(?:\.\d{2})?)\s*-/
    ];
    
    for (const pattern of amountPatterns) {
      const match = subject.match(pattern);
      if (match) {
        return match[1].replace(/,/g, '');
      }
    }
    
    return '';
  }

  static extractCityFromSubject(subject: string): string {
    // Common city patterns in email subjects
    const cityPatterns = [
      /(?:ciudad|city|lugar|ubicación|ubicacion)\s*:?\s*([a-zA-ZñÑáéíóúÁÉÍÓÚ\s]+?)(?:\s*-|\s*\$|$)/i,
      /\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)*)\b(?=\s*\$|\s*-|$)/,
      /\b(TABASCO|CANCUN|CANCÚN|PLAYA|CARMEN|CABOS|COZUMEL|MÉRIDA|MERIDA|GUADALAJARA|MONTERREY|TIJUANA|PUEBLA|QUERETARO|QUERÉTARO|VERACRUZ|ACAPULCO|MAZATLÁN|MAZATLAN|PUERTO\s+VALLARTA|VALLARTA)\b/i,
    ];
    
    for (const pattern of cityPatterns) {
      const match = subject.match(pattern);
      if (match) {
        return match[1].trim();
      }
    }
    
    // Try to extract from common patterns like "RECURSO TABASCO" or "los cabos"
    const locationMatch = subject.match(/\b(?:RECURSO|los|las)\s+([a-zA-ZñÑáéíóúÁÉÍÓÚ\s]+?)\s*\$/i);
    if (locationMatch) {
      return locationMatch[1].trim();
    }
    
    return '';
  }

  private extractDataFromText(text: string): SolicitudParsed {
    const cleanText = text.replace(/\s+/g, ' ').trim();
    
    // Extract only company name from PDF
    const companyName = this.extractCompanyName(text);
    
    return {
      cantidad: '', // Will be extracted from email subject
      domicilioEntrega: companyName || 'No especificada',
      fechaEntrega: '', // Will be extracted from email arrival date
      horario: '', // Not needed per user requirements
      ciudad: '' // Will be extracted from email subject
    };
  }

  private extractField(text: string, fieldName: string, primaryPattern: RegExp, alternativePattern?: RegExp): string {
    // Try primary pattern first
    let match = text.match(primaryPattern);
    if (match && match[1]) {
      logger.debug(`Found ${fieldName} using primary pattern: ${match[1]}`);
      return match[1].trim();
    }

    // Try alternative pattern if provided
    if (alternativePattern) {
      match = text.match(alternativePattern);
      if (match && match[1]) {
        logger.debug(`Found ${fieldName} using alternative pattern: ${match[1]}`);
        return match[1].trim();
      }
    }

    logger.warn(`Could not extract ${fieldName} from PDF text`);
    return '';
  }

  private cleanCantidad(cantidad: string): string {
    if (!cantidad) return '';
    
    // Remove currency symbols and clean up formatting
    return cantidad.replace(/[$,]/g, '').trim();
  }

  private cleanDomicilio(domicilio: string): string {
    if (!domicilio) return '';
    
    // Clean up common formatting issues
    return domicilio
      .replace(/\s+/g, ' ')
      .replace(/[,\s]*$/, '') // Remove trailing commas or spaces
      .trim();
  }

  private cleanFecha(fecha: string): string {
    if (!fecha) return '';
    
    // Normalize date format to DD/MM/YYYY
    const dateMatch = fecha.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dateMatch) {
      const [, day, month, year] = dateMatch;
      const fullYear = year.length === 2 ? `20${year}` : year;
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${fullYear}`;
    }
    
    return fecha.trim();
  }

  private cleanHorario(horario: string): string {
    if (!horario) return '';
    
    // Normalize time format
    return horario
      .replace(/\s+/g, ' ')
      .replace(/–/g, '-') // Replace em dash with hyphen
      .trim();
  }

  private cleanCiudad(ciudad: string): string {
    if (!ciudad) return '';
    
    return ciudad
      .replace(/\s+/g, ' ')
      .replace(/[,\s]*$/, '') // Remove trailing commas or spaces
      .trim();
  }

  private validateParsedData(data: SolicitudParsed): void {
    const missingFields: string[] = [];
    
    if (!data.cantidad) missingFields.push('cantidad');
    if (!data.domicilioEntrega) missingFields.push('domicilioEntrega');
    if (!data.fechaEntrega) missingFields.push('fechaEntrega');
    if (!data.horario) missingFields.push('horario');
    if (!data.ciudad) missingFields.push('ciudad');
    
    if (missingFields.length > 0) {
      logger.warn(`Missing required fields: ${missingFields.join(', ')}`);
      // Don't throw error for missing fields, just log warning
      // The Excel writer can handle empty fields
    }
    
    logger.info('PDF data extraction completed:', {
      cantidad: data.cantidad,
      domicilioEntrega: data.domicilioEntrega.substring(0, 50) + '...',
      fechaEntrega: data.fechaEntrega,
      horario: data.horario,
      ciudad: data.ciudad
    });
  }

  // Helper method to extract ciudad from email subject as fallback
  static extractCiudadFromSubject(subject: string): string {
    // Common patterns for city in email subjects
    const cityPatterns = [
      /solicitud\s+[\-\–]\s+(.+)/i,
      /servicio\s+en\s+(.+)/i,
      /entrega\s+en\s+(.+)/i,
      /(.+)\s+[\-\–]\s+solicitud/i
    ];
    
    for (const pattern of cityPatterns) {
      const match = subject.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
    
    return '';
  }
}

export default PDFParser;
