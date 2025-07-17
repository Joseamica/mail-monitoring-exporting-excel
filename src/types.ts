export interface SolicitudParsed {
  cantidad: string;
  domicilioEntrega: string;
  fechaEntrega: string;
  horario: string;
  ciudad: string;
}

export interface SolicitudData {
  fecha: string;        // FECHA - fecha de recepción en formato corto (24-jun)
  operador: string;     // OPERADOR - nombre del cliente/empresa
  sucursal: string;     // SUCURSAL - ciudad
  monto: string;        // MONTO - cantidad en formato contabilidad
  link: string;         // LINK - URL para acceder al correo específico
  messageId: string;    // Para identificación interna
}

export interface GmailAttachment {
  filename: string;
  mimeType: string;
  attachmentId: string;
  size: number;
}

export interface ProcessedEmail {
  messageId: string;
  subject: string;
  sender: string;
  receivedAt: Date;
  attachments: GmailAttachment[];
}

export interface Config {
  gmail: {
    credentialsPath: string;
    tokenPath: string;
  };
  excel: {
    filePath: string;
    sheetName: string;
  };
  monitoring: {
    pollingIntervalMs: number;
    logLevel: string;
  };
  errorHandling: {
    retryAttempts: number;
    retryDelayMs: number;
  };
}
