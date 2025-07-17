# Monitoreo Gmail - Automated PDF Extraction and Excel Consolidation

Sistema automatizado para extraer datos de solicitudes de servicio recibidas en Gmail (desde archivos PDF adjuntos) y consolidar la información en una hoja de cálculo Excel en tiempo real.

## Características

- ✅ Monitoreo automático de Gmail usando OAuth 2.0
- ✅ Extracción de datos estructurados de PDFs con palabra "CARTA"
- ✅ Consolidación automática en archivo Excel
- ✅ Manejo de errores y logging completo
- ✅ Polling configurable (por defecto cada 5 minutos)
- ✅ Respaldo automático de archivos Excel
- ✅ Prevención de duplicados por Message ID

## Requisitos Previos

### 1. Configuración de Google Cloud Platform

1. **Crear proyecto en Google Cloud Console:**
   - Ir a [Google Cloud Console](https://console.cloud.google.com/)
   - Crear nuevo proyecto o seleccionar existente

2. **Habilitar Gmail API:**
   ```bash
   # En Google Cloud Console
   APIs & Services → Library → Gmail API → Enable
   ```

3. **Configurar pantalla de consentimiento OAuth 2.0:**
   - APIs & Services → OAuth consent screen
   - Configurar como "Externa" si es necesario
   - Completar información básica del proyecto

4. **Crear credenciales OAuth 2.0:**
   - APIs & Services → Credentials
   - Create Credentials → OAuth 2.0 Client ID
   - Tipo: Desktop Application
   - Descargar archivo `credentials.json`

### 2. Requisitos del Sistema

- Node.js 18+ y npm
- Cuenta de Gmail con acceso a API
- Archivos PDF con estructura consistente

## Instalación

1. **Clonar e instalar dependencias:**
```bash
npm install
```

2. **Configurar variables de entorno:**
```bash
cp .env.example .env
# Editar .env con tus configuraciones
```

3. **Colocar credenciales:**
```bash
# Colocar credentials.json descargado de Google Cloud en la raíz del proyecto
```

4. **Compilar TypeScript:**
```bash
npm run build
```

## Configuración

### Archivo `.env`
```env
# Gmail API Configuration
GMAIL_CREDENTIALS_PATH=./credentials.json
GMAIL_TOKEN_PATH=./token.json

# Excel Configuration
EXCEL_FILE_PATH=./solicitudes.xlsx
EXCEL_SHEET_NAME=Solicitudes

# Monitoring Configuration
POLLING_INTERVAL_MS=300000  # 5 minutos
LOG_LEVEL=info

# Error Handling
RETRY_ATTEMPTS=3
RETRY_DELAY_MS=5000
```

## Uso

### Desarrollo
```bash
npm run dev
```

### Producción
```bash
npm run build
npm start
```

### Primera Ejecución
En la primera ejecución, se abrirá un navegador para autorizar el acceso a Gmail:
1. Autorizar la aplicación
2. Copiar el código de autorización
3. Pegar en la terminal cuando se solicite

## Estructura del Proyecto

```
src/
├── index.ts              # Punto de entrada principal
├── gmail/
│   ├── auth.ts          # Autenticación OAuth 2.0
│   └── watcher.ts       # Monitoreo y descarga de emails
├── pdf/
│   └── parser.ts        # Extracción de datos de PDFs
├── excel/
│   └── writer.ts        # Escritura y actualización de Excel
├── utils/
│   ├── config.ts        # Configuración centralizada
│   └── logger.ts        # Sistema de logging
└── types.ts             # Definiciones de tipos TypeScript
```

## Formato de Datos Extraídos

El sistema extrae los siguientes campos de cada PDF:

| Campo | Descripción | Ejemplo |
|-------|-------------|---------|
| `Cantidad` | Monto monetario | `$1,500.00` |
| `Domicilio de Entrega` | Dirección completa | `Calle 123 #45, Colonia Centro` |
| `Fecha de Entrega` | Fecha del servicio | `15/12/2023` |
| `Horario` | Rango de horas | `09:00 - 17:00` |
| `Ciudad` | Ciudad de entrega | `Guadalajara` |

## Archivo Excel Generado

El archivo Excel contiene las siguientes columnas:
- Fecha de Recepción
- Cliente
- Asunto
- Cantidad
- Fecha de Entrega
- Horario
- Domicilio de Entrega
- Ciudad
- Message ID

## Logging y Monitoreo

Los logs se guardan en:
- `logs/combined.log` - Todos los logs
- `logs/error.log` - Solo errores
- Consola (en desarrollo)

Niveles de log disponibles: `error`, `warn`, `info`, `debug`

## Manejo de Errores

El sistema incluye:
- Reintentos automáticos para operaciones de red
- Validación de estructura de PDFs
- Manejo de archivos corruptos
- Prevención de duplicados
- Respaldos automáticos

## Patrones de Extracción PDF

Los patrones regex utilizados para extraer datos:

```typescript
// Cantidad
/(?:cantidad|monto|total|valor):\s*\$?\s*([\d,]+(?:\.\d{2})?)/i

// Domicilio
/(?:domicilio de entrega|dirección de entrega|entrega en):\s*(.+?)(?:\n|$)/i

// Fecha
/(?:fecha de entrega|entrega el|fecha del servicio):\s*([\d]{1,2}[\/\-][\d]{1,2}[\/\-][\d]{2,4})/i

// Horario
/(?:horario|hora):\s*([\d]{1,2}:[\d]{2}\s*(?:am|pm)?\s*[\-\–]\s*[\d]{1,2}:[\d]{2}\s*(?:am|pm)?)/i
```

## Comandos Disponibles

- `npm run dev` - Desarrollo con hot-reload
- `npm run build` - Compilar TypeScript
- `npm start` - Ejecutar en producción
- `npm run watch` - Compilar en modo watch

## Solución de Problemas

### Error: "No se puede encontrar credentials.json"
- Verificar que el archivo esté en la ruta correcta
- Revisar la variable `GMAIL_CREDENTIALS_PATH` en `.env`

### Error: "Token expirado"
- Eliminar `token.json` y volver a autorizar
- Verificar que las credenciales sean válidas

### Error: "No se pueden extraer datos del PDF"
- Verificar que el PDF contenga "CARTA" en el nombre
- Revisar que el formato del PDF sea el esperado
- Consultar logs para patrones regex no coincidentes

### Error: "No se puede escribir en Excel"
- Verificar permisos de escritura en el directorio
- Cerrar Excel si está abierto
- Verificar que la ruta del archivo sea correcta

## Deployment

Para ambiente de producción:
1. Usar PM2 o similar para manejo de procesos
2. Configurar rotación de logs
3. Monitorear el uso de memoria
4. Configurar alertas de error

```bash
# Ejemplo con PM2
pm2 start dist/index.js --name monitoreo-gmail
pm2 logs monitoreo-gmail
```

## Contribuciones

Para contribuir:
1. Fork el proyecto
2. Crear branch para feature
3. Commit con mensajes descriptivos
4. Crear Pull Request

## Licencia

ISC License

## Soporte

Para soporte técnico:
- Revisar logs en `logs/`
- Verificar configuración en `.env`
- Consultar documentación de Gmail API
