/*******************************************************
 * JOB EMAIL COLLECTOR
 *
 * Gmail -> Google Sheets
 *
 * NO clasifica emails.
 * NO interpreta Company/Position/Status.
 *
 * Su única función es:
 *
 * 1. Buscar threads relacionados con Job/Analyze
 * 2. Obtener sus mensajes
 * 3. Agregar solamente mensajes nuevos
 * 4. Guardarlos en Google Sheets
 *
 * La clasificación será realizada posteriormente
 * por Ollama + granite4.2 en Linux.
 *******************************************************/


const CONFIG = {

  LABEL_NAME: 'Job/Analyze',

  FOLDER_NAME: 'Job2026',

  SPREADSHEET_NAME: 'Job_Emails_2026',

  SHEET_NAME: 'Emails',

  START_DATE: new Date(
    '2026-09-14T00:00:00-06:00'
  ),

  TIMEZONE: 'America/Mexico_City',

  TRIGGER_FUNCTION: 'updateJobEmailSheet',

  BATCH_SIZE: 100
};


/*******************************************************
 * SETUP INICIAL
 *
 * Ejecutar UNA sola vez.
 *******************************************************/

function setupJobEmailCollector() {

  console.log(
    '========================================'
  );

  console.log(
    'Configurando Job Email Collector...'
  );


  /*****************************************
   * 1. CARPETA
   *****************************************/

  const folder =
    getOrCreateFolder_(
      CONFIG.FOLDER_NAME
    );


  /*****************************************
   * 2. SPREADSHEET
   *****************************************/

  const properties =
    PropertiesService
      .getScriptProperties();


  let spreadsheetId =
    properties.getProperty(
      'SPREADSHEET_ID'
    );


  let spreadsheet = null;


  if (spreadsheetId) {

    try {

      spreadsheet =
        SpreadsheetApp.openById(
          spreadsheetId
        );

    } catch (error) {

      console.log(
        'Spreadsheet anterior no disponible.'
      );

      spreadsheet = null;
    }
  }


  /*****************************************
   * Crear spreadsheet si no existe
   *****************************************/

  if (!spreadsheet) {

    spreadsheet =
      SpreadsheetApp.create(
        CONFIG.SPREADSHEET_NAME
      );


    spreadsheetId =
      spreadsheet.getId();


    properties.setProperty(
      'SPREADSHEET_ID',
      spreadsheetId
    );


    DriveApp
      .getFileById(
        spreadsheetId
      )
      .moveTo(folder);


    console.log(
      'Spreadsheet creado: ' +
      spreadsheet.getUrl()
    );
  }


  /*****************************************
   * 3. PREPARAR SHEET
   *****************************************/

  prepareSheet_(
    spreadsheet
  );


  /*****************************************
   * 4. PRIMERA IMPORTACIÓN
   *****************************************/

  updateJobEmailSheet();


  /*****************************************
   * 5. TRIGGER
   *****************************************/

  createDailyTrigger_();


  console.log(
    '========================================'
  );

  console.log(
    'CONFIGURACIÓN COMPLETADA'
  );

  console.log(
    'Spreadsheet: ' +
    spreadsheet.getUrl()
  );

  console.log(
    '========================================'
  );
}


/*******************************************************
 * ACTUALIZAR SHEET
 *******************************************************/

function updateJobEmailSheet() {

  console.log(
    'Iniciando actualización de emails...'
  );


  /*****************************************
   * Spreadsheet
   *****************************************/

  const properties =
    PropertiesService
      .getScriptProperties();


  const spreadsheetId =
    properties.getProperty(
      'SPREADSHEET_ID'
    );


  if (!spreadsheetId) {

    throw new Error(
      'No existe SPREADSHEET_ID. ' +
      'Ejecuta setupJobEmailCollector() primero.'
    );
  }


  const spreadsheet =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {

    throw new Error(
      'No existe la hoja ' +
      CONFIG.SHEET_NAME
    );
  }


  /*****************************************
   * Label
   *****************************************/

  const label =
    GmailApp.getUserLabelByName(
      CONFIG.LABEL_NAME
    );


  if (!label) {

    throw new Error(
      'No encontré el label "' +
      CONFIG.LABEL_NAME +
      '".'
    );
  }


  /*****************************************
   * Leer Message IDs existentes
   *****************************************/

  const existingIds =
    getExistingMessageIds_(
      sheet
    );


  console.log(
    'Emails existentes en Sheet: ' +
    existingIds.size
  );


  /*****************************************
   * Obtener threads
   *****************************************/

  const threads = [];

  let start = 0;


  while (true) {

    const batch =
      label.getThreads(
        start,
        CONFIG.BATCH_SIZE
      );


    if (
      !batch ||
      batch.length === 0
    ) {

      break;
    }


    threads.push.apply(
      threads,
      batch
    );


    if (
      batch.length <
      CONFIG.BATCH_SIZE
    ) {

      break;
    }


    start +=
      CONFIG.BATCH_SIZE;
  }


  console.log(
    'Threads encontrados: ' +
    threads.length
  );


  /*****************************************
   * Procesar mensajes
   *****************************************/

  const newRows = [];

  const retrievedAt =
    new Date();


  threads.forEach(
    function(thread) {

      let messages;

      try {

        messages =
          thread.getMessages();

      } catch (error) {

        console.log(
          'No se pudo leer thread: ' +
          thread.getId()
        );

        return;
      }


      messages.forEach(
        function(message) {

          const messageDate =
            message.getDate();


          /**********************************
           * Fecha mínima
           **********************************/

          if (
            messageDate <
            CONFIG.START_DATE
          ) {

            return;
          }


          /**********************************
           * Message ID
           **********************************/

          const messageId =
            message.getId();


          /**********************************
           * Evitar duplicados
           **********************************/

          if (
            existingIds.has(
              messageId
            )
          ) {

            return;
          }


          /**********************************
           * Datos del email
           **********************************/

          let body = '';

          try {

            body =
              message.getPlainBody();

          } catch (error) {

            body = '';
          }


          let threadUrl = '';

          try {

            threadUrl =
              thread.getPermalink();

          } catch (error) {

            threadUrl = '';
          }


          const snippet =
            createSnippet_(
              body,
              500
            );


          /**********************************
           * Nueva fila
           **********************************/

          newRows.push([
            messageId,
            thread.getId(),
            messageDate,
            message.getFrom(),
            message.getTo(),
            message.getSubject(),
            body,
            snippet,
            threadUrl,
            retrievedAt
          ]);


          existingIds.add(
            messageId
          );
        }
      );
    }
  );


  /*****************************************
   * No hay emails nuevos
   *****************************************/

  if (
    newRows.length === 0
  ) {

    console.log(
      'No hay emails nuevos.'
    );

    return;
  }


  /*****************************************
   * Ordenar nuevos emails
   * cronológicamente
   *****************************************/

  newRows.sort(
    function(a, b) {

      return (
        a[2].getTime() -
        b[2].getTime()
      );
    }
  );


  /*****************************************
   * Agregar al final
   *****************************************/

  const startRow =
    sheet.getLastRow() + 1;


  sheet
    .getRange(
      startRow,
      1,
      newRows.length,
      newRows[0].length
    )
    .setValues(
      newRows
    );


  /*****************************************
   * Formato
   *****************************************/

  formatNewRows_(
    sheet,
    startRow,
    newRows.length
  );


  console.log(
    'Emails nuevos agregados: ' +
    newRows.length
  );


  console.log(
    'Actualización completada.'
  );
}


/*******************************************************
 * OBTENER IDS EXISTENTES
 *******************************************************/

function getExistingMessageIds_(
  sheet
) {

  const ids =
    new Set();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return ids;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  values.forEach(
    function(row) {

      if (
        row[0]
      ) {

        ids.add(
          String(row[0])
        );
      }
    }
  );


  return ids;
}


/*******************************************************
 * CREAR SHEET
 *******************************************************/

function prepareSheet_(
  spreadsheet
) {

  let sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        CONFIG.SHEET_NAME
      );
  }


  const headers = [
    'Message_ID',
    'Thread_ID',
    'Date',
    'From',
    'To',
    'Subject',
    'Body',
    'Snippet',
    'Gmail_Link',
    'Retrieved_At'
  ];


  /*****************************************
   * Si está vacío, crear headers
   *****************************************/

  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);
  }


  /*****************************************
   * Formato header
   *****************************************/

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setFontWeight(
      'bold'
    )
    .setBackground(
      '#202124'
    )
    .setFontColor(
      '#ffffff'
    );


  sheet.setFrozenRows(1);


  /*****************************************
   * Anchos
   *****************************************/

  sheet.setColumnWidth(
    1,
    220
  );

  sheet.setColumnWidth(
    2,
    220
  );

  sheet.setColumnWidth(
    3,
    160
  );

  sheet.setColumnWidth(
    4,
    280
  );

  sheet.setColumnWidth(
    5,
    280
  );

  sheet.setColumnWidth(
    6,
    400
  );

  sheet.setColumnWidth(
    7,
    700
  );

  sheet.setColumnWidth(
    8,
    350
  );

  sheet.setColumnWidth(
    9,
    120
  );

  sheet.setColumnWidth(
    10,
    160
  );
}


/*******************************************************
 * FORMATO NUEVAS FILAS
 *******************************************************/

function formatNewRows_(
  sheet,
  startRow,
  rowCount
) {

  /*****************************************
   * Fecha
   *****************************************/

  sheet
    .getRange(
      startRow,
      3,
      rowCount,
      1
    )
    .setNumberFormat(
      'yyyy-mm-dd HH:mm:ss'
    );


  /*****************************************
   * Retrieved At
   *****************************************/

  sheet
    .getRange(
      startRow,
      10,
      rowCount,
      1
    )
    .setNumberFormat(
      'yyyy-mm-dd HH:mm:ss'
    );


  /*****************************************
   * Body
   *****************************************/

  sheet
    .getRange(
      startRow,
      7,
      rowCount,
      4
    )
    .setWrap(
      true
    );


  /*****************************************
   * Link
   *****************************************/

  sheet
    .getRange(
      startRow,
      9,
      rowCount,
      1
    )
    .setFontColor(
      '#1155cc'
    );
}


/*******************************************************
 * SNIPPET
 *******************************************************/

function createSnippet_(
  text,
  maxLength
) {

  if (!text) {
    return '';
  }


  const clean =
    String(text)
      .replace(
        /\s+/g,
        ' '
      )
      .trim();


  if (
    clean.length <= maxLength
  ) {

    return clean;
  }


  return (
    clean.substring(
      0,
      maxLength
    ) +
    '...'
  );
}


/*******************************************************
 * CARPETA
 *******************************************************/

function getOrCreateFolder_(
  folderName
) {

  const folders =
    DriveApp.getFoldersByName(
      folderName
    );


  if (
    folders.hasNext()
  ) {

    return folders.next();
  }


  return DriveApp.createFolder(
    folderName
  );
}


/*******************************************************
 * TRIGGER DIARIO
 *******************************************************/

function createDailyTrigger_() {

  const triggers =
    ScriptApp.getProjectTriggers();


  /*****************************************
   * Eliminar triggers existentes
   *****************************************/

  triggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        CONFIG.TRIGGER_FUNCTION
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );


  /*****************************************
   * Crear nuevo trigger
   *****************************************/

  ScriptApp
    .newTrigger(
      CONFIG.TRIGGER_FUNCTION
    )
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();


  console.log(
    'Trigger diario creado.'
  );
}


/*******************************************************
 * TEST MANUAL
 *******************************************************/

function testUpdate() {

  updateJobEmailSheet();

}
