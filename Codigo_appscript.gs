// Configurações
const SPREADSHEET_ID = '';
const TASK_SHEET_NAME = 'Tarefas';
const LOG_SHEET_NAME = 'Logs';
const CONFIG_SHEET_NAME = 'Configurações';

// Inicializar a planilha
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (!taskSheet) {
    taskSheet = ss.insertSheet(TASK_SHEET_NAME);
    taskSheet.getRange('A1:F1').setValues([['Task ID', 'Título', 'Descrição', 'Estado', 'Prioridade', 'Data Criação']]);
  }
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange('A1:C1').setValues([['Data', 'Tipo', 'Mensagem']]);
  }
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.getRange('A1:B1').setValues([['Chave', 'Valor']]);
    configSheet.getRange('A2:B2').setValues([['estados', JSON.stringify(['To Do', 'In Progress', 'Done'])]]);
  }
}

// Registrar log
function registrarLog(sheet, tipo, mensagem) {
  try {
    sheet.appendRow([new Date(), tipo, mensagem]);
  } catch (error) {
    Logger.log('Erro ao registrar log: ' + error.message);
  }
}

// Receber atualizações (tarefas e estados)
function doPost(e) {
  initializeSheets();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  try {
    const dados = JSON.parse(e.postData.contents);

    // Atualizar tarefas
    if (dados.tarefas) {
      const tarefas = dados.tarefas;
      taskSheet.clear();
      taskSheet.appendRow(['Task ID', 'Título', 'Descrição', 'Estado', 'Prioridade', 'Data Criação']);
      const rows = [];
      for (const taskId in tarefas) {
        const task = tarefas[taskId];
        rows.push([
          taskId,
          task.titulo || '',
          task.descricao || '',
          task.estado || 'To Do',
          task.prioridade || 'Média',
          task.data_criacao || new Date().toISOString()
        ]);
      }
      if (rows.length > 0) {
        taskSheet.getRange(2, 1, rows.length, 6).setValues(rows);
      }
      registrarLog(logSheet, 'Sucesso', `Planilha atualizada com ${rows.length} tarefas`);
    }

    // Atualizar estados
    if (dados.estados) {
      const estados = dados.estados;
      configSheet.getRange('A2:B2').setValues([['estados', JSON.stringify(estados)]]);
      registrarLog(logSheet, 'Sucesso', `Estados atualizados: ${estados.join(', ')}`);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Dados sincronizados com sucesso' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    registrarLog(logSheet, 'Erro', `Falha ao processar doPost: ${error.message}`);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Retornar tarefas e estados
function doGet(e) {
  initializeSheets();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  try {
    // Carregar tarefas
    const rows = taskSheet.getDataRange().getValues();
    rows.shift();
    const tarefas = {};
    rows.forEach(function(row) {
      const taskId = row[0].toString();
      tarefas[taskId] = {
        titulo: row[1] || '',
        descricao: row[2] || '',
        estado: row[3] || 'To Do',
        prioridade: row[4] || 'Média',
        data_criacao: row[5] || new Date().toISOString()
      };
    });

    // Carregar estados
    const configData = configSheet.getRange('A2:B2').getValues();
    const estados = configData[0][0] === 'estados' ? JSON.parse(configData[0][1]) : ['To Do', 'In Progress', 'Done'];

    registrarLog(logSheet, 'Sucesso', `Retornadas ${Object.keys(tarefas).length} tarefas e ${estados.length} estados via doGet`);
    return ContentService.createTextOutput(JSON.stringify({ tarefas, estados }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    registrarLog(logSheet, 'Erro', `Falha ao processar doGet: ${error.message}`);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}