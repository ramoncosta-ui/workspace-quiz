// =========================================================================
// GOOGLE APPS SCRIPT - SERVIDOR DO QUIZ (SEM FIREBASE)
// =========================================================================
// INSTRUÇÕES:
// 1. Na sua Planilha Google, clique em: Extensões > Apps Script
// 2. Apague o código que estiver lá e cole este arquivo inteiro.
// 3. Clique no botão "Salvar" (ícone de disquete).
// 4. Clique em "Implantar" > "Nova Implantação":
//    - Tipo: App da Web
//    - Executar como: Eu (seu email)
//    - Quem tem acesso: Qualquer pessoa
// 5. Clique em "Implantar", autorize as permissões e copie a "URL do App da Web".
// 6. Cole a URL no seu site!
// =========================================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getData';
  var callback = e && e.parameter && e.parameter.callback;
  
  var responseData = {};
  
  if (action === 'getResponses') {
    responseData = {
      result: "success",
      responses: getResponsesFromSheet()
    };
  } else {
    // Retorna perguntas e configurações da aula
    responseData = {
      result: "success",
      settings: getSettings(),
      questions: getQuestionsFromSheet()
    };
  }
  
  var jsonString = JSON.stringify(responseData);
  
  // Suporte a JSONP (evita qualquer bloqueio de CORS em iframes do Google Sites)
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonString + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    var action = data.action || 'submitResponse';
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Salvar ou atualizar perguntas
    if (action === 'saveQuestions') {
      if (data.questions) {
        saveQuestionsToSheet(data.questions);
      }
      if (data.settings) {
        saveSettings(data.settings);
      }
      return createJsonResponse({ result: "success", message: "Perguntas salvas com sucesso" });
    }
    
    // 2. Salvar configurações (Título da aula ou Imagem de Fundo)
    if (action === 'saveSettings') {
      saveSettings(data.settings);
      return createJsonResponse({ result: "success", message: "Configurações salvas" });
    }
    
    // 3. Limpar histórico de respostas
    if (action === 'clearResponses') {
      clearResponsesSheet();
      return createJsonResponse({ result: "success", message: "Respostas limpas" });
    }
    
    // 4. Submissão de respostas de um Aluno
    var sheet = getOrCreateSheet(ss, "Respostas");
    var rowData = [
      data.date || new Date().toLocaleString('pt-BR'),
      data.nome || '',
      data.sobrenome || '',
      data.acertos || '0',
      JSON.stringify(data.respostas || {})
    ];
    sheet.appendRow(rowData);
    
    return createJsonResponse({ result: "success", message: "Resposta registrada" });
  } catch (err) {
    return createJsonResponse({ result: "error", error: err.message });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================
// FUNÇÕES DE AUXÍLIO PARA LEITURA E GRAVAÇÃO NAS ABAS
// =========================================================================

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === "Respostas") {
      sheet.appendRow(["Data", "Nome", "Sobrenome", "Acertos", "JSON Respostas"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#dcfce7");
    } else if (name === "Perguntas") {
      sheet.appendRow(["ID", "Enunciado", "Opcao_A", "Opcao_B", "Opcao_C", "Opcao_D", "Gabarito", "Liberada"]);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e0e7ff");
    }
  }
  return sheet;
}

function getSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    lessonTitle: props.getProperty("LESSON_TITLE") || "Simulado Geral",
    backgroundDriveUrl: props.getProperty("BG_URL") || ""
  };
}

function saveSettings(settings) {
  if (!settings) return;
  var props = PropertiesService.getScriptProperties();
  if (settings.lessonTitle) props.setProperty("LESSON_TITLE", settings.lessonTitle);
  if (settings.backgroundDriveUrl !== undefined) props.setProperty("BG_URL", settings.backgroundDriveUrl);
}

function getQuestionsFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, "Perguntas");
  var values = sheet.getDataRange().getValues();
  
  var questions = [];
  // Pular o cabeçalho (linha 1)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0] && !row[1]) continue;
    
    var visVal = String(row[7]).toLowerCase().trim();
    var isVisible = (visVal === 'true' || visVal === 'sim' || row[7] === true || visVal === '1');
    
    questions.push({
      id: Number(row[0]) || (Date.now() + i),
      title: String(row[1] || ''),
      options: {
        A: String(row[2] || ''),
        B: String(row[3] || ''),
        C: String(row[4] || ''),
        D: String(row[5] || '')
      },
      correct: String(row[6] || 'A').trim().toUpperCase(),
      visible: isVisible
    });
  }
  
  // Se a planilha de perguntas ainda estiver vazia, cria a primeira
  if (questions.length === 0) {
    var defaultQ = {
      id: 1,
      title: "Qual é a capital do Brasil?",
      options: { A: "Rio de Janeiro", B: "São Paulo", C: "Brasília", D: "Salvador" },
      correct: "C",
      visible: true
    };
    saveQuestionsToSheet([defaultQ]);
    return [defaultQ];
  }
  
  return questions;
}

function saveQuestionsToSheet(questions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, "Perguntas");
  
  // Limpar linhas anteriores
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }
  
  if (!questions || questions.length === 0) return;
  
  var rows = questions.map(function(q) {
    return [
      q.id,
      q.title,
      q.options ? q.options.A : '',
      q.options ? q.options.B : '',
      q.options ? q.options.C : '',
      q.options ? q.options.D : '',
      q.correct || 'A',
      q.visible ? 'TRUE' : 'FALSE'
    ];
  });
  
  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
}

function getResponsesFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, "Respostas");
  var values = sheet.getDataRange().getValues();
  
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0] && !row[1]) continue;
    
    var respObj = {};
    try {
      respObj = JSON.parse(row[4]);
    } catch(e) {}
    
    list.push({
      date: String(row[0]),
      nome: String(row[1]),
      sobrenome: String(row[2]),
      acertos: String(row[3]),
      respostas: respObj
    });
  }
  return list;
}

function clearResponsesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, "Respostas");
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
}
