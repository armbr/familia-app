// ╔══════════════════════════════════════════════════════════╗
// ║   FAMÍLIA APP — BACKEND OTIMIZADO (FLUXO)                ║
// ║   1. Cole este código no Google Apps Script              ║
// ║   2. Salve (Ctrl+S)                                      ║
// ║   3. Implantar > Gerenciar implantações > ✏️ editar       ║
// ║   4. Versão: "Nova versão" > Implantar                   ║
// ╚══════════════════════════════════════════════════════════╝

var SPREADSHEET_ID = '174UmeWX3kmj9qjl7Z3I8hACNpx1pjcWYsl1_wBFzgxM';

// ════════════════════════════════════════════════════════════
//  ENTRY POINT — tudo via GET para evitar problemas de CORS
// ════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    if (action === 'paginaInquilino') {
      return paginaInquilino(e);
    }

    var body = {};
    if (params.payload) {
      try { body = JSON.parse(decodeURIComponent(params.payload)); } catch(ex) {}
    }
    var result = processAction(action, body);
    
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handler de Ações Centralizado
function processAction(action, body) {
  if (action === 'boot') return doBoot();
  if (action === 'salvarGasto') return doSalvarGasto(body);
  if (action === 'deletarGasto') return doDeletarGasto(body);
  if (action === 'salvarDoc') return doSalvarDoc(body);
  if (action === 'deletarDoc') return doDeletarDoc(body);
  if (action === 'uploadArquivo') return doUploadArquivo(body);
  return { success: false, error: 'Ação desconhecida: ' + action };
}

// Inicialização e Carga Total de Dados (Evita múltiplas requisições)
function doBoot() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Garantir existência das abas essenciais se não existirem
  var abasNecessarias = ['Gastos', 'Documentos', 'Config', 'Inquilinos'];
  abasNecessarias.forEach(function(nome) {
    if (!ss.getSheetByName(nome)) {
      ss.insertSheet(nome);
    }
  });

  var shGastos = ss.getSheetByName('Gastos');
  var shDocs = ss.getSheetByName('Documentos');
  
  var gastosRaw = shGastos.getDataRange().getValues();
  var docsRaw = shDocs.getDataRange().getValues();
  
  var gastos = [];
  if (gastosRaw.length > 1) {
    for (var i = 1; i < gastosRaw.length; i++) {
      gastos.push({
        id: gastosRaw[i],
        data: gastosRaw[i],
        categoria: gastosRaw[i],
        descricao: gastosRaw[i],
        valor: gastosRaw[i],
        tipo: gastosRaw[i] || 'despesa', // receita ou despesa
        fileUrl: gastosRaw[i] || ''      // Link do Comprovante/Recibo
      });
    }
  }
  
  var docs = [];
  if (docsRaw.length > 1) {
    for (var j = 1; j < docsRaw.length; j++) {
      docs.push({
        id: docsRaw[j],
        pasta: docsRaw[j],
        nome: docsRaw[j],
        url: docsRaw[j]
      });
    }
  }

  return { success: true, gastos: gastos, documentos: docs };
}

// Salvar / Atualizar Gastos e Receitas
function doSalvarGasto(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Gastos');
  var rows = sheet.getDataRange().getValues();
  
  var id = body.id || 'G_' + new Date().getTime();
  var encontrado = false;
  var linhaIdx = -1;

  for (var i = 1; i < rows.length; i++) {
    if (rows[i] == id) {
      encontrado = true;
      linhaIdx = i + 1;
      break;
    }
  }

  var novosDados = [
    id,
    body.data || new Date().toISOString().split('T'),
    body.categoria || '',
    body.descricao || '',
    Number(body.valor) || 0,
    body.tipo || 'despesa',
    body.fileUrl || ''
  ];

  if (encontrado) {
    sheet.getRange(linhaIdx, 1, 1, novosDados.length).setValues([novosDados]);
  } else {
    // Se a planilha estiver totalmente vazia (sem cabeçalho), insere o cabeçalho
    if (rows.length === 1 && rows === '') {
      sheet.getRange(1, 1, 1, 7).setValues([['ID', 'Data', 'Categoria', 'Descrição', 'Valor', 'Tipo', 'FileUrl']]);
    }
    sheet.appendRow(novosDados);
  }
  return { success: true, id: id };
}

// Deletar Gasto
function doDeletarGasto(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Gastos');
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] == body.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Registro não encontrado' };
}

// Salvar Metadados do Documento/Contrato na Planilha
function doSalvarDoc(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Documentos');
  var rows = sheet.getDataRange().getValues();
  
  if (rows.length === 1 && rows === '') {
    sheet.getRange(1, 1, 1, 4).setValues([['ID', 'Pasta', 'Nome', 'URL']]);
  }
  
  var id = 'D_' + new Date().getTime();
  sheet.appendRow([id, body.pasta, body.nome, body.url]);
  return { success: true, id: id };
}

// Deletar Documento
function doDeletarDoc(body) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Documentos');
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] == body.id) {
      // Opcional: Você pode extrair o ID do arquivo do Drive através da URL se quiser deletar do Drive também.
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Documento não encontrado' };
}

// Sistema Integrado de Upload de Arquivos de Mídia (Drive Storage)
function doUploadArquivo(body) {
  try {
    var nomePastaRaiz = "Fluxo_Files";
    var pastas = DriveApp.getFoldersByName(nomePastaRaiz);
    var pastaDestino;
    
    if (pastas.hasNext()) {
      pastaDestino = pastas.next();
    } else {
      pastaDestino = DriveApp.createFolder(nomePastaRaiz);
    }
    
    var bytes = Utilities.base64Decode(body.bytes);
    var blob = Utilities.newBlob(bytes, body.mimeType, body.nome);
    var arquivoSalvo = pastaDestino.createFile(blob);
    
    // Libera o link de visualização pública controlada (essencial para inquilinos e sócios verem)
    arquivoSalvo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: arquivoSalvo.getUrl(), nome: body.nome };
  } catch (err) {
    return { success: false, error: 'Erro no Storage do Drive: ' + err.toString() };
  }
}

// Geração dinâmica da página do inquilino (Visão externa limpa)
function paginaInquilino(e) {
  var params = e.parameter || {};
  var inqId = params.id || '';
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Inquilinos');
  var rows = sheet.getDataRange().getValues();
  var ct = null;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] == inqId) {
      ct = { id: rows[i], inqNome: rows[i], imovel: rows[i], contrato: rows[i] };
      break;
    }
  }
  
  if (!ct) {
    return ContentService.createTextOutput("Acesso não autorizado ou link inválido.")
                         .setMimeType(ContentService.MimeType.TEXT);
  }
  
  // Renderização do HTML compacto do inquilino
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Área do Inquilino</title>' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#0F1923;color:#fff;margin:0;padding:15px}' +
    '.card{background:#17222D;border-radius:12px;padding:20px;margin-bottom:15px;border:1px solid #233240}' +
    'h1{font-size:20px;color:#FF4655;margin:0 0 5px 0}' +
    'p{margin:5px 0;color:#8F9EAC;font-size:14px}' +
    '.btn{display:inline-block;background:#FF4655;color:#fff;text-decoration:none;padding:10px 15px;border-radius:6px;font-weight:bold;font-size:14px;margin-top:10px}' +
    '</style></head><body>' +
    '<div class="card">' +
      '<h1>⚡ Bem-vindo, ' + esc(ct.inqNome) + '</h1>' +
      '<p><strong>Imóvel:</strong> ' + esc(ct.imovel) + '</p>' +
    '</div>' +
    '<div class="card">' +
      '<h3>📄 Seu Contrato</h3>' +
      '<p>Acesse a cópia digitalizada e assinada do seu contrato de locação.</p>' +
      (ct.contrato ? '<a class="btn" href="' + ct.contrato + '" target="_blank">Visualizar Contrato</a>' : '<p><i>Nenhum contrato anexado.</i></p>') +
    '</div>' +
    '</body></html>';
    
  return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
}

function esc(str) {
  if(!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
