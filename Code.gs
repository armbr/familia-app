// ╔══════════════════════════════════════════════════════════╗
// ║   FAMÍLIA APP — BACKEND                                  ║
// ║   1. Cole este código no Google Apps Script              ║
// ║   2. Salve (Ctrl+S)                                      ║
// ║   3. Implantar > Gerenciar implantações > ✏️ editar       ║
// ║   4. Versão: "Nova versão" > Implantar                   ║
// ╚══════════════════════════════════════════════════════════╝

// ★ SUBSTITUA PELO ID DA SUA PLANILHA ★
// (está na URL: docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit)
var SPREADSHEET_ID = '174UmeWX3kmj9qjl7Z3I8hACNpx1pjcWYsl1_wBFzgxM';

// ════════════════════════════════════════════════════════════
//  ENTRY POINT — tudo via GET para evitar problemas de CORS
// ════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';
    var body = {};
    if (params.payload) {
      body = JSON.parse(decodeURIComponent(params.payload));
    }
    var result = doGetInternal(action, body);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGetInternal(action, body) {
  switch (action) {
    case 'getTransactions':   return getTransactions();
    case 'addTransaction':    return addTransaction(body);
    case 'deleteTransaction': return deleteTransaction(body.id);
    case 'getTasks':          return getTasks();
    case 'addTask':           return addTask(body);
    case 'updateTask':        return updateTask(body);
    case 'deleteTask':        return deleteTask(body.id);
    case 'getDocs':           return getDocs();
    case 'addDoc':            return addDoc(body);
    case 'deleteDoc':         return deleteDoc(body.id);
    case 'getRecur':          return getRecur();
    case 'addRecur':          return addRecur(body);
    case 'deleteRecur':       return deleteRecur(body.id);
    case 'uploadComprovante': return uploadComprovante(body);
    case 'ping':              return { ok: true, msg: 'pong' };
    default:                  return { ok: false, error: 'Ação inválida: ' + action };
  }
}

function doPost(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';
    var body = {};

    // POST body pode vir como JSON no postData
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch(ex) {}
    }
    // Também aceita payload na URL
    if (!Object.keys(body).length && params.payload) {
      body = JSON.parse(decodeURIComponent(params.payload));
    }

    var result;
    switch (action) {
      case 'uploadComprovante': result = uploadComprovante(body); break;
      default: result = doGetInternal(action, body);
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════
//  HELPER
// ════════════════════════════════════════════════════════════
function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function fmtDate(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' +
      String(v.getMonth() + 1).padStart(2, '0') + '-' +
      String(v.getDate()).padStart(2, '0');
  }
  return v ? String(v) : '';
}

function sheetRows(name) {
  var sheet = ss().getSheetByName(name);
  if (!sheet) return { sheet: null, headers: [], rows: [] };
  var all = sheet.getDataRange().getValues();
  return {
    sheet:   sheet,
    headers: all[0] || [],
    rows:    all.slice(1)
  };
}

// ════════════════════════════════════════════════════════════
//  TRANSAÇÕES
// ════════════════════════════════════════════════════════════
function getTransactions() {
  var r = sheetRows('Transações');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows
    .map(function(row) {
      var obj = {};
      r.headers.forEach(function(h, i) {
        obj[h] = (h === 'date') ? fmtDate(row[i]) : row[i];
      });
      return obj;
    })
    .filter(function(obj) {
      if(obj.type !== 'inc' && obj.type !== 'exp') return false;
      // Excluir lançamentos automáticos da recorrência (recurId sem fromAgenda)
      var rid = String(obj.recurId || '').trim();
      var fa  = obj.fromAgenda === 'true' || obj.fromAgenda === true;
      if(rid !== '' && !fa) return false;
      return true;
    })
    .map(function(obj) {
      obj.fromAgenda = (obj.fromAgenda === 'true' || obj.fromAgenda === true);
      return obj;
    })
    .reverse();
  return { ok: true, data: data };
}

function addTransaction(body) {
  var sheet = ss().getSheetByName('Transações');
  if (!sheet) {
    sheet = ss().insertSheet('Transações');
    sheet.appendRow(['id','type','desc','value','cat','date','createdAt','fromAgenda','recurId']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFF');
  }
  var id = Date.now();
  sheet.appendRow([id, body.type, body.desc, parseFloat(body.value), body.cat, body.date, new Date().toISOString(), body.fromAgenda ? 'true' : '', body.recurId ? String(body.recurId) : '']);
  return { ok: true, id: id };
}

function deleteTransaction(id) {
  var r = sheetRows('Transações');
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  TAREFAS
// ════════════════════════════════════════════════════════════
function getTasks() {
  var r = sheetRows('Tarefas');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) {
      obj[h] = (h === 'deadline') ? fmtDate(row[i]) : row[i];
    });
    if (!obj.comprovUrl) obj.comprovUrl = '';
    if (!obj.cat)       obj.cat = '';
    if (!obj.recurId)   obj.recurId = '';
    return obj;
  }).reverse();
  return { ok: true, data: data };
}

function addTask(body) {
  var sheet = ss().getSheetByName('Tarefas');
  if (!sheet) {
    sheet = ss().insertSheet('Tarefas');
    sheet.appendRow(['id','desc','type','deadline','value','status','createdAt','cat','recurId']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#0F2438').setFontColor('#FFF');
  }
  var id = Date.now();
  sheet.appendRow([id, body.desc, body.type, body.deadline || '', body.value ? parseFloat(body.value) : '', body.status || 'pend', new Date().toISOString(), body.cat || '', body.recurId ? String(body.recurId) : '']);
  return { ok: true, id: id };
}

function updateTask(body) {
  var r = sheetRows('Tarefas');
  var statusIdx    = r.headers.indexOf('status');
  var comprovIdx   = r.headers.indexOf('comprovUrl');

  // Criar coluna comprovUrl se não existir
  if (comprovIdx === -1 && body.comprovUrl) {
    r.sheet.getRange(1, r.headers.length + 1).setValue('comprovUrl');
    comprovIdx = r.headers.length;
    r.headers.push('comprovUrl');
  }

  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(body.id)) {
      r.sheet.getRange(i + 2, statusIdx + 1).setValue(body.status);
      if (body.comprovUrl && comprovIdx > -1) {
        r.sheet.getRange(i + 2, comprovIdx + 1).setValue(body.comprovUrl);
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Tarefa não encontrada' };
}

function deleteTask(id) {
  var r = sheetRows('Tarefas');
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  DOCUMENTOS
// ════════════════════════════════════════════════════════════
function getDocs() {
  var r = sheetRows('Documentos');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) { obj[String(h).trim()] = row[i]; });
    return {
      id:       String(obj.id       || ''),
      name:     String(obj.name     || ''),
      url:      String(obj.url      || ''),
      type:     String(obj.type     || 'doc'),
      pasta:    String(obj.pasta    || obj.group || 'Outros'),
      subpasta: String(obj.subpasta || '')
    };
  }).filter(function(d) { return d.name && d.url; }).reverse();
  return { ok: true, data: data };
}

function addDoc(body) {
  var sheet = ss().getSheetByName('Documentos');
  if (!sheet) {
    sheet = ss().insertSheet('Documentos');
    sheet.appendRow(['id','name','url','type','pasta','subpasta','createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#1E2D40').setFontColor('#FFF');
  }
  var id = Date.now();
  sheet.appendRow([
    id,
    body.name     || '',
    body.url      || '',
    body.type     || 'doc',
    body.pasta    || 'Outros',
    body.subpasta || '',
    new Date().toISOString()
  ]);
  return { ok: true, id: id };
}

function deleteDoc(id) {
  var r = sheetRows('Documentos');
  if (!r.sheet) return { ok: false, error: 'Aba não encontrada' };
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  RECORRÊNCIA
// ════════════════════════════════════════════════════════════
function getRecur() {
  var r = sheetRows('Recorrentes');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) { obj[String(h).trim()] = row[i]; });
    return {
      id:    String(obj.id    || ''),
      type:  String(obj.type  || 'exp'),
      desc:  String(obj.desc  || ''),
      value: parseFloat(obj.value || 0),
      cat:   String(obj.cat   || ''),
      date:  fmtDate(obj.date) || String(obj.date || '')
    };
  }).filter(function(r) { return r.desc; });
  return { ok: true, data: data };
}

function addRecur(body) {
  var ss2 = ss();
  var sheet = ss2.getSheetByName('Recorrentes');
  if (!sheet) {
    sheet = ss2.insertSheet('Recorrentes');
    sheet.appendRow(['id','type','desc','value','cat','date','createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#162030').setFontColor('#FFF');
  }
  var id = body.id || Date.now();
  sheet.appendRow([id, body.type, body.desc, parseFloat(body.value), body.cat, body.date, new Date().toISOString()]);
  return { ok: true, id: id };
}

function deleteRecur(id) {
  var r = sheetRows('Recorrentes');
  if (!r.sheet) return { ok: false, error: 'Aba não encontrada' };
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  COMPROVANTES — upload para pasta no Google Drive
// ════════════════════════════════════════════════════════════

// ★ Opcional: defina o ID de uma pasta específica no Drive para comprovantes
// Deixe vazio ('') para salvar na raiz do Drive
var COMPROVANTES_FOLDER_ID = '';

function uploadComprovante(body) {
  try {
    var folder;
    if (COMPROVANTES_FOLDER_ID) {
      folder = DriveApp.getFolderById(COMPROVANTES_FOLDER_ID);
    } else {
      // Buscar ou criar pasta "Comprovantes Família" na raiz
      var folders = DriveApp.getFoldersByName('Comprovantes Família');
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Comprovantes Família');
    }

    // Decodificar base64
    var decoded = Utilities.base64Decode(body.data);
    var blob = Utilities.newBlob(decoded, body.mimeType, body.name);

    // Salvar no Drive
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    return { ok: true, url: url, id: file.getId(), name: file.getName() };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ═══════════════════════════════════════════════════
// FUNÇÃO DE LIMPEZA — executar UMA VEZ pelo editor
// do Apps Script para remover transações com data
// futura (geradas automaticamente pela versão antiga).
// Após executar, pode deletar esta função.
// ═══════════════════════════════════════════════════
function limparTransacoesAutomaticas() {
  var sheet = ss().getSheetByName('Transações');
  if (!sheet) { Logger.log('Aba Transações não encontrada'); return; }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf('date');
  var typeIdx = headers.indexOf('type');

  if (dateIdx === -1) { Logger.log('Coluna date não encontrada'); return; }

  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  var deletadas = 0;
  // Percorrer de baixo para cima
  for (var i = data.length - 1; i >= 1; i--) {
    var row  = data[i];
    var type = String(row[typeIdx] || '');
    var dateVal = row[dateIdx];

    // Converter para Date
    var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    d.setHours(0, 0, 0, 0);

    // Deletar se: data futura OU type inválido
    if (d > hoje || (type !== 'inc' && type !== 'exp')) {
      sheet.deleteRow(i + 1);
      deletadas++;
    }
  }
  Logger.log('Transações removidas: ' + deletadas);
}

// ════════════════════════════════════════════════════════════
//  E-MAIL DIÁRIO — RESUMO DE COMPROMISSOS
// ════════════════════════════════════════════════════════════

// ★ CONFIGURE SEU E-MAIL AQUI ★
var EMAIL_DESTINO = 'SEU_EMAIL@gmail.com';

// Normaliza qualquer valor de data para string 'YYYY-MM-DD'
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // Tem T (ISO): pegar só a parte da data
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  // Já é YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/'); return p[2]+'-'+p[1]+'-'+p[0];
  }
  return s;
}

function enviarResumoDiario() {
  var tz      = Session.getScriptTimeZone();
  var hoje    = new Date();
  var todayStr = Utilities.formatDate(hoje, tz, 'yyyy-MM-dd');
  var diaSem  = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'][hoje.getDay()];
  var diaFmt  = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');

  // ── Ler tarefas usando headers por nome ──────────────────
  var sheet = ss().getSheetByName('Tarefas');
  if (!sheet) { Logger.log('Aba Tarefas não encontrada'); return; }

  var all     = sheet.getDataRange().getValues();
  var headers = all[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var rows    = all.slice(1);

  Logger.log('Headers encontrados: ' + JSON.stringify(headers));
  Logger.log('Total de linhas: ' + rows.length);

  var iDesc     = headers.indexOf('desc');
  var iDeadline = headers.indexOf('deadline');
  var iStatus   = headers.indexOf('status');
  var iType     = headers.indexOf('type');
  var iValue    = headers.indexOf('value');
  var iCat      = headers.indexOf('cat');

  Logger.log('Índices — desc:'+iDesc+' deadline:'+iDeadline+' status:'+iStatus+' type:'+iType);

  var tarefas   = [];
  var atrasadas = [];

  rows.forEach(function(row, idx) {
    var status   = String(row[iStatus] || '').trim().toLowerCase();
    var deadline = toDateStr(row[iDeadline]);
    var desc     = String(row[iDesc]   || '').trim();

    Logger.log('Linha '+(idx+1)+': desc="'+desc+'" deadline="'+deadline+'" status="'+status+'"');

    if (status === 'done') return;
    if (!deadline || !desc) return;

    var item = {
      desc:     desc,
      type:     String(row[iType]  || '').trim(),
      value:    parseFloat(row[iValue] || 0),
      cat:      String(row[iCat]   || '').trim(),
      deadline: deadline
    };

    if (deadline === todayStr) {
      tarefas.push(item);
    } else if (deadline < todayStr) {
      atrasadas.push(item);
    }
  });

  Logger.log('Compromissos hoje: ' + tarefas.length);
  Logger.log('Atrasados: ' + atrasadas.length);

  if (!tarefas.length && !atrasadas.length) {
    Logger.log('Nenhum compromisso — e-mail não enviado.');
    return;
  }

  // ── Montar e-mail ────────────────────────────────────────
  function fmtVal(v) {
    if (!v || v <= 0) return '';
    return ' — <strong>R$ ' + v.toFixed(2).replace('.', ',') + '</strong>';
  }
  function tipoIcon(t) {
    if (t === 'exp')  return '🔴';
    if (t === 'inc')  return '🟢';
    return '📌';
  }
  function fmtBRLTotal(v) {
    return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  var totalHoje     = tarefas.reduce(function(s,t){   return s+(t.type!=='task'?t.value:0);}, 0);
  var totalAtrasado = atrasadas.reduce(function(s,t){ return s+(t.type!=='task'?t.value:0);}, 0);

  function linhaTabela(t, cor) {
    var dl = t.deadline ? t.deadline.split('-').reverse().slice(0,2).join('/') : '';
    return '<tr>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a;font-size:18px;width:30px">'+tipoIcon(t.type)+'</td>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a">'
      +   '<div style="font-weight:700;color:'+(cor||'#eee')+'">'+t.desc+'</div>'
      +   '<div style="font-size:12px;color:#888;margin-top:2px">'+(t.cat||'')+(dl&&cor?' · '+dl:'')+'</div>'
      + '</td>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a;text-align:right;font-weight:800;color:'+(t.value>0?(t.type==='inc'?'#2ecc9a':'#FF6B6B'):'#555')+'">'
      +   (t.value>0?fmtBRLTotal(t.value):'')
      + '</td>'
      + '</tr>';
  }

  var secaoHoje = '';
  if (tarefas.length) {
    secaoHoje = '<div style="padding:20px 24px">'
      + '<div style="font-size:10px;font-weight:800;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;display:flex;justify-content:space-between">'
      +   '<span>📅 Compromissos de hoje</span>'
      +   (totalHoje>0?'<span style="color:#FF6B6B;font-size:13px">'+fmtBRLTotal(totalHoje)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + tarefas.map(function(t){ return linhaTabela(t, '#eee'); }).join('')
      + '</table></div>';
  }

  var secaoAtrasado = '';
  if (atrasadas.length) {
    secaoAtrasado = '<div style="padding:20px 24px;border-top:1px solid #222">'
      + '<div style="font-size:10px;font-weight:800;color:#FF6B6B;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;display:flex;justify-content:space-between">'
      +   '<span>⚠️ Atrasados</span>'
      +   (totalAtrasado>0?'<span style="font-size:13px">'+fmtBRLTotal(totalAtrasado)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + atrasadas.map(function(t){ return linhaTabela(t, '#FF6B6B'); }).join('')
      + '</table></div>';
  }

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#111;color:#eee;border-radius:12px;overflow:hidden;border:1px solid #222">'
    + '<div style="background:#0f0f1a;padding:22px 24px;border-bottom:1px solid #222">'
    +   '<div style="font-size:20px;font-weight:800">⚡ Fluxo App</div>'
    +   '<div style="color:#888;margin-top:4px;font-size:13px">'+diaSem+', '+diaFmt+'</div>'
    + '</div>'
    + secaoHoje
    + secaoAtrasado
    + '<div style="padding:14px 24px;border-top:1px solid #222;color:#444;font-size:11px;text-align:center">Resumo automático diário · Fluxo App</div>'
    + '</div>';

  GmailApp.sendEmail(EMAIL_DESTINO,
    '⚡ Fluxo — ' + (tarefas.length ? tarefas.length+' compromisso'+(tarefas.length>1?'s':'') : '') +
    (atrasadas.length ? (tarefas.length?' · ':'') + atrasadas.length+' atrasado'+(atrasadas.length>1?'s':'') : '') +
    ' · ' + diaFmt,
    'Abra no Gmail para ver o e-mail formatado.',
    { htmlBody: html, name: 'Fluxo App' }
  );

  Logger.log('✅ E-mail enviado para ' + EMAIL_DESTINO);
}

// ── Trigger diário às 8h ─────────────────────────────────
function criarTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarResumoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarResumoDiario').timeBased().everyDays(1).atHour(8).create();
  Logger.log('✅ Trigger criado — e-mail todo dia às 8h.');
}

// ── Debug ────────────────────────────────────────────────
function debugEmail() {
  Logger.log('=== DEBUG EMAIL ===');
  enviarResumoDiario();
}

// ════════════════════════════════════════════════════════════
//  ALUGUÉIS — PLANILHA GOOGLE SHEETS
// ════════════════════════════════════════════════════════════

function criarFichaAluguel(body) {
  var ssObj = ss();
  var nomeAba = 'Aluguel - ' + String(body.nome || '').substring(0, 20);

  // Criar aba se não existir
  var sheet = ssObj.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ssObj.insertSheet(nomeAba);
    // Cabeçalho
    sheet.getRange('A1').setValue('FICHA DE ALUGUEL').setFontSize(14).setFontWeight('bold');
    sheet.getRange('A2').setValue('Inquilino: ' + body.nome);
    sheet.getRange('A3').setValue('Referência: ' + (body.ref || ''));
    sheet.getRange('A4').setValue('Telefone: ' + (body.tel || ''));
    sheet.getRange('A5').setValue('Vencimento: dia ' + (body.dia || ''));
    sheet.getRange('A6').setValue('Valor base: R$ ' + parseFloat(body.valorBase || 0).toFixed(2));
    sheet.getRange('A8:E8').setValues([['Mês', 'Valor', 'Data Pagamento', 'Status', 'Código']]);
    sheet.getRange('A8:E8').setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF');
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 150);
  }
  return { ok: true };
}

function registrarPagamentoAluguel(body) {
  var ssObj = ss();
  var nomeAba = 'Aluguel - ' + String(body.nome || '').substring(0, 20);
  var sheet = ssObj.getSheetByName(nomeAba);

  // Criar ficha se não existir
  if (!sheet) criarFichaAluguel(body);
  sheet = ssObj.getSheetByName(nomeAba);
  if (!sheet) return { ok: false, error: 'Aba não encontrada' };

  // Verificar se mês já registrado
  var dados = sheet.getDataRange().getValues();
  var mo = String(body.mesAno || '');
  var partes = mo.split('-');
  var meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var mesLabel = meses[parseInt(partes[1])-1] + '/' + partes[0];

  for (var i = 8; i < dados.length; i++) {
    if (String(dados[i][0]) === mesLabel) {
      // Atualizar linha existente
      sheet.getRange(i+1, 2, 1, 4).setValues([[
        parseFloat(body.valor || 0),
        body.dataPgto || '',
        'PAGO',
        body.codigo || ''
      ]]);
      sheet.getRange(i+1, 1, 1, 5).setBackground('#d4edda');
      return { ok: true };
    }
  }

  // Inserir nova linha
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 5).setValues([[
    mesLabel,
    parseFloat(body.valor || 0),
    body.dataPgto || '',
    'PAGO',
    body.codigo || ''
  ]]);
  sheet.getRange(lastRow, 1, 1, 5).setBackground('#d4edda');

  return { ok: true };
}
