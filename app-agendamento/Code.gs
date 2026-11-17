/**
 * App de Agendamento — backend (Google Apps Script)
 *
 * Lê e grava direto na "Planilha de Agendamento", respeitando o layout dela:
 *
 *   A: Cliente (ex.: "Cliente - retorno")   B: Telefone   C: Valor do serviço
 *   D/E: Segunda (data/hora)  F/G: Terça  H/I: Quarta  J/K: Quinta
 *   L/M: Sexta  N/O: Sábado   P: Serviço (ex.: HIGIENIZAÇÃO)
 *
 * Cada dia é um bloco de linhas separado por uma linha em branco, com uma linha
 * "INTERVALO DE ALMOÇO" no meio. Cada semana começa com a linha de cabeçalho
 * ("Valor do serviço | Segunda-feira | Horas | ...").
 */

// ---- Configuração -----------------------------------------------------------
// Se o script for criado DENTRO da planilha (Extensões > Apps Script), deixe vazio.
// Se for um projeto separado, cole aqui o ID da planilha.
var SPREADSHEET_ID = '';
var SHEET_NAME = 'Página1';
var ALMOCO_TEXTO = '12h10 – 13:40h';
var ALMOCO_INICIO_MIN = 12 * 60 + 10; // 12h10
var NUM_COLS = 16; // A..P

var COL_NOME = 1, COL_TEL = 2, COL_VALOR = 3, COL_SERV = 16;
function colData_(w) { return 4 + 2 * w; }  // w: 0=Seg ... 5=Sáb
function colHora_(w) { return 5 + 2 * w; }

// ---- Web app ----------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Agendamentos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

// ---- Utilitários ------------------------------------------------------------
function sheet_() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  return sh;
}
function tz_() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSpreadsheetTimeZone();
}
function iso_(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function parseIso_(s, tz) { return Utilities.parseDate(s, tz, 'yyyy-MM-dd'); }

/** Segunda-feira (0) ... Sábado (5); Domingo = -1 */
function weekdayFromIso_(s) {
  var p = s.split('-');
  var g = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay();
  return g === 0 ? -1 : g - 1;
}
/** ISO da segunda-feira da semana */
function mondayOf_(s) {
  var p = s.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  var g = d.getUTCDay();
  d.setUTCDate(d.getUTCDate() - (g === 0 ? 6 : g - 1));
  return d.toISOString().slice(0, 10);
}

/** "8h30", "08h31", "14h ", "14:00" -> minutos desde 00:00 */
function horaParaMin_(txt) {
  var m = String(txt || '').trim().match(/^(\d{1,2})\s*[h:]\s*(\d{0,2})/i);
  if (!m) return null;
  return (+m[1]) * 60 + (m[2] ? +m[2] : 0);
}
/** "08:30" -> "8h30"; "14:00" -> "14h" (formato usado na planilha) */
function horaParaPlanilha_(hhmm) {
  var p = String(hhmm).split(':');
  var h = +p[0], m = +(p[1] || 0);
  return h + 'h' + (m ? ('0' + m).slice(-2) : '');
}
/** 130 -> "R$130,00" (formato usado na planilha) */
function valorParaPlanilha_(v) {
  var n = Number(String(v).replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  if (isNaN(n)) return '';
  return 'R$' + n.toFixed(2).replace('.', ',');
}
function valorParaNumero_(txt) {
  var s = String(txt || '').replace(/[^\d,.-]/g, '');
  if (!s) return 0;
  s = s.replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}
function separarNome_(txt) {
  var s = String(txt || '').replace(/\s+/g, ' ').trim();
  var i = s.indexOf(' - ');
  if (i < 0) return { nome: s.replace(/\s*-\s*$/, ''), obs: '' };
  return { nome: s.slice(0, i).trim(), obs: s.slice(i + 3).trim() };
}
function juntarNome_(nome, obs) {
  nome = String(nome || '').trim() || 'Cliente';
  obs = String(obs || '').trim();
  return obs ? nome + ' - ' + obs : nome;
}

/**
 * Lê a planilha inteira e classifica cada linha.
 * Retorna [{row, tipo:'appt'|'lunch'|'header'|'blank'|'other', w, data, ...}]
 */
function lerLinhas_(sh, tz) {
  var last = sh.getLastRow();
  if (last < 1) return [];
  var rng = sh.getRange(1, 1, last, NUM_COLS);
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();
  var out = [];
  for (var i = 0; i < last; i++) {
    var v = vals[i], d = disp[i];
    var a = String(d[0]).trim();
    var linha = { row: i + 1, tipo: 'other' };
    if (/^INTERVALO/i.test(a)) {
      linha.tipo = 'lunch';
    } else if (/segunda/i.test(String(d[3])) && /hora/i.test(String(d[4]))) {
      linha.tipo = 'header';
    } else {
      for (var w = 0; w < 6; w++) {
        var cell = v[colData_(w) - 1];
        if (cell instanceof Date) {
          var n = separarNome_(d[0]);
          var horaTxt = String(d[colHora_(w) - 1]).trim();
          linha.tipo = 'appt';
          linha.w = w;
          linha.data = iso_(cell, tz);
          linha.hora = horaTxt;
          linha.min = horaParaMin_(horaTxt);
          linha.nome = n.nome;
          linha.obs = n.obs;
          linha.telefone = String(d[1]).trim();
          linha.valorTxt = String(d[2]).trim();
          linha.valor = valorParaNumero_(d[2]);
          linha.servico = String(d[COL_SERV - 1]).trim();
          break;
        }
      }
      if (linha.tipo === 'other' && d.join('').trim() === '') linha.tipo = 'blank';
    }
    out.push(linha);
  }
  return out;
}

function publico_(l) {
  return {
    row: l.row, data: l.data, hora: l.hora, min: l.min, nome: l.nome, obs: l.obs,
    telefone: l.telefone, valor: l.valor, valorTxt: l.valorTxt, servico: l.servico,
    chave: chave_(l)
  };
}
/** Impressão digital da linha, para não editar/apagar a linha errada se a planilha mudou */
function chave_(l) {
  return [l.data, l.hora, juntarNome_(l.nome, l.obs), l.telefone].join('|');
}

// ---- API chamada pelo app ---------------------------------------------------

/** Lista agendamentos entre duas datas ISO (inclusive). */
function listarAgendamentos(deIso, ateIso) {
  var sh = sheet_(), tz = tz_();
  return lerLinhas_(sh, tz)
    .filter(function (l) { return l.tipo === 'appt' && l.data >= deIso && l.data <= ateIso; })
    .map(publico_)
    .sort(function (a, b) {
      return a.data < b.data ? -1 : a.data > b.data ? 1 : (a.min || 0) - (b.min || 0) || a.row - b.row;
    });
}

/** Lista nomes de serviços já usados (para sugestões no app). */
function listarServicos() {
  var sh = sheet_(), tz = tz_();
  var set = {};
  lerLinhas_(sh, tz).forEach(function (l) { if (l.tipo === 'appt' && l.servico) set[l.servico] = 1; });
  if (!Object.keys(set).length) set['HIGIENIZAÇÃO'] = 1;
  return Object.keys(set).sort();
}

/**
 * Cria um agendamento.
 * dados = {nome, obs, telefone, valor, data:'yyyy-mm-dd', hora:'HH:MM', servico}
 */
function criarAgendamento(dados) {
  validar_(dados);
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = inserir_(dados);
    SpreadsheetApp.flush();
    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

/** Edita um agendamento existente (row + chave vindos de listarAgendamentos). */
function editarAgendamento(row, chave, dados) {
  validar_(dados);
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), tz = tz_();
    var atual = conferir_(sh, tz, row, chave);
    var novaHora = horaParaPlanilha_(dados.hora);
    if (atual.data === dados.data && horaParaMin_(atual.hora) === horaParaMin_(novaHora)) {
      // Mesmo dia e horário: atualiza a linha no lugar
      sh.getRange(row, COL_NOME, 1, 3).setValues([[
        juntarNome_(dados.nome, dados.obs), String(dados.telefone || ''), valorParaPlanilha_(dados.valor)
      ]]);
      sh.getRange(row, COL_SERV).setValue(String(dados.servico || '').trim().toUpperCase());
      SpreadsheetApp.flush();
      return { ok: true, row: row };
    }
    // Mudou dia/hora: remove e insere de novo na posição certa
    sh.deleteRow(row);
    var nova = inserir_(dados);
    SpreadsheetApp.flush();
    return { ok: true, row: nova };
  } finally {
    lock.releaseLock();
  }
}

/** Apaga um agendamento (remove a linha da planilha). */
function excluirAgendamento(row, chave) {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(), tz = tz_();
    conferir_(sh, tz, row, chave);
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ---- Internos ---------------------------------------------------------------

function validar_(d) {
  if (!d || !d.data || !/^\d{4}-\d{2}-\d{2}$/.test(d.data)) throw new Error('Informe a data.');
  if (!d.hora || !/^\d{1,2}:\d{2}$/.test(d.hora)) throw new Error('Informe o horário.');
  if (weekdayFromIso_(d.data) < 0) throw new Error('A planilha não tem coluna de domingo. Escolha de segunda a sábado.');
}

function conferir_(sh, tz, row, chave) {
  var linhas = lerLinhas_(sh, tz);
  var l = linhas[row - 1];
  if (!l || l.tipo !== 'appt' || chave_(l) !== chave) {
    throw new Error('A planilha foi alterada. Atualize a lista e tente de novo.');
  }
  return l;
}

/** Insere o agendamento no bloco do dia certo, na ordem do horário. Retorna a linha. */
function inserir_(dados) {
  var sh = sheet_(), tz = tz_();
  var linhas = lerLinhas_(sh, tz);
  var iso = dados.data;
  var w = weekdayFromIso_(iso);
  var horaTxt = horaParaPlanilha_(dados.hora);
  var novoMin = horaParaMin_(horaTxt);

  var valores = [];
  for (var c = 0; c < NUM_COLS; c++) valores.push('');
  valores[COL_NOME - 1] = juntarNome_(dados.nome, dados.obs);
  valores[COL_TEL - 1] = String(dados.telefone || '').trim();
  valores[COL_VALOR - 1] = valorParaPlanilha_(dados.valor);
  valores[colData_(w) - 1] = parseIso_(iso, tz);
  valores[colHora_(w) - 1] = horaTxt;
  valores[COL_SERV - 1] = String(dados.servico || 'HIGIENIZAÇÃO').trim().toUpperCase();

  var appts = linhas.filter(function (l) { return l.tipo === 'appt'; });
  var doDia = appts.filter(function (l) { return l.data === iso; });

  // Linhas-modelo para copiar a formatação
  var modeloAppt = (doDia[0] || appts[0] || {}).row;
  var modeloHeader = (linhas.filter(function (l) { return l.tipo === 'header'; })[0] || {}).row;
  var modeloLunch = (linhas.filter(function (l) { return l.tipo === 'lunch'; })[0] || {}).row;

  if (doDia.length) {
    // ---- Já existe bloco para o dia: insere na ordem do horário ----
    var primeiro = doDia[0].row, ultimo = doDia[doDia.length - 1].row;
    var almoco = null;
    for (var r = primeiro; r <= ultimo + 1 && r <= linhas.length; r++) {
      if (linhas[r - 1].tipo === 'lunch') { almoco = r; break; }
    }
    if (!almoco && primeiro > 1 && linhas[primeiro - 2].tipo === 'lunch') almoco = primeiro - 1;

    var antes = doDia.filter(function (l) { return l.min != null && l.min <= novoMin; });
    var depoisDe; // insere depois desta linha
    if (antes.length) depoisDe = antes[antes.length - 1].row;
    else depoisDe = primeiro - 1;

    if (almoco) {
      if (novoMin >= ALMOCO_INICIO_MIN && depoisDe < almoco) depoisDe = almoco;
      if (novoMin < ALMOCO_INICIO_MIN && depoisDe >= almoco) depoisDe = almoco - 1;
    }
    return inserirLinhas_(sh, depoisDe, [{ tipo: 'appt', valores: valores }],
      { appt: modeloAppt, header: modeloHeader, lunch: modeloLunch });
  }

  // ---- Não existe bloco para o dia: cria um novo ----
  var linhaAlmoco = linhaAlmoco_(w);
  var blocoDia = novoMin < ALMOCO_INICIO_MIN
    ? [{ tipo: 'appt', valores: valores }, { tipo: 'lunch', valores: linhaAlmoco }]
    : [{ tipo: 'lunch', valores: linhaAlmoco }, { tipo: 'appt', valores: valores }];
  var semana = mondayOf_(iso);

  var anterior = null, proximo = null;
  appts.forEach(function (l) {
    if (l.data < iso && (!anterior || l.data > anterior.data || (l.data === anterior.data && l.row > anterior.row))) anterior = l;
    if (l.data > iso && (!proximo || l.data < proximo.data || (l.data === proximo.data && l.row < proximo.row))) proximo = l;
  });

  if (anterior && mondayOf_(anterior.data) === semana) {
    // Mesma semana de um dia anterior: coloca depois do bloco dele
    var fim = fimDoBloco_(linhas, anterior.row);
    return inserirLinhas_(sh, fim, [{ tipo: 'blank' }].concat(blocoDia), modelos_());
  }
  if (proximo && mondayOf_(proximo.data) === semana) {
    // Mesma semana de um dia posterior: coloca antes do bloco dele
    var ini = inicioDoBloco_(linhas, proximo.row);
    return inserirLinhas_(sh, ini - 1, blocoDia.concat([{ tipo: 'blank' }]), modelos_());
  }
  // Semana nova: cabeçalho + bloco
  var novaSemana = [{ tipo: 'blank' }, { tipo: 'blank' }, { tipo: 'header' }].concat(blocoDia);
  if (anterior) {
    return inserirLinhas_(sh, fimDoBloco_(linhas, anterior.row), novaSemana, modelos_());
  }
  if (proximo) {
    // Antes da primeira semana existente
    var ini2 = inicioDoBloco_(linhas, proximo.row);
    if (ini2 > 1 && linhas[ini2 - 2].tipo === 'header') ini2--;
    return inserirLinhas_(sh, ini2 - 1, [{ tipo: 'header' }].concat(blocoDia, [{ tipo: 'blank' }, { tipo: 'blank' }]), modelos_());
  }
  // Planilha vazia
  return inserirLinhas_(sh, Math.max(sh.getLastRow(), 2), [{ tipo: 'header' }].concat(blocoDia), modelos_());

  function modelos_() { return { appt: modeloAppt, header: modeloHeader, lunch: modeloLunch }; }
}

function linhaAlmoco_(w) {
  var v = [];
  for (var c = 0; c < NUM_COLS; c++) v.push('');
  v[0] = 'INTERVALO DE ALMOÇO';
  if (w > 0) v[colData_(w) - 3] = 'ALMOÇO';
  v[colData_(w) - 1] = ALMOCO_TEXTO;
  return v;
}

/** Última linha do bloco do dia (inclui almoço colado no fim). */
function fimDoBloco_(linhas, row) {
  var r = row;
  while (r < linhas.length && (linhas[r].tipo === 'appt' || linhas[r].tipo === 'lunch')) {
    var prox = linhas[r];
    if (prox.tipo === 'appt' && prox.data !== linhas[row - 1].data) break;
    r++;
  }
  return r;
}
/** Primeira linha do bloco do dia (inclui almoço colado no início). */
function inicioDoBloco_(linhas, row) {
  var r = row;
  while (r > 1 && (linhas[r - 2].tipo === 'appt' || linhas[r - 2].tipo === 'lunch')) {
    var ant = linhas[r - 2];
    if (ant.tipo === 'appt' && ant.data !== linhas[row - 1].data) break;
    r--;
  }
  return r;
}

/**
 * Insere linhas depois de `depoisDe` e preenche. Copia a formatação das linhas-modelo.
 * Retorna o número da linha do agendamento inserido.
 */
function inserirLinhas_(sh, depoisDe, itens, modelos) {
  var n = itens.length;
  if (depoisDe < 1) { sh.insertRowsBefore(1, n); depoisDe = 0; }
  else sh.insertRowsAfter(depoisDe, n);

  function desloca(r) { return r ? (r > depoisDe ? r + n : r) : null; }
  var mAppt = desloca(modelos.appt), mHead = desloca(modelos.header), mLunch = desloca(modelos.lunch);

  var linhaAppt = null;
  itens.forEach(function (it, i) {
    var r = depoisDe + 1 + i;
    var alvo = sh.getRange(r, 1, 1, NUM_COLS);
    if (it.tipo === 'blank') {
      alvo.clearContent();
      return;
    }
    if (it.tipo === 'header') {
      if (mHead) sh.getRange(mHead, 1, 1, NUM_COLS).copyTo(alvo);
      return;
    }
    var modelo = it.tipo === 'lunch' ? (mLunch || mAppt) : mAppt;
    if (modelo) sh.getRange(modelo, 1, 1, NUM_COLS).copyTo(alvo, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    // Telefone, valor e hora como texto (igual à planilha)
    sh.getRange(r, COL_TEL, 1, 2).setNumberFormat('@');
    alvo.setValues([it.valores]);
    if (it.tipo === 'appt') {
      for (var w = 0; w < 6; w++) {
        if (it.valores[colData_(w) - 1] instanceof Date) {
          sh.getRange(r, colData_(w)).setNumberFormat('dd/mm/yyyy');
          sh.getRange(r, colHora_(w)).setNumberFormat('@').setValue(it.valores[colHora_(w) - 1]);
        }
      }
      linhaAppt = r;
    }
  });
  return linhaAppt;
}
