/**
 * MKW Gestão — Backend "Produtos e Serviços"
 * Base: Google Sheets (primeira aba / gid=0)
 *
 * Colunas da base (linha 1 = cabeçalho):
 *   A Código  | B Situação | C Descrição | D Tipo | E Data de Cadastro |
 *   F Data de Última Alteração | G Unidade de Medida
 *
 * COMO PUBLICAR:
 *   1. Abra a planilha > Extensões > Apps Script.
 *   2. Cole este código no arquivo Code.gs e salve.
 *   3. Implantar > Nova implantação > tipo "App da Web".
 *        - Executar como: Eu
 *        - Quem pode acessar: Qualquer pessoa
 *   4. Copie a URL /exec e cole na constante API_URL dos dois HTML.
 *   5. A cada alteração no código: Implantar > Gerenciar implantações > editar >
 *      Versão = "Nova versão" (senão o código antigo continua no ar).
 */

var SHEET_ID = '1APwsqN8hL0apIk52jIWS3whmJwsoyXCsX3_ecVPThF8';
var TZ = 'America/Sao_Paulo';
var FMT = 'dd/MM/yyyy HH:mm';

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; // primeira aba (gid=0)
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Formata data da planilha em DD/MM/AAAA HH:MM, venha ela como Date ou como texto. */
function fmtData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, FMT);
  return String(v == null ? '' : v).trim();
}

/** Lê todas as linhas de dados (a partir da linha 2) como objetos. */
function lerRegistros_() {
  var sh = getSheet_();
  var ultima = sh.getLastRow();
  if (ultima < 2) return [];
  var valores = sh.getRange(2, 1, ultima - 1, 7).getValues();
  var lista = [];
  for (var i = 0; i < valores.length; i++) {
    var l = valores[i];
    if (!l[0] && !l[2]) continue; // linha vazia
    lista.push({
      codigo:    String(l[0] || '').trim(),
      situacao:  String(l[1] || '').trim(),
      descricao: String(l[2] || '').trim(),
      tipo:      String(l[3] || '').trim(),
      dataCad:   fmtData_(l[4]),
      dataAlt:   fmtData_(l[5]),
      unidade:   String(l[6] || '').trim(),
      _linha:    i + 2
    });
  }
  return lista;
}

/** Converte "000-000" em número inteiro. */
function codigoParaNumero_(codigo) {
  var n = parseInt(String(codigo).replace(/\D/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

/** Converte inteiro em "000-000" (6 dígitos, hífen no meio). */
function numeroParaCodigo_(n) {
  var s = String(n);
  while (s.length < 6) s = '0' + s;
  return s.slice(0, 3) + '-' + s.slice(3);
}

/** Próximo código = maior existente + 1. */
function proximoCodigo_(registros) {
  var max = 0;
  for (var i = 0; i < registros.length; i++) {
    var v = codigoParaNumero_(registros[i].codigo);
    if (v > max) max = v;
  }
  return numeroParaCodigo_(max + 1);
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    var registros = lerRegistros_();

    if (action === 'next') {
      return json_({ ok: true, codigo: proximoCodigo_(registros) });
    }

    if (action === 'get') {
      var id = String((e.parameter.id || '')).trim();
      var achado = null;
      for (var i = 0; i < registros.length; i++) {
        if (registros[i].codigo === id) { achado = registros[i]; break; }
      }
      if (!achado) return json_({ ok: false, error: 'Registro ' + id + ' não encontrado.' });
      delete achado._linha;
      return json_({ ok: true, data: achado });
    }

    // list (padrão) — remove o índice interno de linha
    var limpos = registros.map(function (r) {
      return {
        codigo: r.codigo, situacao: r.situacao, descricao: r.descricao,
        tipo: r.tipo, dataCad: r.dataCad, dataAlt: r.dataAlt, unidade: r.unidade
      };
    });
    return json_({ ok: true, data: limpos });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita códigos duplicados em gravações simultâneas
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var modo = body.mode || 'insert';
    var sh = getSheet_();
    var agora = Utilities.formatDate(new Date(), TZ, FMT);

    var situacao  = String(body.situacao  || 'Ativa').trim();
    var descricao = String(body.descricao || '').trim().toUpperCase();
    var tipo      = String(body.tipo      || 'Produto').trim();
    var unidade   = String(body.unidade   || 'Unidade').trim();

    if (!descricao) return json_({ ok: false, error: 'Descrição obrigatória.' });

    var registros = lerRegistros_();

    if (modo === 'update') {
      var id = String(body.codigo || '').trim();
      var alvo = null;
      for (var i = 0; i < registros.length; i++) {
        if (registros[i].codigo === id) { alvo = registros[i]; break; }
      }
      if (!alvo) return json_({ ok: false, error: 'Registro ' + id + ' não encontrado para alteração.' });

      // Mantém A (código) e E (data de cadastro); atualiza B, C, D, F e G.
      sh.getRange(alvo._linha, 2).setValue(situacao);   // B
      sh.getRange(alvo._linha, 3).setValue(descricao);  // C
      sh.getRange(alvo._linha, 4).setValue(tipo);       // D
      sh.getRange(alvo._linha, 6).setValue(agora);      // F (última alteração)
      sh.getRange(alvo._linha, 7).setValue(unidade);    // G (unidade de medida)
      return json_({ ok: true, codigo: id });
    }

    // insert — servidor sempre gera o código (ignora o enviado pelo form)
    var novo = proximoCodigo_(registros);
    sh.appendRow([novo, situacao, descricao, tipo, agora, agora, unidade]); // A..G
    return json_({ ok: true, codigo: novo });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}