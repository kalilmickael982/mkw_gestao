/**
 * MKW Gestão — Backend "Configuração da Empresa"
 * Base: Google Sheets (primeira aba / gid=0)
 *
 * Colunas da base (linha 1 = cabeçalho), total 18 (A..R):
 *   A Código (0000)        | B Razão Social          | C CPF ou CNPJ
 *   D Inscrição            | E Regime Tributário     | F E-mail
 *   G Número de Contato    | H Situação              | I Data de Cadastro (DD/MM/AAAA)
 *   J Nome de Usuário      | K Logradouro            | L Número
 *   M Complemento          | N Bairro                | O Município
 *   P UF                   | Q CEP                   | R Data de Abertura (DD/MM/AAAA)
 *
 * Gerados pelo servidor: A (código sequencial), I (data de cadastro),
 * J (nome de usuário = "0000 - RAZÃO SOCIAL", em maiúsculo).
 *
 * COMO PUBLICAR:
 *   1. Abra a planilha > Extensões > Apps Script.
 *   2. Cole este código no arquivo Code.gs e salve.
 *   3. Implantar > Nova implantação > tipo "App da Web".
 *        - Executar como: Eu
 *        - Quem pode acessar: Qualquer pessoa
 *   4. Copie a URL /exec e cole na constante API_URL dos dois HTML
 *      (config_empresa.html e incluir_empresa.html).
 *   5. A cada alteração no código: Implantar > Gerenciar implantações > editar >
 *      Versão = "Nova versão" (senão o código antigo continua no ar).
 */

var SHEET_ID = '1df1N2jRRFvK5lVIItNh4PCUi2qmUUMxRW03nnz0LIZ8';
var TZ = 'America/Sao_Paulo';
var FMT_DATA = 'dd/MM/yyyy';

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; // primeira aba (gid=0)
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Formata data da planilha em DD/MM/AAAA, venha ela como Date ou como texto. */
function fmtData_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, FMT_DATA);
  return String(v == null ? '' : v).trim();
}

/** Lê todas as linhas de dados (a partir da linha 2) como objetos. */
function lerRegistros_() {
  var sh = getSheet_();
  var ultima = sh.getLastRow();
  if (ultima < 2) return [];
  var valores = sh.getRange(2, 1, ultima - 1, 18).getValues();
  var lista = [];
  for (var i = 0; i < valores.length; i++) {
    var l = valores[i];
    if (!l[0] && !l[1]) continue; // linha vazia (sem código e sem razão social)
    lista.push({
      codigo:       String(l[0]  || '').trim(),
      razao:        String(l[1]  || '').trim(),
      cpfCnpj:      String(l[2]  || '').trim(),
      inscricao:    String(l[3]  || '').trim(),
      regime:       String(l[4]  || '').trim(),
      email:        String(l[5]  || '').trim(),
      contato:      String(l[6]  || '').trim(),
      situacao:     String(l[7]  || '').trim(),
      dataCad:      fmtData_(l[8]),
      usuario:      String(l[9]  || '').trim(),
      logradouro:   String(l[10] || '').trim(),
      numero:       String(l[11] || '').trim(),
      complemento:  String(l[12] || '').trim(),
      bairro:       String(l[13] || '').trim(),
      municipio:    String(l[14] || '').trim(),
      uf:           String(l[15] || '').trim(),
      cep:          String(l[16] || '').trim(),
      dataAbertura: fmtData_(l[17]),
      _linha:       i + 2
    });
  }
  return lista;
}

/** Converte "0000" em número inteiro. */
function codigoParaNumero_(codigo) {
  var n = parseInt(String(codigo).replace(/\D/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

/** Converte inteiro em "0000" (mínimo 4 dígitos, sem separador). */
function numeroParaCodigo_(n) {
  var s = String(n);
  while (s.length < 4) s = '0' + s;
  return s;
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

/** Nome de usuário = "0000 - RAZÃO SOCIAL", em maiúsculo. */
function montarUsuario_(codigo, razao) {
  return (String(codigo).trim() + ' - ' + String(razao).trim()).toUpperCase();
}

/** Monta a linha A..R a partir de um objeto de campos já normalizados. */
function montarLinha_(c) {
  return [
    c.codigo, c.razao, c.cpfCnpj, c.inscricao, c.regime, c.email,
    c.contato, c.situacao, c.dataCad, c.usuario, c.logradouro, c.numero,
    c.complemento, c.bairro, c.municipio, c.uf, c.cep, c.dataAbertura
  ];
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
      var o = {}; for (var k in r) { if (k !== '_linha') o[k] = r[k]; } return o;
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
    var hoje = Utilities.formatDate(new Date(), TZ, FMT_DATA);

    // Normalização dos campos editáveis
    var razao = String(body.razao || '').trim().toUpperCase();
    if (!razao) return json_({ ok: false, error: 'Razão Social obrigatória.' });

    var campos = {
      razao:        razao,
      cpfCnpj:      String(body.cpfCnpj || '').trim(),
      inscricao:    String(body.inscricao || '').trim().toUpperCase(),
      regime:       String(body.regime  || '').trim(),
      email:        String(body.email   || '').trim(),
      contato:      String(body.contato || '').trim(),
      situacao:     String(body.situacao || 'Ativa').trim(),
      logradouro:   String(body.logradouro  || '').trim().toUpperCase(),
      numero:       String(body.numero       || '').trim().toUpperCase(),
      complemento:  String(body.complemento  || '').trim().toUpperCase(),
      bairro:       String(body.bairro       || '').trim().toUpperCase(),
      municipio:    String(body.municipio    || '').trim().toUpperCase(),
      uf:           String(body.uf           || '').trim().toUpperCase(),
      cep:          String(body.cep          || '').trim(),
      dataAbertura: String(body.dataAbertura || '').trim()
    };

    var registros = lerRegistros_();

    if (modo === 'update') {
      var id = String(body.codigo || '').trim();
      var alvo = null;
      for (var i = 0; i < registros.length; i++) {
        if (registros[i].codigo === id) { alvo = registros[i]; break; }
      }
      if (!alvo) return json_({ ok: false, error: 'Registro ' + id + ' não encontrado para alteração.' });

      // Mantém A (código) e I (data de cadastro); regenera J (nome de usuário).
      campos.codigo  = id;
      campos.dataCad = alvo.dataCad;                       // preserva a data original
      campos.usuario = montarUsuario_(id, razao);          // acompanha a razão social

      sh.getRange(alvo._linha, 1, 1, 18).setValues([montarLinha_(campos)]);
      return json_({ ok: true, codigo: id });
    }

    // insert — servidor gera código (A), data de cadastro (I) e nome de usuário (J)
    campos.codigo  = proximoCodigo_(registros);
    campos.dataCad = hoje;
    campos.usuario = montarUsuario_(campos.codigo, razao);

    sh.appendRow(montarLinha_(campos)); // A..R
    return json_({ ok: true, codigo: campos.codigo });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
