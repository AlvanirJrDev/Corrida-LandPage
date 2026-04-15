/**
 * Google Apps Script — inscrições + Mercado Pago Checkout Pro
 *
 * Erro "UrlFetchApp.fetch sem permissão": no editor, Executar → autorizarAcessoExterno → aceitar tudo.
 * Use appsscript.json SEM lista oauthScopes (o Google pede os escopos na 1ª execução). Implantação: Executar como: Eu.
 * E-mail após pagamento: MailApp na confirmação do pagamento — na 1ª vez o Apps Script pede permissão de envio.
 *
 * Propriedades do script (Projeto ⚙ → Propriedades do projeto → Propriedades do script):
 *   MERCADO_PAGO_ACCESS_TOKEN = Access Token Mercado Pago (aba Produção: APP_USR-... quando config.js → useSandbox: false).
 *   WEB_APP_URL (opcional) = mesma URL do webhookUrl em config.js. Se não definir, usa WEB_APP_URL_FALLBACK no código abaixo.
 *   APROVACAO_SENHA = senha para aprovar pagamento presencial (endpoint /exec/{protocolo}/{senha}).
 *   APROVACAO_EMAILS (opcional) = e-mails autorizados a aprovar (separados por vírgula ou ponto e vírgula).
 *
 * Produção MP: (1) token de produção em MERCADO_PAGO_ACCESS_TOKEN, (2) config.js com useSandbox: false e urlRetorno HTTPS real,
 * (3) Implantar → Gerenciar implantações → Nova versão na Web App.
 *
 * Consulta pública: POST JSON { "tipo": "consulta_inscricao", "email": "…", "telefone": "…" } — busca por e-mail + telefone (só dígitos) na lista oficial e em Inscrições pendentes MP.
 *
 * Duplicata + Mercado Pago: se o POST vier com mercadoPago: true, linhas não pagas com o mesmo e-mail ou telefone
 * (status "Aguardando pagamento online" ou "Pendente (PIX via WhatsApp)") são removidas antes da checagem de duplicata,
 * para o inscrito poder gerar um novo checkout. Inscrição já paga (ex.: "Pago (Mercado Pago)") continua bloqueando.
 *
 * Backup JSON (lista oficial): defina BACKUP_JSON_KEY nas Propriedades do script. GET:
 *   URL_DA_WEB_APP/exec?backup=1&key=SUA_CHAVE
 * Copie o JSON e salve em data/inscritos-confirmados.json no projeto. Ou execute exportarBackupJsonParaDrive() no editor (cria arquivo datado no Drive).
 *
 * Backup automático de segurança: após cada inscrição pendente (presencial ou MP) e após aprovações,
 * o script atualiza um único arquivo JSON no Google Drive (lista oficial + pendentes MP). Na 1ª vez cria o arquivo na raiz do Drive
 * do dono do script e grava BACKUP_DRIVE_FILE_ID nas propriedades. Opcional: BACKUP_DRIVE_FOLDER_ID = ID da pasta onde criar/atualizar.
 * Para desligar: AUTO_BACKUP_DRIVE = 0 nas Propriedades do script.
 *
 * Compartilhamento da planilha (Google Drive): quem tiver permissão de EDITOR pode ver, alterar e apagar
 * todas as linhas. Não use "Qualquer pessoa com o link pode editar" para o público. Restrinja a
 * organizadores (e-mail) ou use "Somente visualização" para quem só precisa consultar.
 *
 * ID da planilha (URL .../spreadsheets/d/ESTE_ID/edit)
 */
var ID_PLANILHA = "1BLVaZLh3Dq64WvUoQ2_XhPXgDAB-uOkW0t-Gek9gaKc";
/** Senha para endpoint manual de confirmação de pagamento presencial. */
var SENHA_MUDANCA_STATUS_PAGAMENTO = "ejcecc@2026@corrida";

/** Alinhe com config.js → nomeEvento (texto do e-mail de confirmação). */
var NOME_EVENTO_EMAIL = "Corrida Mariana em prol do ECC e EJC de Sanharó";
var DATA_EVENTO_EMAIL = "31 de maio de 2026";
var HORARIO_EVENTO_EMAIL = "Largada a partir das 7h";
var LOCAL_EVENTO_EMAIL = "Sanharó, Pernambuco";
var WHATSAPP_ORGANIZACAO_EMAIL = "5587991200165";
var INSTAGRAM_ORGANIZACAO_URL = "https://instagram.com/";
/** Use URL pública da logo (opcional). Ex.: https://seusite.com/assets/logo.png */
var LOGO_EMAIL_URL = "";

/** Mesma URL que config.js → webhookUrl (termina em /exec). Alinhe os dois se mudar a implantação. */
var WEB_APP_URL_FALLBACK =
  "https://script.google.com/a/macros/redealia.com/s/AKfycbwClwfOb9G4AXWFW0tjQAMAkuX_VlmlBHJC6nFPuGczHZZBM0XLv4p36mYF0RvvHCu7/exec";

/** Aba principal: inscrições confirmadas (pagamento aprovado). */
/** Aba "Inscrições pendentes MP": fila de pendentes (Mercado Pago e presencial/PIX). */
var NOME_ABA_PENDENTES = "Inscrições pendentes MP";

/** Limites de lotes: alinhar com config.js */
var LIMITE_LOTE_PROMO = 50;
var ID_LOTE_PROMO = "promo";
var LIMITE_LOTE_REGULAR = 150;
var ID_LOTE_REGULAR = "regular";
/** Índice da coluna "Lote id" (0-based) */
var COL_IX_LOTE = 8;
var COL_IX_EMAIL = 3;
var COL_IX_TELEFONE = 4;
var COL_IX_PROTOCOLO = 1;
var COL_IX_NOME = 2;

var CABECALHOS = [
  "Data",
  "Protocolo",
  "Nome",
  "E-mail",
  "Telefone",
  "Cidade",
  "Camisa",
  "Percurso",
  "Lote id",
  "Lote",
  "Valor (R$)",
  "Forma pagamento",
  "Status pagamento",
];

/** Chaves do objeto em cada linha do backup JSON (mesma ordem que CABECALHOS). */
var CHAVES_JSON_INSCRICAO = [
  "data",
  "protocolo",
  "nome",
  "email",
  "telefone",
  "cidade",
  "camisa",
  "percurso",
  "loteId",
  "lote",
  "valorReais",
  "formaPagamento",
  "statusPagamento",
];

function soDigitos(s) {
  return String(s || "").replace(/\D/g, "");
}

/** Compara telefones BR salvos com ou sem DDI 55 (últimos 11 dígitos). */
function telefonesIguaisBR(a, b) {
  var da = soDigitos(a);
  var db = soDigitos(b);
  if (da === db) return true;
  if (da.length < 10 || db.length < 10) return false;
  return da.slice(-11) === db.slice(-11);
}

function jaTemCadastro(sheet, email, telDigits) {
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][1]) === "Protocolo") {
    start = 1;
  }
  var em = String(email).trim().toLowerCase();
  var td = soDigitos(telDigits);
  for (var i = start; i < values.length; i++) {
    var rowEmail = String(values[i][COL_IX_EMAIL]).trim().toLowerCase();
    if (em && rowEmail && rowEmail === em) return "email";
    var rowTel = soDigitos(values[i][COL_IX_TELEFONE]);
    if (td.length >= 10 && rowTel && rowTel === td) return "telefone";
  }
  return null;
}

function jaTemCadastroQualquerAba(ss, email, telDigits) {
  var m = jaTemCadastro(obterAbaInscricoes(ss), email, telDigits);
  if (m) return m;
  var pend = obterAbaPendentes(ss);
  if (pend.getLastRow() === 0) return null;
  garantirCabecalhosPlanilha(pend);
  return jaTemCadastro(pend, email, telDigits);
}

/**
 * Classifica status da coluna "Status pagamento" para decidir se ainda pode abrir novo checkout MP.
 * nao_pago = aguardando MP ou pendente presencial; pago = confirmado; desconhecido = tratar como já definitivo (bloqueia duplicata).
 */
function classificarStatusPagamento(status) {
  var s = String(status || "").trim().toLowerCase();
  if (!s) return "desconhecido";
  if (s.indexOf("aguardando") !== -1) return "nao_pago";
  if (s.indexOf("pendente") !== -1) return "nao_pago";
  if (s.indexOf("não pago") !== -1 || s.indexOf("nao pago") !== -1) return "nao_pago";
  if (s.indexOf("pago") !== -1) return "pago";
  return "desconhecido";
}

function linhaCombinaEmailOuTelefone(row, email, telDigits) {
  var em = String(email).trim().toLowerCase();
  var rowEmail = String(row[COL_IX_EMAIL] || "").trim().toLowerCase();
  var td = soDigitos(telDigits);
  var rowTel = soDigitos(row[COL_IX_TELEFONE]);
  if (em && rowEmail && rowEmail === em) return true;
  if (td.length >= 10 && rowTel && telefonesIguaisBR(td, rowTel)) return true;
  return false;
}

/**
 * Remove linhas não pagas (mesmo e-mail ou telefone) na lista oficial e em pendentes MP,
 * para o inscrito poder gerar um novo link de pagamento sem erro de duplicata.
 * Só deve ser chamado quando o formulário envia mercadoPago: true.
 */
function removerInscricoesNaoPagasParaNovoCheckoutMp(ss, email, telDigits) {
  function purgeSheet(sh) {
    if (!sh || sh.getLastRow() === 0) return;
    garantirCabecalhosPlanilha(sh);
    var values = sh.getDataRange().getValues();
    var start = 0;
    if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
      start = 1;
    }
    for (var i = values.length - 1; i >= start; i--) {
      var row = values[i];
      if (!linhaCombinaEmailOuTelefone(row, email, telDigits)) continue;
      if (classificarStatusPagamento(row[COL_IX_STATUS]) !== "nao_pago") continue;
      sh.deleteRow(i + 1);
    }
  }
  purgeSheet(obterAbaPendentes(ss));
  purgeSheet(obterAbaInscricoes(ss));
}

function contarInscricoesLote(sheet, loteId) {
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][1]) === "Protocolo") {
    start = 1;
  }
  var n = 0;
  for (var i = start; i < values.length; i++) {
    if (String(values[i][COL_IX_LOTE]).trim() === String(loteId)) {
      n++;
    }
  }
  return n;
}

function contarInscricoesLoteTotal(ss, loteId) {
  var n = contarInscricoesLote(obterAbaInscricoes(ss), loteId);
  var pend = obterAbaPendentes(ss);
  if (pend.getLastRow() > 0) {
    garantirCabecalhosPlanilha(pend);
    n += contarInscricoesLote(pend, loteId);
  }
  return n;
}

function obterAbaInscricoes(ss) {
  var sh = ss.getSheetByName("Lista de inscritos");
  if (sh) return sh;
  sh = ss.getSheetByName("Inscrições");
  if (sh) return sh;
  sh = ss.getSheetByName("Página1");
  if (sh) return sh;
  sh = ss.getSheetByName("Sheet1");
  if (sh) return sh;
  var sheets = ss.getSheets();
  return sheets.length ? sheets[0] : null;
}

function obterAbaPendentes(ss) {
  var sh = ss.getSheetByName(NOME_ABA_PENDENTES);
  if (!sh) {
    sh = ss.insertSheet(NOME_ABA_PENDENTES);
  }
  return sh;
}

function garantirCabecalhosPlanilha(sheet) {
  if (sheet.getLastRow() > 0) {
    var lc = sheet.getLastColumn();
    if (lc < CABECALHOS.length) {
      sheet.getRange(1, CABECALHOS.length).setValue(CABECALHOS[CABECALHOS.length - 1]);
    }
    return;
  }
  sheet.getRange(1, 1, 1, CABECALHOS.length).setValues([CABECALHOS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, CABECALHOS.length).setFontWeight("bold");
}

function montarLinhaInscricao(data, statusPagamento) {
  return [
    data.criadoEm || new Date().toISOString(),
    data.protocolo || "",
    data.nome || "",
    data.email || "",
    data.telefone || "",
    data.cidade || "",
    data.camisa || "",
    data.percurso || "",
    data.lote || "",
    data.loteNome || "",
    data.valorReais !== undefined && data.valorReais !== null ? data.valorReais : "",
    data.formaPagamento || "",
    statusPagamento || "Pendente",
  ];
}

function encontrarLinhaPorProtocolo(sheet, protocolo) {
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
    start = 1;
  }
  var p = String(protocolo).trim();
  for (var i = start; i < values.length; i++) {
    if (String(values[i][COL_IX_PROTOCOLO]).trim() === p) {
      return i + 1;
    }
  }
  return -1;
}

/** Índice coluna "Status pagamento" (0-based), alinhado a CABECALHOS */
var COL_IX_STATUS = 12;

/**
 * Busca inscrição na lista principal ou em pendentes MP — exige e-mail + telefone (mesmos da inscrição; telefone comparado só com dígitos).
 */
function buscarInscricaoPorEmailETelefone(ss, email, telefoneDigitos) {
  var em = String(email || "")
    .trim()
    .toLowerCase();
  var td = soDigitos(telefoneDigitos);
  if (!em || td.length < 10) return null;

  var alvo = [obterAbaInscricoes(ss), obterAbaPendentes(ss)];
  var nomes = ["lista_oficial", "pendente_mp"];
  for (var s = 0; s < alvo.length; s++) {
    var sheet = alvo[s];
    if (!sheet || sheet.getLastRow() === 0) continue;
    garantirCabecalhosPlanilha(sheet);
    var values = sheet.getDataRange().getValues();
    var start = 0;
    if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
      start = 1;
    }
    for (var i = start; i < values.length; i++) {
      var row = values[i];
      if (String(row[COL_IX_EMAIL]).trim().toLowerCase() !== em) continue;
      if (!telefonesIguaisBR(row[COL_IX_TELEFONE], td)) continue;
      return { row: row, origem: nomes[s] };
    }
  }
  return null;
}

function linhaParaRespostaConsulta(row, origem) {
  var valor = row[10];
  if (valor !== "" && valor !== null && valor !== undefined) {
    if (typeof valor === "number") {
      valor =
        "R$ " +
        valor.toFixed(2).replace(".", ",");
    } else {
      valor = String(valor);
      if (valor !== "" && valor.indexOf("R$") !== 0) {
        valor = "R$ " + valor;
      }
    }
  } else {
    valor = "";
  }
  var situacaoLista =
    origem === "lista_oficial"
      ? "Inscrição na lista oficial do evento"
      : "Inscrição pendente aguardando confirmação do pagamento";
  return {
    protocolo: String(row[COL_IX_PROTOCOLO] || "").trim(),
    nome: String(row[COL_IX_NOME] || "").trim(),
    cidade: String(row[5] || "").trim(),
    camisa: String(row[6] || "").trim(),
    percurso: String(row[7] || "").trim(),
    loteNome: String(row[9] || "").trim(),
    valorReais: valor,
    formaPagamento: String(row[11] || "").trim(),
    statusPagamento: String(row[COL_IX_STATUS] || "").trim(),
    situacaoLista: situacaoLista,
  };
}

function celulaParaJsonBackup(v) {
  if (v === null || typeof v === "undefined") return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    try {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
    } catch (e1) {
      return String(v);
    }
  }
  if (typeof v === "number") return v;
  return String(v);
}

function linhaParaObjetoBackup(row) {
  var o = {};
  for (var c = 0; c < CHAVES_JSON_INSCRICAO.length && c < row.length; c++) {
    var key = CHAVES_JSON_INSCRICAO[c];
    var val = row[c];
    if (key === "valorReais" && typeof val === "number") {
      o[key] = val;
    } else if (key === "valorReais" && val !== "" && val !== null && typeof val !== "undefined") {
      var n = parseFloat(String(val).replace(",", "."));
      o[key] = isNaN(n) ? celulaParaJsonBackup(val) : n;
    } else {
      o[key] = celulaParaJsonBackup(val);
    }
  }
  return o;
}

function listaBackupDaAba(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return [];
  garantirCabecalhosPlanilha(sheet);
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
    start = 1;
  }
  var lista = [];
  for (var i = start; i < values.length; i++) {
    lista.push(linhaParaObjetoBackup(values[i]));
  }
  return lista;
}

/**
 * Snapshot da aba Lista de inscritos (cabeçalho ignorado). Uso: backup local em JSON.
 */
function gerarPayloadBackupListaInscritos() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = obterAbaInscricoes(ss);
  if (!sheet) throw new Error("Aba Lista de inscritos não encontrada.");
  garantirCabecalhosPlanilha(sheet);
  var lista = listaBackupDaAba(sheet);
  return {
    meta: {
      descricao: "Backup da aba Lista de inscritos (inscrições na lista oficial do evento).",
      geradoEm: new Date().toISOString(),
      fonte: "Google Sheets",
      planilhaId: ID_PLANILHA,
      aba: sheet.getName(),
      totalInscricoes: lista.length,
    },
    inscricoes: lista,
  };
}

/**
 * Snapshot completo para backup automático: lista oficial + aba de pendentes.
 */
function gerarPayloadBackupCompleto() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = obterAbaInscricoes(ss);
  if (!sheet) throw new Error("Aba Lista de inscritos não encontrada.");
  garantirCabecalhosPlanilha(sheet);
  var listaOficial = listaBackupDaAba(sheet);
  var pend = obterAbaPendentes(ss);
  var listaPend = listaBackupDaAba(pend);
  return {
    meta: {
      descricao: "Backup automático: lista oficial + inscrições pendentes (Mercado Pago e presencial/PIX).",
      geradoEm: new Date().toISOString(),
      fonte: "Google Sheets",
      planilhaId: ID_PLANILHA,
      totalInscricoesOficial: listaOficial.length,
      totalPendentesMercadoPago: listaPend.length,
    },
    inscricoes: listaOficial,
    pendentesMercadoPago: listaPend,
  };
}

var NOME_ARQUIVO_BACKUP_AUTO = "backup-seguranca-inscricoes-corrida-mariana.json";

/**
 * Atualiza (ou cria) um JSON no Drive com lista oficial + pendentes. Não interrompe inscrição se falhar.
 */
function sincronizarBackupSegurancaNoDrive() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (String(props.getProperty("AUTO_BACKUP_DRIVE") || "") === "0") return;

    var payload = gerarPayloadBackupCompleto();
    var json = JSON.stringify(payload, null, 2);
    var fileId = props.getProperty("BACKUP_DRIVE_FILE_ID");
    var folderId = String(props.getProperty("BACKUP_DRIVE_FOLDER_ID") || "").trim();
    var file;

    if (fileId) {
      file = DriveApp.getFileById(fileId);
      file.setContent(json);
    } else if (folderId) {
      var folder = DriveApp.getFolderById(folderId);
      var itF = folder.getFilesByName(NOME_ARQUIVO_BACKUP_AUTO);
      if (itF.hasNext()) {
        file = itF.next();
        file.setContent(json);
      } else {
        file = folder.createFile(NOME_ARQUIVO_BACKUP_AUTO, json, MimeType.PLAIN_TEXT);
      }
      props.setProperty("BACKUP_DRIVE_FILE_ID", file.getId());
    } else {
      var root = DriveApp.getRootFolder();
      var itR = root.getFilesByName(NOME_ARQUIVO_BACKUP_AUTO);
      if (itR.hasNext()) {
        file = itR.next();
        file.setContent(json);
      } else {
        file = DriveApp.createFile(NOME_ARQUIVO_BACKUP_AUTO, json, MimeType.PLAIN_TEXT);
      }
      props.setProperty("BACKUP_DRIVE_FILE_ID", file.getId());
    }
  } catch (err) {
    Logger.log("sincronizarBackupSegurancaNoDrive: " + err);
  }
}

/**
 * Executar no editor (▶ Executar): cria um arquivo .json no Google Drive com o backup.
 * Abra o arquivo → copie o conteúdo para data/inscritos-confirmados.json no repositório, se quiser.
 */
function exportarBackupJsonParaDrive() {
  var payload = gerarPayloadBackupListaInscritos();
  var json = JSON.stringify(payload, null, 2);
  var nome =
    "inscritos-confirmados-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd-HHmm") + ".json";
  var file = DriveApp.createFile(nome, json, MimeType.PLAIN_TEXT);
  Logger.log("Criado: " + file.getName() + " — " + file.getUrl());
  return file.getUrl();
}

function executarConsultaInscricao(data) {
  var email = data.email;
  var tel = data.telefone != null ? data.telefone : data.telefoneDigitos;
  tel = tel != null ? String(tel) : "";
  if (!String(email || "").trim() || !String(tel || "").trim()) {
    return { ok: false, error: "Informe o e-mail e o telefone (WhatsApp) usados na inscrição." };
  }
  if (soDigitos(tel).length < 10) {
    return { ok: false, error: "Telefone inválido: use DDD + número (ex.: (87) 99999-9999)." };
  }
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var achado = buscarInscricaoPorEmailETelefone(ss, email, tel);
  if (!achado) {
    return {
      ok: true,
      encontrado: false,
      error:
        "Não encontramos inscrição com este e-mail e telefone. Confira se são os mesmos do formulário ou fale com a organização.",
    };
  }
  return {
    ok: true,
    encontrado: true,
    dados: linhaParaRespostaConsulta(achado.row, achado.origem),
  };
}

function extrairPaymentIdNotificacao(obj, e) {
  if (e && e.parameter) {
    var par = e.parameter;
    if (par.topic === "payment" && par.id) {
      return String(par.id);
    }
    if (par["data.id"]) {
      return String(par["data.id"]);
    }
  }
  if (!obj || typeof obj !== "object") return null;
  if (obj.tipo === "inscricao_corrida") return null;
  if (obj.data && obj.data.id) {
    return String(obj.data.id);
  }
  if (obj.id && !obj.nome) {
    return String(obj.id);
  }
  if (obj.resource && typeof obj.resource === "string") {
    var m = obj.resource.match(/\/payments\/(\d+)/);
    if (m) return m[1];
  }
  return null;
}

function escapeHtmlEmail(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * E-mail inicial da aplicação: enviado ao receber a inscrição.
 * Para online, informa que está aguardando aprovação do pagamento.
 */
function enviarEmailInscricaoRecebida(rowData) {
  if (!rowData || !rowData.length) {
    Logger.log("enviarEmailInscricaoRecebida: rowData indefinido/vazio");
    return;
  }
  var email = String(rowData[COL_IX_EMAIL] || "").trim();
  if (!email || email.indexOf("@") < 0) return;
  var nome = String(rowData[COL_IX_NOME] || "").trim();
  var proto = String(rowData[COL_IX_PROTOCOLO] || "").trim();
  var cidade = String(rowData[5] || "").trim();
  var camisa = String(rowData[6] || "").trim();
  var percurso = String(rowData[7] || "").trim();
  var loteNome = String(rowData[9] || "").trim();
  var valor = rowData[10];
  var formaPagamento = String(rowData[11] || "").trim();
  var statusPagamento = String(rowData[COL_IX_STATUS] || "").trim();
  var aguardandoMp = statusPagamento.toLowerCase().indexOf("aguardando pagamento online") !== -1;
  var waLink = "https://wa.me/" + String(WHATSAPP_ORGANIZACAO_EMAIL || "").replace(/\D/g, "");

  var valorFmt = "";
  if (valor !== "" && valor !== null && valor !== undefined) {
    if (typeof valor === "number") {
      valorFmt = "R$ " + valor.toFixed(2).replace(".", ",");
    } else {
      valorFmt = String(valor).trim();
      if (valorFmt && valorFmt.indexOf("R$") !== 0) valorFmt = "R$ " + valorFmt;
    }
  }

  var textoStatus = aguardandoMp
    ? "Sua inscrição foi recebida e está aguardando a aprovação do pagamento no Mercado Pago."
    : "Sua inscrição foi recebida e está pendente até a confirmação do pagamento pela organização.";

  var corpoTexto =
    "Olá" +
    (nome ? " " + nome : "") +
    ",\n\n" +
    textoStatus +
    "\n\n" +
    "Protocolo: " +
    (proto || "—") +
    "\n" +
    "Camisa: " +
    (camisa || "—") +
    "\n" +
    "Percurso: " +
    (percurso || "—") +
    "\n" +
    "Lote: " +
    (loteNome || "—") +
    "\n" +
    "Valor: " +
    (valorFmt || "—") +
    "\n" +
    "Forma de pagamento: " +
    (formaPagamento || "—") +
    "\n" +
    "Status: " +
    (statusPagamento || "—") +
    "\n\n" +
    "Evento: " +
    NOME_EVENTO_EMAIL +
    "\nData: " +
    DATA_EVENTO_EMAIL +
    "\nHorário: " +
    HORARIO_EVENTO_EMAIL +
    "\nLocal: " +
    LOCAL_EVENTO_EMAIL +
    "\n\nWhatsApp da organização: " +
    WHATSAPP_ORGANIZACAO_EMAIL +
    "\n\n— Organização";

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;background:#f5f8fa;padding:24px;">' +
    '<div style="max-width:680px;margin:0 auto;background:#ffffff;border:1px solid #d7e3ea;border-radius:12px;overflow:hidden;">' +
    '<div style="background:linear-gradient(90deg,#004C74,#00BEE2);padding:18px 24px;color:#fff;">' +
    '<h1 style="margin:0;font-size:22px;line-height:1.2;">Inscrição recebida</h1>' +
    '<p style="margin:8px 0 0;font-size:14px;opacity:.95;">' +
    NOME_EVENTO_EMAIL +
    "</p>" +
    "</div>" +
    '<div style="padding:24px;">' +
    '<p style="margin:0 0 16px;color:#16384a;font-size:15px;">Olá <strong>' +
    escapeHtmlEmail(nome || "atleta") +
    "</strong>. " +
    escapeHtmlEmail(textoStatus) +
    "</p>" +
    '<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 16px;">' +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Protocolo</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;font-weight:700;">' +
    escapeHtmlEmail(proto || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Cidade</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(cidade || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Camisa</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(camisa || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Percurso</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(percurso || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Lote</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(loteNome || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Valor</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;font-weight:700;">' +
    escapeHtmlEmail(valorFmt || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Forma de pagamento</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(formaPagamento || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;color:#456;">Status</td><td style="padding:9px 10px;color:#123;">' +
    escapeHtmlEmail(statusPagamento || "—") +
    "</td></tr>" +
    "</table>" +
    '<p style="margin:0 0 12px;font-size:13px;color:#355264;">Data: ' +
    escapeHtmlEmail(DATA_EVENTO_EMAIL) +
    " · Horário: " +
    escapeHtmlEmail(HORARIO_EVENTO_EMAIL) +
    " · Local: " +
    escapeHtmlEmail(LOCAL_EVENTO_EMAIL) +
    "</p>" +
    '<p style="margin:0;text-align:center;">' +
    '<a href="' +
    escapeHtmlEmail(waLink) +
    '" style="display:inline-block;background:#128c7e;color:#fff;text-decoration:none;font-weight:700;font-size:14px;padding:10px 16px;border-radius:8px;">Falar com a organização no WhatsApp</a>' +
    "</p>" +
    "</div>" +
    "</div>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: email,
      subject: "Inscrição recebida — " + NOME_EVENTO_EMAIL,
      body: corpoTexto,
      htmlBody: htmlBody,
    });
  } catch (err) {
    Logger.log("enviarEmailInscricaoRecebida: " + err);
  }
}

/**
 * Envia e-mail ao inscrito quando o pagamento MP é confirmado (webhook).
 * 1ª vez: o Apps Script pedirá permissão para enviar e-mail. Limite diário do Gmail se aplicam.
 */
function enviarEmailPagamentoConfirmado(rowData) {
  if (!rowData || !rowData.length) {
    Logger.log("enviarEmailPagamentoConfirmado: rowData indefinido/vazio");
    return;
  }
  while (rowData.length < CABECALHOS.length) {
    rowData.push("");
  }
  var email = String(rowData[COL_IX_EMAIL] || "").trim();
  var nome = String(rowData[COL_IX_NOME] || "").trim();
  var proto = String(rowData[COL_IX_PROTOCOLO] || "").trim();
  if (!email || email.indexOf("@") < 0) return;
  var cidade = String(rowData[5] || "").trim();
  var camisa = String(rowData[6] || "").trim();
  var percurso = String(rowData[7] || "").trim();
  var loteNome = String(rowData[9] || "").trim();
  var valor = rowData[10];
  var formaPagamento = String(rowData[11] || "").trim();
  var statusPagamento = String(rowData[COL_IX_STATUS] || "").trim();
  var waLink = "https://wa.me/" + String(WHATSAPP_ORGANIZACAO_EMAIL || "").replace(/\D/g, "");
  var logoHtml = "";
  if (String(LOGO_EMAIL_URL || "").trim()) {
    logoHtml =
      '<p style="margin:0 0 12px;text-align:center;">' +
      '<img src="' +
      escapeHtmlEmail(LOGO_EMAIL_URL) +
      '" alt="Logo do evento" style="max-width:180px;height:auto;border:0;"/>' +
      "</p>";
  }

  var valorFmt = "";
  if (valor !== "" && valor !== null && valor !== undefined) {
    if (typeof valor === "number") {
      valorFmt = "R$ " + valor.toFixed(2).replace(".", ",");
    } else {
      valorFmt = String(valor).trim();
      if (valorFmt && valorFmt.indexOf("R$") !== 0) valorFmt = "R$ " + valorFmt;
    }
  }

  var corpoTexto =
    "Olá" +
    (nome ? " " + nome : "") +
    ",\n\n" +
    "Seu pagamento foi confirmado e sua inscrição está na lista oficial do evento.\n\n" +
    "Resumo da inscrição:\n" +
    "- Protocolo: " +
    (proto || "—") +
    "\n" +
    "- Nome: " +
    (nome || "—") +
    "\n" +
    "- E-mail: " +
    (email || "—") +
    "\n" +
    "- Cidade: " +
    (cidade || "—") +
    "\n" +
    "- Camisa: " +
    (camisa || "—") +
    "\n" +
    "- Percurso: " +
    (percurso || "—") +
    "\n" +
    "- Lote: " +
    (loteNome || "—") +
    "\n" +
    "- Valor: " +
    (valorFmt || "—") +
    "\n" +
    "- Forma de pagamento: " +
    (formaPagamento || "Mercado Pago") +
    "\n" +
    "- Status: " +
    (statusPagamento || "Pago (Mercado Pago)") +
    "\n\n" +
    "Dados do evento:\n" +
    "- Evento: " +
    NOME_EVENTO_EMAIL +
    "\n" +
    "- Data: " +
    DATA_EVENTO_EMAIL +
    "\n" +
    "- Horário: " +
    HORARIO_EVENTO_EMAIL +
    "\n" +
    "- Local: " +
    LOCAL_EVENTO_EMAIL +
    "\n" +
    "- WhatsApp: " +
    WHATSAPP_ORGANIZACAO_EMAIL +
    "\n\n" +
    "Guarde este e-mail e o protocolo para qualquer consulta.\n\n" +
    "— Organização";

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;background:#f5f8fa;padding:24px;">' +
    '<div style="max-width:680px;margin:0 auto;background:#ffffff;border:1px solid #d7e3ea;border-radius:12px;overflow:hidden;">' +
    '<div style="background:linear-gradient(90deg,#004C74,#00BEE2);padding:18px 24px;color:#fff;">' +
    '<h1 style="margin:0;font-size:22px;line-height:1.2;">Inscricao confirmada</h1>' +
    '<p style="margin:8px 0 0;font-size:14px;opacity:.95;">' +
    NOME_EVENTO_EMAIL +
    "</p>" +
    "</div>" +
    '<div style="padding:24px;">' +
    logoHtml +
    '<p style="margin:0 0 16px;color:#16384a;font-size:15px;">Olá <strong>' +
    escapeHtmlEmail(nome || "atleta") +
    "</strong>, seu pagamento foi confirmado. Sua inscrição está na lista oficial.</p>" +
    '<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 18px;">' +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Protocolo</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;font-weight:700;">' +
    escapeHtmlEmail(proto || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Nome</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(nome || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">E-mail</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(email || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Cidade</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(cidade || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Camisa</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(camisa || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Percurso</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(percurso || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Lote</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(loteNome || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Valor</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;font-weight:700;">' +
    escapeHtmlEmail(valorFmt || "—") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#456;">Forma de pagamento</td><td style="padding:9px 10px;border-bottom:1px solid #e5eef3;color:#123;">' +
    escapeHtmlEmail(formaPagamento || "Mercado Pago") +
    "</td></tr>" +
    '<tr><td style="padding:9px 10px;color:#456;">Status</td><td style="padding:9px 10px;color:#0b5f3b;font-weight:700;">' +
    escapeHtmlEmail(statusPagamento || "Pago (Mercado Pago)") +
    "</td></tr>" +
    "</table>" +
    '<div style="background:#f7fcff;border:1px solid #dceef7;border-radius:10px;padding:14px 16px;margin-bottom:14px;">' +
    '<p style="margin:0 0 8px;color:#0d3a55;font-weight:700;">Informacoes do evento</p>' +
    '<p style="margin:0;color:#1a4f6a;font-size:14px;line-height:1.6;">' +
    "<strong>Data:</strong> " +
    DATA_EVENTO_EMAIL +
    "<br/>" +
    "<strong>Horario:</strong> " +
    HORARIO_EVENTO_EMAIL +
    "<br/>" +
    "<strong>Local:</strong> " +
    escapeHtmlEmail(LOCAL_EVENTO_EMAIL) +
    "</p>" +
    "</div>" +
    '<p style="margin:0 0 14px;color:#355264;font-size:13px;">Guarde este e-mail e seu protocolo para qualquer consulta.</p>' +
    '<p style="margin:0 0 8px;text-align:center;">' +
    '<a href="' +
    escapeHtmlEmail(waLink) +
    '" style="display:inline-block;background:#128c7e;color:#fff;text-decoration:none;font-weight:700;font-size:14px;padding:10px 16px;border-radius:8px;">Falar com a organização no WhatsApp</a>' +
    "</p>" +
    '<p style="margin:0;text-align:center;font-size:12px;color:#5a7889;">' +
    'Instagram: <a href="' +
    escapeHtmlEmail(INSTAGRAM_ORGANIZACAO_URL) +
    '" style="color:#004c74;text-decoration:none;">' +
    escapeHtmlEmail(INSTAGRAM_ORGANIZACAO_URL) +
    "</a>" +
    "</p>" +
    "</div>" +
    '<div style="padding:12px 24px;background:#f2f7fa;color:#5a7889;font-size:12px;">Organizacao — ' +
    NOME_EVENTO_EMAIL +
    "</div>" +
    "</div>" +
    "</div>";
  try {
    MailApp.sendEmail({
      to: email,
      subject: "Inscrição confirmada — " + NOME_EVENTO_EMAIL,
      body: corpoTexto,
      htmlBody: htmlBody,
    });
  } catch (err) {
    Logger.log("enviarEmailPagamentoConfirmado: " + err);
  }
}

/**
 * Teste manual: envia e-mail de confirmação usando a última linha da aba principal.
 * Execute esta função no editor Apps Script para validar template e permissões do MailApp.
 */
function testeEnvioEmailPagamentoConfirmado() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var main = obterAbaInscricoes(ss);
  if (!main) throw new Error("Aba principal não encontrada.");
  garantirCabecalhosPlanilha(main);

  var lastRow = main.getLastRow();
  if (lastRow < 2) {
    throw new Error("Sem inscrições na aba principal para testar o envio de e-mail.");
  }

  var rowData = main.getRange(lastRow, 1, 1, CABECALHOS.length).getValues()[0];
  enviarEmailPagamentoConfirmado(rowData);
}

function obterPagamentoMercadoPago(paymentId) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("MERCADO_PAGO_ACCESS_TOKEN");
  if (!token) return null;
  var res = UrlFetchApp.fetch("https://api.mercadopago.com/v1/payments/" + encodeURIComponent(paymentId), {
    method: "get",
    headers: {
      Authorization: "Bearer " + token,
    },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    Logger.log("GET payment " + paymentId + " HTTP " + res.getResponseCode() + " " + res.getContentText());
    return null;
  }
  try {
    return JSON.parse(res.getContentText());
  } catch (err) {
    return null;
  }
}

function processarNotificacaoPagamentoMercadoPago(paymentId) {
  var pay = obterPagamentoMercadoPago(paymentId);
  if (!pay) return;
  if (pay.status !== "approved") {
    Logger.log("Pagamento " + paymentId + " status=" + pay.status);
    return;
  }
  var ref = String(pay.external_reference || "").trim();
  if (!ref) {
    Logger.log("Pagamento sem external_reference");
    return;
  }

  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var pend = obterAbaPendentes(ss);
  garantirCabecalhosPlanilha(pend);
  var rowIndex = encontrarLinhaPorProtocolo(pend, ref);
  if (rowIndex < 0) {
    Logger.log("Protocolo pendente não encontrado: " + ref);
    return;
  }

  var main = obterAbaInscricoes(ss);
  if (!main) return;
  garantirCabecalhosPlanilha(main);

  var rowData = pend.getRange(rowIndex, 1, 1, CABECALHOS.length).getValues()[0];
  while (rowData.length < CABECALHOS.length) {
    rowData.push("");
  }
  rowData[CABECALHOS.length - 1] = "Pago (Mercado Pago)";

  if (jaTemCadastro(main, rowData[COL_IX_EMAIL], rowData[COL_IX_TELEFONE])) {
    pend.deleteRow(rowIndex);
    sincronizarBackupSegurancaNoDrive();
    return;
  }

  main.appendRow(rowData);
  pend.deleteRow(rowIndex);
  enviarEmailPagamentoConfirmado(rowData);
  sincronizarBackupSegurancaNoDrive();
}

/**
 * Cria preferência Checkout Pro.
 * Retorno: { url: string|null, erroCodigo: string|null }
 */
function criarPreferenciaMercadoPago(data) {
  var out = { url: null, erroCodigo: null };
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("MERCADO_PAGO_ACCESS_TOKEN");
  if (!token) {
    out.erroCodigo = "no_token";
    return out;
  }
  if (!data.mercadoPago) {
    return out;
  }

  var preco = Number(data.valorReais);
  if (!preco || preco <= 0) {
    out.erroCodigo = "preco_invalido";
    return out;
  }

  var base = String(data.urlRetorno || "").trim().replace(/\/$/, "");
  if (!base) {
    base = "https://www.mercadopago.com.br";
  }

  var webApp = String(props.getProperty("WEB_APP_URL") || WEB_APP_URL_FALLBACK || "").trim().replace(/\?.*$/, "");

  var pref = {
    items: [
      {
        title: "Inscrição — " + (data.evento || "Corrida"),
        description: "Protocolo " + (data.protocolo || ""),
        quantity: 1,
        unit_price: preco,
        currency_id: "BRL",
      },
    ],
    payer: {
      email: String(data.email || "").trim(),
    },
    external_reference: String(data.protocolo || ""),
    /** Não exclui tipos/meios — no BR o checkout pode oferecer PIX, cartão, boleto etc. conforme a conta MP. */
    payment_methods: {
      excluded_payment_types: [],
      excluded_payment_methods: [],
    },
    back_urls: {
      success: base + (base.indexOf("?") >= 0 ? "&" : "?") + "mp=ok",
      failure: base + (base.indexOf("?") >= 0 ? "&" : "?") + "mp=erro",
      pending: base + (base.indexOf("?") >= 0 ? "&" : "?") + "mp=pendente",
    },
    auto_return: "approved",
  };

  if (webApp) {
    pref.notification_url = webApp;
  } else {
    Logger.log("Defina WEB_APP_URL nas propriedades do script (mesma URL /exec do webhook) para o MP confirmar pagamento e mover a inscrição para a lista principal.");
  }

  var res = UrlFetchApp.fetch("https://api.mercadopago.com/checkout/preferences", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
    },
    payload: JSON.stringify(pref),
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200 && code !== 201) {
    Logger.log("Mercado Pago preferences HTTP " + code + ": " + body);
    out.erroCodigo = "api_" + code;
    return out;
  }
  var j;
  try {
    j = JSON.parse(body);
  } catch (err) {
    Logger.log("Mercado Pago JSON parse: " + err);
    out.erroCodigo = "parse";
    return out;
  }
  var link = null;
  if (data.useSandbox && j.sandbox_init_point) {
    link = j.sandbox_init_point;
  } else {
    link = j.init_point || j.sandbox_init_point || null;
  }
  if (!link) {
    Logger.log("Mercado Pago sem init_point: " + body);
    out.erroCodigo = "sem_init_point";
    return out;
  }
  out.url = link;
  return out;
}

/** Texto quando UrlFetchApp não foi autorizado (comum na 1ª vez ou conta Workspace). */
function mensagemErroUrlFetchPermissao() {
  return (
    "Permissão negada para acessar a internet (UrlFetch). No Apps Script, faça login com a MESMA conta que implantou a Web App. " +
      "Menu Executar → escolha a função autorizarAcessoExterno → Executar → aceite TODAS as permissões (incl. serviços externos). " +
      "Depois: Implantar → Nova versão. Se usar Google Workspace, o administrador pode precisar liberar Apps Script. " +
      "Alternativa: myaccount.google.com/permissions → remova o acesso do Apps Script → execute autorizarAcessoExterno de novo."
  );
}

function mensagemErroCheckoutMercadoPago(codigo) {
  if (codigo === "no_token") {
    return (
      "Falta o Access Token no Apps Script. Engrenagem → Configurações do projeto → Propriedades do SCRIPT (não use Propriedades do usuário). Adicione: nome MERCADO_PAGO_ACCESS_TOKEN (exato) e valor = Access Token em mercadopago.com.br/developers. Salve o projeto. Nova implantação só se alterou o código do script. O token não vai no config.js."
    );
  }
  if (codigo === "preco_invalido") {
    return "Valor da inscrição inválido para o Mercado Pago.";
  }
  if (codigo === "sem_init_point" || codigo === "parse") {
    return "Resposta inválida do Mercado Pago. Verifique se o token é da mesma conta (produção ou teste) e se useSandbox no config bate com o tipo de token.";
  }
  if (codigo && String(codigo).indexOf("api_") === 0) {
    return "Mercado Pago recusou a criação do pagamento (HTTP). Confira o Access Token e as credenciais em developers.mercadopago.com.br — veja o log do Apps Script (Execuções).";
  }
  return "Não foi possível gerar o link de pagamento no Mercado Pago.";
}

function respostaJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function respostaTexto(msg) {
  return ContentService.createTextOutput(String(msg || "")).setMimeType(ContentService.MimeType.TEXT);
}

function montarMensagemLeigaAtualizacaoPagamento(out) {
  if (!out || out.ok !== true) {
    var erro = out && out.error ? String(out.error) : "Não foi possível atualizar agora.";
    return (
      "Nao foi possivel confirmar o pagamento.\n\n" +
      "Motivo: " +
      erro +
      "\n\n" +
      "Confira o protocolo e a senha e tente novamente. Se continuar, fale com a organizacao."
    );
  }
  if (out.atualizado === false) {
    return (
      "Pagamento ja estava confirmado.\n\n" +
      "Protocolo: " +
      out.protocolo +
      "\n" +
      "Status atual: " +
      (out.statusNovo || "Pago")
    );
  }
  return (
    "Pagamento confirmado com sucesso!\n\n" +
    "Protocolo: " +
    out.protocolo +
    "\n" +
    "Status anterior: " +
    (out.statusAnterior || "—") +
    "\n" +
    "Novo status: " +
    (out.statusNovo || "Pago")
  );
}

function obterSenhaAprovacaoConfigurada() {
  var props = PropertiesService.getScriptProperties();
  var senhaProp = String(props.getProperty("APROVACAO_SENHA") || "").trim();
  if (senhaProp) return senhaProp;
  return String(SENHA_MUDANCA_STATUS_PAGAMENTO || "").trim();
}

function validarAprovadorAutorizado() {
  var props = PropertiesService.getScriptProperties();
  var bruto = String(props.getProperty("APROVACAO_EMAILS") || "").trim();
  /** Se não configurar APROVACAO_EMAILS, aprovação funciona só com senha. */
  if (!bruto) return { ok: true };

  var emailAtual = String(Session.getActiveUser().getEmail() || "")
    .trim()
    .toLowerCase();
  if (!emailAtual) {
    return { ok: false, error: "Nao foi possivel identificar o e-mail da conta que abriu o link." };
  }

  var lista = bruto
    .split(/[;,]/)
    .map(function (v) {
      return String(v || "")
        .trim()
        .toLowerCase();
    })
    .filter(function (v) {
      return !!v;
    });

  if (lista.indexOf(emailAtual) === -1) {
    return { ok: false, error: "Conta sem permissao para aprovar pagamento.", email: emailAtual };
  }
  return { ok: true, email: emailAtual };
}

/**
 * Corrige dados legados: move inscrições presenciais pendentes da lista oficial para a aba de pendentes.
 * Útil quando a Web App antiga ainda gravou presencial na aba principal.
 */
function migrarPresenciaisPendentesParaAbaPendentes() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var main = obterAbaInscricoes(ss);
  var pend = obterAbaPendentes(ss);
  if (!main || !pend) return { ok: false, error: "Abas não encontradas." };
  garantirCabecalhosPlanilha(main);
  garantirCabecalhosPlanilha(pend);

  var moved = 0;
  var values = main.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
    start = 1;
  }

  for (var i = values.length - 1; i >= start; i--) {
    var row = values[i];
    var forma = String(row[11] || "").toLowerCase();
    var status = String(row[COL_IX_STATUS] || "").toLowerCase();
    var protocolo = String(row[COL_IX_PROTOCOLO] || "").trim();
    var ehPresencial =
      forma.indexOf("presencial") !== -1 ||
      forma.indexOf("secretaria") !== -1 ||
      forma.indexOf("pix") !== -1 ||
      forma.indexOf("whatsapp") !== -1;
    var naoPago = classificarStatusPagamento(status) === "nao_pago";
    if (!ehPresencial || !naoPago || !protocolo) continue;

    if (encontrarLinhaPorProtocolo(pend, protocolo) < 0) {
      pend.appendRow(row);
    }
    main.deleteRow(i + 1);
    moved++;
  }

  if (moved > 0) sincronizarBackupSegurancaNoDrive();
  return { ok: true, movidas: moved };
}

/**
 * Endpoint manual (GET /exec/{protocolo}/{senha}) para virar status de pagamento presencial.
 */
function confirmarPagamentoPresencialPorProtocolo(protocolo, senha) {
  var p = String(protocolo || "").trim();
  var s = String(senha || "").trim();
  if (!p) {
    return { ok: false, error: "Informe o protocolo na URL." };
  }
  var senhaEsperada = obterSenhaAprovacaoConfigurada();
  if (!senhaEsperada) {
    return { ok: false, error: "Senha de aprovacao nao configurada (APROVACAO_SENHA)." };
  }
  if (!s || s !== senhaEsperada) {
    return { ok: false, error: "Senha inválida." };
  }
  var autorizacao = validarAprovadorAutorizado();
  if (!autorizacao.ok) {
    return { ok: false, error: autorizacao.error };
  }

  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var main = obterAbaInscricoes(ss);
  var pend = obterAbaPendentes(ss);
  if (!main) {
    return { ok: false, error: "Aba principal de inscrições não encontrada." };
  }
  garantirCabecalhosPlanilha(main);
  garantirCabecalhosPlanilha(pend);

  var rowMain = encontrarLinhaPorProtocolo(main, p);
  if (rowMain > 0) {
    var statusMain = String(main.getRange(rowMain, COL_IX_STATUS + 1).getValue() || "").trim();
    if (classificarStatusPagamento(statusMain) === "pago") {
      return { ok: true, protocolo: p, statusAnterior: statusMain, statusNovo: statusMain, atualizado: false };
    }
    var novoStatusMain = "Pago (presencial confirmado)";
    main.getRange(rowMain, COL_IX_STATUS + 1).setValue(novoStatusMain);
    var rowMainData = main.getRange(rowMain, 1, 1, CABECALHOS.length).getValues()[0];
    enviarEmailPagamentoConfirmado(rowMainData);
    sincronizarBackupSegurancaNoDrive();
    return { ok: true, protocolo: p, statusAnterior: statusMain || "—", statusNovo: novoStatusMain, atualizado: true };
  }

  var rowPend = encontrarLinhaPorProtocolo(pend, p);
  if (rowPend < 0) {
    return { ok: false, error: "Protocolo não encontrado.", protocolo: p };
  }

  var rowData = pend.getRange(rowPend, 1, 1, CABECALHOS.length).getValues()[0];
  while (rowData.length < CABECALHOS.length) {
    rowData.push("");
  }
  var statusPend = String(rowData[COL_IX_STATUS] || "").trim();
  var novoStatus = "Pago (presencial confirmado)";
  rowData[COL_IX_STATUS] = novoStatus;

  if (!jaTemCadastro(main, rowData[COL_IX_EMAIL], rowData[COL_IX_TELEFONE])) {
    main.appendRow(rowData);
    enviarEmailPagamentoConfirmado(rowData);
  }
  pend.deleteRow(rowPend);
  sincronizarBackupSegurancaNoDrive();
  return { ok: true, protocolo: p, statusAnterior: statusPend || "—", statusNovo: novoStatus, atualizado: true };
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var raw = e.postData && e.postData.contents ? e.postData.contents : "";
    var data = {};
    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch (parseErr) {
        data = {};
      }
    }

    if (data.tipo === "consulta_inscricao") {
      var outConsulta = executarConsultaInscricao(data);
      return ContentService.createTextOutput(JSON.stringify(outConsulta)).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.tipo !== "inscricao_corrida") {
      var payId = extrairPaymentIdNotificacao(data, e);
      if (payId) {
        processarNotificacaoPagamentoMercadoPago(payId);
        return ContentService.createTextOutput(JSON.stringify({ ok: true, webhook: "mp_payment" })).setMimeType(
          ContentService.MimeType.JSON
        );
      }
      if (Object.keys(data).length > 0) {
        return ContentService.createTextOutput(JSON.stringify({ ok: true, ignored: true })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Corpo vazio ou notificação não reconhecida." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.openById(ID_PLANILHA);
    var sheet = obterAbaInscricoes(ss);
    if (!sheet) {
      throw new Error("Nenhuma aba encontrada na planilha.");
    }
    garantirCabecalhosPlanilha(sheet);

    data.mercadoPago = data.mercadoPago === true || String(data.mercadoPago || "").toLowerCase() === "true";

    if (!String(data.nome || "").trim() || !String(data.email || "").trim() || !String(data.telefone || "").trim()) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, error: "Preencha todos os campos obrigatórios." })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var telDup = data.telefoneDigitos || data.telefone;
    if (data.mercadoPago) {
      removerInscricoesNaoPagasParaNovoCheckoutMp(ss, data.email, telDup);
    }

    var dup = jaTemCadastroQualquerAba(ss, data.email, telDup);
    if (dup === "email") {
      return ContentService.createTextOutput(
        JSON.stringify({
          ok: false,
          error: "Já existe inscrição com este e-mail. Use outro e-mail ou fale com a organização.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    if (dup === "telefone") {
      return ContentService.createTextOutput(
        JSON.stringify({
          ok: false,
          error: "Já existe inscrição com este telefone. Se for você, fale com a organização.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (String(data.lote).trim() === ID_LOTE_PROMO) {
      var ja = contarInscricoesLoteTotal(ss, ID_LOTE_PROMO);
      if (ja >= LIMITE_LOTE_PROMO) {
        return ContentService.createTextOutput(
          JSON.stringify({
            ok: false,
            error: "Lote promocional esgotado (limite de " + LIMITE_LOTE_PROMO + " inscrições).",
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    } else if (String(data.lote).trim() === ID_LOTE_REGULAR) {
      var jaReg = contarInscricoesLoteTotal(ss, ID_LOTE_REGULAR);
      if (jaReg >= LIMITE_LOTE_REGULAR) {
        return ContentService.createTextOutput(
          JSON.stringify({
            ok: false,
            error: "Lote regular esgotado (limite de " + LIMITE_LOTE_REGULAR + " inscrições).",
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    var out = { ok: true };

    if (data.mercadoPago) {
      var pend = obterAbaPendentes(ss);
      garantirCabecalhosPlanilha(pend);
      var rowPend = montarLinhaInscricao(data, "Aguardando pagamento online");
      pend.appendRow(rowPend);

      var mpRes;
      try {
        mpRes = criarPreferenciaMercadoPago(data);
      } catch (mpErr) {
        if (pend.getLastRow() > 0) {
          pend.deleteRow(pend.getLastRow());
        }
        var em = String(mpErr);
        if (em.indexOf("UrlFetchApp") !== -1 || em.indexOf("external_request") !== -1) {
          return ContentService.createTextOutput(
            JSON.stringify({ ok: false, error: mensagemErroUrlFetchPermissao(), codigo: "urlfetch_auth" })
          ).setMimeType(ContentService.MimeType.JSON);
        }
        throw mpErr;
      }

      if (mpRes.url) {
        out.checkoutUrl = mpRes.url;
        out.aguardandoPagamento = true;
        sincronizarBackupSegurancaNoDrive();
        enviarEmailInscricaoRecebida(rowPend);
      } else {
        if (pend.getLastRow() > 0) {
          pend.deleteRow(pend.getLastRow());
        }
        out.checkoutFalhou = true;
        out.erroCheckout = mensagemErroCheckoutMercadoPago(mpRes.erroCodigo);
      }
    } else {
      var row = montarLinhaInscricao(data, "Pendente (PIX via WhatsApp)");
      var pendPresencial = obterAbaPendentes(ss);
      garantirCabecalhosPlanilha(pendPresencial);
      pendPresencial.appendRow(row);
      sincronizarBackupSegurancaNoDrive();
      enviarEmailInscricaoRecebida(row);
    }

    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var es = String(err);
    if (es.indexOf("UrlFetchApp") !== -1 || es.indexOf("external_request") !== -1) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, error: mensagemErroUrlFetchPermissao(), codigo: "urlfetch_auth" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: es })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var protocoloQs = e && e.parameter && e.parameter.protocolo ? String(e.parameter.protocolo) : "";
  var senhaQs = e && e.parameter && e.parameter.senha ? String(e.parameter.senha) : "";
  if (protocoloQs && senhaQs) {
    var outManualQs = confirmarPagamentoPresencialPorProtocolo(protocoloQs, senhaQs);
    return respostaTexto(montarMensagemLeigaAtualizacaoPagamento(outManualQs));
  }

  var path = e && e.pathInfo != null ? String(e.pathInfo) : "";
  path = path.replace(/^\/+|\/+$/g, "");
  if (path) {
    var parts = path.split("/");
    if (parts.length === 2) {
      var outManual = confirmarPagamentoPresencialPorProtocolo(parts[0], parts[1]);
      return respostaTexto(montarMensagemLeigaAtualizacaoPagamento(outManual));
    }
  }

  if (e && e.parameter && e.parameter.backup === "1") {
    var keyEsperada = PropertiesService.getScriptProperties().getProperty("BACKUP_JSON_KEY");
    var keyRecebida = e.parameter.key != null ? String(e.parameter.key) : "";
    if (!keyEsperada || keyRecebida !== keyEsperada) {
      return respostaJson({ ok: false, error: "Não autorizado. Defina BACKUP_JSON_KEY nas Propriedades do script e use ?backup=1&key=..." });
    }
    try {
      var payloadBackup = gerarPayloadBackupListaInscritos();
      return respostaJson(payloadBackup);
    } catch (errBackup) {
      return respostaJson({ ok: false, error: String(errBackup) });
    }
  }

  if (e && e.parameter && e.parameter.topic === "payment" && e.parameter.id) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      processarNotificacaoPagamentoMercadoPago(String(e.parameter.id));
    } finally {
      lock.releaseLock();
    }
    return ContentService.createTextOutput("OK");
  }
  return respostaTexto(
    "Use POST JSON para inscricoes.\n\n" +
      "Para aprovar pagamento presencial, use um destes formatos:\n" +
      "1) /exec/PROTOCOLO/SENHA\n" +
      "2) /exec?protocolo=PROTOCOLO&senha=SENHA"
  );
}

function colocarCabecalhosNaPlanilha() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = obterAbaInscricoes(ss);
  if (!sheet) throw new Error("Aba não encontrada.");
  if (sheet.getRange("A1").getValue() !== "") {
    throw new Error('A célula A1 já tem conteúdo. Limpe A1 ou insira uma linha em branco no topo.');
  }
  if (sheet.getLastRow() === 0) {
    garantirCabecalhosPlanilha(sheet);
    return;
  }
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, CABECALHOS.length).setValues([CABECALHOS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, CABECALHOS.length).setFontWeight("bold");
}

/**
 * Execute UMA VEZ no editor (Executar → autorizarAcessoExterno) como dono do projeto.
 * Aceite "Acessar dados em serviços externos" / conexão com URL. Depois nova implantação se pedir.
 */
function autorizarAcessoExterno() {
  UrlFetchApp.fetch("https://api.mercadopago.com", { muteHttpExceptions: true });
}

/**
 * Execute no editor para autorizar e validar envio de e-mail da aplicação.
 * Opcional: defina EMAIL_TESTE nas Propriedades do script para receber em outro endereço.
 */
function testarEnvioEmailAplicacao() {
  var props = PropertiesService.getScriptProperties();
  var destino = String(props.getProperty("EMAIL_TESTE") || Session.getActiveUser().getEmail() || "").trim();
  if (!destino || destino.indexOf("@") < 0) {
    throw new Error("Defina EMAIL_TESTE nas Propriedades do script com um e-mail válido.");
  }

  var assunto = "Teste de envio — " + NOME_EVENTO_EMAIL;
  var corpoTexto =
    "Este é um teste de envio de e-mail da aplicação.\n\n" +
    "Se você recebeu, o MailApp está autorizado e funcionando.\n\n" +
    "Evento: " +
    NOME_EVENTO_EMAIL +
    "\nData: " +
    DATA_EVENTO_EMAIL +
    "\nHorário: " +
    HORARIO_EVENTO_EMAIL +
    "\nLocal: " +
    LOCAL_EVENTO_EMAIL +
    "\n\n— Organização";

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;background:#f5f8fa;padding:24px;">' +
    '<div style="max-width:680px;margin:0 auto;background:#fff;border:1px solid #d7e3ea;border-radius:12px;overflow:hidden;">' +
    '<div style="background:linear-gradient(90deg,#004C74,#00BEE2);padding:18px 24px;color:#fff;">' +
    '<h1 style="margin:0;font-size:22px;">Teste de envio concluído</h1>' +
    '<p style="margin:8px 0 0;font-size:14px;opacity:.95;">' +
    escapeHtmlEmail(NOME_EVENTO_EMAIL) +
    "</p>" +
    "</div>" +
    '<div style="padding:24px;color:#16384a;">' +
    '<p style="margin:0 0 12px;">Se você recebeu este e-mail, a sua aplicação está autorizada para enviar mensagens automáticas aos inscritos.</p>' +
    '<p style="margin:0;font-size:13px;color:#355264;">Destino do teste: <strong>' +
    escapeHtmlEmail(destino) +
    "</strong></p>" +
    "</div></div></div>";

  MailApp.sendEmail({
    to: destino,
    subject: assunto,
    body: corpoTexto,
    htmlBody: htmlBody,
  });

  Logger.log("Teste de e-mail enviado para: " + destino);
}
