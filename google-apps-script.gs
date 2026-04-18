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
 *   E-mail: o endereço "De" é sempre a conta Google que implanta/autoriza o script. Opcional nas Propriedades do script:
 *     EMAIL_REMETENTE_NOME = nome exibido no remetente (ex.: Corrida Mariana).
 *     EMAIL_REPLY_TO = e-mail para onde vão as respostas (Reply-To).
 *
 * Produção MP: (1) token de produção em MERCADO_PAGO_ACCESS_TOKEN, (2) config.js com useSandbox: false e urlRetorno HTTPS real,
 * (3) Implantar → Gerenciar implantações → Nova versão na Web App.
 *
 * Consulta pública: POST JSON { "tipo": "consulta_inscricao", "email": "…", "telefone": "…" } — busca por e-mail + telefone (só dígitos) na lista oficial e em todas as abas de fila pendentes (nomes em NOMES_RESERVADOS_ABA_PENDENTES).
 *
 * PIX via WhatsApp (presencial): a linha fica na fila de pendentes até aprovar. GET /exec?protocolo=…&senha=… (ou /exec/PROTO/SENHA) procura o protocolo na lista oficial e em todas essas abas; ao aprovar, move para a lista oficial, status "Pago (presencial confirmado)" e envia e-mail de confirmação (MailApp autorizado).
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
var HORARIO_EVENTO_EMAIL = "Concentração às 5h30 · Largada às 6h da manhã";
var LOCAL_EVENTO_EMAIL = "Sanharó, Pernambuco";
var WHATSAPP_ORGANIZACAO_EMAIL = "5587991200165";
var INSTAGRAM_ORGANIZACAO_URL = "https://instagram.com/";
/** Use URL pública da logo (opcional). Ex.: https://seusite.com/assets/logo.png */
var LOGO_EMAIL_URL = "";

/**
 * Nome exibido como remetente nos e-mails (MailApp → campo "name"). O endereço continua sendo a conta que autoriza o script.
 * Sobrescreve com a propriedade do script EMAIL_REMETENTE_NOME. Deixe "" para usar o padrão do Gmail.
 */
var EMAIL_REMETENTE_NOME = "";

/**
 * Reply-To: para onde o inscrito responde ao e-mail. Útil se o envio é pela sua conta mas as respostas devem ir à organização.
 * Sobrescreve com EMAIL_REPLY_TO nas Propriedades do script. Deixe "" para não definir.
 */
var EMAIL_REPLY_TO = "";

/** Mesma URL que config.js → webhookUrl (termina em /exec). Alinhe os dois se mudar a implantação. */
var WEB_APP_URL_FALLBACK =
  "https://script.google.com/a/macros/redealia.com/s/AKfycbwClwfOb9G4AXWFW0tjQAMAkuX_VlmlBHJC6nFPuGczHZZBM0XLv4p36mYF0RvvHCu7/exec";

/** Aba principal: inscrições confirmadas (pagamento aprovado). */
/** Aba de fila: Mercado Pago e presencial/PIX. O script também reconhece nomes alternativos (ver obterAbaPendentes). */
var NOME_ABA_PENDENTES = "Inscrições pendentes MP";
/** Nomes de aba que são só fila de pendentes — nunca usar como lista oficial (fallback sheets[0]). */
var NOMES_RESERVADOS_ABA_PENDENTES = ["Inscrições pendentes MP", "Pendentes", "Inscrições pendentes", "Pendências"];

/** Limites de lotes: alinhar com config.js */
var LIMITE_LOTE_PROMO = 50;
var ID_LOTE_PROMO = "promo";
var LIMITE_LOTE_REGULAR = 100;
var ID_LOTE_REGULAR = "regular";
/** Preços em produção (alinhar com config.js PRECO_PROMO / PRECO_REGULAR). Usados ao reclassificar lote automaticamente. */
var VALOR_LOTE_PROMO_REAIS = 50;
var VALOR_LOTE_REGULAR_REAIS = 55;
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
  var listaPend = listarAbasPorNomesPendentes(ss);
  for (var i = 0; i < listaPend.length; i++) {
    var pend = listaPend[i];
    if (pend.getLastRow() === 0) continue;
    garantirCabecalhosPlanilha(pend);
    var r = jaTemCadastro(pend, email, telDigits);
    if (r) return r;
  }
  return null;
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
  var listaPendPurge = listarAbasPorNomesPendentes(ss);
  for (var pi = 0; pi < listaPendPurge.length; pi++) {
    purgeSheet(listaPendPurge[pi]);
  }
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
  var listaPendLote = listarAbasPorNomesPendentes(ss);
  for (var li = 0; li < listaPendLote.length; li++) {
    var pend = listaPendLote[li];
    if (pend.getLastRow() > 0) {
      garantirCabecalhosPlanilha(pend);
      n += contarInscricoesLote(pend, loteId);
    }
  }
  return n;
}

function nomeAbaEhReservadaPendentes(nome) {
  var n = String(nome || "").trim();
  for (var i = 0; i < NOMES_RESERVADOS_ABA_PENDENTES.length; i++) {
    if (n === NOMES_RESERVADOS_ABA_PENDENTES[i]) return true;
  }
  return false;
}

/** Abre a planilha do evento ou lança erro claro (evita TypeError em ss.getSheetByName). */
function obterPlanilhaCorridaOuErro() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  if (!ss || typeof ss.getSheetByName !== "function") {
    throw new Error(
      "Não foi possível abrir a planilha. Confira ID_PLANILHA no script e permissões da conta que executa a Web App."
    );
  }
  return ss;
}

function obterAbaInscricoes(ss) {
  if (ss == null || typeof ss.getSheetByName !== "function") {
    throw new Error("obterAbaInscricoes: planilha (ss) inválida ou não informada.");
  }
  var sh = ss.getSheetByName("Lista de inscritos");
  if (sh && !nomeAbaEhReservadaPendentes(sh.getName())) return sh;
  sh = ss.getSheetByName("Inscrições");
  if (sh && !nomeAbaEhReservadaPendentes(sh.getName())) return sh;
  sh = ss.getSheetByName("Página1");
  if (sh && !nomeAbaEhReservadaPendentes(sh.getName())) return sh;
  sh = ss.getSheetByName("Sheet1");
  if (sh && !nomeAbaEhReservadaPendentes(sh.getName())) return sh;
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    if (!nomeAbaEhReservadaPendentes(sheets[j].getName())) return sheets[j];
  }
  return null;
}

/**
 * Todas as abas da planilha que são fila de pendentes (por nome), exceto a lista oficial.
 * Ordem = NOMES_RESERVADOS_ABA_PENDENTES (para escrita preferir a primeira vazia ou com dados).
 */
function listarAbasPorNomesPendentes(ss) {
  if (ss == null || typeof ss.getSheetByName !== "function") {
    throw new Error("listarAbasPorNomesPendentes: planilha (ss) inválida ou não informada.");
  }
  var main = obterAbaInscricoes(ss);
  var list = [];
  var seen = {};
  for (var k = 0; k < NOMES_RESERVADOS_ABA_PENDENTES.length; k++) {
    var sh = ss.getSheetByName(NOMES_RESERVADOS_ABA_PENDENTES[k]);
    if (!sh) continue;
    var id = sh.getSheetId();
    if (seen[id]) continue;
    if (main && id === main.getSheetId()) continue;
    seen[id] = true;
    list.push(sh);
  }
  return list;
}

/**
 * Aba onde gravar nova pendência: se existir mais de uma aba de nomes conhecidos, prefere a que já tem linhas de dado
 * (evita gravar numa aba vazia enquanto o protocolo continua noutra — o link de aprovação deixaria de achar a linha).
 */
function obterAbaPendentes(ss) {
  if (ss == null || typeof ss.insertSheet !== "function") {
    throw new Error("obterAbaPendentes: planilha (ss) inválida ou não informada.");
  }
  var candidates = listarAbasPorNomesPendentes(ss);
  if (candidates.length === 0) {
    /** Nova aba no fim da planilha (índice = quantidade de abas), para não virar a “primeira aba” à esquerda. */
    var idxFim = ss.getSheets().length;
    return ss.insertSheet(NOME_ABA_PENDENTES, idxFim);
  }
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].getLastRow() > 1) return candidates[i];
  }
  return candidates[0];
}

/** Retorna { sheet, rowIndex } ou null se o protocolo não estiver em nenhuma aba de pendentes conhecida. */
function encontrarProtocoloNasAbasPendentes(ss, protocolo) {
  if (ss == null || typeof ss.getSheetByName !== "function") {
    throw new Error("encontrarProtocoloNasAbasPendentes: planilha (ss) inválida ou não informada.");
  }
  var lista = listarAbasPorNomesPendentes(ss);
  var p = String(protocolo || "").trim();
  for (var i = 0; i < lista.length; i++) {
    garantirCabecalhosPlanilha(lista[i]);
    var r = encontrarLinhaPorProtocolo(lista[i], p);
    if (r > 0) return { sheet: lista[i], rowIndex: r };
  }
  return null;
}

/**
 * Garante que lista oficial e fila de pendentes são abas diferentes (evita gravar presencial na “lista” por engano).
 */
function assertAbasInscricaoDistintas(main, pend) {
  if (!main || !pend) return;
  if (main.getSheetId() === pend.getSheetId()) {
    throw new Error(
      "A aba principal e a de pendentes coincidem. Na planilha, use uma aba com nome \"Lista de inscritos\" (oficial) " +
        "e outra para a fila: \"" +
        NOME_ABA_PENDENTES +
        "\" ou \"Pendentes\". Não use o mesmo nome para as duas funções."
    );
  }
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

  var mainB = obterAbaInscricoes(ss);
  var listaPB = listarAbasPorNomesPendentes(ss);
  var alvo = [mainB];
  var nomes = ["lista_oficial"];
  for (var pb = 0; pb < listaPB.length; pb++) {
    alvo.push(listaPB[pb]);
    nomes.push("pendente_fila");
  }
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
    /** Alinhado aos e-mails (NOME_EVENTO_EMAIL, DATA_EVENTO_EMAIL, …) para exibir na consulta pública. */
    nomeEvento: NOME_EVENTO_EMAIL,
    dataEvento: DATA_EVENTO_EMAIL,
    horarioEvento: HORARIO_EVENTO_EMAIL,
    localEvento: LOCAL_EVENTO_EMAIL,
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
  var ss = obterPlanilhaCorridaOuErro();
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
  var ss = obterPlanilhaCorridaOuErro();
  var sheet = obterAbaInscricoes(ss);
  if (!sheet) throw new Error("Aba Lista de inscritos não encontrada.");
  garantirCabecalhosPlanilha(sheet);
  var listaOficial = listaBackupDaAba(sheet);
  var listaPend = [];
  var shPendBackup = listarAbasPorNomesPendentes(ss);
  for (var bi = 0; bi < shPendBackup.length; bi++) {
    garantirCabecalhosPlanilha(shPendBackup[bi]);
    var trecho = listaBackupDaAba(shPendBackup[bi]);
    for (var bj = 0; bj < trecho.length; bj++) listaPend.push(trecho[bj]);
  }
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
  var ss = obterPlanilhaCorridaOuErro();
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

/** Opções extras do MailApp (nome do remetente, reply-to). Propriedades do script têm prioridade sobre as variáveis do código. */
function opcoesRemetenteMailApp() {
  var props = PropertiesService.getScriptProperties();
  var nome = String(props.getProperty("EMAIL_REMETENTE_NOME") || EMAIL_REMETENTE_NOME || "").trim();
  var replyTo = String(props.getProperty("EMAIL_REPLY_TO") || EMAIL_REPLY_TO || "").trim();
  var out = {};
  if (nome) out.name = nome;
  if (replyTo && replyTo.indexOf("@") > 0) out.replyTo = replyTo;
  return out;
}

function enviarMailAppComRemetente(opts) {
  var extra = opcoesRemetenteMailApp();
  for (var k in extra) {
    if (Object.prototype.hasOwnProperty.call(extra, k)) opts[k] = extra[k];
  }
  MailApp.sendEmail(opts);
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
    enviarMailAppComRemetente({
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
 * Dados variáveis para o template único de confirmação de pagamento (texto + chip + cores).
 * payJson: GET /v1/payments/:id só no webhook MP.
 */
function detalheConfirmacaoPagamentoParaTemplate(rowData, payJson) {
  var status = String(rowData[COL_IX_STATUS] || "").toLowerCase();
  if (status.indexOf("presencial") !== -1) {
    return {
      chipLabel: "Presencial",
      chipCor: "#004C74",
      txt:
        "Meio: confirmação presencial pela organização.\n\n" +
        "Seu pagamento foi confirmado pela organização e sua inscrição está na lista oficial do evento.",
      fraseHtml:
        "A organização <strong>confirmou seu pagamento</strong>. Sua inscrição já consta na <strong>lista oficial</strong> do evento — é só comparecer no dia e curtir a corrida.",
    };
  }
  if (payJson && typeof payJson === "object") {
    var pmId = String(payJson.payment_method_id || "").toLowerCase();
    var ptId = String(payJson.payment_type_id || "").toLowerCase();
    if (pmId === "pix") {
      return {
        chipLabel: "PIX",
        chipCor: "#00b1a5",
        txt:
          "Meio: PIX (Mercado Pago).\n\n" +
          "Seu pagamento via PIX pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
        fraseHtml:
          "Identificamos e <strong>aprovamos</strong> seu pagamento via <strong>PIX</strong> pelo Mercado Pago. Sua vaga está <strong>garantida</strong> na lista oficial.",
      };
    }
    if (ptId === "credit_card") {
      return {
        chipLabel: "Cartão de crédito",
        chipCor: "#004C74",
        txt:
          "Meio: cartão de crédito (Mercado Pago).\n\n" +
          "Seu pagamento com cartão de crédito pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
        fraseHtml:
          "Seu pagamento com <strong>cartão de crédito</strong> pelo Mercado Pago foi <strong>aprovado</strong>. Sua inscrição já está na <strong>lista oficial</strong>.",
      };
    }
    if (ptId === "debit_card") {
      return {
        chipLabel: "Cartão de débito",
        chipCor: "#006799",
        txt:
          "Meio: cartão de débito (Mercado Pago).\n\n" +
          "Seu pagamento com cartão de débito pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
        fraseHtml:
          "Seu pagamento com <strong>cartão de débito</strong> pelo Mercado Pago foi <strong>aprovado</strong>. Sua inscrição já está na <strong>lista oficial</strong>.",
      };
    }
    if (ptId === "ticket" || pmId.indexOf("bol") === 0) {
      return {
        chipLabel: "Boleto",
        chipCor: "#5b4b8a",
        txt:
          "Meio: boleto (Mercado Pago).\n\n" +
          "Seu pagamento com boleto pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
        fraseHtml:
          "Seu <strong>boleto</strong> foi compensado pelo Mercado Pago e o pagamento está <strong>confirmado</strong>. Sua inscrição já está na <strong>lista oficial</strong>.",
      };
    }
    return {
      chipLabel: "Mercado Pago",
      chipCor: "#009ee3",
      txt:
        "Meio: Mercado Pago.\n\n" +
        "Seu pagamento pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
      fraseHtml:
        "Seu pagamento pelo <strong>Mercado Pago</strong> foi <strong>confirmado</strong>. Sua inscrição já está na <strong>lista oficial</strong> do evento.",
    };
  }
  if (status.indexOf("mercado") !== -1) {
    return {
      chipLabel: "Mercado Pago",
      chipCor: "#009ee3",
      txt:
        "Meio: Mercado Pago.\n\n" +
        "Seu pagamento pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
      fraseHtml:
        "Seu pagamento pelo <strong>Mercado Pago</strong> foi <strong>confirmado</strong>. Sua inscrição já está na <strong>lista oficial</strong>.",
    };
  }
  return {
    chipLabel: "Confirmado",
    chipCor: "#0d7a4f",
    txt: "Seu pagamento foi confirmado e sua inscrição está na lista oficial do evento.",
    fraseHtml:
      "Seu pagamento foi <strong>confirmado</strong>. Sua inscrição já está na <strong>lista oficial</strong> do evento.",
  };
}

/**
 * Envia e-mail ao inscrito quando o pagamento MP é confirmado (webhook).
 * payJsonOpcional: objeto do pagamento MP (GET /v1/payments) para personalizar PIX vs cartão etc.
 * 1ª vez: o Apps Script pedirá permissão para enviar e-mail. Limite diário do Gmail se aplicam.
 */
function enviarEmailPagamentoConfirmado(rowData, payJsonOpcional) {
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
  var det = detalheConfirmacaoPagamentoParaTemplate(rowData, payJsonOpcional);
  var waLink = "https://wa.me/" + String(WHATSAPP_ORGANIZACAO_EMAIL || "").replace(/\D/g, "");
  var logoHtml = "";
  if (String(LOGO_EMAIL_URL || "").trim()) {
    logoHtml =
      '<p style="margin:0 0 20px;text-align:center;">' +
      '<img src="' +
      escapeHtmlEmail(LOGO_EMAIL_URL) +
      '" alt="Logo do evento" style="max-width:200px;height:auto;border:0;display:inline-block;"/>' +
      "</p>";
  }
  var logoRow = logoHtml
    ? '<tr><td style="padding:24px 28px 0;text-align:center;background:#ffffff;">' + logoHtml + "</td></tr>"
    : "";
  var chipHtml =
    '<span style="display:inline-block;margin-top:4px;padding:8px 18px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#ffffff;background:' +
    det.chipCor +
    ';">' +
    escapeHtmlEmail(det.chipLabel) +
    "</span>";

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
    det.txt +
    "\n\n" +
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
    '<div style="margin:0;padding:0;background:#dfe8ef;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#dfe8ef;font-family:Arial,Helvetica,sans-serif;">' +
    '<tr><td style="padding:28px 14px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 16px 48px rgba(0,52,84,.15);">' +
    '<tr><td style="background:linear-gradient(135deg,#003d5c 0%,#005a7a 40%,#00a3c4 100%);padding:30px 26px 28px;text-align:center;color:#ffffff;">' +
    '<p style="margin:0 0 8px;font-size:11px;letter-spacing:.26em;text-transform:uppercase;opacity:.88;">Inscrição confirmada</p>' +
    '<h1 style="margin:0;font-size:23px;line-height:1.25;font-weight:700;">Tudo certo — você está na lista oficial</h1>' +
    '<p style="margin:14px 0 0;font-size:15px;line-height:1.45;opacity:.93;">' +
    escapeHtmlEmail(NOME_EVENTO_EMAIL) +
    "</p>" +
    "</td></tr>" +
    logoRow +
    '<tr><td style="padding:22px 26px 10px;background:#ffffff;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f0fdf4;border:1px solid #86efac;border-radius:16px;">' +
    '<tr><td style="padding:26px 22px;text-align:center;">' +
    '<div style="display:inline-block;width:50px;height:50px;line-height:50px;border-radius:50%;background:' +
    det.chipCor +
    ';color:#ffffff;font-size:22px;font-weight:bold;">&#10003;</div>' +
    '<div style="margin-top:14px;">' +
    chipHtml +
    "</div>" +
    '<p style="margin:20px 0 0;font-size:17px;color:#0f172a;line-height:1.45;">Olá <strong>' +
    escapeHtmlEmail(nome || "atleta") +
    "</strong>,</p>" +
    '<p style="margin:12px 0 0;font-size:15px;color:#334155;line-height:1.55;">' +
    det.fraseHtml +
    "</p>" +
    "</td></tr></table>" +
    "</td></tr>" +
    '<tr><td style="padding:6px 26px 22px;background:#ffffff;">' +
    '<p style="margin:0 0 14px;font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#64748b;">Seus dados</p>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;">' +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;width:38%;font-size:13px;color:#64748b;background:#f8fafc;">Protocolo</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;font-weight:700;">' +
    escapeHtmlEmail(proto || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Nome</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(nome || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">E-mail</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(email || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Cidade</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(cidade || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Camisa</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(camisa || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Percurso</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(percurso || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Lote</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(loteNome || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Valor</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;font-weight:700;">' +
    escapeHtmlEmail(valorFmt || "—") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#64748b;background:#f8fafc;">Forma de pagamento</td><td style="padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:14px;color:#0f172a;">' +
    escapeHtmlEmail(formaPagamento || "Mercado Pago") +
    "</td></tr>" +
    '<tr><td style="padding:11px 14px;font-size:13px;color:#64748b;background:#f8fafc;">Status</td><td style="padding:11px 14px;font-size:14px;color:#047857;font-weight:700;">' +
    escapeHtmlEmail(statusPagamento || "Pago (Mercado Pago)") +
    "</td></tr>" +
    "</table>" +
    "</td></tr>" +
    '<tr><td style="padding:0 26px 22px;background:#ffffff;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(180deg,#f0f9ff 0%,#e0f2fe 100%);border:1px solid #bae6fd;border-radius:14px;">' +
    '<tr><td style="padding:18px 18px 16px;">' +
    '<p style="margin:0 0 10px;font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#0369a1;">Evento</p>' +
    '<p style="margin:0;font-size:14px;line-height:1.65;color:#0c4a6e;">' +
    "<strong>Data:</strong> " +
    escapeHtmlEmail(DATA_EVENTO_EMAIL) +
    "<br/><strong>Horário:</strong> " +
    escapeHtmlEmail(HORARIO_EVENTO_EMAIL) +
    "<br/><strong>Local:</strong> " +
    escapeHtmlEmail(LOCAL_EVENTO_EMAIL) +
    "</p>" +
    "</td></tr></table>" +
    "</td></tr>" +
    '<tr><td style="padding:0 26px 26px;background:#ffffff;text-align:center;">' +
    '<p style="margin:0 0 16px;font-size:13px;color:#475569;line-height:1.5;">Guarde este e-mail e o <strong>protocolo</strong> para qualquer dúvida.</p>' +
    '<a href="' +
    escapeHtmlEmail(waLink) +
    '" style="display:inline-block;background:#128c7e;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 22px;border-radius:999px;">Falar com a organização no WhatsApp</a>' +
    '<p style="margin:18px 0 0;font-size:12px;color:#64748b;">Instagram: <a href="' +
    escapeHtmlEmail(INSTAGRAM_ORGANIZACAO_URL) +
    '" style="color:#0369a1;text-decoration:none;">' +
    escapeHtmlEmail(INSTAGRAM_ORGANIZACAO_URL) +
    "</a></p>" +
    "</td></tr>" +
    '<tr><td style="padding:16px 26px;background:#f1f5f9;border-top:1px solid #e2e8f0;text-align:center;font-size:12px;color:#64748b;line-height:1.5;">' +
    escapeHtmlEmail(NOME_EVENTO_EMAIL) +
    "<br/>Organização do evento</td></tr>" +
    "</table></td></tr></table></div>";
  try {
    enviarMailAppComRemetente({
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
  var ss = obterPlanilhaCorridaOuErro();
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

  var ss = obterPlanilhaCorridaOuErro();
  var achadoMp = encontrarProtocoloNasAbasPendentes(ss, ref);
  if (!achadoMp) {
    Logger.log("Protocolo pendente não encontrado: " + ref);
    return;
  }
  var pend = achadoMp.sheet;
  var rowIndex = achadoMp.rowIndex;

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
  enviarEmailPagamentoConfirmado(rowData, pay);
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
  var ss = obterPlanilhaCorridaOuErro();
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

  var ss = obterPlanilhaCorridaOuErro();
  var main = obterAbaInscricoes(ss);
  if (!main) {
    return { ok: false, error: "Aba principal de inscrições não encontrada." };
  }
  garantirCabecalhosPlanilha(main);

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

  var achadoPres = encontrarProtocoloNasAbasPendentes(ss, p);
  if (!achadoPres) {
    return { ok: false, error: "Protocolo não encontrado.", protocolo: p };
  }
  var sheetPend = achadoPres.sheet;
  var rowPend = achadoPres.rowIndex;

  var rowData = sheetPend.getRange(rowPend, 1, 1, CABECALHOS.length).getValues()[0];
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
  sheetPend.deleteRow(rowPend);
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

    var ss = obterPlanilhaCorridaOuErro();
    var sheet = obterAbaInscricoes(ss);
    if (!sheet) {
      throw new Error(
        "Nenhuma aba válida para a lista oficial. Crie/renomeie uma aba para \"Lista de inscritos\" (confirmados) " +
          "e mantenha a fila de pendentes com outro nome (ex.: \"" +
          NOME_ABA_PENDENTES +
          "\" ou \"Pendentes\")."
      );
    }
    garantirCabecalhosPlanilha(sheet);
    var abaPendInscricao = obterAbaPendentes(ss);
    assertAbasInscricaoDistintas(sheet, abaPendInscricao);
    garantirCabecalhosPlanilha(abaPendInscricao);

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

    var loteAjustadoAutomatico = false;
    var loteEfetivo = String(data.lote || "").trim();

    if (loteEfetivo === ID_LOTE_PROMO) {
      var jaPromo = contarInscricoesLoteTotal(ss, ID_LOTE_PROMO);
      if (jaPromo >= LIMITE_LOTE_PROMO) {
        var jaRegCheio = contarInscricoesLoteTotal(ss, ID_LOTE_REGULAR);
        if (jaRegCheio >= LIMITE_LOTE_REGULAR) {
          return ContentService.createTextOutput(
            JSON.stringify({
              ok: false,
              error:
                "Lotes esgotados: o promocional já encerrou e o lote regular também atingiu o limite de " +
                LIMITE_LOTE_REGULAR +
                " inscrições.",
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }
        data.lote = ID_LOTE_REGULAR;
        data.loteNome = "Lote regular";
        data.valorReais = VALOR_LOTE_REGULAR_REAIS;
        loteEfetivo = ID_LOTE_REGULAR;
        loteAjustadoAutomatico = true;
      }
    }

    if (loteEfetivo === ID_LOTE_REGULAR) {
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
    if (loteAjustadoAutomatico) {
      out.loteAjustadoAutomatico = true;
      out.lote = data.lote;
      out.loteNome = data.loteNome;
      out.valorReais = data.valorReais;
    }

    if (data.mercadoPago) {
      var pend = abaPendInscricao;
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
      var pendPresencial = abaPendInscricao;
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
  var ss = obterPlanilhaCorridaOuErro();
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

  enviarMailAppComRemetente({
    to: destino,
    subject: assunto,
    body: corpoTexto,
    htmlBody: htmlBody,
  });

  Logger.log("Teste de e-mail enviado para: " + destino);
}
