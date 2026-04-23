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
 * Consulta pública: POST JSON { "tipo": "consulta_inscricao", "email": "…", "telefone": "…" } — busca na lista oficial e em abas de pendentes (nomes reservados) e em qualquer outra aba que tenha o cabeçalho padrão (linha 1, coluna "Protocolo").
 * Alterar forma de pagamento (só se ainda não pago): POST { "tipo": "alterar_forma_pagamento", "email", "telefone", "protocolo", "novaForma": "mercado_pago_online" | "presencial_secretaria", "urlRetorno"?, "useSandbox"? } — atualiza colunas na planilha; para online devolve checkoutUrl.
 *
 * PIX via WhatsApp: fila de pendentes até aprovar pelo GET /exec?protocolo=…&senha=… — status "Pago (presencial confirmado)" + e-mail.
 *
 * Mercado Pago (checkout): preferência **sem PIX nem boleto** — só cartão (crédito/débito). Quando o pagamento fica **approved**, o webhook move para a lista oficial, status "Pago (Mercado Pago)", linha verde e e-mail.
 *
 * Limites de lote (camisas): contam só inscrições com status de pagamento confirmado (texto com "Pago" — ex.: Pago (Mercado Pago), Pago (presencial confirmado)). Pendências e "Aguardando pagamento online" não consomem vaga do lote. Exige coluna Protocolo preenchida.
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
 * Cores das linhas: o script já pinta fundo vermelho (não pago) e verde (pago) ao gravar — não depende de rodar formatação condicional.
 * Opcional: **instalarFormatacaoCoresStatusPlanilha** duplica o efeito por regras; evite acumular regras repetidas.
 * Coluna "Link aprovar PIX": exige WEB_APP_URL (Propriedades) ou WEB_APP_URL_FALLBACK no código; senha em APROVACAO_SENHA ou SENHA no código.
 * Se a célula do link estiver vazia ou com aviso, configure e rode **reaplicarLinksAprovarPixOndeFaltam**.
 *
 * Compartilhamento da planilha (Google Drive): quem tiver permissão de EDITOR pode ver, alterar e apagar
 * todas as linhas. Não use "Qualquer pessoa com o link pode editar" para o público. Restrinja a
 * organizadores (e-mail) ou use "Somente visualização" para quem só precisa consultar.
 *
 * ID da planilha (URL .../spreadsheets/d/ESTE_ID/edit). Tem que ser a mesma planilha que você abre no Drive;
 * senão as inscrições “somem” (gravam em outro arquivo).
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
  "https://script.google.com/macros/s/AKfycbxL_3csfdYsIsimppYjGf4rK44kRuArvItMUs2miQtVy8FusEVzmwCe-glgNrTAYiBi/exec";

/** Aba principal: inscrições confirmadas (pagamento aprovado). */
/** Aba de fila: Mercado Pago e presencial/PIX. O script também reconhece nomes alternativos (ver obterAbaPendentes). */
var NOME_ABA_PENDENTES = "Inscrições pendentes MP";
/** Nomes de aba que são só fila de pendentes — nunca usar como lista oficial (fallback sheets[0]). */
var NOMES_RESERVADOS_ABA_PENDENTES = [
  "Inscrições pendentes MP",
  "Pendentes",
  "Inscrições pendentes",
  "Pendências",
  "Fila PIX",
  "PIX pendentes",
];

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
/** Coluna "Status pagamento" (0-based), alinhado a CABECALHOS */
var COL_IX_STATUS = 12;
/** Coluna com fórmula HYPERLINK para aprovação GET (PIX pendente). */
var COL_IX_LINK_APROVACAO = 13;
/** Coluna "Forma pagamento" (texto exibido na planilha). */
var COL_IX_FORMA = 11;

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
  "Link aprovar PIX",
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
  "linkAprovarPix",
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
  var mainJ = obterAbaInscricoes(ss);
  var m = jaTemCadastro(mainJ, email, telDigits);
  if (m) return m;
  var listaPend = listarAbasParaBuscaInscricao(ss);
  for (var i = 0; i < listaPend.length; i++) {
    var pend = listaPend[i];
    if (mainJ && pend.getSheetId() === mainJ.getSheetId()) continue;
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
  if (s.indexOf("análise") !== -1 || s.indexOf("analise") !== -1) return "nao_pago";
  if (s.indexOf("processando") !== -1) return "nao_pago";
  if (s.indexOf("pago") !== -1) return "pago";
  return "desconhecido";
}

/**
 * Pode exibir / permitir troca de forma: status claramente em aberto, ou inscrição na fila de pendentes sem status "pago".
 */
function inscricaoPermiteAlterarFormaPagamento(stTexto, origem) {
  var cls = classificarStatusPagamento(stTexto);
  if (cls === "pago") return false;
  if (cls === "nao_pago") return true;
  if (origem === "pendente_fila") return true;
  return false;
}

/**
 * Alinha com os rótulos do site (inscricao.js): mercado_pago_online | presencial_secretaria | "".
 */
function inferirCodigoFormaPagamentoLinha(row) {
  if (!row || !row.length) return "";
  var st = String(row[COL_IX_STATUS] || "").toLowerCase();
  var fp = String(row[COL_IX_FORMA] || "").toLowerCase();
  if (st.indexOf("aguardando") !== -1) return "mercado_pago_online";
  if (st.indexOf("pendente (pix") !== -1 || (st.indexOf("pendente") !== -1 && st.indexOf("pix") !== -1))
    return "presencial_secretaria";
  if (fp.indexOf("mercado pago") !== -1 && fp.indexOf("online") !== -1) return "mercado_pago_online";
  if (fp.indexOf("mercado pago") !== -1) return "mercado_pago_online";
  if (fp.indexOf("pix") !== -1 || fp.indexOf("whatsapp") !== -1 || fp.indexOf("secretaria") !== -1)
    return "presencial_secretaria";
  return "";
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
  var mainPurge = obterAbaInscricoes(ss);
  var listaPendPurge = listarAbasParaBuscaInscricao(ss);
  for (var pi = 0; pi < listaPendPurge.length; pi++) {
    var shP = listaPendPurge[pi];
    if (mainPurge && shP.getSheetId() === mainPurge.getSheetId()) continue;
    purgeSheet(shP);
  }
  purgeSheet(mainPurge);
}

/** Quantas vagas do lote já foram efetivamente pagas (para limites de camisas / promo vs regular). */
function contarInscricoesLote(sheet, loteId) {
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][COL_IX_PROTOCOLO]) === "Protocolo") {
    start = 1;
  }
  var n = 0;
  for (var i = start; i < values.length; i++) {
    var proto = String(values[i][COL_IX_PROTOCOLO] || "").trim();
    if (!proto) continue;
    if (classificarStatusPagamento(values[i][COL_IX_STATUS]) !== "pago") continue;
    if (String(values[i][COL_IX_LOTE]).trim() === String(loteId)) {
      n++;
    }
  }
  return n;
}

function contarInscricoesLoteTotal(ss, loteId) {
  var main = obterAbaInscricoes(ss);
  var n = contarInscricoesLote(main, loteId);
  var listaTodas = listarAbasParaBuscaInscricao(ss);
  for (var li = 0; li < listaTodas.length; li++) {
    var pend = listaTodas[li];
    if (main && pend.getSheetId() === main.getSheetId()) continue;
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

/** Linha 1 tem "Protocolo" na coluna certa — mesma estrutura da lista de inscrições. */
function sheetTemCabecalhoInscricao(sh) {
  if (!sh || sh.getLastRow() < 1) return false;
  var v = String(sh.getRange(1, COL_IX_PROTOCOLO + 1).getValue() || "").trim();
  return v === "Protocolo";
}

/**
 * Pendentes: abas com nomes reservados primeiro; depois qualquer outra aba (exceto lista oficial) com cabeçalho de inscrição.
 * Evita "Protocolo não encontrado" quando a fila está em aba com nome personalizado.
 */
function encontrarProtocoloNasAbasPendentes(ss, protocolo) {
  if (ss == null || typeof ss.getSheetByName !== "function") {
    throw new Error("encontrarProtocoloNasAbasPendentes: planilha (ss) inválida ou não informada.");
  }
  var main = obterAbaInscricoes(ss);
  var p = String(protocolo || "").trim();
  if (!p) return null;

  var listaNomeada = listarAbasPorNomesPendentes(ss);
  var seen = {};
  var i;
  var sh;
  var r;
  for (i = 0; i < listaNomeada.length; i++) {
    sh = listaNomeada[i];
    seen[sh.getSheetId()] = true;
    garantirCabecalhosPlanilha(sh);
    r = encontrarLinhaPorProtocolo(sh, p);
    if (r > 0) return { sheet: sh, rowIndex: r };
  }

  var todas = ss.getSheets();
  for (i = 0; i < todas.length; i++) {
    sh = todas[i];
    if (main && sh.getSheetId() === main.getSheetId()) continue;
    if (seen[sh.getSheetId()]) continue;
    if (!sheetTemCabecalhoInscricao(sh)) continue;
    garantirCabecalhosPlanilha(sh);
    r = encontrarLinhaPorProtocolo(sh, p);
    if (r > 0) return { sheet: sh, rowIndex: r };
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
      for (var c = lc; c < CABECALHOS.length; c++) {
        sheet.getRange(1, c + 1).setValue(CABECALHOS[c]);
      }
      sheet.getRange(1, 1, 1, CABECALHOS.length).setFontWeight("bold");
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
    "",
  ];
}

/** Vírgula ou ponto e vírgula nas fórmulas (HYPERLINK etc.), conforme idioma da planilha. */
function separadorArgumentosFormulas(ss) {
  var loc = String(ss.getSpreadsheetLocale() || "en_US").toLowerCase();
  if (/^pt/.test(loc) || /^(de|at|ch|fr|es|it|nl|pl|cs|ro|be|lu|dk|no|se|fi|ru)/.test(loc)) {
    return ";";
  }
  return ",";
}

/** Mesmas regras da formatação condicional opcional — aplicadas na gravação (pendente = vermelho visível, pago = verde). */
var COR_FUNDO_LINHA_PAGO = "#c8e6c9";
var COR_FUNDO_LINHA_NAO_PAGO = "#ffcdd2";

/**
 * Pinta o fundo da linha (A até última coluna de CABECALHOS) conforme o status de pagamento.
 * @param {string} [statusOuVazio] — se omitido, lê da coluna Status na linha.
 */
function aplicarCorFundoLinhaInscricao(sheet, rowIndex, statusOuVazio) {
  if (!sheet || rowIndex < 2) return;
  garantirCabecalhosPlanilha(sheet);
  var st =
    statusOuVazio != null && String(statusOuVazio) !== ""
      ? String(statusOuVazio).trim()
      : String(sheet.getRange(rowIndex, COL_IX_STATUS + 1).getValue() || "").trim();
  var r = sheet.getRange(rowIndex, 1, rowIndex, CABECALHOS.length);
  var cls = classificarStatusPagamento(st);
  if (cls === "pago") {
    r.setBackground(COR_FUNDO_LINHA_PAGO);
  } else if (cls === "nao_pago") {
    r.setBackground(COR_FUNDO_LINHA_NAO_PAGO);
  } else {
    r.setBackground(null);
  }
}

/** URL do GET de aprovação PIX/presencial (mesma Web App que recebe o POST de inscrição). */
function montarUrlAprovacaoPagamentoPix(protocolo) {
  var props = PropertiesService.getScriptProperties();
  var base = String(props.getProperty("WEB_APP_URL") || WEB_APP_URL_FALLBACK || "")
    .trim()
    .replace(/\?.*$/, "")
    .replace(/\/+$/, "");
  if (!base) return "";
  var senha = obterSenhaAprovacaoConfigurada();
  if (!senha) return "";
  var p = String(protocolo || "").trim();
  if (!p) return "";
  return base + "?protocolo=" + encodeURIComponent(p) + "&senha=" + encodeURIComponent(senha);
}

/** Detecta se a célula já tem link de aprovação (fórmula HYPERLINK ou texto com URL em RichText). */
function celulaJaTemLinkAprovacao(cell) {
  var f = String(cell.getFormula() || "");
  if (f.indexOf("HYPERLINK") !== -1) return true;
  try {
    var rt = cell.getRichTextValue();
    if (rt) {
      var runs = rt.getRuns();
      for (var ri = 0; ri < runs.length; ri++) {
        if (String(runs[ri].getLinkUrl() || "").trim()) return true;
      }
    }
  } catch (errR) {}
  return false;
}

/**
 * Grava link clicável na coluna "Link aprovar PIX" (RichText — funciona em qualquer idioma da planilha).
 * Se WEB_APP_URL ou senha faltar, grava aviso. Fallback: fórmula HYPERLINK, depois URL em texto puro.
 */
function aplicarHyperlinkAprovacaoPixNaLinha(sheet, rowIndex, protocolo) {
  if (!sheet || rowIndex < 2) return;
  garantirCabecalhosPlanilha(sheet);
  var cell = sheet.getRange(rowIndex, COL_IX_LINK_APROVACAO + 1);
  var url = montarUrlAprovacaoPagamentoPix(protocolo);
  if (!url) {
    cell.setValue(
      "Defina WEB_APP_URL e APROVACAO_SENHA nas Propriedades do script (Projeto) e rode reaplicarLinksAprovarPixOndeFaltam()."
    );
    return;
  }
  try {
    cell.clear();
    cell.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText("Aprovar PIX").setLinkUrl(url).build()
    );
  } catch (errRich) {
    try {
      var esc = url.replace(/"/g, '""');
      var sep = separadorArgumentosFormulas(sheet.getParent());
      cell.setFormula('=HYPERLINK("' + esc + '"' + sep + '"Aprovar PIX")');
    } catch (errForm) {
      cell.setValue(url);
    }
  }
  /** Garante fundo da linha (vermelho pendente / verde pago) depois de alterar a célula do link. */
  aplicarCorFundoLinhaInscricao(sheet, rowIndex, "");
}

/**
 * Escreve o link de aprovação com tolerância a falha na Web App (RichText/flush).
 * Se falhar, grava a URL em texto (continua clicável no Planilhas Google).
 */
function aplicarLinkAprovacaoPixGarantido(sheet, rowIndex, protocoloOuVazio) {
  var proto = String(protocoloOuVazio || "").trim();
  if (!proto && sheet && rowIndex >= 2) {
    proto = String(sheet.getRange(rowIndex, COL_IX_PROTOCOLO + 1).getValue() || "").trim();
  }
  try {
    aplicarHyperlinkAprovacaoPixNaLinha(sheet, rowIndex, proto);
  } catch (err1) {
    // segue para URL em texto se a célula continuar sem link
  }
  var url = montarUrlAprovacaoPagamentoPix(proto);
  if (url) {
    var c = sheet.getRange(rowIndex, COL_IX_LINK_APROVACAO + 1);
    if (!celulaJaTemLinkAprovacao(c)) {
      c.setValue(url);
      aplicarCorFundoLinhaInscricao(sheet, rowIndex, "");
    }
  }
  SpreadsheetApp.flush();
}

/**
 * Para linhas ainda não pagas sem fórmula HYPERLINK na coluna do link, tenta gerar o link de aprovação de novo.
 * Útil após configurar WEB_APP_URL / APROVACAO_SENHA ou quando linhas foram criadas sem URL.
 */
function reaplicarLinksAprovarPixOndeFaltam() {
  var ss = obterPlanilhaCorridaOuErro();
  var sheets = ss.getSheets();
  var n = 0;
  var linkCol = COL_IX_LINK_APROVACAO + 1;
  for (var si = 0; si < sheets.length; si++) {
    var sheet = sheets[si];
    if (!sheetTemCabecalhoInscricao(sheet)) continue;
    garantirCabecalhosPlanilha(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var values = sheet.getRange(2, 1, lastRow, CABECALHOS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var proto = String(row[COL_IX_PROTOCOLO] || "").trim();
      if (!proto) continue;
      if (classificarStatusPagamento(row[COL_IX_STATUS]) !== "nao_pago") continue;
      var cLink = sheet.getRange(i + 2, linkCol);
      if (celulaJaTemLinkAprovacao(cLink)) continue;
      aplicarLinkAprovacaoPixGarantido(sheet, i + 2, proto);
      n++;
    }
  }
  Logger.log("reaplicarLinksAprovarPixOndeFaltam: tentativas em linhas (nao pagas sem link): " + n);
  return { ok: true, linhasAtualizadas: n };
}

/**
 * Repinta todas as linhas de inscrição (a partir da coluna "Status pagamento"): vermelho = não pago, verde = pago.
 * Execute no editor após mudar cores ou se linhas antigas ficaram sem cor.
 */
function reaplicarCoresFundoConformeStatus() {
  var ss = obterPlanilhaCorridaOuErro();
  var sheets = ss.getSheets();
  var n = 0;
  for (var si = 0; si < sheets.length; si++) {
    var sheet = sheets[si];
    if (!sheetTemCabecalhoInscricao(sheet)) continue;
    garantirCabecalhosPlanilha(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var values = sheet.getRange(2, 1, lastRow, CABECALHOS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      if (!String(values[i][COL_IX_PROTOCOLO] || "").trim()) continue;
      aplicarCorFundoLinhaInscricao(sheet, i + 2, "");
      n++;
    }
  }
  Logger.log("reaplicarCoresFundoConformeStatus: linhas repintadas: " + n);
  return { ok: true, linhasRepintadas: n };
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

/** Ordem: lista oficial → abas com nome de pendentes → outras abas com cabeçalho de inscrição (ex.: fila PIX). */
function listarAbasParaBuscaInscricao(ss) {
  var main = obterAbaInscricoes(ss);
  var ordered = [];
  var seen = {};
  var i;
  var sh;
  if (main) {
    ordered.push(main);
    seen[main.getSheetId()] = true;
  }
  var nomeadas = listarAbasPorNomesPendentes(ss);
  for (i = 0; i < nomeadas.length; i++) {
    sh = nomeadas[i];
    if (seen[sh.getSheetId()]) continue;
    ordered.push(sh);
    seen[sh.getSheetId()] = true;
  }
  var todas = ss.getSheets();
  for (i = 0; i < todas.length; i++) {
    sh = todas[i];
    if (seen[sh.getSheetId()]) continue;
    if (!sheetTemCabecalhoInscricao(sh)) continue;
    ordered.push(sh);
    seen[sh.getSheetId()] = true;
  }
  return ordered;
}

/**
 * Busca inscrição na lista principal ou em pendentes — exige e-mail + telefone (mesmos da inscrição; telefone comparado só com dígitos).
 * Retorna aba, índice da linha (1-based) e linha de valores para atualização na planilha.
 */
function buscarInscricaoPorEmailETelefoneComLocal(ss, email, telefoneDigitos) {
  var em = String(email || "")
    .trim()
    .toLowerCase();
  var td = soDigitos(telefoneDigitos);
  if (!em || td.length < 10) return null;

  var mainB = obterAbaInscricoes(ss);
  var alvo = listarAbasParaBuscaInscricao(ss);
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
      var origem = mainB && sheet.getSheetId() === mainB.getSheetId() ? "lista_oficial" : "pendente_fila";
      return { sheet: sheet, rowIndex: i + 1, row: row, origem: origem };
    }
  }
  return null;
}

function buscarInscricaoPorEmailETelefone(ss, email, telefoneDigitos) {
  var loc = buscarInscricaoPorEmailETelefoneComLocal(ss, email, telefoneDigitos);
  if (!loc) return null;
  return { row: loc.row, origem: loc.origem };
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
    origemConsulta: origem,
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
  var shPendBackup = listarAbasParaBuscaInscricao(ss);
  for (var bi = 0; bi < shPendBackup.length; bi++) {
    var shBk = shPendBackup[bi];
    if (sheet && shBk.getSheetId() === sheet.getSheetId()) continue;
    garantirCabecalhosPlanilha(shBk);
    var trecho = listaBackupDaAba(shBk);
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

function linhaPlanilhaParaPayloadMp(row, urlRetorno, useSandbox) {
  var vr = row[10];
  if (typeof vr === "string") {
    vr = parseFloat(String(vr).replace(/[^\d.,]/g, "").replace(",", "."));
    if (isNaN(vr)) vr = 0;
  } else {
    vr = Number(vr);
    if (isNaN(vr)) vr = 0;
  }
  return {
    protocolo: String(row[COL_IX_PROTOCOLO] || "").trim(),
    nome: String(row[COL_IX_NOME] || "").trim(),
    email: String(row[COL_IX_EMAIL] || "").trim(),
    telefone: String(row[COL_IX_TELEFONE] || "").trim(),
    cidade: String(row[5] || "").trim(),
    camisa: String(row[6] || "").trim(),
    percurso: String(row[7] || "").trim(),
    lote: String(row[COL_IX_LOTE] || "").trim(),
    loteNome: String(row[9] || "").trim(),
    valorReais: vr,
    evento: NOME_EVENTO_EMAIL,
    mercadoPago: true,
    urlRetorno: urlRetorno,
    useSandbox: useSandbox === true,
  };
}

/**
 * Atualiza forma/status na linha encontrada por e-mail+telefone, com protocolo como confirmação.
 * Só permite enquanto o pagamento não estiver confirmado (mesma regra da consulta).
 */
function executarAlterarFormaPagamento(data) {
  var email = String(data.email || "")
    .trim()
    .toLowerCase();
  var tel = data.telefone != null ? String(data.telefone) : "";
  var protocolo = String(data.protocolo || "").trim();
  var novaForma = String(data.novaForma || "").trim();
  if (!email || !tel || !protocolo) {
    return { ok: false, error: "Informe e-mail, telefone e protocolo." };
  }
  if (novaForma !== "mercado_pago_online" && novaForma !== "presencial_secretaria") {
    return { ok: false, error: "Forma de pagamento inválida." };
  }
  if (soDigitos(tel).length < 10) {
    return { ok: false, error: "Telefone inválido: use DDD + número." };
  }

  var ss;
  try {
    ss = obterPlanilhaCorridaOuErro();
  } catch (err0) {
    return { ok: false, error: String(err0) };
  }

  var achado = buscarInscricaoPorEmailETelefoneComLocal(ss, email, tel);
  if (!achado) {
    return { ok: false, error: "Inscrição não encontrada com este e-mail e telefone." };
  }

  var protoLinha = String(achado.row[COL_IX_PROTOCOLO] || "").trim();
  if (!protoLinha || protoLinha !== protocolo) {
    return { ok: false, error: "O protocolo não confere com o e-mail e telefone informados." };
  }

  var stAtual = String(achado.row[COL_IX_STATUS] || "").trim();
  if (!inscricaoPermiteAlterarFormaPagamento(stAtual, achado.origem)) {
    return {
      ok: false,
      error:
        "Não é possível alterar: o pagamento já foi confirmado ou o status não permite mudança automática pela consulta.",
    };
  }

  var codAtual = inferirCodigoFormaPagamentoLinha(achado.row);
  if (codAtual && novaForma === codAtual) {
    return { ok: false, error: "Essa já é a forma de pagamento registrada na sua inscrição." };
  }

  var sheet = achado.sheet;
  var rowIndex = achado.rowIndex;
  var urlRetorno = String(data.urlRetorno || "").trim();
  var useSandbox = data.useSandbox === true || String(data.useSandbox || "").toLowerCase() === "true";

  if (novaForma === "presencial_secretaria") {
    sheet.getRange(rowIndex, COL_IX_FORMA + 1).setValue("PIX via WhatsApp");
    sheet.getRange(rowIndex, COL_IX_STATUS + 1).setValue("Pendente (PIX via WhatsApp)");
    aplicarLinkAprovacaoPixGarantido(sheet, rowIndex, protocolo);
    aplicarCorFundoLinhaInscricao(sheet, rowIndex, "Pendente (PIX via WhatsApp)");
    sincronizarBackupSegurancaNoDrive();
    return {
      ok: true,
      mensagem:
        "Forma de pagamento atualizada para PIX via WhatsApp. Fale com a organização para receber os dados do PIX.",
    };
  }

  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("MERCADO_PAGO_ACCESS_TOKEN");
  if (!token) {
    return { ok: false, error: "Pagamento online indisponível no servidor (configure MERCADO_PAGO_ACCESS_TOKEN)." };
  }
  if (!urlRetorno) {
    return {
      ok: false,
      error:
        "Falta a URL de retorno do site. Em config.js defina mercadoPago.urlRetorno (HTTPS) e tente de novo.",
    };
  }

  var payloadMp = linhaPlanilhaParaPayloadMp(achado.row, urlRetorno, useSandbox);
  var mpRes;
  try {
    mpRes = criarPreferenciaMercadoPago(payloadMp);
  } catch (mpErr) {
    return { ok: false, error: "Erro ao falar com o Mercado Pago: " + String(mpErr) };
  }
  if (!mpRes.url) {
    return {
      ok: false,
      error: mensagemErroCheckoutMercadoPago(mpRes.erroCodigo) || "Não foi possível gerar o link de pagamento online.",
    };
  }

  sheet.getRange(rowIndex, COL_IX_FORMA + 1).setValue("Mercado Pago (pagamento online)");
  sheet.getRange(rowIndex, COL_IX_STATUS + 1).setValue("Aguardando pagamento online");
  aplicarLinkAprovacaoPixGarantido(sheet, rowIndex, protocolo);
  aplicarCorFundoLinhaInscricao(sheet, rowIndex, "Aguardando pagamento online");
  sincronizarBackupSegurancaNoDrive();

  return {
    ok: true,
    checkoutUrl: mpRes.url,
    mensagem: "Forma de pagamento atualizada. Use o link do Mercado Pago para concluir o pagamento com cartão.",
  };
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
  var achado = buscarInscricaoPorEmailETelefoneComLocal(ss, email, tel);
  if (!achado) {
    return {
      ok: true,
      encontrado: false,
      error:
        "Não encontramos inscrição com este e-mail e telefone. Confira se são os mesmos do formulário ou fale com a organização.",
    };
  }
  var dados = linhaParaRespostaConsulta(achado.row, achado.origem);
  dados.permiteAlterarFormaPagamento = inscricaoPermiteAlterarFormaPagamento(
    String(achado.row[COL_IX_STATUS] || "").trim(),
    achado.origem
  );
  dados.formaPagamentoCodigo = inferirCodigoFormaPagamentoLinha(achado.row);
  return {
    ok: true,
    encontrado: true,
    dados: dados,
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
        chipLabel: "Mercado Pago",
        chipCor: "#009ee3",
        txt:
          "Meio: pagamento online (Mercado Pago).\n\n" +
          "Seu pagamento pelo Mercado Pago foi confirmado e sua inscrição está na lista oficial do evento.",
        fraseHtml:
          "Seu pagamento pelo <strong>Mercado Pago</strong> foi <strong>confirmado</strong>. Sua vaga está <strong>garantida</strong> na lista oficial.",
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
 * payJsonOpcional: objeto do pagamento MP (GET /v1/payments) para personalizar meio (cartão, boleto, etc.).
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
  rowData[COL_IX_STATUS] = "Pago (Mercado Pago)";
  rowData[COL_IX_LINK_APROVACAO] = "";

  if (jaTemCadastro(main, rowData[COL_IX_EMAIL], rowData[COL_IX_TELEFONE])) {
    pend.deleteRow(rowIndex);
    sincronizarBackupSegurancaNoDrive();
    return;
  }

  main.appendRow(rowData);
  aplicarCorFundoLinhaInscricao(main, main.getLastRow(), "Pago (Mercado Pago)");
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
    /** Checkout só com cartão: sem PIX e sem boleto (ticket) na preferência. */
    payment_methods: {
      excluded_payment_types: [{ id: "ticket" }],
      excluded_payment_methods: [{ id: "pix" }],
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
      aplicarCorFundoLinhaInscricao(pend, pend.getLastRow(), "");
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
    var stMainLower = String(statusMain || "").toLowerCase();
    var novoStatusMain =
      stMainLower.indexOf("aguardando pagamento online") !== -1
        ? "Pago (Mercado Pago)"
        : "Pago (presencial confirmado)";
    main.getRange(rowMain, COL_IX_STATUS + 1).setValue(novoStatusMain);
    aplicarCorFundoLinhaInscricao(main, rowMain, novoStatusMain);
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
  var stPendLower = statusPend.toLowerCase();
  var novoStatus =
    stPendLower.indexOf("aguardando pagamento online") !== -1
      ? "Pago (Mercado Pago)"
      : "Pago (presencial confirmado)";
  rowData[COL_IX_STATUS] = novoStatus;
  rowData[COL_IX_LINK_APROVACAO] = "";

  if (!jaTemCadastro(main, rowData[COL_IX_EMAIL], rowData[COL_IX_TELEFONE])) {
    main.appendRow(rowData);
    aplicarCorFundoLinhaInscricao(main, main.getLastRow(), novoStatus);
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

    if (data.tipo === "alterar_forma_pagamento") {
      var outAlt = executarAlterarFormaPagamento(data);
      return ContentService.createTextOutput(JSON.stringify(outAlt)).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.tipo === "estado_lotes") {
      try {
        var ssLotes = obterPlanilhaCorridaOuErro();
        var pPromo = contarInscricoesLoteTotal(ssLotes, ID_LOTE_PROMO);
        var pReg = contarInscricoesLoteTotal(ssLotes, ID_LOTE_REGULAR);
        return ContentService.createTextOutput(
          JSON.stringify({
            ok: true,
            limitePromo: LIMITE_LOTE_PROMO,
            limiteRegular: LIMITE_LOTE_REGULAR,
            promoPagos: pPromo,
            regularPagos: pReg,
            promoEsgotado: pPromo >= LIMITE_LOTE_PROMO,
            regularEsgotado: pReg >= LIMITE_LOTE_REGULAR,
          })
        ).setMimeType(ContentService.MimeType.JSON);
      } catch (errL) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(errL) })).setMimeType(
          ContentService.MimeType.JSON
        );
      }
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
        aplicarLinkAprovacaoPixGarantido(pend, pend.getLastRow(), data.protocolo);
        aplicarCorFundoLinhaInscricao(pend, pend.getLastRow(), "Aguardando pagamento online");
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
      aplicarLinkAprovacaoPixGarantido(pendPresencial, pendPresencial.getLastRow(), data.protocolo);
      aplicarCorFundoLinhaInscricao(pendPresencial, pendPresencial.getLastRow(), "Pendente (PIX via WhatsApp)");
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

function normalizarParametroUrlGet_(valor) {
  var s = String(valor != null ? valor : "").trim();
  if (!s) return "";
  try {
    return decodeURIComponent(s.replace(/\+/g, " "));
  } catch (err1) {
    return s;
  }
}

function obterParametroGet_(e, chaves) {
  if (!e || !e.parameter) return "";
  var p = e.parameter;
  for (var i = 0; i < chaves.length; i++) {
    var k = chaves[i];
    if (p[k] == null) continue;
    var v = normalizarParametroUrlGet_(p[k]);
    if (v) return v;
  }
  return "";
}

function doGet(e) {
  var protocoloQs = obterParametroGet_(e, ["protocolo", "Protocolo", "PROTOCOLO"]);
  var senhaQs = obterParametroGet_(e, ["senha", "Senha", "SENHA", "password"]);
  if (protocoloQs && senhaQs) {
    var outManualQs = confirmarPagamentoPresencialPorProtocolo(protocoloQs, senhaQs);
    return respostaTexto(montarMensagemLeigaAtualizacaoPagamento(outManualQs));
  }

  var path = e && e.pathInfo != null ? String(e.pathInfo) : "";
  path = path.replace(/^\/+|\/+$/g, "");
  if (path) {
    var parts = path.split("/");
    if (parts.length === 2) {
      var outManual = confirmarPagamentoPresencialPorProtocolo(
        normalizarParametroUrlGet_(parts[0]),
        normalizarParametroUrlGet_(parts[1])
      );
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

/**
 * Converte índice de coluna 1-based (1=A) para letra(s), ex.: 13 → "M".
 */
function colunaLetra1Based(col1Based) {
  var t = "";
  var n = col1Based;
  while (n > 0) {
    var r = (n - 1) % 26;
    t = String.fromCharCode(65 + r) + t;
    n = Math.floor((n - 1) / 26);
  }
  return t;
}

/**
 * Verde / vermelho na FC: evita REGEXMATCH (falhava só no vermelho em algumas planilhas).
 * pt_BR: ESQUERDA+MINÚSCULA, LOCALIZAR+SEERRO. Demais: LEFT+LOWER, SEARCH+IFERROR. Argumentos: ";" ou "," conforme o idioma.
 */
function formulasFormatacaoCondicionalStatus(letraStatus, ss) {
  var loc = String(ss.getSpreadsheetLocale() || "en_US").toLowerCase();
  var isPt = /^pt/.test(loc);
  var useSemi =
    isPt || /^(de|at|ch|fr|es|it|nl|pl|cs|ro|be|lu|dk|no|se|fi|ru)/.test(loc);
  var sep = useSemi ? ";" : ",";
  var c = "$" + letraStatus + "2";
  if (isPt) {
    return {
      fVerde: "=ESQUERDA(MINÚSCULA(" + c + ")" + sep + '4)="pago"',
      fVermelhoPend: "=SEERRO(LOCALIZAR(\"pendente\"" + sep + c + ");0)>0",
      fVermelhoAg: "=SEERRO(LOCALIZAR(\"aguardando\"" + sep + c + ");0)>0",
    };
  }
  return {
    fVerde: "=LEFT(LOWER(" + c + ")" + sep + '4)="pago"',
    fVermelhoPend: "=IFERROR(SEARCH(\"pendente\"" + sep + c + ");0)>0",
    fVermelhoAg: "=IFERROR(SEARCH(\"aguardando\"" + sep + c + ");0)>0",
  };
}

/**
 * Formatação condicional na aba: fundo vermelho claro = ainda não pago; verde claro = pago.
 * Usa a coluna "Status pagamento" (COL_IX_STATUS). Execute no editor: instalarFormatacaoCoresStatusPlanilha()
 * Executar **uma vez** por aba (ou apague regras duplicadas em Formatar → Formatação condicional).
 */
function aplicarFormatacaoCoresNaAba(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  garantirCabecalhosPlanilha(sheet);
  var ss = sheet.getParent();
  var colSt = COL_IX_STATUS + 1;
  var letra = colunaLetra1Based(colSt);
  var numCols = CABECALHOS.length;
  var endRow = Math.min(10000, Math.max(sheet.getMaxRows(), 500));
  var rangeRows = sheet.getRange(2, 1, endRow, numCols);

  var F = formulasFormatacaoCondicionalStatus(letra, ss);

  var verde = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(F.fVerde)
    .setBackground("#d7eeda")
    .setRanges([rangeRows])
    .build();

  var vermelhoPend = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(F.fVermelhoPend)
    .setBackground("#ffe3e3")
    .setRanges([rangeRows])
    .build();

  var vermelhoAg = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(F.fVermelhoAg)
    .setBackground("#ffe3e3")
    .setRanges([rangeRows])
    .build();

  var regras = sheet.getConditionalFormatRules();
  regras.push(vermelhoPend);
  regras.push(vermelhoAg);
  regras.push(verde);
  sheet.setConditionalFormatRules(regras);
  Logger.log("Formatacao condicional adicionada na aba: " + sheet.getName());
}

/** Instala cores em todas as abas com formato de inscrição (lista oficial + filas de pendentes). */
function instalarFormatacaoCoresStatusPlanilha() {
  var ss = obterPlanilhaCorridaOuErro();
  var abas = listarAbasParaBuscaInscricao(ss);
  for (var i = 0; i < abas.length; i++) {
    aplicarFormatacaoCoresNaAba(abas[i]);
  }
}
