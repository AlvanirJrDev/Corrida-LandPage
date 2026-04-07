/**
 * Google Apps Script — inscrições + Mercado Pago Checkout Pro
 *
 * Erro "UrlFetchApp.fetch sem permissão": no editor, Executar → autorizarAcessoExterno → aceitar tudo.
 * Use appsscript.json SEM lista oauthScopes (o Google pede os escopos na 1ª execução). Implantação: Executar como: Eu.
 * E-mail após pagamento: MailApp na confirmação do pagamento — na 1ª vez o Apps Script pede permissão de envio.
 *
 * Propriedades do script (Projeto ⚙ → Propriedades do projeto → Propriedades do script):
 *   MERCADO_PAGO_ACCESS_TOKEN = Access Token (mercadopago.com.br/developers)
 *   WEB_APP_URL (opcional) = mesma URL do webhookUrl em config.js. Se não definir, usa WEB_APP_URL_FALLBACK no código abaixo.
 *
 * Compartilhamento da planilha (Google Drive): quem tiver permissão de EDITOR pode ver, alterar e apagar
 * todas as linhas. Não use "Qualquer pessoa com o link pode editar" para o público. Restrinja a
 * organizadores (e-mail) ou use "Somente visualização" para quem só precisa consultar.
 *
 * ID da planilha (URL .../spreadsheets/d/ESTE_ID/edit)
 */
var ID_PLANILHA = "1BLVaZLh3Dq64WvUoQ2_XhPXgDAB-uOkW0t-Gek9gaKc";

/** Alinhe com config.js → nomeEvento (texto do e-mail de confirmação). */
var NOME_EVENTO_EMAIL = "2ª Corrida Sagrado Coração de Jesus";

/** Mesma URL que config.js → webhookUrl (termina em /exec). Alinhe os dois se mudar a implantação. */
var WEB_APP_URL_FALLBACK =
  "https://script.google.com/a/macros/redealia.com/s/AKfycbwClwfOb9G4AXWFW0tjQAMAkuX_VlmlBHJC6nFPuGczHZZBM0XLv4p36mYF0RvvHCu7/exec";

/** Aba principal: inscrições confirmadas (presencial) ou pagamento MP já aprovado */
/** Aba "Inscrições pendentes MP": só pagamento online antes de confirmar no MP */
var NOME_ABA_PENDENTES = "Inscrições pendentes MP";

/** Limite lote promocional (50): alinhar com config.js */
var LIMITE_LOTE_PROMO = 50;
var ID_LOTE_PROMO = "promo";
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

function soDigitos(s) {
  return String(s || "").replace(/\D/g, "");
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

function contarInscricoesLotePromo(sheet) {
  var values = sheet.getDataRange().getValues();
  var start = 0;
  if (values.length > 0 && String(values[0][1]) === "Protocolo") {
    start = 1;
  }
  var n = 0;
  for (var i = start; i < values.length; i++) {
    if (String(values[i][COL_IX_LOTE]).trim() === ID_LOTE_PROMO) {
      n++;
    }
  }
  return n;
}

function contarInscricoesLotePromoTotal(ss) {
  var n = contarInscricoesLotePromo(obterAbaInscricoes(ss));
  var pend = obterAbaPendentes(ss);
  if (pend.getLastRow() > 0) {
    garantirCabecalhosPlanilha(pend);
    n += contarInscricoesLotePromo(pend);
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

/**
 * Envia e-mail ao inscrito quando o pagamento MP é confirmado (webhook).
 * 1ª vez: o Apps Script pedirá permissão para enviar e-mail. Limite diário do Gmail se aplicam.
 */
function enviarEmailPagamentoConfirmado(rowData) {
  var email = String(rowData[COL_IX_EMAIL] || "").trim();
  var nome = String(rowData[COL_IX_NOME] || "").trim();
  var proto = String(rowData[COL_IX_PROTOCOLO] || "").trim();
  if (!email || email.indexOf("@") < 0) return;
  try {
    MailApp.sendEmail({
      to: email,
      subject: "Pagamento confirmado — " + NOME_EVENTO_EMAIL,
      body:
        "Olá" +
        (nome ? " " + nome : "") +
        ",\n\n" +
        "Seu pagamento foi confirmado pelo Mercado Pago e sua inscrição está na lista oficial.\n\n" +
        "Protocolo: " +
        proto +
        "\n\n" +
        "Guarde esse número. Em caso de dúvida, fale com a organização.\n\n" +
        "— Organização",
    });
  } catch (err) {
    Logger.log("enviarEmailPagamentoConfirmado: " + err);
  }
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

  var rowData = pend.getRange(rowIndex, 1, rowIndex, CABECALHOS.length).getValues()[0];
  while (rowData.length < CABECALHOS.length) {
    rowData.push("");
  }
  rowData[CABECALHOS.length - 1] = "Pago (Mercado Pago)";

  if (jaTemCadastro(main, rowData[COL_IX_EMAIL], rowData[COL_IX_TELEFONE])) {
    pend.deleteRow(rowIndex);
    return;
  }

  main.appendRow(rowData);
  pend.deleteRow(rowIndex);
  enviarEmailPagamentoConfirmado(rowData);
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

    var dup = jaTemCadastroQualquerAba(ss, data.email, data.telefoneDigitos || data.telefone);
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
      var ja = contarInscricoesLotePromoTotal(ss);
      if (ja >= LIMITE_LOTE_PROMO) {
        return ContentService.createTextOutput(
          JSON.stringify({
            ok: false,
            error: "Lote promocional esgotado (limite de " + LIMITE_LOTE_PROMO + " inscrições).",
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
      } else {
        if (pend.getLastRow() > 0) {
          pend.deleteRow(pend.getLastRow());
        }
        out.checkoutFalhou = true;
        out.erroCheckout = mensagemErroCheckoutMercadoPago(mpRes.erroCodigo);
      }
    } else {
      var row = montarLinhaInscricao(data, "Pendente (presencial)");
      sheet.appendRow(row);
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
  return ContentService.createTextOutput("Use POST JSON para inscrições.").setMimeType(ContentService.MimeType.TEXT);
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
