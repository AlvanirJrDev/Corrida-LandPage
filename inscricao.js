(function () {
  "use strict";

  var cfg = window.CORRIDA_CONFIG || {};
  var form = document.getElementById("form-inscricao");
  if (!form) return;

  var mpRetornoMsg = document.getElementById("mp-retorno-msg");
  var mpLoadingOverlay = document.getElementById("inscricao-mp-loading");
  var qsMp = window.location.search || "";
  if (mpRetornoMsg && (qsMp.indexOf("mp=ok") !== -1 || qsMp.indexOf("mp=pendente") !== -1 || qsMp.indexOf("mp=erro") !== -1)) {
    mpRetornoMsg.hidden = false;
    if (qsMp.indexOf("mp=ok") !== -1) {
      mpRetornoMsg.textContent =
        "Você voltou ao site. Quando o Mercado Pago confirmar o pagamento, sua inscrição passa para a lista oficial e você recebe um e-mail (se o sistema já estiver autorizado a enviar).";
    } else if (qsMp.indexOf("mp=pendente") !== -1) {
      mpRetornoMsg.textContent = "Pagamento em análise. Quando for confirmado, sua inscrição será atualizada.";
    } else {
      mpRetornoMsg.textContent = "O pagamento não foi concluído. Você pode tentar de novo ou escolher pagamento na secretaria.";
    }
    var secInsc = document.getElementById("inscricao");
    if (secInsc) {
      setTimeout(function () {
        secInsc.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 200);
    }
  }

  function showMpLoading() {
    if (!mpLoadingOverlay) return;
    mpLoadingOverlay.hidden = false;
    mpLoadingOverlay.setAttribute("aria-hidden", "false");
  }

  function hideMpLoading() {
    if (!mpLoadingOverlay) return;
    mpLoadingOverlay.hidden = true;
    mpLoadingOverlay.setAttribute("aria-hidden", "true");
  }

  function setFormSubmitLoading(btn, loading) {
    if (!btn) return;
    var labelEl = btn.querySelector(".form-submit__label");
    if (loading) {
      if (!btn.dataset.labelOriginal) {
        btn.dataset.labelOriginal = labelEl ? labelEl.textContent.trim() : btn.textContent.trim();
      }
      btn.classList.add("form-submit--loading");
      btn.disabled = true;
      if (labelEl) labelEl.textContent = "Enviando…";
    } else {
      btn.classList.remove("form-submit--loading");
      btn.disabled = false;
      if (labelEl) labelEl.textContent = btn.dataset.labelOriginal || "Confirmar inscrição";
    }
  }

  var statusEl = document.getElementById("inscricao-status");
  var successPanel = document.getElementById("inscricao-sucesso");
  var webhookAviso = document.getElementById("webhook-aviso");
  var confirmacaoPendenteTexto = document.getElementById("inscricao-confirmacao-pendente-texto");
  var btnWa = document.getElementById("btn-wa-comprovante");
  var protocoloEl = document.getElementById("protocolo-numero");
  var copyPixBtn = document.getElementById("btn-copy-pix");

  function gerarProtocolo() {
    var y = new Date().getFullYear();
    var rand = Math.random().toString(36).substring(2, 8).toUpperCase();
    return "INSC-" + y + "-" + rand;
  }

  function formatarMoedaBR(num) {
    var n = Number(num);
    if (isNaN(n)) return "R$ 0,00";
    return "R$ " + n.toFixed(2).replace(".", ",");
  }

  function lotesAtivos() {
    var L = cfg.lotes;
    if (!L || !L.length) return [];
    return L.filter(function (l) {
      return l.ativo !== false;
    });
  }

  function acharLote(id) {
    var list = cfg.lotes || [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === id) return list[i];
    }
    return null;
  }

  function montarMensagemWhatsApp(dados) {
    var nomeEvento = cfg.nomeEvento || "Corrida Mariana em prol do ECC e EJC de Sanharó";
    var valor = dados.valorFormatado || cfg.valorInscricao || "";
    var linhas = [
      "Olá! Segue minha inscrição na *" + nomeEvento + "*.",
      "",
      "*Protocolo:* " + dados.protocolo,
      "*Nome:* " + dados.nome,
      "*E-mail:* " + dados.email,
      "*Telefone:* " + dados.telefone,
      "*Cidade:* " + dados.cidade,
      "*Tam. camisa:* " + dados.camisa,
      "*Percurso:* " + dados.percurso,
      "*Lote:* " + (dados.loteNome || dados.lote || "—"),
      "*Forma de pagamento:* " + dados.formaPagamento,
      "",
      "Valor da inscrição: " + valor,
      "",
    ];
    if (dados.formaPagamentoCodigo === "mercado_pago_online") {
      linhas.push("Vou concluir o pagamento pelo Mercado Pago (pagamento online). Se precisar, envio o comprovante em anexo.");
    } else {
      linhas.push("Quero finalizar minha inscrição. Peço que me enviem os dados do PIX por aqui para eu realizar o pagamento.");
    }
    return linhas.join("\n");
  }

  function urlWhatsApp(texto) {
    var num = (cfg.whatsappNumero || "").replace(/\D/g, "");
    if (!num) num = "5511999999999";
    return "https://wa.me/" + num + "?text=" + encodeURIComponent(texto);
  }

  function abrirWhatsAppAposInscricao(urlWa) {
    if (!urlWa) return;
    var aba = window.open(urlWa, "_blank", "noopener,noreferrer");
    if (!aba && statusEl) {
      statusEl.hidden = false;
      statusEl.className = "form-status form-status--error";
      statusEl.textContent =
        "Seu navegador bloqueou a abertura do WhatsApp. Toque no botão 'Falar no WhatsApp da organização' para continuar.";
    }
  }

  function redirecionarDiretoParaWhatsApp(urlWa) {
    if (!urlWa) return;
    window.location.href = urlWa;
  }

  function apenasDigitos(str) {
    return String(str || "").replace(/\D/g, "");
  }

  function formatarTelefoneBR(valor) {
    var d = apenasDigitos(valor).slice(0, 11);
    if (d.length === 0) return "";
    if (d.length <= 2) return "(" + d;
    if (d.length <= 6) return "(" + d.slice(0, 2) + ") " + d.slice(2);
    if (d.length <= 10) return "(" + d.slice(0, 2) + ") " + d.slice(2, 6) + "-" + d.slice(6);
    return "(" + d.slice(0, 2) + ") " + d.slice(2, 7) + "-" + d.slice(7);
  }

  function labelFormaPagamento(codigo) {
    if (codigo === "mercado_pago_online") return "Mercado Pago (pagamento online)";
    if (codigo === "presencial_secretaria") return "PIX via WhatsApp";
    return String(codigo || "").trim() || "—";
  }

  function coletarDados(formEl) {
    var fd = new FormData(formEl);
    var fpCod = (fd.get("forma_pagamento") || "").trim();
    var loteId = (fd.get("lote") || "").trim();
    var loteCfg = acharLote(loteId);
    var valorReais = loteCfg ? Number(loteCfg.valorReais) : NaN;
    if (lotesAtivos().length === 0) {
      var fallback = cfg.valorInscricao || "0";
      var m = String(fallback).replace(/[^\d,]/g, "").replace(",", ".");
      valorReais = parseFloat(m) || 0;
      loteCfg = { nome: "Inscrição", valorReais: valorReais };
      if (!loteId) loteId = "unico";
    }
    if (isNaN(valorReais) && loteCfg) valorReais = Number(loteCfg.valorReais) || 0;
    if (isNaN(valorReais)) valorReais = 0;
    return {
      nome: (fd.get("nome") || "").trim().replace(/\s+/g, " "),
      email: (fd.get("email") || "").trim().toLowerCase(),
      telefone: (fd.get("telefone") || "").trim(),
      cidade: (fd.get("cidade") || "").trim().replace(/\s+/g, " "),
      camisa: fd.get("camisa") || "",
      percurso: fd.get("percurso") || "",
      lote: loteId,
      loteNome: loteCfg ? loteCfg.nome : "",
      valorReais: valorReais,
      valorFormatado: formatarMoedaBR(valorReais),
      formaPagamento: labelFormaPagamento(fpCod),
      formaPagamentoCodigo: fpCod,
      statusPagamento: "Pendente",
      protocolo: formEl.dataset.protocolo || gerarProtocolo(),
      criadoEm: new Date().toISOString(),
    };
  }

  function popularSelectLotes() {
    var sel = document.getElementById("select-lote");
    if (!sel) return;
    sel.innerHTML = "";
    var list = lotesAtivos();
    if (list.length === 0) {
      var fallback = cfg.valorInscricao || "R$ 0,00";
      var opt = document.createElement("option");
      opt.value = "unico";
      opt.textContent = "Inscrição — " + fallback;
      sel.appendChild(opt);
      sel.selectedIndex = 0;
      atualizarUiLote();
      return;
    }
    var ph = document.createElement("option");
    ph.value = "";
    ph.disabled = true;
    ph.selected = true;
    ph.textContent = "Selecione o lote";
    sel.appendChild(ph);
    list.forEach(function (l) {
      var o = document.createElement("option");
      o.value = l.id;
      o.textContent = l.nome + " — " + formatarMoedaBR(l.valorReais);
      sel.appendChild(o);
    });
    sel.disabled = false;
    sel.selectedIndex = 1;
    atualizarUiLote();
  }

  function atualizarUiLote() {
    var sel = document.getElementById("select-lote");
    var hint = document.getElementById("lote-descricao");
    var asideNome = document.getElementById("aside-lote-nome");
    if (!sel) return;
    var id = sel.value;
    var l = acharLote(id);
    if (hint) {
      var h = l && l.descricao ? l.descricao : "";
      if (l && l.limite > 0) {
        h = (h ? h + " " : "") + "Máximo " + l.limite + " inscrições neste lote.";
      }
      hint.textContent = h;
    }
    if (asideNome) {
      if (l && l.nome) {
        asideNome.textContent = "Lote: " + l.nome;
        asideNome.hidden = false;
      } else {
        asideNome.textContent = "";
        asideNome.hidden = true;
      }
    }
    var valor = l ? formatarMoedaBR(l.valorReais) : cfg.valorInscricao || "R$ 0,00";
    document.querySelectorAll("[data-valor-inscricao]").forEach(function (el) {
      el.textContent = valor;
    });
    var menorValor = valor;
    var ativos = lotesAtivos();
    if (ativos.length) {
      var menor = Number(ativos[0].valorReais);
      for (var i = 1; i < ativos.length; i++) {
        var v = Number(ativos[i].valorReais);
        if (!isNaN(v) && (isNaN(menor) || v < menor)) menor = v;
      }
      if (!isNaN(menor)) menorValor = formatarMoedaBR(menor);
    }
    document.querySelectorAll("[data-hero-preco]").forEach(function (el) {
      el.textContent = menorValor;
    });
  }

  var selectLote = document.getElementById("select-lote");
  if (selectLote) {
    selectLote.addEventListener("change", atualizarUiLote);
  }
  popularSelectLotes();

  function atualizarPainelPagamento() {
    var sel = document.getElementById("select-forma-pagamento");
    var blocoOnline = document.getElementById("aside-bloco-online");
    var blocoPres = document.getElementById("aside-bloco-presencial");
    var passoFinal = document.getElementById("aside-passo-final");
    var mpCfg = cfg.mercadoPago || {};
    if (!sel) return;
    var v = sel.value;
    var mpOk = mpCfg.ativo;
    if (blocoOnline) blocoOnline.hidden = v !== "mercado_pago_online" || !mpOk;
    if (blocoPres) blocoPres.hidden = v !== "presencial_secretaria";
    if (passoFinal) {
      if (v === "mercado_pago_online" && mpOk) {
        passoFinal.textContent =
          "Ao enviar, você será levado ao Mercado Pago para concluir o pagamento online.";
      } else if (v === "presencial_secretaria") {
        passoFinal.textContent =
          "Depois, abra o WhatsApp com a mensagem pronta. A equipe vai enviar os dados do PIX por lá.";
      } else {
        passoFinal.textContent =
          "Escolha no formulário como vai pagar — as instruções aparecem aqui em cima.";
      }
    }
  }

  function configurarFormaPagamento() {
    var sel = document.getElementById("select-forma-pagamento");
    if (!sel) return;
    var mpCfg = cfg.mercadoPago || {};
    var optMp = sel.querySelector('option[value="mercado_pago_online"]');
    if (optMp && !mpCfg.ativo) {
      optMp.remove();
    }
    if (!sel.value && sel.querySelector('option[value="presencial_secretaria"]')) {
      sel.value = "presencial_secretaria";
    }
    atualizarPainelPagamento();
  }

  var selectForma = document.getElementById("select-forma-pagamento");
  if (selectForma) {
    selectForma.addEventListener("change", atualizarPainelPagamento);
  }
  configurarFormaPagamento();

  var inputTel = document.getElementById("input-telefone");
  if (inputTel) {
    inputTel.addEventListener("input", function () {
      inputTel.value = formatarTelefoneBR(inputTel.value);
      var len = inputTel.value.length;
      inputTel.setSelectionRange(len, len);
    });
    inputTel.addEventListener("blur", function () {
      inputTel.value = formatarTelefoneBR(inputTel.value);
    });
  }

  var inputEmail = document.getElementById("input-email");
  if (inputEmail) {
    inputEmail.addEventListener("blur", function () {
      inputEmail.value = inputEmail.value.trim().toLowerCase();
    });
  }

  var inputNome = document.getElementById("input-nome");
  if (inputNome) {
    inputNome.addEventListener("blur", function () {
      inputNome.value = inputNome.value.trim().replace(/\s+/g, " ");
    });
  }

  var inputCidade = document.getElementById("input-cidade");
  if (inputCidade) {
    inputCidade.addEventListener("blur", function () {
      inputCidade.value = inputCidade.value.trim().replace(/\s+/g, " ");
    });
  }

  function limparErro() {
    if (!statusEl) return;
    statusEl.hidden = true;
    statusEl.textContent = "";
  }

  function isErroLotePromoEsgotado(msg) {
    var m = String(msg || "").toLowerCase();
    return m.indexOf("lote promocional esgotado") !== -1;
  }

  function tentarMigrarParaLoteRegular(dados, payload) {
    if (String(dados.lote || "").trim() !== "promo") return false;
    var loteRegular = acharLote("regular");
    if (!loteRegular) return false;

    var selectLoteEl = document.getElementById("select-lote");
    if (selectLoteEl) {
      selectLoteEl.value = "regular";
      atualizarUiLote();
    }

    dados.lote = "regular";
    dados.loteNome = loteRegular.nome || "Lote regular";
    dados.valorReais = Number(loteRegular.valorReais) || 0;
    dados.valorFormatado = formatarMoedaBR(dados.valorReais);

    payload.lote = dados.lote;
    payload.loteNome = dados.loteNome;
    payload.valorReais = dados.valorReais;
    return true;
  }

  /** Quando o servidor reclassifica promo → regular num único POST (loteAjustadoAutomatico). */
  function aplicarAjusteLoteDoServidor(dados, payload, result) {
    if (!result || !result.loteAjustadoAutomatico) return false;
    if (result.lote) dados.lote = String(result.lote).trim();
    if (result.loteNome) dados.loteNome = String(result.loteNome).trim();
    if (result.valorReais !== undefined && result.valorReais !== null && !isNaN(Number(result.valorReais))) {
      dados.valorReais = Number(result.valorReais);
      dados.valorFormatado = formatarMoedaBR(dados.valorReais);
    }
    payload.lote = dados.lote;
    payload.loteNome = dados.loteNome;
    payload.valorReais = dados.valorReais;
    var sel = document.getElementById("select-lote");
    if (sel && dados.lote) {
      sel.value = dados.lote;
      atualizarUiLote();
    }
    return true;
  }

  async function enviarWebhook(payload) {
    var url = (cfg.webhookUrl || "").trim();
    if (!url) {
      if (payload.mercadoPago === true) {
        return {
          ok: false,
          error:
            "Para abrir o Mercado Pago é preciso da URL do Apps Script em config.js (webhookUrl). Sem isso só é possível inscrição com pagamento na secretaria.",
        };
      }
      return { ok: true, skipped: true };
    }

    var esperaCheckout = payload.mercadoPago === true;
    var body = JSON.stringify(payload);
    /** text/plain evita preflight CORS que costuma falhar com o Apps Script; o script lê o JSON em postData.contents. */
    var headersPlain = { "Content-Type": "text/plain;charset=utf-8" };

    try {
      var res = await fetch(url, {
        method: "POST",
        headers: headersPlain,
        body: body,
        mode: "cors",
        credentials: "omit",
      });
      var text = (await res.text()).trim().replace(/^\uFEFF/, "");
      var json = null;
      try {
        json = JSON.parse(text);
      } catch (parseErr) {
        json = null;
      }
      if (json && json.ok === false) {
        return { ok: false, error: json.error || "Inscrição não aceita no servidor." };
      }
      if (!res.ok) return { ok: false, error: "Servidor respondeu " + res.status };
      if (!json) {
        return {
          ok: false,
          error:
            "Resposta do servidor não é JSON válido (pode ser página de erro do Google). No Apps Script: Implantar → Gerenciar implantações → Nova versão → Implantar, e confira a URL em config.js.",
        };
      }
      var checkoutUrl = json.checkoutUrl ? String(json.checkoutUrl).trim() : null;
      var checkoutFalhou = !!json.checkoutFalhou;
      var erroCheckout = json.erroCheckout ? String(json.erroCheckout) : null;
      if (esperaCheckout && !checkoutUrl && !checkoutFalhou) {
        checkoutFalhou = true;
        erroCheckout =
          erroCheckout ||
          "O Apps Script em produção parece antigo (só devolve ok:true sem checkoutUrl). Copie o código atual de google-apps-script.gs, salve, depois Implantar → Gerenciar implantações → ícone lápis → Nova versão → Implantar. Em Propriedades do script, defina MERCADO_PAGO_ACCESS_TOKEN.";
      }
      return {
        ok: true,
        checkoutUrl: checkoutUrl,
        checkoutFalhou: checkoutFalhou,
        erroCheckout: erroCheckout,
        loteAjustadoAutomatico: !!json.loteAjustadoAutomatico,
        lote: json.lote != null ? String(json.lote) : "",
        loteNome: json.loteNome != null ? String(json.loteNome) : "",
        valorReais: json.valorReais !== undefined && json.valorReais !== null ? Number(json.valorReais) : undefined,
      };
    } catch (e) {
      if (esperaCheckout) {
        return {
          ok: false,
          error:
            "Não foi possível abrir o pagamento online (falha de conexão ou bloqueio). Tente outra rede, desative bloqueador ou pague na secretaria. Não envie o formulário de novo se já apareceu mensagem de sucesso antes — fale com a organização.",
        };
      }
      try {
        await fetch(url, {
          method: "POST",
          body: body,
          mode: "no-cors",
          headers: headersPlain,
        });
        return { ok: true };
      } catch (e2) {
        return { ok: false, error: "Não foi possível salvar na planilha agora. Use o WhatsApp abaixo." };
      }
    }
  }

  form.addEventListener("submit", async function (e) {
    e.preventDefault();
    limparErro();

    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }

    var telDigits = apenasDigitos(inputTel ? inputTel.value : "");
    if (telDigits.length < 10 || telDigits.length > 11) {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status form-status--error";
        statusEl.textContent = "Informe um telefone válido com DDD: (87) 99999-9999 ou (87) 9999-9999.";
      }
      return;
    }

    var protocolo = gerarProtocolo();
    form.dataset.protocolo = protocolo;

    var dados = coletarDados(form);
    dados.protocolo = protocolo;
    dados.telefoneDigitos = apenasDigitos(dados.telefone);

    if (!dados.lote && lotesAtivos().length > 0) {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status form-status--error";
        statusEl.textContent = "Selecione o lote da inscrição.";
      }
      return;
    }

    var mpCfgPre = cfg.mercadoPago || {};
    if (dados.formaPagamentoCodigo === "mercado_pago_online" && !mpCfgPre.ativo) {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status form-status--error";
        statusEl.textContent = "Pagamento online não está disponível. Escolha presencial na secretaria.";
      }
      return;
    }

    var submitBtn = form.querySelector('[type="submit"]');
    setFormSubmitLoading(submitBtn, true);

    var mpCfg = cfg.mercadoPago || {};
    var urlRetorno = (mpCfg.urlRetorno || "").trim();
    if (!urlRetorno && typeof window !== "undefined" && window.location) {
      urlRetorno = window.location.origin + window.location.pathname;
    }

    var payload = Object.assign({}, dados, {
      tipo: "inscricao_corrida",
      evento: cfg.nomeEvento,
      mercadoPago: !!(mpCfg.ativo && dados.formaPagamentoCodigo === "mercado_pago_online"),
      urlRetorno: urlRetorno,
      useSandbox: !!mpCfg.useSandbox,
    });
    delete payload.formaPagamentoCodigo;

    if (payload.mercadoPago) {
      showMpLoading();
    }

    var result = await enviarWebhook(payload);
    var trocouParaRegular = false;

    if (!result.ok && !result.skipped && isErroLotePromoEsgotado(result.error)) {
      trocouParaRegular = tentarMigrarParaLoteRegular(dados, payload);
      if (trocouParaRegular) {
        result = await enviarWebhook(payload);
      }
    }

    if (result.ok && result.loteAjustadoAutomatico) {
      if (aplicarAjusteLoteDoServidor(dados, payload, result)) {
        trocouParaRegular = true;
      }
    }

    var msg = montarMensagemWhatsApp(dados);
    var urlWa = urlWhatsApp(msg);
    if (btnWa) btnWa.href = urlWa;

    if (result.checkoutUrl) {
      abrirWhatsAppAposInscricao(urlWa);
      window.location.href = result.checkoutUrl;
      return;
    }

    hideMpLoading();

    setFormSubmitLoading(submitBtn, false);

    if (!result.ok && !result.skipped) {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status form-status--error";
        statusEl.textContent =
          result.error || "Não foi possível concluir a inscrição. Tente de novo ou fale com a organização.";
      }
      return;
    }

    if (webhookAviso) {
      webhookAviso.hidden = true;
      webhookAviso.textContent = "";
    }

    /** Mercado Pago escolhido mas sem link: inscrição pode já estar na planilha — não reenviar o formulário. */
    if (
      dados.formaPagamentoCodigo === "mercado_pago_online" &&
      !result.skipped &&
      result.checkoutFalhou
    ) {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status form-status--error";
        statusEl.textContent =
          (result.erroCheckout ||
            "Não foi possível abrir o Mercado Pago.") +
          " Seu protocolo (guarde): " +
          protocolo +
          ". Não envie o formulário de novo — a inscrição pode já estar salva. Pague na secretaria ou fale com a organização.";
      }
      return;
    }

    var sucessoTexto = document.getElementById("inscricao-sucesso-texto");
    if (sucessoTexto) {
      if (dados.formaPagamentoCodigo === "mercado_pago_online") {
        sucessoTexto.innerHTML =
          "Não foi possível abrir o checkout do Mercado Pago agora. Guarde o protocolo e fale no <strong>WhatsApp da organização</strong> — a mensagem já vem com seus dados. Você também pode pagar na secretaria.";
      } else {
        sucessoTexto.innerHTML =
          "Guarde esse número. Agora fale no <strong>WhatsApp da organização</strong> — a mensagem já vem com seus dados para a equipe te enviar o PIX.";
      }
      if (trocouParaRegular) {
        sucessoTexto.innerHTML =
          "O <strong>lote promocional esgotou</strong> e sua inscrição foi ajustada automaticamente para o <strong>lote regular (" +
          (dados.valorFormatado || formatarMoedaBR(dados.valorReais)) +
          ")</strong>. " +
          sucessoTexto.innerHTML;
      }
    }
    if (confirmacaoPendenteTexto) {
      if (dados.formaPagamentoCodigo === "mercado_pago_online") {
        confirmacaoPendenteTexto.textContent =
          "Aguarde: sua inscrição será confirmada quando o pagamento online for aprovado.";
      } else {
        confirmacaoPendenteTexto.textContent =
          "Aguarde: sua inscrição será confirmada após o contato da equipe no WhatsApp e validação do pagamento.";
      }
    }

    if (protocoloEl) protocoloEl.textContent = protocolo;

    if (dados.formaPagamentoCodigo === "presencial_secretaria") {
      if (statusEl) {
        statusEl.hidden = false;
        statusEl.className = "form-status";
        statusEl.textContent = "Aguarde... abrindo WhatsApp para finalizar seu atendimento.";
      }
      setFormSubmitLoading(submitBtn, true);
      setTimeout(function () {
        redirecionarDiretoParaWhatsApp(urlWa);
      }, 180);
      return;
    }

    var headInsc = document.querySelector(".inscricao-head");
    if (headInsc) headInsc.hidden = true;
    var layoutInsc = document.getElementById("inscricao-formulario");
    if (layoutInsc) layoutInsc.hidden = true;

    form.reset();
    popularSelectLotes();
    configurarFormaPagamento();

    if (successPanel) {
      successPanel.hidden = false;
      successPanel.scrollIntoView({ behavior: "smooth", block: "start" });
      successPanel.focus();
    }
    abrirWhatsAppAposInscricao(urlWa);
  });

  if (copyPixBtn) {
    var copyPixLabel = copyPixBtn.textContent;
    copyPixBtn.addEventListener("click", function () {
      var chave = cfg.pixChave || "";
      if (!chave || chave.indexOf("(") === 0) {
        copyPixBtn.textContent = "Chave ainda não configurada";
        setTimeout(function () {
          copyPixBtn.textContent = copyPixLabel;
        }, 2200);
        return;
      }
      navigator.clipboard.writeText(chave).then(
        function () {
          copyPixBtn.textContent = "Copiado!";
          setTimeout(function () {
            copyPixBtn.textContent = copyPixLabel;
          }, 2000);
        },
        function () {
          copyPixBtn.textContent = "Copie manualmente acima";
          setTimeout(function () {
            copyPixBtn.textContent = copyPixLabel;
          }, 2500);
        }
      );
    });
  }

  function aplicarNumeroWhatsApp() {
    var num = (cfg.whatsappNumero || "").replace(/\D/g, "");
    if (!num) return;
    document.querySelectorAll('a[href*="wa.me/"]').forEach(function (a) {
      try {
        var u = new URL(a.href);
        u.pathname = "/" + num;
        a.href = u.toString();
      } catch (err) {
        a.href = a.href.replace(/wa\.me\/\d+/, "wa.me/" + num);
      }
    });
  }
  aplicarNumeroWhatsApp();

  var testModeBanner = document.getElementById("test-mode-banner");
  if (testModeBanner) {
    testModeBanner.hidden = !cfg.modoTeste;
  }

  var pixDisplay = document.getElementById("pix-chave-display");
  if (pixDisplay && cfg.pixChave) pixDisplay.textContent = cfg.pixChave;
  if (!lotesAtivos().length) {
    document.querySelectorAll("[data-valor-inscricao]").forEach(function (el) {
      el.textContent = cfg.valorInscricao || el.textContent;
    });
  }

  var patTel = document.getElementById("patrocinio-tel");
  if (patTel && cfg.patrocinioTelefone) patTel.textContent = cfg.patrocinioTelefone;
  var patLink = document.getElementById("patrocinio-tel-link");
  if (patLink && cfg.patrocinioTelefoneDigits) {
    var d = String(cfg.patrocinioTelefoneDigits).replace(/\D/g, "");
    patLink.href = "tel:+" + d;
  }
  var patWa = document.getElementById("patrocinio-wa-link");
  if (patWa && cfg.patrocinioTelefoneDigits) {
    var d2 = String(cfg.patrocinioTelefoneDigits).replace(/\D/g, "");
    var msg =
      "Olá! Gostaria de informações sobre *patrocínio* na " +
      (cfg.nomeEvento || "corrida") +
      ".";
    patWa.href = "https://wa.me/" + d2 + "?text=" + encodeURIComponent(msg);
  }
})();

/** Consulta de inscrição por e-mail + telefone (Apps Script tipo consulta_inscricao). */
(function () {
  "use strict";
  var cfg = window.CORRIDA_CONFIG || {};
  var formConsulta = document.getElementById("form-consulta");
  var resultadoEl = document.getElementById("consulta-resultado");
  var erroEl = document.getElementById("consulta-erro");
  var btnConsulta = document.getElementById("btn-consulta");
  if (!formConsulta || !resultadoEl) return;

  function apenasDigitosTel(str) {
    return String(str || "").replace(/\D/g, "");
  }

  function formatarTelefoneBR(valor) {
    var d = apenasDigitosTel(valor).slice(0, 11);
    if (d.length === 0) return "";
    if (d.length <= 2) return "(" + d;
    if (d.length <= 6) return "(" + d.slice(0, 2) + ") " + d.slice(2);
    if (d.length <= 10) return "(" + d.slice(0, 2) + ") " + d.slice(2, 6) + "-" + d.slice(6);
    return "(" + d.slice(0, 2) + ") " + d.slice(2, 7) + "-" + d.slice(7);
  }

  function escapeHtml(s) {
    var d = document.createElement("div");
    d.textContent = s == null ? "" : String(s);
    return d.innerHTML;
  }

  function setConsultaLoading(loading) {
    if (!btnConsulta) return;
    var label = btnConsulta.querySelector(".consulta-form__label");
    if (loading) {
      btnConsulta.disabled = true;
      btnConsulta.classList.add("consulta-form__submit--loading");
      if (label && !btnConsulta.dataset.labelOriginal) btnConsulta.dataset.labelOriginal = label.textContent;
      if (label) label.textContent = "Consultando…";
    } else {
      btnConsulta.disabled = false;
      btnConsulta.classList.remove("consulta-form__submit--loading");
      if (label) label.textContent = btnConsulta.dataset.labelOriginal || "Consultar";
    }
  }

  function limparConsultaUi() {
    if (erroEl) {
      erroEl.hidden = true;
      erroEl.textContent = "";
    }
    resultadoEl.hidden = true;
    resultadoEl.innerHTML = "";
  }

  function mostrarErro(msg) {
    limparConsultaUi();
    if (erroEl) {
      erroEl.textContent = msg;
      erroEl.hidden = false;
    }
  }

  function renderDados(d) {
    var linhas = [
      ["Protocolo", d.protocolo || "—"],
      ["Situação", d.situacaoLista || "—"],
      ["Status do pagamento", d.statusPagamento || "—"],
      ["Nome", d.nome || "—"],
      ["Cidade", d.cidade || "—"],
      ["Lote", d.loteNome || "—"],
      ["Valor", d.valorReais || "—"],
      ["Forma de pagamento", d.formaPagamento || "—"],
      ["Percurso", d.percurso || "—"],
      ["Tam. camisa", d.camisa || "—"],
      ["Evento", d.nomeEvento || "—"],
      ["Data do evento", d.dataEvento || "—"],
      ["Horário (concentração e largada)", d.horarioEvento || "—"],
      ["Local", d.localEvento || "—"],
    ];
    var parts = ['<h3 class="consulta-resultado__titulo">Inscrição encontrada</h3>', '<dl class="consulta-dl">'];
    linhas.forEach(function (pair) {
      parts.push(
        "<dt>" +
          escapeHtml(pair[0]) +
          "</dt><dd>" +
          escapeHtml(pair[1]) +
          "</dd>"
      );
    });
    parts.push("</dl>");
    resultadoEl.innerHTML = parts.join("");
    resultadoEl.hidden = false;
    resultadoEl.focus();
  }

  async function enviarConsulta(payload) {
    var url = (cfg.webhookUrl || "").trim();
    if (!url) {
      return { ok: false, error: "Consulta indisponível: configure webhookUrl (URL do Apps Script) em config.js." };
    }
    var body = JSON.stringify(payload);
    var headersPlain = { "Content-Type": "text/plain;charset=utf-8" };
    try {
      var res = await fetch(url, {
        method: "POST",
        headers: headersPlain,
        body: body,
        mode: "cors",
        credentials: "omit",
      });
      var text = (await res.text()).trim().replace(/^\uFEFF/, "");
      var json = null;
      try {
        json = JSON.parse(text);
      } catch (e) {
        json = null;
      }
      if (!json) {
        return {
          ok: false,
          error:
            "Resposta inválida do servidor. No Google Apps Script: salve o código atualizado (função consulta), depois Implantar → Nova versão.",
        };
      }
      return json;
    } catch (e) {
      return { ok: false, error: "Não foi possível consultar agora. Verifique a conexão ou tente mais tarde." };
    }
  }

  formConsulta.addEventListener("submit", async function (e) {
    e.preventDefault();
    limparConsultaUi();
    var email = (document.getElementById("consulta-email") || {}).value;
    var telRaw = (document.getElementById("consulta-telefone") || {}).value;
    email = email ? email.trim().toLowerCase() : "";
    var telDigits = apenasDigitosTel(telRaw);
    if (!email || !telRaw || !telRaw.trim()) {
      mostrarErro("Preencha o e-mail e o telefone (WhatsApp) usados na inscrição.");
      return;
    }
    if (telDigits.length < 10) {
      mostrarErro("Informe o telefone com DDD (ex.: (87) 99999-9999).");
      return;
    }

    setConsultaLoading(true);
    var out = await enviarConsulta({
      tipo: "consulta_inscricao",
      email: email,
      telefone: telRaw,
    });
    setConsultaLoading(false);

    if (!out.ok) {
      mostrarErro(out.error || "Erro ao consultar.");
      return;
    }
    if (out.encontrado && out.dados) {
      renderDados(out.dados);
      return;
    }
    mostrarErro(out.error || "Inscrição não encontrada. Confira e-mail e telefone.");
  });

  var emailConsulta = document.getElementById("consulta-email");
  if (emailConsulta) {
    emailConsulta.addEventListener("blur", function () {
      emailConsulta.value = emailConsulta.value.trim().toLowerCase();
    });
  }
  var telConsulta = document.getElementById("consulta-telefone");
  if (telConsulta) {
    telConsulta.addEventListener("input", function () {
      telConsulta.value = formatarTelefoneBR(telConsulta.value);
      var len = telConsulta.value.length;
      telConsulta.setSelectionRange(len, len);
    });
    telConsulta.addEventListener("blur", function () {
      telConsulta.value = formatarTelefoneBR(telConsulta.value);
    });
  }
})();
