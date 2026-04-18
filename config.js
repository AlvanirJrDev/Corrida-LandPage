/**
 * Ajuste estes valores antes de publicar o site.
 *
 * Planilha: use webhookUrl com Google Apps Script ou Make.com → Google Sheets.
 * A planilha fica no Google Drive da conta que criou o Sheets (menu Compartilhar).
 * Coluna "Status pagamento": deixe "Pendente" e altere manualmente para "Pago" ao confirmar PIX/dinheiro.
 *
 * Mercado Pago: com mercadoPago.ativo = true, o formulário exibe "Mercado Pago" como opção e,
 * só se o inscrito escolher pagamento online, o Apps Script cria o checkout no Mercado Pago.
 * Configure MERCADO_PAGO_ACCESS_TOKEN no projeto (veja google-apps-script.gs).
 *
 * Planilha (Drive): só compartilhe com quem for organizador. "Pode editar" = pode ver todos os
 * inscritos e apagar linhas. Evite link público com permissão de edição.
 *
 * Pagamento online + consulta de inscrição: webhookUrl = URL da Web App (/exec). O google-apps-script.gs tem WEB_APP_URL_FALLBACK igual a ela;
 * opcionalmente defina WEB_APP_URL nas Propriedades do script para sobrescrever.
 */
var MODO_TESTE = false;

var PRECO_PROMO = MODO_TESTE ? 1 : 50;
var PRECO_REGULAR = MODO_TESTE ? 1 : 55;

window.CORRIDA_CONFIG = {
  modoTeste: MODO_TESTE,
  whatsappNumero: "5587991200165",
  /**
   * ORDEM (não existe link no site que “cria” a planilha — você cria no Google e depois cola aqui):
   * 1) Acesse https://sheets.google.com e faça login com a conta em que a planilha deve ficar.
   * 2) Planilha em branco → crie uma aba chamada exatamente "Inscrições" (veja cabeçalhos em google-apps-script.gs).
   * 3) Menu Extensões → Apps Script → cole o código do arquivo google-apps-script.gs → Salvar.
   * 4) No Apps Script: Implantar → Nova implantação → tipo “App da web” → Executar como: Eu → Acesso: Qualquer pessoa.
   * 5) Copie a URL que o Google mostrar (começa com https://script.google.com/...) e cole em webhookUrl abaixo.
   * Enquanto não tiver essa URL, deixe "" — o formulário ainda funciona (só WhatsApp, sem gravar na planilha).
   */
  webhookUrl:
    "https://script.google.com/a/macros/redealia.com/s/AKfycbwClwfOb9G4AXWFW0tjQAMAkuX_VlmlBHJC6nFPuGczHZZBM0XLv4p36mYF0RvvHCu7/exec",
  /** Usado só se a lista lotes estiver vazia (fallback). */
  valorInscricao: "R$ " + PRECO_REGULAR.toFixed(2).replace(".", ","),
  pixChave: "(informe a chave PIX ou orientação de pagamento)",
  nomeEvento: "Corrida Mariana em prol do ECC e EJC de Sanharó",

  /**
   * Mercado Pago — PRODUÇÃO (Checkout Pro)
   * - useSandbox: false + Access Token da aba "Credenciais de produção" (APP_USR-...) no Apps Script.
   * - urlRetorno: URL HTTPS pública do site no ar (Netlify, Pages, domínio próprio).
   * - Apps Script: Propriedades → MERCADO_PAGO_ACCESS_TOKEN = token produção | Implantar → Nova versão.
   */
  mercadoPago: {
    ativo: true,
    urlRetorno: "https://superlative-babka-667cbe.netlify.app/",
    useSandbox: false,
  },

  /**
   * Lotes ativos:
   * - Promo: 50 camisas (R$ 1,00 no teste | R$ 50,00 em produção)
   * - Regular: 100 camisas (R$ 1,00 no teste | R$ 55,00 em produção)
   * Alinhar com google-apps-script.gs (IDs, valores e limites).
   */
  lotes: [
    {
      id: "promo",
      nome: "Lote promocional",
      valorReais: PRECO_PROMO,
      descricao:
        "Primeiras 50 camisas por R$ " +
        PRECO_PROMO.toFixed(2).replace(".", ",") +
        (MODO_TESTE ? " (modo teste)." : "."),
      limite: 50,
    },
    {
      id: "regular",
      nome: "Lote regular",
      valorReais: PRECO_REGULAR,
      descricao:
        "Após o promocional: 100 camisas por R$ " +
        PRECO_REGULAR.toFixed(2).replace(".", ",") +
        (MODO_TESTE ? " (modo teste)." : "."),
      limite: 100,
    },
  ],

  patrocinioTelefone: "(87) 99120-0165",
  patrocinioTelefoneDigits: "5587991200165",
};
