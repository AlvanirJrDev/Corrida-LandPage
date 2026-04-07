/**
 * Ajuste estes valores antes de publicar o site.
 *
 * Planilha: use webhookUrl com Google Apps Script ou Make.com → Google Sheets.
 * A planilha fica no Google Drive da conta que criou o Sheets (menu Compartilhar).
 * Coluna "Status pagamento": deixe "Pendente" e altere manualmente para "Pago" ao confirmar PIX/dinheiro.
 *
 * Mercado Pago: com mercadoPago.ativo = true, o formulário exibe "Mercado Pago" como opção e,
 * só se o inscrito escolher pagamento online, o Apps Script cria o checkout (PIX/cartão).
 * Configure MERCADO_PAGO_ACCESS_TOKEN no projeto (veja google-apps-script.gs).
 *
 * Planilha (Drive): só compartilhe com quem for organizador. "Pode editar" = pode ver todos os
 * inscritos e apagar linhas. Evite link público com permissão de edição.
 *
 * Pagamento online + consulta de inscrição: webhookUrl = URL da Web App (/exec). O google-apps-script.gs tem WEB_APP_URL_FALLBACK igual a ela;
 * opcionalmente defina WEB_APP_URL nas Propriedades do script para sobrescrever.
 */
window.CORRIDA_CONFIG = {
  whatsappNumero: "5511999999999",
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
  valorInscricao: "R$ 45,00",
  pixChave: "(informe a chave PIX ou orientação de pagamento)",
  nomeEvento: "Corrida Mariana em prol do ECC e EJC de Sanharó",

  /**
   * Mercado Pago (Checkout Pro — redireciona para página segura de pagamento).
   * ativo: true para exibir a opção "Mercado Pago" e abrir o checkout ao escolher pagamento online.
   * urlRetorno: URL HTTPS da página de inscrição após pagar (deixe "" para usar a página atual). Obrigatório em produção com domínio fixo para o botão "voltar" do MP funcionar bem.
   * useSandbox: true = mesmo ambiente das "Credenciais de teste" no painel MP. false = "Credenciais de produção".
   * O Access Token NÃO vai aqui — só no Apps Script → Propriedades → MERCADO_PAGO_ACCESS_TOKEN (teste OU produção, alinhado a useSandbox).
   * Depois de colar o token ou alterar google-apps-script.gs: Implantar → Nova versão (senão o site usa código antigo).
   */
  mercadoPago: {
    ativo: true,
    urlRetorno: "",
    useSandbox: true,
  },

  /**
   * Só lote promocional ativo. limite = máximo de inscrições (50 camisas = 50 linhas neste lote).
   * O mesmo limite está no google-apps-script.gs (LIMITE_LOTE_PROMO) — altere nos dois se mudar.
   */
  lotes: [
    {
      id: "promo",
      nome: "Lote promocional",
      valorReais: 45,
      descricao: "Limite de 50 inscrições (50 camisas). Quando esgotar, o sistema bloqueia novas inscrições.",
      limite: 50,
    },
  ],

  patrocinioTelefone: "(87) 99999-0000",
  patrocinioTelefoneDigits: "5587999990000",
};
