// ID da planilha Google (substitua pelo seu ID real)
const SPREADSHEET_ID = '13b-Fxii7FBsjiHP36t2t53uD5-fYCF0V9mmNqZPrRpQ'; // <<<<< ATENÇÃO: SUBSTITUA Pelo ID da sua planilha aqui!

const MAPEAMENTO_COLUNAS = {
  // Campos básicos (colunas 1-10)
  DATA: "Carimbo de data/hora", // Coluna 1
  COLABORADOR: "Colaborador DS", // Coluna 2
  SOLICITANTE: "Solicitante", // Coluna 3
  PART_NUMBER: "Part Number:", // Coluna 4
  IMAGEM: "Imagem", // Coluna 5
  TIPO_PROJETO: "Tipo de Projeto", // Coluna 6
  LINHA_PRODUTO: "Qual a linha de produto?", // Coluna 7
  PART_NUMBER_DS: "PartNumber DS", // Coluna 8
  TIPO_DESENVOLVIMENTO: "Tipo de Desenvolvimento", // Coluna 9
  OBSERVACAO_PROJETO: "Observação", // Coluna 10

  // Campos Complemento de Linha (colunas 11-26)
  INV_FERRAMENTAL_COMP: "Investimento Ferramental (R$): Complemento", // Coluna 11
  RESP_FERRAMENTAL_COMP: "Responsável pela informação do investimento ferramental: Complemento", // Coluna 12
  INV_AMOSTRA_COMP: "Investimento da Amostra (R$): Complemento", // Coluna 13
  RESP_AMOSTRA_COMP: "Responsável pela informação do investimento da amostra: Complemento", // Coluna 14
  DT_ORCAMENTO_COMP: "Data do orçamento do custo da amostra: Complemento", // Coluna 15
  CURVA_VENDA_COMP: "Curva de venda:", // Coluna 16
  TIME_DEV_COMP: "Tempo de desenvolvimento: Complemento", // Coluna 17
  PRECO_REF_COMP: "Preço de referência (R$): Complemento", // Coluna 18
  VOL_MENSAL_COMP: "Volume mensal (und): Complemento", // Coluna 19
  MARKUP_COMP: "Mark Up Complemento", // Coluna 26
  
  // Campos Mercado (colunas 20-27)
  OEM_MERCADO: "Possui peças OEM no mercado?", // Coluna 20
  FROTA_APLICACAO: "Tamanho da frota de veículos da aplicação:", // Coluna 21
  FROTA_VEICULO: "Tamanho da Frota de veículos:", // Coluna 22
  QTD_VENDIDAS_FROTA: "Quantidade de peças vendidas (deste item próximo) pela DS:", // Coluna 23
  TIME_REST_FROTA: "Tempo restante de rodagem da frota:", // Coluna 24
  VOLUME_MENSAL: "Volume mensal:", // Coluna 25
  
  // Concorrentes (colunas 27-36)
  NOME_CONCORRENTE1: "Concorrente 1:", // Coluna 27
  PRC_CONCORRENTE1: "Preço do concorrente 1 (R$):", // Coluna 28
  NOME_CONCORRENTE2: "Concorrente 2:", // Coluna 29
  PRC_CONCORRENTE2: "Preço do concorrente 2 (R$):", // Coluna 30
  NOME_CONCORRENTE3: "Concorrente 3:", // Coluna 31
  PRC_CONCORRENTE3: "Preço do concorrente 3 (R$):", // Coluna 32
  NOME_CONCORRENTE4: "Concorrente 4:", // Coluna 33
  PRC_CONCORRENTE4: "Preço do concorrente 4 (R$):", // Coluna 34
  NOME_CONCORRENTE5: "Concorrente 5:", // Coluna 35
  PRC_CONCORRENTE5: "Preço do concorrente 5 (R$):", // Coluna 36

  // Campos de Linha Nova (colunas 76-86)
  INV_FERRAMENTAL_NOVA: "Investimento Ferramental (R$): Nova", // Coluna 76
  RESP_FERRAMENTAL_NOVA: "Responsável pela informação do investimento ferramental: Nova", // Coluna 77
  INV_AMOSTRA_NOVA: "Investimento da Amostra (R$): Nova", // Coluna 78
  RESP_AMOSTRA_NOVA: "responsável pela informação do investimento da amostra: Nova", // Coluna 79
  DT_ORCAMENTO_NOVA: "Data do orçamento do custo da amostra: Nova", // Coluna 80
  TIME_DEV_NOVA: "Tempo de desenvolvimento: Nova", // Coluna 81
  VOL_MENSAL_NOVA: "Volume mensal (und): Nova", // Coluna 82
  PRECO_REF_NOVA: "Preço de referência (R$): Nova", // Coluna 83
  MARKUP_NOVA: "Mark Up Nova", // Coluna 84
  MARGEM_CONTRIBUICAO_NOVA: "Margem de Contribuição:", // Coluna 85
  NOME_CONCORRENTE1: "Concorrente 1:", // Coluna 86

  // Campos adicionais que podem estar sendo usados
  PREVISAO_FATURAMENTO: "Previsão de Faturamento",
  FATURAMENTO_MENSAL_TOTAL: "Faturamento Mensal Total",
  PARTICIPACAO_FATURAMENTO: "Participação no Faturamento",
  PAYBACK: "Payback",
  PAYBACK_COMP: "Payback Complemento",
  
  // Campos para concorrentes (mantendo compatibilidade)
  CONCORRENTE_1_NOME: "Concorrente 1:",
  CONCORRENTE_1_OBS: "Concorrente 1 Observação",
  CONCORRENTE_2_NOME: "Concorrente 2:",
  CONCORRENTE_2_OBS: "Concorrente 2 Observação", 
  CONCORRENTE_3_NOME: "Concorrente 3:",
  CONCORRENTE_3_OBS: "Concorrente 3 Observação",
  CONCORRENTE_4_NOME: "Concorrente 4:",
  CONCORRENTE_4_OBS: "Concorrente 4 Observação",
  CONCORRENTE_5_NOME: "Concorrente 5:",
  CONCORRENTE_5_OBS: "Concorrente 5 Observação"
};

/**
 * Serve o arquivo HTML para o aplicativo web.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('FormularioHTML').evaluate()
    .setTitle('Formulário de Pré-Análise de Projetos');
}

/**
 * Processa os dados do formulário enviados do cliente.
 * @param {Object} formData - Os dados do formulário enviados.
 * @returns {Object} Um objeto com sucesso (true/false) e uma mensagem.
 */
function processForm(formData) {
  logExecucao("FormularioHTML", "processForm", "debug", "Inicio", "Dados recebidos: " + JSON.stringify(formData));

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Define nome da aba conforme o tipo de projeto
    let sheetName = "";
    if (formData.tipoProjeto === "Linha Nova") {
      sheetName = "PLANILHA DE LINHA NOVA";
    } else if (formData.tipoProjeto === "Complemento de Linha") {
      sheetName = "PLANILHA DE COMPLEMENTO";
    } else {
      logExecucao("FormularioHTML", "processForm", "error", "Erro", "Tipo de projeto não reconhecido.");
      return { success: false, message: "Erro: Tipo de projeto não reconhecido." };
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      logExecucao("FormularioHTML", "processForm", "error", "Erro", `Aba "${sheetName}" não encontrada.`);
      return { success: false, message: `Erro: Aba "${sheetName}" não encontrada.` };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let response;

    if (formData.tipoProjeto === "Linha Nova") {
      response = onFormSubmitLinhaNova(formData, sheet, headers);
    } else if (formData.tipoProjeto === "Complemento de Linha") {
      response = onFormSubmitComplementoLinha(formData, sheet, headers);
    }

    return response;

  } catch (e) {
    logExecucao("FormularioHTML", "processForm", "error", "Erro geral", e.message);
    return { success: false, message: `Erro ao processar o formulário: ${e.message}` };
  }
}


/**
 * Função para calcular o investimento total para Linha Nova.
 * @param {number} investimentoFerramental
 * @param {number} investimentoAmostra
 * @returns {number} Investimento total
 */
function calcInvestimento(investimentoFerramental, investimentoAmostra) {
    return investimentoFerramental + investimentoAmostra;
}

/**
 * Função para calcular o investimento total para Complemento de Linha.
 * @param {number} investimentoFerramentalComp
 * @param {number} investimentoAmostraComp
 * @returns {number} Investimento total
 */
function calcInvestimentoComp(investimentoFerramentalComp, investimentoAmostraComp) {
    return investimentoFerramentalComp + investimentoAmostraComp;
}

/**
 * Função para calcular o Payback para Linha Nova.
 * Payback = Investimento Total / (Volume Mensal * Preço de Referência * Markup)
 * @param {number} investimentoFerramental
 * @param {number} investimentoAmostra
 * @param {number} volumeMensal
 * @param {number} precoReferencia
 * @param {number} markup (em decimal)
 * @returns {number} Payback em meses
 */
function calcPayback(investimentoFerramental, investimentoAmostra, volumeMensal, precoReferencia, markup) {
  const investimentoTotal = calcInvestimento(investimentoFerramental, investimentoAmostra);
  if (markup > 0 && volumeMensal > 0 && precoReferencia > 0) {
    return investimentoTotal / (volumeMensal * precoReferencia * markup);
  }
  return 0;
}

/**
 * Função para calcular o Payback para Complemento de Linha.
 * Payback = Investimento Total / (Volume Mensal * Preço de Referência * Markup)
 * @param {number} investimentoFerramentalComp
 * @param {number} investimentoAmostraComp
 * @param {number} volumeMensalComp
 * @param {number} precoReferenciaComp
 * @param {number} markupComp (em decimal)
 * @returns {number} Payback em meses
 */
function calcPaybackComp(investimentoFerramentalComp, investimentoAmostraComp, volumeMensalComp, precoReferenciaComp, markupComp) {
  const investimentoTotalComp = calcInvestimentoComp(investimentoFerramentalComp, investimentoAmostraComp);
  if (markupComp > 0 && volumeMensalComp > 0 && precoReferenciaComp > 0) {
    return investimentoTotalComp / (volumeMensalComp * precoReferenciaComp * markupComp);
  }
  return 0;
}

/**
 * Calcula Participação no Faturamento
 */
function calcParticipacao(faturamentoPrevisto, faturamentoMensalTotal) {
  if (faturamentoMensalTotal > 0) {
    return faturamentoPrevisto / faturamentoMensalTotal;
  }
  return 0;
}

/**
 * Registra LOG na aba LOG
 */
function logExecucao(tipoProjeto, partNumber, destinatario, status, erro = "") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const abaLog = sheet.getSheetByName("LOG");
  if (abaLog) {
    abaLog.appendRow([new Date(), tipoProjeto, partNumber, destinatario, status, erro]);
  } else {
    // Se a aba LOG não existe, loga no console (para debug)
    console.log(`LOG - Tipo: ${tipoProjeto}, PN: ${partNumber}, Dest: ${destinatario}, Status: ${status}, Erro: ${erro}`);
  }
}

/**
 * Função auxiliar para obter índice da coluna com verificação
 */
function getColumnIndex(headers, columnName) {
  const index = headers.indexOf(columnName);
  if (index === -1) {
    console.warn(`Coluna "${columnName}" não encontrada nos cabeçalhos`);
  }
  return index;
}

function getFormHTML(tipo) {
  if (tipo === "Complemento de Linha") {
    return HtmlService.createHtmlOutputFromFile("formComplementoLinha").getContent();
  } else if (tipo === "Linha Nova") {
    return HtmlService.createHtmlOutputFromFile("formLinhaNova").getContent();
  }
  return HtmlService.createHtmlOutput("Tipo de projeto inválido.");
}


/**
 * Função para processar o formulário quando o Tipo de Projeto for "Linha Nova".
 * @param {Object} formData - Os dados do formulário enviados.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} aba - A aba de destino para salvar os dados.
 * @param {Object} headers - O cabeçalho da aba de destino (mapeamento).
 */
function onFormSubmitLinhaNova(formData, aba, headers) {
  logExecucao("FormularioHTML", "onFormSubmitLinhaNova", "debug", "Inicio", "Dados de Linha Nova: " + JSON.stringify(formData));

  try {
    const rowData = [];
    for (let i = 0; i < headers.length; i++) {
      rowData[i] = '';
    }

    const dataIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.DATA);
    if (dataIndex !== -1) rowData[dataIndex] = new Date();

    const colaboradorIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.COLABORADOR);
    if (colaboradorIndex !== -1) rowData[colaboradorIndex] = formData.colaboradorDS;

    const solicitanteIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.SOLICITANTE);
    if (solicitanteIndex !== -1) rowData[solicitanteIndex] = formData.solicitante;

    const partNumberIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PART_NUMBER);
    if (partNumberIndex !== -1) rowData[partNumberIndex] = formData.partNumber;

    const imagemIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.IMAGEM);
    if (imagemIndex !== -1) rowData[imagemIndex] = formData.imagem || '';

    const tipoProjetoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIPO_PROJETO);
    if (tipoProjetoIndex !== -1) rowData[tipoProjetoIndex] = formData.tipoProjeto;

    const linhaProdutoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.LINHA_PRODUTO);
    if (linhaProdutoIndex !== -1) rowData[linhaProdutoIndex] = formData.linhaProduto;

    const partNumberDSIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PART_NUMBER_DS);
    if (partNumberDSIndex !== -1) {
      if (formData.partNumberDS === "Sim") {
        rowData[partNumberDSIndex] = formData.partNumberDS_valor;
      } else {
        rowData[partNumberDSIndex] = formData.partNumberDS;
      }
    }

    const tipoDesenvolvimentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIPO_DESENVOLVIMENTO);
    if (tipoDesenvolvimentoIndex !== -1) rowData[tipoDesenvolvimentoIndex] = formData.tipoDesenvolvimento;

    const observacaoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.OBSERVACAO_PROJETO);
    if (observacaoIndex !== -1) rowData[observacaoIndex] = formData.observacao;

    const investimentoFerramental = parseFloat(formData.investimento_ferramental || 0);
    const investimentoAmostra = parseFloat(formData.investimento_amostra || 0);
    const volumeMensal = parseFloat(formData.volume_mensal || 0);
    const precoReferencia = parseFloat(formData.preco_referencia || 0);
    const markup = parseFloat(formData.markup || 0) / 100;

    const invFerramentalIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.INV_FERRAMENTAL_NOVA);
    if (invFerramentalIndex !== -1) rowData[invFerramentalIndex] = investimentoFerramental;

    const respFerramentalIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.RESP_FERRAMENTAL_NOVA);
    if (respFerramentalIndex !== -1) rowData[respFerramentalIndex] = formData.resp_inv_ferramental;

    const invAmostraIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.INV_AMOSTRA_NOVA);
    if (invAmostraIndex !== -1) rowData[invAmostraIndex] = investimentoAmostra;

    const dtOrcamentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.DT_ORCAMENTO_NOVA);
    if (dtOrcamentoIndex !== -1) rowData[dtOrcamentoIndex] = formData.data_orcamento_amostra;

    const timeDevIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIME_DEV_NOVA);
    if (timeDevIndex !== -1) rowData[timeDevIndex] = parseFloat(formData.tempo_desenvolvimento || 0);

    const precoRefIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PRECO_REF_NOVA);
    if (precoRefIndex !== -1) rowData[precoRefIndex] = precoReferencia;

    const volMensalIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.VOL_MENSAL_NOVA);
    if (volMensalIndex !== -1) rowData[volMensalIndex] = volumeMensal;

    const markupIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.MARKUP_NOVA);
    if (markupIndex !== -1) rowData[markupIndex] = markup * 100;

    const payback = calcPayback(investimentoFerramental, investimentoAmostra, volumeMensal, precoReferencia, markup);
    const paybackIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PAYBACK);
    if (paybackIndex !== -1) rowData[paybackIndex] = payback;

    const previsaoFaturamento = parseFloat(formData.previsao_faturamento || 0);
    const faturamentoMensalTotal = parseFloat(formData.faturamento_mensal_total || 0);

    const previsaoFaturamentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PREVISAO_FATURAMENTO);
    if (previsaoFaturamentoIndex !== -1) rowData[previsaoFaturamentoIndex] = previsaoFaturamento;

    const faturamentoMensalTotalIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FATURAMENTO_MENSAL_TOTAL);
    if (faturamentoMensalTotalIndex !== -1) rowData[faturamentoMensalTotalIndex] = faturamentoMensalTotal;

    const participacaoFaturamento = calcParticipacao(previsaoFaturamento, faturamentoMensalTotal);
    const participacaoFaturamentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PARTICIPACAO_FATURAMENTO);
    if (participacaoFaturamentoIndex !== -1) rowData[participacaoFaturamentoIndex] = participacaoFaturamento;

    const oemMercadoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.OEM_MERCADO);
    if (oemMercadoIndex !== -1) {
      if (formData.mercadoOEM === "Sim") {
        rowData[oemMercadoIndex] = formData.mercadoOEM_valor;
      } else {
        rowData[oemMercadoIndex] = formData.mercadoOEM;
      }
    }

    const frotaAplicacaoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FROTA_APLICACAO);
    if (frotaAplicacaoIndex !== -1) rowData[frotaAplicacaoIndex] = parseFloat(formData.frota_aplicacao || 0);

    const frotaVeiculoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FROTA_VEICULO);
    if (frotaVeiculoIndex !== -1) rowData[frotaVeiculoIndex] = parseFloat(formData.frota_veiculos || 0);

    const qtdVendidasFrotaIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.QTD_VENDIDAS_FROTA);
    if (qtdVendidasFrotaIndex !== -1) rowData[qtdVendidasFrotaIndex] = parseFloat(formData.pecasDS_vendidas || 0);

    const timeRestFrotaIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIME_REST_FROTA);
    if (timeRestFrotaIndex !== -1) rowData[timeRestFrotaIndex] = parseFloat(formData.frota_rodagem || 0);

    const existeConcorrenteIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.EXISTE_CONCORRENTE);
    if (existeConcorrenteIndex !== -1) rowData[existeConcorrenteIndex] = formData.existe_concorrente;

    for (let i = 1; i <= 5; i++) {
      const concorrenteNomeIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS[`CONCORRENTE_${i}_NOME`]);
      if (concorrenteNomeIndex !== -1) {
        rowData[concorrenteNomeIndex] = formData[`concorrente_${i}_nome`] || '';
      }
      const concorrenteObsIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS[`CONCORRENTE_${i}_OBS`]);
      if (concorrenteObsIndex !== -1) {
        rowData[concorrenteObsIndex] = formData[`concorrente_${i}_obs`] || '';
      }
    }

    // cálculos
    const investimentoTotal = investimentoFerramental + investimentoAmostra;
    const participacao = calcParticipacao(previsaoFaturamento, faturamentoMensalTotal);

    // atribuição ao template
    const emailTemplate = HtmlService.createTemplateFromFile('templateLinhaNova');
    emailTemplate.partNumber = formData.partNumber;
    emailTemplate.volumeMensal = volumeMensal;
    emailTemplate.precoReferencia = precoReferencia;
    emailTemplate.markup = markup;
    emailTemplate.investimentoNova = investimentoTotal;
    emailTemplate.faturamentoPrevisto = previsaoFaturamento;
    emailTemplate.payback = payback;
    emailTemplate.participacao = participacao;

    const htmlBody = emailTemplate.evaluate().getContent();

    const destinatario = formData.engenheiroResponsavelLinhaNova;
    const assuntoEmail = "Stage in Gate - Linha Nova";

    GmailApp.sendEmail(destinatario, assuntoEmail, "", {
      htmlBody: htmlBody
    });


    aba.appendRow(rowData);
    logExecucao("FormularioHTML", "onFormSubmitLinhaNova", "success", "Dados de Linha Nova salvos.");

    return { success: true, message: "Formulário de Linha Nova enviado com sucesso!" };

  } catch (e) {
    logExecucao("FormularioHTML", "onFormSubmitLinhaNova", "error", "Erro", e.message);
    throw new Error(`Erro ao processar formulário Linha Nova: ${e.message}`);
  }
}

/**
 * Função para processar o formulário quando o Tipo de Projeto for "Complemento de Linha".
 * @param {Object} formData - Os dados do formulário enviados.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} aba - A aba de destino para salvar os dados.
 * @param {Object} headers - O cabeçalho da aba de destino (mapeamento).
 */
function onFormSubmitComplementoLinha(formData, aba, headers) {
  console.log("Processando Complemento de Linha");

  try {
    const rowData = [];
    // Inicializa rowData com valores vazios para todas as colunas
    for (let i = 0; i < headers.length; i++) {
        rowData[i] = '';
    }

    // === CAMPOS BÁSICOS (Colunas 1-10) ===
    const dataIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.DATA);
    if (dataIndex !== -1) rowData[dataIndex] = new Date();
    
    const colaboradorIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.COLABORADOR);
    if (colaboradorIndex !== -1) rowData[colaboradorIndex] = formData.colaboradorDS;
    
    const solicitanteIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.SOLICITANTE);
    if (solicitanteIndex !== -1) rowData[solicitanteIndex] = formData.solicitante;
    
    const partNumberIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PART_NUMBER);
    if (partNumberIndex !== -1) rowData[partNumberIndex] = formData.partNumber;
    
    const imagemIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.IMAGEM);
    if (imagemIndex !== -1) rowData[imagemIndex] = formData.imagem || '';
    
    const tipoProjetoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIPO_PROJETO);
    if (tipoProjetoIndex !== -1) rowData[tipoProjetoIndex] = formData.tipoProjeto;
    
    const linhaProdutoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.LINHA_PRODUTO);
    if (linhaProdutoIndex !== -1) rowData[linhaProdutoIndex] = formData.linhaProduto;

    // Part Number DS
    const partNumberDSIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PART_NUMBER_DS);
    if (partNumberDSIndex !== -1) {
      if (formData.partNumberDS === "Sim") {
          rowData[partNumberDSIndex] = formData.partNumberDS_valor;
      } else {
          rowData[partNumberDSIndex] = formData.partNumberDS;
      }
    }

    const tipoDesenvolvimentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIPO_DESENVOLVIMENTO);
    if (tipoDesenvolvimentoIndex !== -1) rowData[tipoDesenvolvimentoIndex] = formData.tipoDesenvolvimento;
    
    const observacaoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.OBSERVACAO_PROJETO);
    if (observacaoIndex !== -1) rowData[observacaoIndex] = formData.observacao;

    // === INVESTIMENTOS E CUSTOS COMPLEMENTO (Colunas 11-19, 25-26) ===
    const investimentoFerramentalComp = parseFloat(formData.investimento_ferramental_comp || 0);
    const investimentoAmostraComp = parseFloat(formData.investimento_amostra_comp || 0);
    const volumeMensalComp = parseFloat(formData.volume_mensal_comp || 0);
    const precoReferenciaComp = parseFloat(formData.preco_referencia_comp || 0);
    const markupComp = parseFloat(formData.markup_comp || 0) / 100; // Converte para decimal

    // Coluna 11 - Investimento Ferramental Complemento
    const invFerramentalCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.INV_FERRAMENTAL_COMP);
    if (invFerramentalCompIndex !== -1) rowData[invFerramentalCompIndex] = investimentoFerramentalComp;
    
    // Coluna 12 - Responsável Investimento Ferramental Complemento
    const respFerramentalCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.RESP_FERRAMENTAL_COMP);
    if (respFerramentalCompIndex !== -1) rowData[respFerramentalCompIndex] = formData.resp_inv_ferramental_comp;
    
    // Coluna 13 - Investimento Amostra Complemento
    const invAmostraCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.INV_AMOSTRA_COMP);
    if (invAmostraCompIndex !== -1) rowData[invAmostraCompIndex] = investimentoAmostraComp;
    
    // Coluna 14 - Responsável Investimento Amostra Complemento
    const respAmostraCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.RESP_AMOSTRA_COMP);
    if (respAmostraCompIndex !== -1) rowData[respAmostraCompIndex] = formData.resp_inv_amostra_comp; // Assumindo que é o mesmo responsável
    
    // Coluna 15 - Data Orçamento Amostra Complemento
    const dtOrcamentoCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.DT_ORCAMENTO_COMP);
    if (dtOrcamentoCompIndex !== -1) rowData[dtOrcamentoCompIndex] = formData.data_orcamento_amostra_comp;
    
    // Coluna 16 - Curva de Venda
    const curvaVendaCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.CURVA_VENDA_COMP);
    if (curvaVendaCompIndex !== -1) rowData[curvaVendaCompIndex] = formData.curva_venda_comp;
    
    // Coluna 17 - Tempo Desenvolvimento Complemento
    const timeDevCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIME_DEV_COMP);
    if (timeDevCompIndex !== -1) rowData[timeDevCompIndex] = formData.tempo_desenvolvimento_comp;
    
    // Coluna 18 - Preço Referência Complemento
    const precoRefCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PRECO_REF_COMP);
    if (precoRefCompIndex !== -1) rowData[precoRefCompIndex] = precoReferenciaComp;
    
    // Coluna 19 - Volume Mensal Complemento
    const volMensalCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.VOL_MENSAL_COMP);
    if (volMensalCompIndex !== -1) rowData[volMensalCompIndex] = volumeMensalComp;

    // Coluna 26 - Markup Complemento
    const markupCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.MARKUP_COMP);
    if (markupCompIndex !== -1) rowData[markupCompIndex] = markupComp * 100; // Salva como porcentagem

    // === CAMPOS DE MERCADO (Colunas 20-24) ===
    // Coluna 20 - Mercado OEM
    const oemMercadoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.OEM_MERCADO);
    if (oemMercadoIndex !== -1) {
      if (formData.mercadoOEM === "Sim") {
          rowData[oemMercadoIndex] = formData.mercadoOEM_valor;
      } else {
          rowData[oemMercadoIndex] = formData.mercadoOEM;
      }
    }

    // Coluna 21 - Frota Aplicação
    const frotaAplicacaoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FROTA_APLICACAO);
    if (frotaAplicacaoIndex !== -1) rowData[frotaAplicacaoIndex] = parseFloat(formData.frota_aplicacao || 0);
    
    // Coluna 22 - Frota Veículos
    const frotaVeiculoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FROTA_VEICULO);
    if (frotaVeiculoIndex !== -1) rowData[frotaVeiculoIndex] = parseFloat(formData.frota_veiculos || 0);
    
    // Coluna 23 - Peças DS Vendidas
    const qtdVendidasFrotaIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.QTD_VENDIDAS_FROTA);
    if (qtdVendidasFrotaIndex !== -1) rowData[qtdVendidasFrotaIndex] = parseFloat(formData.pecasDS_vendidas || 0);
    
    // Coluna 24 - Tempo Restante Frota
    const timeRestFrotaIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.TIME_REST_FROTA);
    if (timeRestFrotaIndex !== -1) rowData[timeRestFrotaIndex] = parseFloat(formData.frota_rodagem || 0);
    
    // === CONCORRENTES (Colunas 27-36) ===
    // Coluna 27 - Existe Concorrente
    const existeConcorrenteIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.EXISTE_CONCORRENTE);
    if (existeConcorrenteIndex !== -1) rowData[existeConcorrenteIndex] = formData.existe_concorrente;

    // Observação: A estrutura da planilha não tem "Concorrente 1 Nome" na coluna 28, mas sim "Preço do concorrente 1"
    // Vamos mapear os nomes dos concorrentes para as colunas corretas baseadas na estrutura
    
    // Assumindo que os nomes dos concorrentes vão para as colunas de "Concorrente X:" (29, 31, 33, 35)
    // E que existe uma coluna "Concorrente 1:" que não está listada mas deve existir
    
    // Mapeamento dos concorrentes baseado na estrutura real
    const concorrenteNomes = [
      { campo: 'concorrente_1_nome', coluna: MAPEAMENTO_COLUNAS.NOME_CONCORRENTE1 }, // Assumindo coluna "Concorrente 1:"
      { campo: 'concorrente_2_nome', coluna: MAPEAMENTO_COLUNAS.NOME_CONCORRENTE2 }, // Coluna 29
      { campo: 'concorrente_3_nome', coluna: MAPEAMENTO_COLUNAS.NOME_CONCORRENTE3 }, // Coluna 31
      { campo: 'concorrente_4_nome', coluna: MAPEAMENTO_COLUNAS.NOME_CONCORRENTE4 }, // Coluna 33
      { campo: 'concorrente_5_nome', coluna: MAPEAMENTO_COLUNAS.NOME_CONCORRENTE5 }  // Coluna 35
    ];

    concorrenteNomes.forEach(concorrente => {
      const index = getColumnIndex(headers, concorrente.coluna);
      if (index !== -1) {
        rowData[index] = formData[concorrente.campo] || '';
      }
    });

    // === CÁLCULOS ===
    const paybackComp = calcPaybackComp(investimentoFerramentalComp, investimentoAmostraComp, volumeMensalComp, precoReferenciaComp, markupComp);
    
    // Se houver uma coluna específica para Payback Complemento
    const paybackCompIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PAYBACK_COMP);
    if (paybackCompIndex !== -1) rowData[paybackCompIndex] = paybackComp;

    // === DADOS ADICIONAIS (Se existirem na planilha) ===
    const previsaoFaturamento = parseFloat(formData.previsao_faturamento || 0);
    const faturamentoMensalTotal = parseFloat(formData.faturamento_mensal_total || 0);
    
    const previsaoFaturamentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PREVISAO_FATURAMENTO);
    if (previsaoFaturamentoIndex !== -1) rowData[previsaoFaturamentoIndex] = previsaoFaturamento;
    
    const faturamentoMensalTotalIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.FATURAMENTO_MENSAL_TOTAL);
    if (faturamentoMensalTotalIndex !== -1) rowData[faturamentoMensalTotalIndex] = faturamentoMensalTotal;

    const participacaoFaturamento = calcParticipacao(previsaoFaturamento, faturamentoMensalTotal);
    const participacaoFaturamentoIndex = getColumnIndex(headers, MAPEAMENTO_COLUNAS.PARTICIPACAO_FATURAMENTO);
    if (participacaoFaturamentoIndex !== -1) rowData[participacaoFaturamentoIndex] = participacaoFaturamento;

    // === SALVAR DADOS NA PLANILHA ===
    aba.appendRow(rowData);
    console.log("Dados salvos na planilha");

    // === ENVIO DE EMAIL ===
    const linhaProdutoSelecionada = formData.linhaProduto;
    const sheetEmails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails");

    if (!sheetEmails) {
        throw new Error("Aba 'Emails' não encontrada. Verifique o nome da aba.");
    }

    const emailValues = sheetEmails.getDataRange().getValues();
    const emailHeaders = emailValues[0];
    const colLinhaProdutoIndex = emailHeaders.indexOf("Linha de Produto");
    const colNomeEngenheiroIndex = emailHeaders.indexOf("Nome do Engenheiro");
    const colEmailEngenheiroIndex = emailHeaders.indexOf("E-mail do Engenheiro");

    if (colLinhaProdutoIndex === -1 || colNomeEngenheiroIndex === -1 || colEmailEngenheiroIndex === -1) {
        throw new Error("Cabeçalhos 'Linha de Produto', 'Nome do Engenheiro' ou 'E-mail do Engenheiro' não encontrados na aba 'Emails'. Verifique a aba 'Emails'.");
    }

    let engenheiroInfo = null;
    for (let i = 1; i < emailValues.length; i++) {
        const row = emailValues[i];
        if (row[colLinhaProdutoIndex] === linhaProdutoSelecionada) {
            engenheiroInfo = {
                email: row[colEmailEngenheiroIndex],
                nome: row[colNomeEngenheiroIndex]
            };
            break;
        }
    }

    if (!engenheiroInfo || !engenheiroInfo.email) {
      throw new Error(`Nenhum engenheiro ou e-mail encontrado na aba 'Emails' para a Linha de Produto: ${linhaProdutoSelecionada}. Verifique a aba 'Emails'.`);
    }

    const destinatario = engenheiroInfo.email;
    const nomeEngenheiro = engenheiroInfo.nome;

    const emailTemplate = HtmlService.createTemplateFromFile('templateComplemento');
    emailTemplate.partNumber = formData.partNumber;
    emailTemplate.paybackComp = paybackComp;
    emailTemplate.investimentoComp = investimentoFerramentalComp + investimentoAmostraComp;
    emailTemplate.investimentoAmostraComp = investimentoAmostraComp;
    emailTemplate.volumeMensalComp = volumeMensalComp;
    emailTemplate.precoReferenciaComp = precoReferenciaComp;
    emailTemplate.markupComp = markupComp * 100;
    emailTemplate.nomeEngenheiro = nomeEngenheiro;

    const htmlBody = emailTemplate.evaluate().getContent();
    const assuntoEmail = "Stage in Gate - Complemento de Linha";
    
    GmailApp.sendEmail(destinatario, assuntoEmail, "", {
      htmlBody: htmlBody
    });

    console.log("Email enviado para:", destinatario);
    logExecucao("FormularioHTML", "onFormSubmitComplementoLinha", "success", "Dados de Complemento de Linha salvos e email enviado.");

    return { success: true, message: "Formulário de Complemento de Linha enviado com sucesso!" };

  } catch (e) {
    console.log("Erro:", e.message);
    logExecucao("FormularioHTML", "onFormSubmitComplementoLinha", "error", "Erro", e.message);
    throw new Error(`Erro ao processar formulário Complemento de Linha: ${e.message}`);
  }
}