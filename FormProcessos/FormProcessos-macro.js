/**
 * @fileoverview Backend para Gestão de Processos SDP
 * Baseado na estrutura: Id, Processo, Requerente, Requerido, Procurador, Local, Resumo, Ementa, Provas
 */

function formProcessos_abrirModal() {
  const template = HtmlService.createTemplateFromFile('FormProcessos-layout');
  const html = template.evaluate()
    .setTitle('SDP - Gestão de Processos')
    .setWidth(1200)
    .setHeight(850);
  
  DocumentApp.getUi().showModalDialog(html, ' ');
}

/**
 * Função para ser chamada pelo Painel Lateral
 * IMPORTANTE: Esta função é executada no contexto do servidor quando chamada
 * via google.script.run do painel lateral, então tem acesso à UI
 */
function PainelLateral_abrirProcessos() {
  // Fecha o painel lateral antes de abrir o modal (opcional)
  // SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(''));

  formProcessos_abrirModal();
  return true;
}

/**
 * Busca dados para popular o formulário e carregar listas de apoio
 */
/**
 * Abre o formulário de processos com largura ampliada.
 */
function sdp_abrirFormProcessos() {
  const template = HtmlService.createTemplateFromFile('FormProcessos-layout');
  const html = template.evaluate()
    .setTitle('SDP - Gestão de Processos')
    .setWidth(1200) // Aumentado conforme solicitado
    .setHeight(850);
  
  DocumentApp.getUi().showModalDialog(html, ' ');
}

/**
 * Busca dados de processos e cruza com a data mais recente da tabHistorico
 */
/**
 * @fileoverview Backend para Gestão de Processos SDP
 * Integra tabProcessos com a última data de tabHistorico
 */

function sdp_abrirFormProcessos() {
  const template = HtmlService.createTemplateFromFile('FormProcessos-layout');
  const html = template.evaluate()
    .setTitle('SDP - Gestão de Processos')
    .setWidth(1200) 
    .setHeight(850);
  
  DocumentApp.getUi().showModalDialog(html, ' ');
}

/**
 * Busca dados de processos e cruza com a data mais recente da tabHistorico
 */
function formProcessos_buscarDadosCompletos() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    
    // 1. Acesso à aba Processos
    const sheetProc = ss.getSheetByName("tabProcessos");
    if (!sheetProc) throw new Error("Aba 'tabProcessos' não encontrada.");
    const mapaProc = getMapaColunas(sheetProc);
    const dadosProc = sheetProc.getDataRange().getValues();
    dadosProc.shift(); // Remove cabeçalho

    // 2. Acesso à aba Histórico
    const sheetHist = ss.getSheetByName("tabHistorico");
    let ultimoEventoMap = {};
    if (sheetHist) {
      const mapaHist = getMapaColunas(sheetHist);
      const dadosHist = sheetHist.getDataRange().getValues();
      dadosHist.shift();

      dadosHist.forEach(h => {
        const idProc = h[mapaHist['idprocesso'] - 1];
        const data = h[mapaHist['datahora'] - 1];
        if (idProc && data) {
          const dataObjeto = new Date(data);
          if (!ultimoEventoMap[idProc] || dataObjeto > ultimoEventoMap[idProc]) {
            ultimoEventoMap[idProc] = dataObjeto;
          }
        }
      });
    }

    // 3. Montagem da lista final
    const listaFinal = dadosProc.map(linha => {
      const id = linha[mapaProc['id'] - 1];
      let dataFormatada = "Sem histórico";
      
      if (ultimoEventoMap[id]) {
        dataFormatada = Utilities.formatDate(ultimoEventoMap[id], "GMT-3", "dd/MM/yyyy");
      }

      return {
        id: id,
        processo: linha[mapaProc['processo'] - 1] || "S/N",
        requerente: linha[mapaProc['requerente'] - 1] || "Não informado",
        requerido: linha[mapaProc['requerido'] - 1] || "N/C",
        procurador: linha[mapaProc['procurador'] - 1] || "Pendente",
        relator: linha[mapaProc['relator'] - 1] || "Não designado",
        ementa: linha[mapaProc['ementa'] - 1] || "",
        ultimaData: dataFormatada
      };
    });

    return { sucesso: true, dados: listaFinal };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

function formProcessos_salvarRegistro(obj) {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const sheet = ss.getSheetByName("tabProcessos");
  const dados = sheet.getDataRange().getValues();
  const colunas = dados[0].map(c => c.toLowerCase());
  
  let rowIndex = -1;
  if (obj.id) {
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === obj.id) { rowIndex = i + 1; break; }
    }
  }

  // Monta a linha baseada na ordem exata das colunas da planilha
  const linhaParaSalvar = colunas.map(col => obj[col] || "");

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 1, 1, linhaParaSalvar.length).setValues([linhaParaSalvar]);
  } else {
    obj.id = "PRC-" + new Date().getTime(); // Gerador simples de ID
    linhaParaSalvar[0] = obj.id;
    sheet.appendRow(linhaParaSalvar);
  }
  
  return { sucesso: true };
}