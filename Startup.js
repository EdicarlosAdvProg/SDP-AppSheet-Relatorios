/**
 * Função acionada ao abrir o Google Docs.
 */
function onOpen(e) {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu('SDP-OAB')
        .addItem('Abrir Painel Lateral', 'PainelLateral_exibirSidebar')
        .addSeparator()
        .addItem('Sobre o Sistema', 'PainelLateral_exibirSobre')
        .addToUi();
  } catch (err) {
    Logger.log('Erro ao carregar UI no Docs: ' + err.toString());
  }
}

/**
 * Função utilitária para inclusão de componentes HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ID da planilha de dados (Google Sheets)
const PLANILHA_DADOS_ID = '1Xp8qKvWGIJRLZyCSWkm8erxtF5s6Ya5txOTkCxEmB34';

/**
 * Mapeia os cabeçalhos da planilha para facilitar o acesso aos dados
 * @param {Sheet} sheet - A aba da planilha
 * @param {number} linhaCabecalho - Linha onde está o cabeçalho (padrão: 1)
 * @return {Object} Objeto com mapeamento {nomeColuna: índice}
 */
function getMapaColunas(sheet, linhaCabecalho = 1) {
  const cabecalhos = sheet.getRange(linhaCabecalho, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mapa = {};
  
  cabecalhos.forEach((cabecalho, indice) => {
    if (cabecalho) {
      const chave = cabecalho.toString().trim().toLowerCase();
      mapa[chave] = indice + 1;
    }
  });
  
  return mapa;
}