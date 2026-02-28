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

function novoIdTimeStamp() {
  const ano = new Date().getFullYear() - 2025;
  const ms = Date.now() % 46656000000;
  return ano.toString(36) + ms.toString(36).padStart(5, '0');
}

/**
 * Recebe um ID Base36 existente e gera o próximo ID sequencial
 * @param {string} ultimoId O último ID gerado (ex: "1a4f3h")
 * @return {string} Novo ID incrementado em 1 unidade de tempo
 */
function gerarProximoIdIncremental(ultimoId) {
  if (!ultimoId || ultimoId.length < 2) return novoIdTimeStamp();

  // 1. Separa o prefixo do ano (primeiro caractere) e o corpo do timestamp
  const prefixoAno = ultimoId.substring(0, 1);
  const corpoMsBase36 = ultimoId.substring(1);

  // 2. Converte o corpo de Base 36 para Decimal (número inteiro)
  let msDecimal = parseInt(corpoMsBase36, 36);

  // 3. Incrementa 1 unidade
  msDecimal++;

  // 4. Converte de volta para Base 36 e garante o preenchimento de 5 dígitos
  const novoCorpoMs = msDecimal.toString(36).padStart(5, '0');

  return prefixoAno + novoCorpoMs;
}