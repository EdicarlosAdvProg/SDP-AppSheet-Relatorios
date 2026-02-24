/**
 * Instancia e exibe a barra lateral no Google Docs.
 */
function PainelLateral_exibirSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('PainelLateral-layout')
        .evaluate()
        .setTitle('Relatórios SDP-OAB')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    DocumentApp.getUi().showSidebar(html);
  } catch (err) {
    const ui = DocumentApp.getUi();
    ui.alert('Erro ao abrir painel: ' + err.message);
  }
}

/**
 * Exibe modal de informações no contexto do Docs.
 */
function PainelLateral_exibirSobre() {
  const ui = DocumentApp.getUi();
  ui.alert('Sistema de Relatórios SDP-OAB\nVersão 1.0\nAmbiente: Google Docs');
}

/**
 * Trata as ações rápidas disparadas pelo botão suspenso.
 */
function PainelLateral_processarAcaoRapida(acaoId) {
  const ui = DocumentApp.getUi();
  // No futuro, aqui abriremos os modais específicos para cada ação
  Logger.log('Ação rápida solicitada: ' + acaoId);
  ui.alert('Ação iniciada: ' + acaoId);
}

// ==========================================================
// CONFIGURAÇÃO DAS ABAS E COLUNAS
// Ajuste os nomes aqui caso divirjam da planilha real
// ==========================================================
const PL_ABA_SESSOES     = 'tabSessoes';
const PL_ABA_FICHAS      = 'tabFichas';
const PL_ABA_PROCURADORES = 'tabProcuradores';

/**
 * Retorna a planilha de dados pelo ID definido em Startup.gs
 */
function PainelLateral_getPlanilha() {
  return SpreadsheetApp.openById(PLANILHA_DADOS_ID);
}

/**
 * Retorna todas as sessões em ordem decrescente de data
 * para popular o seletor no topo da barra lateral.
 * @returns {Array<{id, data, orgao}>}
 */
function PainelLateral_listarSessoes() {
  try {
    const ss    = PainelLateral_getPlanilha();
    const sheet = ss.getSheetByName(PL_ABA_SESSOES);
    if (!sheet) throw new Error('Aba ' + PL_ABA_SESSOES + ' não encontrada.');

    const mapa   = getMapaColunas(sheet);
    const dados  = sheet.getDataRange().getValues();
    const header = dados[0];
    const linhas = dados.slice(1);

    // Índices base-0 para leitura direta no array de valores
    const iId    = (mapa['id']     || mapa['id sessão'] || 1) - 1;
    const iData  = (mapa['data']   || mapa['data da sessão'] || 2) - 1;
    const iOrgao = (mapa['órgão']  || mapa['orgao'] || mapa['órgao'] || 3) - 1;

    const sessoes = linhas
      .filter(row => row[iId] !== '')
      .map(row => ({
        id:    row[iId],
        data:  row[iData] ? Utilities.formatDate(new Date(row[iData]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        orgao: row[iOrgao] || ''
      }))
      .sort((a, b) => {
        // Ordena decrescente pela data original (parse simples)
        const da = new Date(a.data.split('/').reverse().join('-'));
        const db = new Date(b.data.split('/').reverse().join('-'));
        return db - da;
      });

    return sessoes;

  } catch (err) {
    throw new Error('PainelLateral_listarSessoes: ' + err.message);
  }
}

/**
 * Carrega todos os dados necessários para renderizar a pauta
 * de uma sessão específica.
 * @param {string|number} sessaoId
 * @returns {Object} { sessao, fichas, membros, procuradores, expediente }
 */
function PainelLateral_carregarPauta(sessaoId) {
  try {
    const ss = PainelLateral_getPlanilha();

    // --- Dados da Sessão ---
    const sheetSessoes  = ss.getSheetByName(PL_ABA_SESSOES);
    if (!sheetSessoes) throw new Error('Aba ' + PL_ABA_SESSOES + ' não encontrada.');
    const mapaSessoes   = getMapaColunas(sheetSessoes);
    const dadosSessoes  = sheetSessoes.getDataRange().getValues();

    const iSId        = (mapaSessoes['id']          || mapaSessoes['id sessão'] || 1) - 1;
    const iSData      = (mapaSessoes['data']         || mapaSessoes['data da sessão'] || 2) - 1;
    const iSOrgao     = (mapaSessoes['órgão']        || mapaSessoes['orgao'] || mapaSessoes['órgao'] || 3) - 1;
    const iSMembros   = (mapaSessoes['membros']      || 4) - 1;
    const iSProcurad  = (mapaSessoes['procuradores'] || 5) - 1;
    const iSExpediente= (mapaSessoes['expediente']   || 6) - 1;

    // Converte sessaoId para o tipo correto antes de comparar
    const idBusca = isNaN(sessaoId) ? String(sessaoId) : Number(sessaoId);
    const linhaSessao = dadosSessoes.slice(1).find(row => {
      const v = isNaN(row[iSId]) ? String(row[iSId]) : Number(row[iSId]);
      return v == idBusca;
    });

    if (!linhaSessao) throw new Error('Sessão ID ' + sessaoId + ' não encontrada.');

    const sessao = {
      id:    linhaSessao[iSId],
      data:  linhaSessao[iSData]
        ? Utilities.formatDate(new Date(linhaSessao[iSData]), Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : '',
      orgao: linhaSessao[iSOrgao] || ''
    };

    const membros = PainelLateral_parseLista(linhaSessao[iSMembros]);
    const procuradoresSessao = PainelLateral_parseLista(linhaSessao[iSProcurad]);
    const expediente = linhaSessao[iSExpediente] || '';

    // --- Fichas da Sessão ---
    const sheetFichas = ss.getSheetByName(PL_ABA_FICHAS);
    if (!sheetFichas) throw new Error('Aba ' + PL_ABA_FICHAS + ' não encontrada.');
    const mapaFichas  = getMapaColunas(sheetFichas);
    const dadosFichas = sheetFichas.getDataRange().getValues();

    const iFId        = (mapaFichas['id']           || mapaFichas['id ficha'] || 1) - 1;
    const iFSessaoId  = (mapaFichas['id sessão']    || mapaFichas['sessao']   || mapaFichas['sessão'] || 2) - 1;
    const iFOrdem     = (mapaFichas['ordem']         || 3) - 1;
    const iFProcesso  = (mapaFichas['processo']      || mapaFichas['nº processo'] || mapaFichas['numero'] || 4) - 1;
    const iFRequer    = (mapaFichas['requerente']    || 5) - 1;
    const iFProcurad  = (mapaFichas['procurador']    || 6) - 1;
    const iFRelator   = (mapaFichas['relator']       || 7) - 1;

    const fichas = dadosFichas.slice(1)
      .filter(row => {
        const v = isNaN(row[iFSessaoId]) ? String(row[iFSessaoId]) : Number(row[iFSessaoId]);
        return v == idBusca && row[iFId] !== '';
      })
      .map(row => ({
        id:        row[iFId],
        ordem:     row[iFOrdem]    || '',
        processo:  row[iFProcesso] || '',
        requerente:row[iFRequer]   || '',
        procurador:row[iFProcurad] || '',
        relator:   row[iFRelator]  || ''
      }))
      .sort((a, b) => Number(a.ordem) - Number(b.ordem));

    return { sessao, fichas, membros, procuradores: procuradoresSessao, expediente };

  } catch (err) {
    throw new Error('PainelLateral_carregarPauta: ' + err.message);
  }
}

/**
 * Atualiza o campo [Membros] de uma sessão.
 * @param {string|number} sessaoId
 * @param {Array<string>} listaMembros
 */
function PainelLateral_salvarMembros(sessaoId, listaMembros) {
  PainelLateral_salvarCampoSessao(sessaoId, 'membros', listaMembros.join(';'));
}

/**
 * Atualiza o campo [Procuradores] de uma sessão.
 * @param {string|number} sessaoId
 * @param {Array<string>} listaProcuradores
 */
function PainelLateral_salvarProcuradores(sessaoId, listaProcuradores) {
  PainelLateral_salvarCampoSessao(sessaoId, 'procuradores', listaProcuradores.join(';'));
}

/**
 * Atualiza o campo [Expediente] de uma sessão.
 * @param {string|number} sessaoId
 * @param {string} texto
 */
function PainelLateral_salvarExpediente(sessaoId, texto) {
  PainelLateral_salvarCampoSessao(sessaoId, 'expediente', texto);
}

/**
 * Retorna os nomes cadastrados em tabProcuradores para autocomplete.
 * @returns {Array<string>}
 */
function PainelLateral_listarProcuradoresCadastrados() {
  try {
    const ss    = PainelLateral_getPlanilha();
    const sheet = ss.getSheetByName(PL_ABA_PROCURADORES);
    if (!sheet) throw new Error('Aba ' + PL_ABA_PROCURADORES + ' não encontrada.');

    const mapa  = getMapaColunas(sheet);
    const dados = sheet.getDataRange().getValues();
    const iNome = (mapa['nome'] || mapa['procurador'] || 1) - 1;

    return dados.slice(1)
      .map(row => (row[iNome] || '').toString().trim())
      .filter(nome => nome !== '');

  } catch (err) {
    throw new Error('PainelLateral_listarProcuradoresCadastrados: ' + err.message);
  }
}

// ------------------------------------------------------------------
// Funções auxiliares privadas
// ------------------------------------------------------------------

/**
 * Converte uma string "Nome1;Nome2;Nome3" em array limpo.
 * @param {string} valor
 * @returns {Array<string>}
 */
function PainelLateral_parseLista(valor) {
  if (!valor) return [];
  return valor.toString().split(';').map(s => s.trim()).filter(s => s !== '');
}

/**
 * Localiza a linha de uma sessão pelo ID e grava um campo específico.
 * @param {string|number} sessaoId
 * @param {string} nomeCampo   - chave usada em getMapaColunas (lowercase)
 * @param {string} novoValor
 */
function PainelLateral_salvarCampoSessao(sessaoId, nomeCampo, novoValor) {
  try {
    const ss    = PainelLateral_getPlanilha();
    const sheet = ss.getSheetByName(PL_ABA_SESSOES);
    if (!sheet) throw new Error('Aba ' + PL_ABA_SESSOES + ' não encontrada.');

    const mapa      = getMapaColunas(sheet);
    const dados     = sheet.getDataRange().getValues();
    const iId       = (mapa['id'] || mapa['id sessão'] || 1) - 1;
    const iCampo    = (mapa[nomeCampo] || null);
    if (!iCampo) throw new Error('Coluna "' + nomeCampo + '" não encontrada em tabSessoes.');

    const idBusca = isNaN(sessaoId) ? String(sessaoId) : Number(sessaoId);
    let linhaEncontrada = -1;

    for (let i = 1; i < dados.length; i++) {
      const v = isNaN(dados[i][iId]) ? String(dados[i][iId]) : Number(dados[i][iId]);
      if (v == idBusca) { linhaEncontrada = i + 1; break; } // +1 base-1 do Sheets
    }

    if (linhaEncontrada === -1) throw new Error('Sessão ID ' + sessaoId + ' não encontrada para gravação.');

    sheet.getRange(linhaEncontrada, iCampo).setValue(novoValor);

  } catch (err) {
    throw new Error('PainelLateral_salvarCampoSessao: ' + err.message);
  }
}