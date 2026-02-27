/**
 * Instancia e exibe a barra lateral no Google Docs.
 */
function PainelLateral_exibirSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('PainelLateral-layout');
    
    // Coleta o pacote de dados (Sessão recente, Fichas e Votos dela)
    const dadosIniciais = PainelLateral_obterPacoteInicial();
    
    // Injeta na variável que o HTML vai ler
    html.dadosIniciaisJSON = JSON.stringify(dadosIniciais);
    
    const display = html.evaluate()
        .setTitle('SDP-OAB')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    DocumentApp.getUi().showSidebar(display);
  } catch (err) {
    DocumentApp.getUi().alert('Erro: ' + err.message);
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

    // 1. DADOS DA SESSÃO
    const sheetSessoes = ss.getSheetByName('tabSessoes');
    const dadosSessoes = sheetSessoes.getDataRange().getValues();
    const mapaSessoes  = getMapaColunas(sheetSessoes);
    
    const iSId = (mapaSessoes['id'] || 1) - 1;
    const idBusca = isNaN(sessaoId) ? String(sessaoId) : Number(sessaoId);
    const linhaSessao = dadosSessoes.slice(1).find(row => row[iSId] == idBusca);
    if (!linhaSessao) throw new Error('Sessão não encontrada.');

    const sessao = {
      id:    linhaSessao[iSId],
      data:  linhaSessao[(mapaSessoes['datasessao'] || 2) - 1] ? Utilities.formatDate(new Date(linhaSessao[(mapaSessoes['datasessao'] || 2) - 1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      orgao: linhaSessao[(mapaSessoes['órgão'] || 3) - 1] || ''
    };

    // 2. BUSCA DE PROCESSOS (Cache para lookup rápido)
    const sheetProcs = ss.getSheetByName('tabProcessos');
    const dadosProcs = sheetProcs.getDataRange().getValues();
    const mapaProcs  = getMapaColunas(sheetProcs);
    const iPrId      = (mapaProcs['id'] || 1) - 1;
    const iPrNum     = (mapaProcs['processo'] || 2) - 1;
    const iPrReq     = (mapaProcs['requerente'] || 3) - 1;
    const iPrProc    = (mapaProcs['procurador'] || 5) - 1; // Coluna E na sua planilha

    // 3. FICHAS DA SESSÃO
    const sheetFichas = ss.getSheetByName('tabFichas');
    const dadosFichas = sheetFichas.getDataRange().getValues();
    const mapaFichas  = getMapaColunas(sheetFichas);
    const iFSessaoId  = (mapaFichas['idsessao'] || 2) - 1;
    const iFProcId    = (mapaFichas['idprocesso'] || 3) - 1;

    const fichas = dadosFichas.slice(1)
      .filter(row => row[iFSessaoId] == idBusca)
      .map(row => {
        const idProcessoReferencia = row[iFProcId];
        // Busca na tabProcessos pelo ID único (Ex: 1ks3cbsy)
        const infoProc = dadosProcs.find(p => p[iPrId] == idProcessoReferencia);

        return {
          id:         row[(mapaFichas['id'] || 1) - 1],
          ordem:      row[(mapaFichas['ordem'] || 4) - 1] || '',
          processo:   infoProc ? infoProc[iPrNum] : 'N/D',
          requerente: infoProc ? infoProc[iPrReq] : 'N/D',
          procurador: infoProc ? infoProc[iPrProc] : 'N/D',
          relator:    row[(mapaFichas['relator'] || 5) - 1] || ''
        };
      })
      .sort((a, b) => Number(a.ordem) - Number(b.ordem));

    return { 
      sessao, 
      fichas, 
      membros: PainelLateral_parseLista(linhaSessao[(mapaSessoes['membros'] || 7) - 1]), 
      procuradores: PainelLateral_parseLista(linhaSessao[(mapaSessoes['procuradores'] || 8) - 1]), 
      expediente: linhaSessao[(mapaSessoes['expediente'] || 9) - 1] || '' 
    };

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

function PainelLateral_obterPacoteInicial() {
  const sessoes = PainelLateral_listarSessoes();
  const idRecente = sessoes.length > 0 ? sessoes[0].id : null;
  
  // Busca pauta (fichas) e votos apenas da sessão recente
  const pautaAtiva = idRecente ? PainelLateral_carregarPauta(idRecente) : null;
  const votosSessao = idRecente ? PainelLateral_obterVotosDaSessao(idRecente) : [];

  return {
    config: { sessaoAtivaId: idRecente },
    sessoes: sessoes,
    pautaAtiva: pautaAtiva, 
    votosSessao: votosSessao,
    cadastros: {
      membros: PainelLateral_listarMembrosCompleto(), 
      procuradores: PainelLateral_listarProcuradoresCadastrados()
    }
  };
}

function PainelLateral_obterVotosDaSessao(sessaoId) {
  const ss = PainelLateral_getPlanilha();
  const sheetVotos = ss.getSheetByName('tabVotos');
  const dadosVotos = sheetVotos.getDataRange().getValues();
  const mapaVotos = getMapaColunas(sheetVotos);
  
  // Pegamos os votos onde o ID do Processo ou ID da Ficha pertença a esta sessão
  // Nota: Na sua pautaAtiva já temos os IDs das fichas carregados
  return dadosVotos.slice(1).map(v => ({
    idFicha: v[(mapaVotos['idfichavotacao'] || 2) - 1],
    voto: v[(mapaVotos['voto'] || 6) - 1]
  })).filter(v => v.voto !== ""); // Filtro simples para não enviar lixo
}

/**
 * Retorna a lista completa de nomes da tabMembros para o cache de autocompletes.
 */
function PainelLateral_listarMembrosCompleto() {
  try {
    const ss = PainelLateral_getPlanilha();
    const sheet = ss.getSheetByName('tabMembros');
    if (!sheet) return [];

    const dados = sheet.getDataRange().getValues();
    const mapa = getMapaColunas(sheet);
    const iNome = (mapa['nome'] || 2) - 1;

    return dados.slice(1)
      .map(row => (row[iNome] || '').toString().trim())
      .filter(nome => nome !== '');
  } catch (err) {
    Logger.log('Erro em listarMembrosCompleto: ' + err.message);
    return [];
  }
}

/**
 * Busca votos em tabVotos filtrando apenas pelos IDs das Fichas da sessão atual.
 * Versão otimizada para grandes volumes.
 */
function PainelLateral_obterVotosDaSessao(sessaoId) {
  try {
    const ss = PainelLateral_getPlanilha();
    
    // 1. Primeiro identificamos os IDs das fichas que pertencem a esta sessão
    const sheetFichas = ss.getSheetByName('tabFichas');
    const dadosFichas = sheetFichas.getDataRange().getValues();
    const mapaFichas = getMapaColunas(sheetFichas);
    const iFid = (mapaFichas['id'] || 1) - 1;
    const iFsessao = (mapaFichas['idsessao'] || 2) - 1;
    
    const idsFichasDaSessao = dadosFichas.slice(1)
      .filter(r => r[iFsessao] == sessaoId)
      .map(r => r[iFid]);

    if (idsFichasDaSessao.length === 0) return [];

    // 2. Filtramos a tabVotos apenas para estas fichas
    const sheetVotos = ss.getSheetByName('tabVotos');
    const dadosVotos = sheetVotos.getDataRange().getValues();
    const mapaVotos = getMapaColunas(sheetVotos);
    const iVidFicha = (mapaVotos['idfichavotacao'] || 2) - 1;
    const iVvoto = (mapaVotos['voto'] || 6) - 1;

    return dadosVotos.slice(1)
      .filter(v => idsFichasDaSessao.indexOf(v[iVidFicha]) !== -1)
      .map(v => ({
        idFicha: v[iVidFicha],
        voto: v[iVvoto] || ''
      }));
  } catch (err) {
    Logger.log('Erro em obterVotosDaSessao: ' + err.message);
    return [];
  }
}