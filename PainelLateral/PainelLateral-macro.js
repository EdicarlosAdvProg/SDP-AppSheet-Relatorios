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
        .setTitle('Ferramentas da Sessão')
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
 * Carrega todos os campos necessários para geração do documento,
 * incluindo Requerido, Ementa (tabProcessos) e votos (tabVotos).
 * O cliente ficará com tudo em memória — a geração não precisará
 * tocar a planilha novamente.
 */
function PainelLateral_carregarPauta(sessaoId) {
  try {
    const ss       = PainelLateral_getPlanilha();
    const idBusca  = isNaN(sessaoId) ? String(sessaoId) : Number(sessaoId);

    // ── 1. SESSÃO ────────────────────────────────────────────────
    const sheetSessoes = ss.getSheetByName('tabSessoes');
    const dadosSessoes = sheetSessoes.getDataRange().getValues();
    const mS           = getMapaColunas(sheetSessoes);

    const linhaSessao = dadosSessoes.slice(1).find(r => r[(mS['id'] || 1) - 1] == idBusca);
    if (!linhaSessao) throw new Error('Sessão não encontrada.');

    const sessao = {
      id:         linhaSessao[(mS['id']          || 1) - 1],
      data:       linhaSessao[(mS['datasessao']  || 2) - 1]
                    ? Utilities.formatDate(new Date(linhaSessao[(mS['datasessao'] || 2) - 1]),
                        Session.getScriptTimeZone(), 'dd/MM/yyyy')
                    : '',
      orgao:      linhaSessao[(mS['órgão']        || 3) - 1] || '',
      presidente: linhaSessao[(mS['presidente']  || 5) - 1] || '',
      secretario: linhaSessao[(mS['secretário']  || 6) - 1] || ''
    };

    // ── 2. PROCESSOS — mapa id→linha para lookup O(1) ───────────
    const sheetProcs = ss.getSheetByName('tabProcessos');
    const dadosProcs = sheetProcs.getDataRange().getValues();
    const mP         = getMapaColunas(sheetProcs);
    const iPrId      = (mP['id']         || 1) - 1;

    const procMap = {};
    dadosProcs.slice(1).forEach(r => { procMap[r[iPrId]] = r; });

    // ── 3. VOTOS — mapa idFicha→linha para lookup O(1) ──────────
    const sheetVotos = ss.getSheetByName('tabVotos');
    const dadosVotos = sheetVotos.getDataRange().getValues();
    const mV         = getMapaColunas(sheetVotos);
    const iVFicha    = (mV['idfichavotacao'] || 2) - 1;

    const votosMap = {};
    dadosVotos.slice(1).forEach(r => { votosMap[r[iVFicha]] = r; });

    // ── 4. FICHAS — enriquecidas com todos os campos do documento ─
    const sheetFichas = ss.getSheetByName('tabFichas');
    const dadosFichas = sheetFichas.getDataRange().getValues();
    const mF          = getMapaColunas(sheetFichas);

    const iFId       = (mF['id']         || 1) - 1;
    const iFSessao   = (mF['idsessao']   || 2) - 1;
    const iFProcId   = (mF['idprocesso'] || 3) - 1;
    const iFOrdem    = (mF['ordem']      || 4) - 1;
    const iFRelator  = (mF['relator']    || 5) - 1;

    const fichas = dadosFichas.slice(1)
      .filter(r => r[iFSessao] == idBusca)
      .map(r => {
        const idFicha   = r[iFId];
        const idProc    = r[iFProcId];
        const linhaProc = procMap[idProc] || null;
        const linhaVoto = votosMap[idFicha] || null;

        return {
          id:             idFicha,
          idProcesso:     idProc,
          ordem:          r[iFOrdem]   || '',
          relator:        r[iFRelator] || '',

          // tabProcessos — campos completos para o documento
          processo:       linhaProc ? linhaProc[(mP['processo']   || 2) - 1] : 'N/D',
          requerente:     linhaProc ? linhaProc[(mP['requerente'] || 3) - 1] : 'N/D',
          requerido:      linhaProc ? linhaProc[(mP['requerido']  || 4) - 1] : 'N/D',
          procurador:     linhaProc ? linhaProc[(mP['procurador'] || 5) - 1] : 'N/D',
          ementa:         linhaProc ? linhaProc[(mP['ementa']     || 6) - 1] : 'N/D',

          // tabVotos — campos completos para o documento
          voto:           linhaVoto ? linhaVoto[(mV['voto']            || 6) - 1] : '',
          resultado:      linhaVoto ? linhaVoto[(mV['resultado']        || 7) - 1] : '',
          votosRelator:   linhaVoto ? linhaVoto[(mV['votosrelator']     || 8) - 1] : '0',
          totalVotantes:  linhaVoto ? linhaVoto[(mV['totalvotantes']    || 9) - 1] : '0'
        };
      })
      .sort((a, b) => Number(a.ordem) - Number(b.ordem));

    return {
      sessao,
      fichas,
      membros:      PainelLateral_parseLista(linhaSessao[(mS['membros']      || 7) - 1]),
      procuradores: PainelLateral_parseLista(linhaSessao[(mS['procuradores'] || 8) - 1]),
      expediente:   linhaSessao[(mS['expediente'] || 9) - 1] || ''
    };

  } catch (err) {
    throw new Error('PainelLateral_carregarPauta: ' + err.message);
  }
}

/** Função para persistir a mesa diretora */
function PainelLateral_salvarMesa(sessaoId, presidente, secretario) {
  PainelLateral_salvarCampoSessao(sessaoId, 'presidente', presidente);
  PainelLateral_salvarCampoSessao(sessaoId, 'secretário', secretario);
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

/**
 * Agrega apenas os dados estáticos/cadastros para o carregamento inicial.
 */
function PainelLateral_obterPacoteInicial() {
  try {
    // 1. Lista de sessões para o Dropdown (sempre necessária no início)
    const sessoes = PainelLateral_listarSessoes();
    
    // 2. Cadastros para Autocomplete (Dados que não mudam a cada segundo)
    return {
      sessoes: sessoes,
      cadastros: {
        membros: PainelLateral_listarMembrosCompleto(), 
        procuradores: PainelLateral_listarProcuradoresCadastrados()
      }
    };
  } catch (err) {
    Logger.log('Erro em obterPacoteInicial: ' + err.message);
    return { sessoes: [], cadastros: { membros: [], procuradores: [] } };
  }
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

// =================================================================================
// [BLOCO] CONFIGURAÇÃO DE TEMPLATES E CONSTANTES
// =================================================================================
const TEMPLATES_IDS = {
  DELIBERATIVO: '1Viq_bKZstJ4EharqLSn5HQf_HgWZVcITHDD3UJ5DeBI',
  PLENO: '1IkcUTLayOYuu4IKbiROgTygLk4y3Xs_ELV-KqvXGIr0'
};

// =================================================================================
// [BLOCO] DOCS ENGINE - GERAÇÃO VIA TEMPLATE
// =================================================================================

/**
 * Função principal: Clona o template e substitui as tags {{Tabela.Campo}}
 * Responsabilidade única: operações de Drive e Docs.
 * NÃO acessa a planilha. Recebe o mapa de substituição completo do cliente.
 *
 * @param {string} templateId  - ID do Google Doc template
 * @param {Object} subs        - Objeto { 'chave.campo': 'valor', ... }
 */
function PainelLateral_gerarDocumentoFicha(templateId, subs) {
  let tempFileId = null;
  try {
    // 1. Cópia do template
    tempFileId = DriveApp.getFileById(templateId)
                   .makeCopy('__temp_sdp__')
                   .getId();

    // 2. Substituições — single open, sem save intermediário
    const docTemp  = DocumentApp.openById(tempFileId);
    const bodyTemp = docTemp.getBody();

    for (const chave in subs) {
      // Padrão: chave = "tabProcessos.Processo" → tag = {{tabProcessos.Processo}}
      const pattern = '\\{\\{' + chave.replace(/\./g, '\\.') + '\\}\\}';
      bodyTemp.replaceText(pattern, String(subs[chave] ?? '') || ' ');
    }

    // Limpa tags não resolvidas
    bodyTemp.replaceText('\\{\\{[^}]+\\}\\}', ' ');

    // 3. Importa para o documento ativo
    const bodyAtivo = DocumentApp.getActiveDocument().getBody();
    bodyAtivo.clear();

    for (let i = 0; i < bodyTemp.getNumChildren(); i++) {
      const el   = bodyTemp.getChild(i).copy();
      const tipo = el.getType();
      if      (tipo === DocumentApp.ElementType.PARAGRAPH)  bodyAtivo.appendParagraph(el.asParagraph());
      else if (tipo === DocumentApp.ElementType.TABLE)       bodyAtivo.appendTable(el.asTable());
      else if (tipo === DocumentApp.ElementType.LIST_ITEM)   bodyAtivo.appendListItem(el.asListItem());
    }

    return 'Documento gerado com sucesso.';

  } catch (err) {
    throw new Error('PainelLateral_gerarDocumentoFicha: ' + err.message);
  } finally {
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e) {}
    }
  }
}

// =================================================================================
// [BLOCO] AUXILIARES DE BUSCA PARA DOCUMENTOS
// =================================================================================

/**
 * Busca o gênero do membro para flexão gramatical no documento.
 * @param {string} nomeCompleto
 * @returns {string} 'Masculino' | 'Feminino'
 */
function PainelLateral_obterGeneroMembro(nomeCompleto) {
  if (!nomeCompleto) return "Masculino";
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName("tabMembros");
    const dados = sheet.getDataRange().getValues();
    const mapa = getMapaColunas(sheet);
    const iNome = mapa["nome"] - 1;
    const iGen = mapa["gênero"] - 1;

    const membro = dados.find(r => r[iNome] === nomeCompleto.trim());
    return membro ? membro[iGen] : "Masculino";
  } catch (e) {
    return "Masculino";
  }
}

/**
 * Utilitário para evitar erro de texto vazio no Google Docs
 */
function validarTexto(valor, padrao = " ") {
  if (valor === null || valor === undefined || valor === "") return padrao;
  return String(valor);
}