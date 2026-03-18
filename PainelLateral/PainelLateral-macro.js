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

// ═══════════════════════════════════════════════════════
// GESTÃO DO DOCUMENTO DE SESSÃO
// ═══════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
// CONSTANTES DE FORMATAÇÃO
// ═══════════════════════════════════════════════════════════════════
const _F   = 'Lora';
const _T1  = 12;   // Título 1 — FICHA DE VOTAÇÃO
const _T2  = 11;   // Título 2 — VOTO / ACÓRDÃO
const _NRM = 10;   // Texto normal
const _SM  =  9;   // Quórum e assinaturas

const _AL  = DocumentApp.HorizontalAlignment.LEFT;
const _AC  = DocumentApp.HorizontalAlignment.CENTER;
const _AJ  = DocumentApp.HorizontalAlignment.JUSTIFY;

const _EMENTA_INDENT = 226.8; // 8 cm em pontos (1 cm ≈ 28.35 pt)


// ═══════════════════════════════════════════════════════════════════
// PONTO DE ENTRADA — SEM TEMPLATE, SEM MAKECOPY
// ═══════════════════════════════════════════════════════════════════

/**
 * Gera a ficha inteiramente via código no documento ativo.
 * Recebe o objeto de dados completo montado pelo cliente.
 * Zero acesso à planilha. Zero operações de Drive.
 * @param {Object} d - Dados completos para geração
 */
function PainelLateral_gerarDocumentoFicha(d) {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    body.clear();

    if (d.isPleno) {
      _PL_fichaPleno(body, d);
    } else {
      _PL_fichaDeliberativo(body, d);
    }

    return 'Documento gerado.';
  } catch (err) {
    throw new Error('PainelLateral_gerarDocumentoFicha: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// ESTRUTURA DELIBERATIVO
// ═══════════════════════════════════════════════════════════════════

function _PL_fichaDeliberativo(body, d) {
  _PL_T1(body, 'FICHA DE VOTAÇÃO');
  _PL_refs(body, d, false);
  _PL_expediente(body, d.expediente);
  _PL_ementa(body, d.ementa, false);
  _PL_T2(body, 'VOTO');
  _PL_normal(body, d.voto);
  _PL_data(body, d.dataExtenso);
  _PL_orgao(body, 'Órgão Deliberativo do SDP da OAB-GO', false);
  _PL_quorumDelib(body, d);
}


// ═══════════════════════════════════════════════════════════════════
// ESTRUTURA PLENO
// ═══════════════════════════════════════════════════════════════════

function _PL_fichaPleno(body, d) {
  _PL_T1(body, 'FICHA DE VOTAÇÃO');
  _PL_refs(body, d, true);
  _PL_ementa(body, d.ementa, true);
  _PL_T2(body, 'ACÓRDÃO');
  _PL_normal(body, d.voto);
  _PL_expediente(body, d.expediente);
  _PL_resultado(body, d);
  _PL_data(body, d.dataExtenso);
  _PL_orgao(body, 'Pleno do Sistema de Defesa das Prerrogativas da OAB-GO', true);
  _PL_assinaturas(body, d);
  _PL_membrosPleno(body, d.membros);
}


// ═══════════════════════════════════════════════════════════════════
// BLOCOS DE CONTEÚDO
// ═══════════════════════════════════════════════════════════════════

// Título 1: upcase, centralizado, bold, antes 0, depois 12
function _PL_T1(body, texto) {
  var p = body.appendParagraph(texto.toUpperCase());
  p.setAlignment(_AC);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(0);
  p.setSpacingAfter(12);
  _PL_fmt(p.editAsText(), _T1, true);
}

// Título 2: upcase, centralizado, bold, antes 12, depois 6
function _PL_T2(body, texto) {
  var p = body.appendParagraph(texto.toUpperCase());
  p.setAlignment(_AC);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(12);
  p.setSpacingAfter(6);
  _PL_fmt(p.editAsText(), _T2, true);
}

// Bloco de referências completo
function _PL_refs(body, d, ePleno) {
  _PL_rl(body, 'Referências:',  '',               'all',   0);
  _PL_rl(body, 'Processo nº ',  _s(d.processo),   'valor', 0);
  _PL_rl(body, 'Requerente: ',  _s(d.requerente), 'label', 0);
  _PL_rl(body, 'Requerido: ',   _s(d.requerido),  'label', 0);
  if (ePleno) {
    _PL_rl(body, 'Relator: ',    _s(d.relator),    'label', 12);
  } else {
    _PL_rl(body, 'Procurador: ', _s(d.procurador), 'label', 12);
  }
}

/**
 * Linha de referência com formatação mista.
 * mode: 'all' = tudo bold | 'label' = label bold | 'valor' = valor bold
 */
function _PL_rl(body, label, valor, mode, spaceAfter) {
  var texto = label + valor;
  var p = body.appendParagraph(texto);
  p.setAlignment(_AL);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(0);
  p.setSpacingAfter(spaceAfter || 0);

  var t  = p.editAsText();
  var ll = label.length;
  var tl = texto.length;

  _PL_fmt(t, _NRM, false);

  if (mode === 'all'   && tl > 0)           t.setBold(0,  tl - 1, true);
  if (mode === 'label' && ll > 0)           t.setBold(0,  ll - 1, true);
  if (mode === 'valor' && tl > ll)          t.setBold(ll, tl - 1, true);
}

// Ementa: label bold, recuo 8 cm, justificado
function _PL_ementa(body, texto, ePleno) {
  var sep   = ePleno ? '. ' : ': ';
  var label = 'EMENTA' + sep;
  var full  = label + _s(texto);

  var p = body.appendParagraph(full);
  p.setAlignment(_AJ);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(0);
  p.setSpacingAfter(6);
  p.setIndentFirstLine(_EMENTA_INDENT);  // ← primeira linha
  p.setIndentStart(_EMENTA_INDENT);      // ← demais linhas

  var t = p.editAsText();
  _PL_fmt(t, _NRM, false);
  if (label.length > 0) t.setBold(0, label.length - 1, true);
}

// Expediente: Texto normal, label bold
function _PL_expediente(body, texto) {
  var label = 'Expediente: ';
  var full  = label + _s(texto);

  var p = body.appendParagraph(full);
  p.setAlignment(_AJ);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(0);
  p.setSpacingAfter(6);

  var t = p.editAsText();
  _PL_fmt(t, _NRM, false);
  t.setBold(0, label.length - 1, true);
}

// Texto normal justificado (voto / acórdão) — suporte a múltiplas linhas
function _PL_normal(body, texto) {
  var linhas = _s(texto).split('\n');
  if (linhas.length === 0) linhas = [''];
  linhas.forEach(function(linha, i) {
    var p = body.appendParagraph(linha);
    p.setAlignment(_AJ);
    p.setLineSpacing(1.0);
    p.setSpacingBefore(0);
    p.setSpacingAfter(i === linhas.length - 1 ? 6 : 0);
    _PL_fmt(p.editAsText(), _NRM, false);
  });
}

// Resultado da votação: linha separadora + título bold + placar normal
function _PL_resultado(body, d) {
  // Linha horizontal como "borda superior" da seção
  var hr = body.appendHorizontalRule();
  try {
    var hrPara = hr.getParent().asParagraph();
    hrPara.setSpacingBefore(6);
    hrPara.setSpacingAfter(0);
  } catch (e) {}

  var pTit = body.appendParagraph('Resultado da votação');
  pTit.setAlignment(_AL);
  pTit.setLineSpacing(1.0);
  pTit.setSpacingBefore(0);
  pTit.setSpacingAfter(0);
  _PL_fmt(pTit.editAsText(), _NRM, true);

  var placar = 'Voto com o relator: ' + _s(d.votosRelator) +
               ' | ' + '0' +
               ' | ' + '0' +
               ' | Total votantes: '  + _s(d.totalVotantes);

  var pNum = body.appendParagraph(placar);
  pNum.setAlignment(_AL);
  pNum.setLineSpacing(1.0);
  pNum.setSpacingBefore(0);
  pNum.setSpacingAfter(6);
  _PL_fmt(pNum.editAsText(), _NRM, false);
}

// Data por extenso: bold, esquerda, antes 12, depois 0
function _PL_data(body, dataExtenso) {
  var p = body.appendParagraph(_s(dataExtenso));
  p.setAlignment(_AL);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(12);
  p.setSpacingAfter(0);
  _PL_fmt(p.editAsText(), _NRM, true);
}

// Nome do órgão: centralizado, bold, espaçamentos conforme tipo
function _PL_orgao(body, nome, ePleno) {
  var p = body.appendParagraph(nome);
  p.setAlignment(_AC);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(ePleno ? 12 : 18);
  p.setSpacingAfter(ePleno  ? 42 :  6);
  _PL_fmt(p.editAsText(), _NRM, true);
}

// Quórum do Deliberativo: 4 linhas, 9pt, dados em bold, flex de gênero
function _PL_quorumDelib(body, d) {
  var secLabel = d.generoSecretario === 'Feminino'
    ? 'secretária da mesa: '
    : 'secretário da mesa: ';
  var relLabel = d.generoRelator === 'Feminino'
    ? 'relatora: '
    : 'relator: ';

  var linhas = [
    { pre: 'Sessão deliberativa presidida por ',    dado: _s(d.presidente) + ',' },
    { pre: secLabel,                                dado: _s(d.secretario) + ',' },
    { pre: relLabel,                                dado: _s(d.relator)    + ',' },
    { pre: 'com participação dos demais membros: ', dado: _s(d.membros)          }
  ];

  linhas.forEach(function(l) {
    var full = l.pre + l.dado;
    var p = body.appendParagraph(full);
    p.setAlignment(_AL);
    p.setLineSpacing(1.0);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);

    var t  = p.editAsText();
    var pl = l.pre.length;
    var dl = l.dado.length;

    _PL_fmt(t, _SM, false);
    if (dl > 0) t.setBold(pl, pl + dl - 1, true);
  });
}

// Campos de assinatura do Pleno: tabela 2×2, 9pt
function _PL_assinaturas(body, d) {
  var LARGURA_TOTAL  = 453;
  var LARGURA_ESPACO = 30;
  var LARGURA_COL    = (LARGURA_TOTAL - LARGURA_ESPACO) / 2;

  // Tabela de UMA linha e 3 colunas
  var tabela = body.appendTable([['', '', '']]);

  // Remove todas as bordas da tabela
  var attrs = {};
  attrs[DocumentApp.Attribute.BORDER_COLOR] = '#ffffff';
  attrs[DocumentApp.Attribute.BORDER_WIDTH] = 0;
  tabela.setAttributes(attrs);

  tabela.setColumnWidth(0, LARGURA_COL);
  tabela.setColumnWidth(1, LARGURA_ESPACO);
  tabela.setColumnWidth(2, LARGURA_COL);

  var linha = tabela.getRow(0);

  [
    { col: 0, label: 'Assinatura do presidente da sessão', nome: _s(d.presidente) },
    { col: 2, label: 'Assinatura do secretário da mesa',   nome: _s(d.secretario) }
  ].forEach(function(info) {
    var cel = linha.getCell(info.col);
    cel.setPaddingTop(0);
    cel.setPaddingBottom(2);
    cel.setPaddingLeft(0);
    cel.setPaddingRight(0);

    // Insere o HR no parágrafo que já existe na célula (índice 0)
    // Não cria parágrafo extra — elimina o espaço acima da linha
    var pHr = cel.getChild(0).asParagraph();
    pHr.appendHorizontalRule();
    pHr.setSpacingBefore(0);
    pHr.setSpacingAfter(0);
    pHr.setLineSpacing(1.0);

    // Label: "Assinatura do ..."
    var pLabel = cel.appendParagraph(info.label);
    pLabel.setAlignment(_AC);
    pLabel.setLineSpacing(1.0);
    pLabel.setSpacingBefore(3);
    pLabel.setSpacingAfter(0);
    _PL_fmt(pLabel.editAsText(), _SM, false);

    // Nome em negrito
    var pNome = cel.appendParagraph(info.nome);
    pNome.setAlignment(_AC);
    pNome.setLineSpacing(1.0);
    pNome.setSpacingBefore(0);
    pNome.setSpacingAfter(0);
    _PL_fmt(pNome.editAsText(), _SM, true);
  });

  // Célula espaçadora central invisível
  var celEspaco = linha.getCell(1);
  celEspaco.setPaddingTop(0);
  celEspaco.setPaddingBottom(0);
  celEspaco.setPaddingLeft(0);
  celEspaco.setPaddingRight(0);
}

// Membros participantes do Pleno: 9pt, dados bold, antes 6
function _PL_membrosPleno(body, membros) {
  var label = 'Membros participantes: ';
  var valor = _s(membros);
  var full  = label + valor;

  var p = body.appendParagraph(full);
  p.setAlignment(_AL);
  p.setLineSpacing(1.0);
  p.setSpacingBefore(6);
  p.setSpacingAfter(0);

  var t  = p.editAsText();
  var ll = label.length;
  var vl = valor.length;

  _PL_fmt(t, _SM, false);
  if (vl > 0) t.setBold(ll, ll + vl - 1, true);
}


// ═══════════════════════════════════════════════════════════════════
// AUXILIARES
// ═══════════════════════════════════════════════════════════════════

/** Aplica fonte Lora + tamanho + bold ao objeto Text */
function _PL_fmt(textObj, size, bold) {
  textObj.setFontFamily(_F);
  textObj.setFontSize(size);
  textObj.setBold(bold);
}

/** Retorna string segura, nunca null/undefined */
function _s(val) {
  return (val != null && String(val).trim() !== '') ? String(val) : '';
}

/** Converte "dd/MM/yyyy" → "Goiânia, D de mês de YYYY" */
function _PL_formatDataExtenso(dataStr) {
  if (!dataStr) return '';
  var partes = dataStr.split('/');
  if (partes.length !== 3) return dataStr;
  var meses = ['janeiro','fevereiro','março','abril','maio','junho',
               'julho','agosto','setembro','outubro','novembro','dezembro'];
  return 'Goiânia, ' + parseInt(partes[0], 10) +
         ' de ' + (meses[parseInt(partes[1], 10) - 1] || '') +
         ' de ' + partes[2];
}

/** Lê tabMembros e retorna mapa { nome: gênero } */
function _PL_buildGeneroMap(ss) {
  var mapa = {};
  try {
    var sheet = ss.getSheetByName('tabMembros');
    if (!sheet) return mapa;
    var dados = sheet.getDataRange().getValues();
    var m     = getMapaColunas(sheet);
    var iNome = (m['nome']    || 2) - 1;
    var iGen  = (m['gênero'] || 3) - 1;
    dados.slice(1).forEach(function(r) {
      var nome = (r[iNome] || '').toString().trim();
      if (nome) mapa[nome] = (r[iGen] || 'Masculino').toString();
    });
  } catch (e) {
    Logger.log('_PL_buildGeneroMap: ' + e.message);
  }
  return mapa;
}





// ── Auxiliares privadas ──────────────────────────────────────────

function _rep(body, pattern, replacement) {
  body.replaceText(pattern, replacement);
}

function _v(val) {
  return (val != null && val !== '') ? String(val) : ' ';
}