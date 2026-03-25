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
    const iFExpediente = (mF['expediente'] || 6) - 1;
    const iFMembros  = (mF['membros']    ? mF['membros'] - 1 : null); // ← NOVO

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
          expediente:     r[iFExpediente] || '',
          membros:        iFMembros !== null ? (r[iFMembros] || '') : null, // ← string original, sem parse

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

/**
 * Retorna a lista completa de nomes da tabMembros para o cache de autocompletes.
 */
function PainelLateral_listarMembrosCompleto() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabMembros');
    if (!sheet) return [];

    const dados = sheet.getDataRange().getValues();
    const mapa = getMapaColunas(sheet);
    const iNome = (mapa['nome'] || 2) - 1;
    const iGenero = (mapa['gênero'] || mapa['genero'] || 3) - 1; // ajuste conforme sua planilha

    return dados.slice(1)
      .map(row => ({
        nome: (row[iNome] || '').toString().trim(),
        genero: (row[iGenero] || 'Masculino').toString().trim()
      }))
      .filter(item => item.nome !== '');
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

/**
 * Salva a lista de membros participantes diretamente na ficha selecionada.
 * @param {string|number} fichaId
 * @param {Array<string>} listaMembros
 */
/**
 * Salva a lista de membros participantes diretamente na ficha selecionada.
 * @param {string|number} fichaId
 * @param {Array<string>} listaMembros
 */
function PainelLateral_salvarMembrosNaFicha(fichaId, listaMembros) {
  try {
    const ss = PainelLateral_getPlanilha();
    const sheetFichas = ss.getSheetByName('tabFichas');
    if (!sheetFichas) throw new Error('Aba tabFichas não encontrada.');

    const mapa = getMapaColunas(sheetFichas);
    const dados = sheetFichas.getDataRange().getValues();
    
    const iId = (mapa['id'] || 1) - 1;
    const iMembros = mapa['membros']; // Certifique-se que existe a coluna [membros] na tabFichas
    
    if (!iMembros) throw new Error('Coluna "membros" não encontrada na tabFichas.');

    // Localiza a linha da ficha
    let linhaFicha = -1;
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][iId] == fichaId) {
        linhaFicha = i + 1;
        break;
      }
    }

    if (linhaFicha === -1) throw new Error('Ficha não encontrada para salvar membros.');

    // Salva concatenado por ponto e vírgula (padronizado)
    const listaTexto = listaMembros.join(';');
    sheetFichas.getRange(linhaFicha, iMembros).setValue(listaTexto);

    return { sucesso: true, listaFormatada: listaTexto };
  } catch (err) {
    throw new Error('Erro ao salvar membros na ficha: ' + err.message);
  }
}

// =================================================================================
// [BLOCO] CONFIGURAÇÃO DE TEMPLATES E CONSTANTES
// =================================================================================
const TEMPLATES_IDS = {
  DELIBERATIVO: '14K8XDmy92dpSOKRRi14G4dNSPl6cPkSNA19RCJg1--U',
  PLENO: '1IkcUTLayOYuu4IKbiROgTygLk4y3Xs_ELV-KqvXGIr0'
};

// =================================================================================
// [BLOCO] DOCS ENGINE - GERAÇÃO DIRETA DO TEMPLATE (SEM CACHE)
// =================================================================================

/**
 * Copia todo o conteúdo de um documento de origem para o documento ativo.
 * @param {string} sourceDocId - ID do documento de origem
 */
function copiarDocumentoParaAtivo(sourceDocId) {
  const source = DocumentApp.openById(sourceDocId);
  const bodySource = source.getBody();
  const bodyAtivo = DocumentApp.getActiveDocument().getBody();

  bodyAtivo.clear();
  for (let i = 0; i < bodySource.getNumChildren(); i++) {
    const el = bodySource.getChild(i).copy();
    const tipo = el.getType();
    if (tipo === DocumentApp.ElementType.PARAGRAPH) {
      bodyAtivo.appendParagraph(el.asParagraph());
    } else if (tipo === DocumentApp.ElementType.TABLE) {
      bodyAtivo.appendTable(el.asTable());
    } else if (tipo === DocumentApp.ElementType.LIST_ITEM) {
      bodyAtivo.appendListItem(el.asListItem());
    }
    // Outros tipos (HEADING, etc.) podem ser adicionados se necessário
  }

  // Remove possível parágrafo vazio no início (caso exista)
  if (bodyAtivo.getNumChildren() > 0) {
    const first = bodyAtivo.getChild(0);
    if (first.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = first.asParagraph().getText();
      if (text.trim() === '') {
        first.removeFromParent();
      }
    }
  }
}

/**
 * Substitui todos os placeholders no documento ativo usando replaceText.
 * @param {Object} subs - Mapa de substituição { 'chave': 'valor' }
 */
function preencherDocumentoComSubs(subs) {
  const body = DocumentApp.getActiveDocument().getBody();
  for (const chave in subs) {
    const placeholder = '{{' + chave + '}}';
    const valor = String(subs[chave] ?? '');
    // Escapa caracteres especiais no placeholder para uso em regex
    const placeholderEscaped = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(placeholderEscaped, valor);
  }
  return true;
}

/**
 * Função principal chamada pelo front-end para cada ficha.
 * @param {Object} subs - Mapa de substituição
 * @param {string} templateId - ID do template (Deliberativo ou Pleno)
 */
function gerarDocumentoParaFicha(subs, templateId) {
  copiarDocumentoParaAtivo(templateId);
  return preencherDocumentoComSubs(subs);
}

/**
 * Atualiza a lista de membros no documento ativo preservando toda a formatação.
 * @param {string} novaLista - String com os nomes dos membros separados por vírgula e espaço.
 */
function atualizarMembrosNoDocumento(novaLista) {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const padroes = [
      /Membros participantes:\s*/i,
      /com participação dos demais membros:\s*/i
    ];

    let encontrado = false;
    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = child.asParagraph();
        const textElement = paragraph.editAsText();
        const texto = textElement.getText();
        
        for (let padrao of padroes) {
          const match = padrao.exec(texto);
          if (match) {
            const startPos = match.index + match[0].length; // posição após os dois pontos
            const endPos = texto.length;

            // Captura a formatação do primeiro caractere do texto antigo (se existir)
            let fontFamily = null;
            let fontSize = null;
            let isBold = null;
            let isItalic = null;
            
            if (endPos > startPos) {
              // Pega a formatação do primeiro caractere da lista antiga
              const oldStart = startPos;
              fontFamily = textElement.getFontFamily(oldStart);
              fontSize = textElement.getFontSize(oldStart);
              isBold = textElement.isBold(oldStart);
              isItalic = textElement.isItalic(oldStart);
              
              // Remove o texto antigo
              textElement.deleteText(oldStart, endPos - 1);
            }
            
            // Insere o novo texto (pode ser vazio)
            const newStart = startPos;
            if (novaLista !== '') {
              textElement.insertText(newStart, novaLista);
              const newEnd = newStart + novaLista.length - 1;
              
              // Aplica a formatação capturada (ou usa fallback do caractere anterior)
              if (fontFamily) textElement.setFontFamily(newStart, newEnd, fontFamily);
              if (fontSize) textElement.setFontSize(newStart, newEnd, fontSize);
              if (isBold !== null) textElement.setBold(newStart, newEnd, isBold);
              if (isItalic !== null) textElement.setItalic(newStart, newEnd, isItalic);
              
              // Se não havia texto antigo (primeira vez), copia do caractere anterior
              if (fontFamily === null && newStart > 0) {
                const prevChar = newStart - 1;
                const fallbackFont = textElement.getFontFamily(prevChar);
                const fallbackSize = textElement.getFontSize(prevChar);
                const fallbackBold = textElement.isBold(prevChar);
                const fallbackItalic = textElement.isItalic(prevChar);
                if (fallbackFont) textElement.setFontFamily(newStart, newEnd, fallbackFont);
                if (fallbackSize) textElement.setFontSize(newStart, newEnd, fallbackSize);
                if (fallbackBold !== null) textElement.setBold(newStart, newEnd, fallbackBold);
                if (fallbackItalic !== null) textElement.setItalic(newStart, newEnd, fallbackItalic);
              }
            } else {
              // Se a nova lista é vazia, não inserimos nada; o campo permanece vazio
              // A formatação já foi preservada com a remoção
            }
            
            encontrado = true;
            break;
          }
        }
        if (encontrado) break;
      }
    }
    if (!encontrado) {
      throw new Error('Não foi possível localizar o local da lista de membros no documento.');
    }
    return true;
  } catch (err) {
    throw new Error('Erro ao atualizar membros: ' + err.message);
  }
}

// =================================================================================
// [BLOCO] FUNÇÕES DE SALVAMENTO DE FICHA
// =================================================================================

/**
 * Extrai o texto do expediente do documento ativo (após "Expediente:").
 */
function extrairExpedienteDoDocumento() {
  const body = DocumentApp.getActiveDocument().getBody();
  const encontrado = body.findText("Expediente:");
  if (!encontrado) return '';
  const elem = encontrado.getElement();
  const textoCompleto = elem.asText().getText();
  const pos = textoCompleto.indexOf('Expediente:');
  if (pos === -1) return '';
  return textoCompleto.substring(pos + 'Expediente:'.length).trim();
}

/**
 * Extrai o texto do voto do documento ativo.
 * Localiza o título "VOTO" e captura o texto até a linha que contém a data (ex: "Goiânia,").
 */
function extrairVotoDoDocumento() {
  const body = DocumentApp.getActiveDocument().getBody();
  const numChildren = body.getNumChildren();
  let voto = '';
  let capturando = false;
  for (let i = 0; i < numChildren; i++) {
    const elem = body.getChild(i);
    if (elem.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const texto = elem.asParagraph().getText();
      if (!capturando) {
        if (texto.trim().toUpperCase() === 'VOTO' || texto.trim().toUpperCase().startsWith('VOTO')) {
          capturando = true;
        }
      } else {
        if (texto.trim().startsWith('Goiânia,')) break;
        if (voto) voto += '\n';
        voto += texto;
      }
    }
  }
  return voto.trim();
}

/**
 * Extrai a lista de membros do documento ativo.
 * @returns {string} Texto após "Membros participantes:" ou "com participação dos demais membros:".
 */
function extrairMembrosDoDocumento() {
  const body = DocumentApp.getActiveDocument().getBody();
  const padroes = [
    /Membros participantes:\s*/i,
    /com participação dos demais membros:\s*/i
  ];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const texto = child.asParagraph().getText();
      for (let padrao of padroes) {
        const match = padrao.exec(texto);
        if (match) {
          const startPos = match.index + match[0].length;
          return texto.substring(startPos).trim();
        }
      }
    }
  }
  return '';
}

/**
 * Salva a ficha atual: Expediente, Voto e Membros, e consolida os membros da sessão.
 * @param {string|number} fichaId
 * @returns {Object} Status da operação.
 */
function PainelLateral_salvarFichaCompleta(fichaId) {
  try {
    // Extrai dados do documento
    const expediente = extrairExpedienteDoDocumento();
    const voto = extrairVotoDoDocumento();
    const membrosTexto = extrairMembrosDoDocumento();

    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetFichas = ss.getSheetByName('tabFichas');
    if (!sheetFichas) throw new Error('Aba tabFichas não encontrada.');
    const mapaF = getMapaColunas(sheetFichas);
    const dadosF = sheetFichas.getDataRange().getValues();
    const iFId = (mapaF['id'] || 1) - 1;
    const iFExpediente = mapaF['expediente'];
    const iFMembros = mapaF['membros'];
    const iFSessao = mapaF['idsessao'] || 2;

    if (!iFExpediente) throw new Error('Coluna "expediente" não encontrada em tabFichas.');
    if (!iFMembros) throw new Error('Coluna "membros" não encontrada em tabFichas.');

    // Localiza a linha da ficha
    let linhaFicha = -1;
    for (let i = 1; i < dadosF.length; i++) {
      if (dadosF[i][iFId] == fichaId) {
        linhaFicha = i + 1;
        break;
      }
    }
    if (linhaFicha === -1) throw new Error('Ficha ID ' + fichaId + ' não encontrada.');

    // Atualiza Expediente e Membros
    sheetFichas.getRange(linhaFicha, iFExpediente).setValue(expediente);
    sheetFichas.getRange(linhaFicha, iFMembros).setValue(membrosTexto);

    // Salva o Voto em tabVotos
    const sheetVotos = ss.getSheetByName('tabVotos');
    if (!sheetVotos) throw new Error('Aba tabVotos não encontrada.');
    const mapaV = getMapaColunas(sheetVotos);
    const dadosV = sheetVotos.getDataRange().getValues();
    const iVIdFicha = mapaV['idfichavotacao'] || 2;
    const iVVoto = mapaV['voto'] || 6;
    const iVTipoVoto = mapaV['tipovoto'];
    const iVRelator = mapaV['relator'];

    const iFRelator = mapaF['relator'] || 5;
    const relator = dadosF[linhaFicha - 1][iFRelator - 1] || '';

    let linhaVoto = -1;
    for (let i = 1; i < dadosV.length; i++) {
      if (dadosV[i][iVIdFicha - 1] == fichaId) {
        linhaVoto = i + 1;
        break;
      }
    }
    if (linhaVoto !== -1) {
      sheetVotos.getRange(linhaVoto, iVVoto).setValue(voto);
      if (iVTipoVoto) sheetVotos.getRange(linhaVoto, iVTipoVoto).setValue('Voto do relator');
      if (iVRelator) sheetVotos.getRange(linhaVoto, iVRelator).setValue(relator);
    } else {
      const maxCol = Math.max(...Object.values(mapaV));
      const novaLinha = new Array(maxCol).fill('');
      novaLinha[iVIdFicha - 1] = fichaId;
      novaLinha[iVVoto - 1] = voto;
      if (iVTipoVoto) novaLinha[iVTipoVoto - 1] = 'Voto do relator';
      if (iVRelator) novaLinha[iVRelator - 1] = relator;
      sheetVotos.appendRow(novaLinha);
    }

    // Consolida os membros da sessão
    const sessaoId = dadosF[linhaFicha - 1][iFSessao - 1];
    const consolidacao = PainelLateral_consolidarMembrosSessao(sessaoId);
    let membrosConsolidados = [];
    if (consolidacao.sucesso) {
      membrosConsolidados = consolidacao.membros ? consolidacao.membros.split(';') : [];
    } else {
      Logger.log('Consolidação falhou: ' + consolidacao.erro);
    }

    return {
      sucesso: true,
      fichaId: fichaId,
      expediente: expediente,
      voto: voto,
      membros: membrosTexto,
      membrosConsolidados: membrosConsolidados
    };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/**
 * Consolida os membros de todas as fichas de uma sessão,
 * atualizando o campo [Membros] da sessão em tabSessoes.
 * @param {string|number} sessaoId
 * @returns {Object} status da operação
 */
function PainelLateral_consolidarMembrosSessao(sessaoId) {
  try {
    const ss = PainelLateral_getPlanilha();
    const idBusca = isNaN(sessaoId) ? String(sessaoId) : Number(sessaoId);

    // 1. Obter todas as fichas da sessão
    const sheetFichas = ss.getSheetByName('tabFichas');
    if (!sheetFichas) throw new Error('Aba tabFichas não encontrada.');
    const mapaF = getMapaColunas(sheetFichas);
    const dadosF = sheetFichas.getDataRange().getValues();

    const iFSessao = (mapaF['idsessao'] || 2) - 1;
    const iFMembros = mapaF['membros']; // coluna membros (se existir)
    if (!iFMembros) throw new Error('Coluna "membros" não encontrada em tabFichas.');

    const membrosSet = new Set();
    for (let i = 1; i < dadosF.length; i++) {
      const linha = dadosF[i];
      if (linha[iFSessao] == idBusca) {
        const membrosStr = linha[iFMembros - 1] || '';
        if (membrosStr.trim() !== '') {
          // Divide por ponto e vírgula ou vírgula (compatibilidade)
          const nomes = membrosStr.split(/[;,]/).map(n => n.trim()).filter(n => n !== '');
          nomes.forEach(n => membrosSet.add(n));
        }
      }
    }

    // 2. Ordenar alfabeticamente
    const membrosConsolidados = Array.from(membrosSet).sort((a, b) => a.localeCompare(b));
    const membrosTexto = membrosConsolidados.join(';'); // sem espaços

    // 3. Atualizar a sessão em tabSessoes
    const sheetSessoes = ss.getSheetByName('tabSessoes');
    if (!sheetSessoes) throw new Error('Aba tabSessoes não encontrada.');
    const mapaS = getMapaColunas(sheetSessoes);
    const iSId = (mapaS['id'] || mapaS['id sessão'] || 1) - 1;
    const iSMembros = mapaS['membros'];
    if (!iSMembros) throw new Error('Coluna "membros" não encontrada em tabSessoes.');

    const dadosS = sheetSessoes.getDataRange().getValues();
    let linhaSessao = -1;
    for (let i = 1; i < dadosS.length; i++) {
      if (dadosS[i][iSId] == idBusca) {
        linhaSessao = i + 1;
        break;
      }
    }
    if (linhaSessao === -1) throw new Error('Sessão não encontrada para atualização.');

    sheetSessoes.getRange(linhaSessao, iSMembros).setValue(membrosTexto);
    return { sucesso: true, membros: membrosTexto };
  } catch (err) {
    Logger.log('Erro ao consolidar membros da sessão: ' + err.message);
    return { sucesso: false, erro: err.message };
  }
}

/**
 * Converte o documento ativo em PDF e salva na pasta correspondente.
 */
function PainelLateral_salvarDocumentoComoPDF(contexto) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const docFile = DriveApp.getFileById(doc.getId());
    const paiFolder = docFile.getParents().next(); // Pasta onde o projeto está

    // 1. Identifica o Órgão e define nomes
    const ePleno = contexto.orgao.toLowerCase().includes("pleno");
    const nomePastaAlvo = ePleno ? "Fichas pleno" : "Fichas deliberativo";
    const sufixoNome = ePleno ? "ficha_pleno" : "ficha_deliberativo";

    // 2. Busca a pasta
    let pastaAlvo;
    const folders = paiFolder.getFoldersByName(nomePastaAlvo);
    if (folders.hasNext()) {
      pastaAlvo = folders.next();
    } else {
      throw new Error("A pasta '" + nomePastaAlvo + "' não foi encontrada. Verifique o nome na mesma pasta deste projeto.");
    }

    // 3. Extrai o número do processo do conteúdo do documento (primeira ocorrência)
    const corpo = doc.getBody().getText();
    const matchProcesso = corpo.match(/\d{9,}/);
    const numProcesso = matchProcesso ? matchProcesso[0] : "000000000";

    // 4. Formata a data para o nome do arquivo usando a função auxiliar
    const dataFormatada = _formatarDataRelatorio(contexto.data);
    const nomeArquivo = `${numProcesso}_${sufixoNome}_${dataFormatada}`;

    // 5. Gera o PDF
    const pdfBlob = docFile.getAs('application/pdf');
    const novoPdf = pastaAlvo.createFile(pdfBlob).setName(nomeArquivo);

    return novoPdf.getUrl();
  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Exibe um alerta com os IDs da ficha e da sessão.
 * @param {string|number} fichaId
 * @param {string|number} sessaoId
 */
function PainelLateral_mostrarModalVotantes(fichaId, sessaoId) {
  const ui = DocumentApp.getUi();
  const mensagem = `ID da Ficha: ${fichaId}\nID da Sessão: ${sessaoId}`;
  ui.alert('Modificar Votantes', mensagem, ui.ButtonSet.OK);
}