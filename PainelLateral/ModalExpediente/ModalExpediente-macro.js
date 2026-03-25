/**
 * @fileoverview Backend — Modal de Expediente Predefinido (standalone)
 * Gerencia abertura do modal, leitura/escrita em tabConfiguracoes e
 * inserção do texto resolvido no documento ativo do Google Docs.
 *
 * Funções expostas ao front-end (google.script.run):
 *   modalExpediente_abrir()                   ← chamada do FAB no PainelLateral
 *   modalExpediente_buscarDados()             ← carregamento inicial (configs + membros)
 *   modalExpediente_salvarConfig(chave, valor)
 *   modalExpediente_inserirNoDocumento(texto)
 */

// ─────────────────────────────────────────────────────────────────────────────
// ABERTURA DO MODAL
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Abre o modal de expediente predefinido, recebendo os dados da ficha atual.
 * @param {string} fichaId  ID da ficha selecionada no painel lateral.
 * @param {string} sessaoId ID da sessão selecionada.
 * @param {string} orgao    Órgão da sessão (ex: "Pleno", "1ª Turma", etc).
 */
function modalExpediente_abrir(fichaId, sessaoId, orgao) {
  const template = HtmlService.createTemplateFromFile('ModalExpediente-front');
  template.fichaId = fichaId || '';
  template.sessaoId = sessaoId || '';
  template.orgao = orgao || '';
  const html = template.evaluate()
    .setTitle('Expedientes Predefinidos')
    .setWidth(500)
    .setHeight(310);
  DocumentApp.getUi().showModalDialog(html, ' ');
}

// ─────────────────────────────────────────────────────────────────────────────
// CARGA INICIAL — configs + lista de membros em uma única chamada
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Retorna em uma única chamada:
 *   - configs : Array<{chave, valor}> de tabConfiguracoes (chave começa com "Expediente")
 *   - membros : Array<{nome, genero}> de tabMembros
 *   - relator : obtido do documento atual
 *
 * @returns {{ sucesso: boolean, configs: Array, membros: Array }}
 */
function modalExpediente_buscarDados() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);

    // --- Configurações (expedientes) ---
    const sheetCfg = ss.getSheetByName('tabConfiguracoes');
    if (!sheetCfg) throw new Error("Aba 'tabConfiguracoes' não encontrada.");
    const mapaCfg = getMapaColunas(sheetCfg);
    const dadosCfg = sheetCfg.getDataRange().getValues();
    const iChave = (mapaCfg['chave'] || 1) - 1;
    const iValor = (mapaCfg['valor'] || 2) - 1;
    const configs = [];
    for (let i = 1; i < dadosCfg.length; i++) {
      const chave = String(dadosCfg[i][iChave] || '').trim();
      const valor = String(dadosCfg[i][iValor] || '').trim();
      if (chave && chave.toLowerCase().startsWith('expediente')) {
        configs.push({ chave, valor });
      }
    }

    // --- Membros ---
    const sheetMem = ss.getSheetByName('tabMembros');
    const membros = [];
    if (sheetMem) {
      const mapaMem = getMapaColunas(sheetMem);
      const dadosMem = sheetMem.getDataRange().getValues();
      const iNome = (mapaMem['nome'] || 2) - 1;
      const iGenero = (mapaMem['gênero'] || mapaMem['genero'] || 4) - 1;
      for (let i = 1; i < dadosMem.length; i++) {
        const nome = String(dadosMem[i][iNome] || '').trim();
        const genero = String(dadosMem[i][iGenero] || 'Masculino').trim();
        if (nome) membros.push({ nome, genero });
      }
    }

    // --- Relator do documento ativo ---
    let relatorDocumento = '';
    try {
      relatorDocumento = extrairRelatorDoDocumento();
    } catch (e) {
      console.error('Falha ao extrair relator (ignorado):', e);
      relatorDocumento = '';
    }

    return { 
      sucesso: true, 
      configs, 
      membros,
      relatorDocumento 
    };
  } catch (e) {
    console.error('modalExpediente_buscarDados: ' + e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Extrai o nome do relator do documento ativo (após a etiqueta "Relator:").
 */
function extrairRelatorDoDocumento() {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      Logger.log('Nenhum documento ativo.');
      return '';
    }
    const body = doc.getBody();
    if (!body) {
      Logger.log('Corpo do documento indisponível.');
      return '';
    }

    const numChildren = body.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      const elemento = body.getChild(i);
      if (elemento.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const texto = elemento.asParagraph().getText();
        const textoLower = texto.toLowerCase();

        // Procura por "relator:" ou "relatora:"
        if (textoLower.includes('relator:')) {
          Logger.log('Encontrado "relator:" no parágrafo: ' + texto);
          const partes = texto.split(/relator:\s*/i);
          if (partes.length > 1) {
            let nome = partes[1].trim();
            nome = nome.replace(/[.,;!?]+$/, '').trim();
            Logger.log('Nome extraído: ' + nome);
            return nome;
          }
        } else if (textoLower.includes('relatora:')) {
          Logger.log('Encontrado "relatora:" no parágrafo: ' + texto);
          const partes = texto.split(/relatora:\s*/i);
          if (partes.length > 1) {
            let nome = partes[1].trim();
            nome = nome.replace(/[.,;!?]+$/, '').trim();
            Logger.log('Nome extraído: ' + nome);
            return nome;
          }
        }
      }
    }
    Logger.log('Nenhuma etiqueta "relator:" ou "relatora:" encontrada.');
    return '';
  } catch (e) {
    Logger.log('Erro em extrairRelatorDoDocumento: ' + e.toString());
    return '';
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// GRAVAÇÃO
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Cria ou atualiza uma linha em tabConfiguracoes (busca case-insensitive pela chave).
 *
 * @param {string} chave  Ex: "Expediente voto relator aprovado unanimidade"
 * @param {string} valor  Texto do expediente
 * @returns {{ sucesso: boolean, tipo: 'criado'|'atualizado' }}
 */
function modalExpediente_salvarConfig(chave, valor) {
  try {
    if (!chave || !chave.trim()) throw new Error('Chave não pode ser vazia.');
    if (!valor || !valor.trim()) throw new Error('Valor não pode ser vazio.');

    const ss        = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet     = ss.getSheetByName('tabConfiguracoes');
    if (!sheet) throw new Error("Aba 'tabConfiguracoes' não encontrada.");

    const mapa      = getMapaColunas(sheet);
    const dados     = sheet.getDataRange().getValues();
    const iChave    = (mapa['chave'] || 1) - 1;
    const iValor    = (mapa['valor'] || 2) - 1;
    const chaveBusca = chave.trim().toLowerCase();

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][iChave] || '').trim().toLowerCase() === chaveBusca) {
        sheet.getRange(i + 1, iValor + 1).setValue(valor.trim());
        return { sucesso: true, tipo: 'atualizado' };
      }
    }

    // Não encontrado: cria nova linha
    const numCols   = Object.keys(mapa).length;
    const novaLinha = new Array(numCols).fill('');
    novaLinha[iChave] = chave.trim();
    novaLinha[iValor] = valor.trim();
    sheet.appendRow(novaLinha);

    return { sucesso: true, tipo: 'criado' };

  } catch (e) {
    throw new Error('Erro ao salvar configuração: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// INSERÇÃO NO DOCUMENTO ATIVO
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Localiza a etiqueta "Expediente:" no documento ativo do Google Docs e
 * substitui todo o conteúdo da mesma linha pelo texto recebido.
 *
 * body.replaceText() usa Java regex; "." não cruza newlines, portanto
 * apenas o parágrafo da etiqueta é afetado.**/
 /**
 * Insere o texto no documento, substituindo a linha do "Expediente:"
 * e flexionando {{do}} e {{relator}} com base no relator encontrado.
 */
/**
 * Insere o texto já flexionado no documento, substituindo a linha do "Expediente:".
 */
function modalExpediente_inserirNoDocumento(textoFinal) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const ETIQUETA = 'Expediente:';

    const encontrado = body.findText(ETIQUETA);
    if (!encontrado) {
      return { sucesso: false, erro: 'Etiqueta "Expediente:" não encontrada.' };
    }

    const elemento = encontrado.getElement();
    const textObj = elemento.asText();
    const textoCompleto = textObj.getText();
    const posEtiqueta = textoCompleto.indexOf(ETIQUETA);
    if (posEtiqueta === -1) {
      return { sucesso: false, erro: 'Etiqueta não encontrada no texto.' };
    }

    // Captura a formatação de fonte do primeiro caractere (ou de um ponto de referência)
    // Vamos usar a formatação do caractere na posição 0 do elemento
    const fontFamily = textObj.getFontFamily(0);
    const fontSize = textObj.getFontSize(0);
    const isBold = textObj.isBold(0);

    // Constrói o novo texto
    const prefixo = textoCompleto.substring(0, posEtiqueta + ETIQUETA.length);
    const novoTexto = prefixo + ' ' + textoFinal;

    // Substitui
    textObj.setText(novoTexto);

    // Aplica a formatação de fonte base a todo o texto
    if (fontFamily) textObj.setFontFamily(0, novoTexto.length - 1, fontFamily);
    if (fontSize) textObj.setFontSize(0, novoTexto.length - 1, fontSize);

    // Ajusta negrito: etiqueta em negrito, conteúdo normal
    const novaPosEtiqueta = novoTexto.indexOf(ETIQUETA);
    if (novaPosEtiqueta !== -1) {
      textObj.setBold(novaPosEtiqueta, novaPosEtiqueta + ETIQUETA.length - 1, true);
    }
    const inicioConteudo = novaPosEtiqueta + ETIQUETA.length + 1;
    const fimConteudo = novoTexto.length - 1;
    if (inicioConteudo <= fimConteudo) {
      textObj.setBold(inicioConteudo, fimConteudo, false);
    }

    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}