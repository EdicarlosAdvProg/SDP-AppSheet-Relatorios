/**
 * @fileoverview Backend para Gestão de Sessões SDP — OAB/GO
 * Funções: CRUD sessões, fichas de votação, votos, upload de relatório em PDF.
 *
 * Tabelas envolvidas:
 *   tabSessoes  : Id | DataSessao | Órgão | Local/Sala | Presidente | Secretário | Membros | Procuradores | Expediente
 *   tabFichas   : Id | IdSessao | IdProcesso | Relator | Membros | Procuradores | Expediente
 *   tabVotos    : Id | IdFichaVotacao | IdProcesso | TipoVoto | Relator | Voto | Resultado | URL Voto
 *   tabProcessos: Id | Processo | Requerente | Requerido | Procurador | ... | Status
 *   tabMembros  : Id | Nome | Email | Gênero | Cargo
 */

// ─────────────────────────────────────────────────────────────────────────────
// ABERTURA DO FORMULÁRIO
// ─────────────────────────────────────────────────────────────────────────────

function formSessoes_abrirModal() {
  const template = HtmlService.createTemplateFromFile('FormSessoes-layout');
  const html = template.evaluate()
    .setTitle('SDP - Gestão de Sessões')
    .setWidth(1200)
    .setHeight(850);
  DocumentApp.getUi().showModalDialog(html, ' ');
}

// ─────────────────────────────────────────────────────────────────────────────
// CARGA INICIAL DE DADOS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Retorna sessões com contagem de processos pautados, membros para autocomplete
 * e órgãos únicos para o filtro/select.
 */
function formSessoes_buscarDadosCompletos() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);

    // ── tabSessoes ────────────────────────────────────────────────────────────
    const sheetSess = ss.getSheetByName('tabSessoes');
    if (!sheetSess) throw new Error("Aba 'tabSessoes' não encontrada.");
    const mapaSess  = getMapaColunas(sheetSess);
    const dadosSess = sheetSess.getDataRange().getValues();
    dadosSess.shift(); // remove cabeçalho

    // ── tabFichas: contagem por sessão ────────────────────────────────────────
    const sheetFichas = ss.getSheetByName('tabFichas');
    const contagemFichas = {};
    if (sheetFichas) {
      const mapaF   = getMapaColunas(sheetFichas);
      const dadosF  = sheetFichas.getDataRange().getValues();
      dadosF.shift();
      dadosF.forEach(f => {
        const idSess = String(f[mapaF['idsessao'] - 1] || '').trim();
        if (idSess) contagemFichas[idSess] = (contagemFichas[idSess] || 0) + 1;
      });
    }

    // ── tabMembros: autocomplete ──────────────────────────────────────────────
    const sheetMembros = ss.getSheetByName('tabMembros');
    const membrosAuto  = {};
    if (sheetMembros) {
      const mapaM  = getMapaColunas(sheetMembros);
      const dadosM = sheetMembros.getDataRange().getValues();
      dadosM.shift();
      dadosM.forEach(m => {
        const nome = m[mapaM['nome'] - 1];
        if (nome) membrosAuto[nome.toString().trim()] = null;
      });
    }

    // ── Chaves do mapa (com acentos) ──────────────────────────────────────────
    const chaveOrgao    = _encontrarChave(mapaSess, ['órgão', 'orgao', 'orgão']);
    const chaveLocal    = _encontrarChave(mapaSess, ['local/sala', 'local', 'sala']);
    const chavePresid   = _encontrarChave(mapaSess, ['presidente']);
    const chaveSecret   = _encontrarChave(mapaSess, ['secretário', 'secretario']);
    const chaveMembros  = _encontrarChave(mapaSess, ['membros']);
    const chaveProc     = _encontrarChave(mapaSess, ['procuradores']);
    const chaveExped    = _encontrarChave(mapaSess, ['expediente']);

    // ── Monta lista de sessões ────────────────────────────────────────────────
    const listaSessoes = dadosSess.map(linha => {
      const id      = String(linha[mapaSess['id'] - 1] || '');
      const rawDate = linha[mapaSess['datasessao'] - 1];
      const dObj    = rawDate ? new Date(rawDate) : null;
      const dataOk  = dObj && !isNaN(dObj.getTime());

      return {
        id:          id,
        datasessao:  dataOk ? Utilities.formatDate(dObj, 'GMT-3', 'dd/MM/yyyy') : 'Data não informada',
        datasort:    dataOk ? dObj.getTime() : 0,
        dataiso:     dataOk ? Utilities.formatDate(dObj, 'GMT-3', 'yyyy-MM-dd') : '',
        orgao:       chaveOrgao  ? String(linha[mapaSess[chaveOrgao]  - 1] || '') : '',
        local:       chaveLocal  ? String(linha[mapaSess[chaveLocal]  - 1] || '') : '',
        presidente:  chavePresid ? String(linha[mapaSess[chavePresid] - 1] || '') : '',
        secretario:  chaveSecret ? String(linha[mapaSess[chaveSecret] - 1] || '') : '',
        membros:     chaveMembros ? String(linha[mapaSess[chaveMembros] - 1] || '') : '',
        procuradores:chaveProc   ? String(linha[mapaSess[chaveProc]   - 1] || '') : '',
        expediente:  chaveExped  ? String(linha[mapaSess[chaveExped]  - 1] || '') : '',
        totalFichas: contagemFichas[id] || 0
      };
    }).sort((a, b) => b.datasort - a.datasort);

    return { sucesso: true, sessoes: listaSessoes, membros: membrosAuto };

  } catch (e) {
    console.error('Erro em formSessoes_buscarDadosCompletos: ' + e.message);
    return { sucesso: false, erro: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CRUD DE SESSÕES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Salva ou atualiza uma sessão.
 * @param {Object} obj  { id?, datasessao, orgao, local, presidente, secretario }
 */
function formSessoes_salvarRegistro(obj) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabSessoes');
    if (!sheet) throw new Error("Aba 'tabSessoes' não encontrada.");

    const mapa    = getMapaColunas(sheet);
    const dados   = sheet.getDataRange().getValues();
    const numCols = Object.keys(mapa).length;

    const chaveOrgao  = _encontrarChave(mapa, ['órgão', 'orgao', 'orgão']);
    const chaveLocal  = _encontrarChave(mapa, ['local/sala', 'local', 'sala']);
    const chavePresid = _encontrarChave(mapa, ['presidente']);
    const chaveSecret = _encontrarChave(mapa, ['secretário', 'secretario']);

    if (!obj.id) obj.id = novoIdTimeStamp();

    // Converte data ISO para Date
    let dataObj = '';
    if (obj.datasessao) {
      const p = obj.datasessao.split('-');
      if (p.length === 3) {
        dataObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 12, 0, 0);
      }
    }

    // Localiza linha existente
    let linhaAlvo = -1;
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][mapa['id'] - 1]).trim() === String(obj.id).trim()) {
        linhaAlvo = i + 1;
        break;
      }
    }

    const linha = linhaAlvo !== -1
      ? dados[linhaAlvo - 1].slice()
      : new Array(numCols).fill('');

    linha[mapa['id'] - 1]                              = obj.id;
    linha[mapa['datasessao'] - 1]                      = dataObj || '';
    if (chaveOrgao)  linha[mapa[chaveOrgao]  - 1]     = obj.orgao  || '';
    if (chaveLocal)  linha[mapa[chaveLocal]  - 1]      = obj.local  || '';
    if (chavePresid) linha[mapa[chavePresid] - 1]      = obj.presidente || '';
    if (chaveSecret) linha[mapa[chaveSecret] - 1]      = obj.secretario || '';

    if (linhaAlvo !== -1) {
      sheet.getRange(linhaAlvo, 1, 1, linha.length).setValues([linha]);
    } else {
      sheet.appendRow(linha);
    }

    return { sucesso: true };
  } catch (e) {
    throw new Error('Erro ao salvar sessão: ' + e.message);
  }
}

/**
 * Exclui uma sessão. Fichas e votos vinculados são preservados.
 */
function formSessoes_excluirRegistro(id) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabSessoes');
    if (!sheet) throw new Error("Aba 'tabSessoes' não encontrada.");

    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return { sucesso: true };
      }
    }
    throw new Error('Sessão não encontrada para exclusão.');
  } catch (e) {
    throw new Error(e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FICHAS DE VOTAÇÃO
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Retorna as fichas de uma sessão enriquecidas com dados do processo.
 * @param {string} idSessao
 */
function formSessoes_buscarFichas(idSessao) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);

    const sheetFichas = ss.getSheetByName('tabFichas');
    if (!sheetFichas) throw new Error("Aba 'tabFichas' não encontrada.");
    const mapaF   = getMapaColunas(sheetFichas);
    const dadosF  = sheetFichas.getDataRange().getValues();
    dadosF.shift();

    const sheetProc = ss.getSheetByName('tabProcessos');
    const cacheProc = {};
    if (sheetProc) {
      const mapaP  = getMapaColunas(sheetProc);
      const dadosP = sheetProc.getDataRange().getValues();
      dadosP.shift();
      const chaveLocal = _encontrarChave(mapaP, ['local da ocorrência', 'local']);
      dadosP.forEach(p => {
        const pid = String(p[mapaP['id'] - 1] || '').trim();
        if (pid) {
          cacheProc[pid] = {
            numero:     String(p[mapaP['processo']   - 1] || 'S/N'),
            requerente: String(p[mapaP['requerente'] - 1] || ''),
            requerido:  String(p[mapaP['requerido']  - 1] || ''),
            status:     mapaP['status'] ? String(p[mapaP['status'] - 1] || '') : '',
            local:      chaveLocal ? String(p[mapaP[chaveLocal] - 1] || '') : ''
          };
        }
      });
    }

    const chaveRelator  = _encontrarChave(mapaF, ['relator']);
    const chaveMembros  = _encontrarChave(mapaF, ['membros']);
    const chaveProcs    = _encontrarChave(mapaF, ['procuradores']);
    const chaveExped    = _encontrarChave(mapaF, ['expediente']);

    const fichas = dadosF
      .filter(f => String(f[mapaF['idsessao'] - 1]).trim() === String(idSessao).trim())
      .map(f => {
        const idFicha = String(f[mapaF['id'] - 1] || '');
        const idProc  = String(f[mapaF['idprocesso'] - 1] || '');
        const proc    = cacheProc[idProc] || { numero: 'S/N', requerente: '', requerido: '', status: '', local: '' };
        return {
          id:          idFicha,
          idprocesso:  idProc,
          relator:     chaveRelator ? String(f[mapaF[chaveRelator] - 1] || '') : '',
          membros:     chaveMembros ? String(f[mapaF[chaveMembros] - 1] || '') : '',
          procuradores:chaveProcs   ? String(f[mapaF[chaveProcs]   - 1] || '') : '',
          expediente:  chaveExped   ? String(f[mapaF[chaveExped]   - 1] || '') : '',
          proc: proc
        };
      });

    return { sucesso: true, fichas: fichas };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva ou atualiza campos da ficha (relator, expediente).
 */
function formSessoes_salvarFicha(obj) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabFichas');
    if (!sheet) throw new Error("Aba 'tabFichas' não encontrada.");

    const mapa  = getMapaColunas(sheet);
    const dados = sheet.getDataRange().getValues();

    const chaveRelator = _encontrarChave(mapa, ['relator']);
    const chaveExped   = _encontrarChave(mapa, ['expediente']);

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][mapa['id'] - 1]).trim() === String(obj.id).trim()) {
        if (chaveRelator) sheet.getRange(i + 1, mapa[chaveRelator]).setValue(obj.relator || '');
        if (chaveExped)   sheet.getRange(i + 1, mapa[chaveExped]).setValue(obj.expediente || '');
        return { sucesso: true };
      }
    }
    throw new Error('Ficha não encontrada.');
  } catch (e) {
    throw new Error('Erro ao salvar ficha: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VOTOS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Retorna os votos de uma ficha.
 * @param {string} idFicha
 */
function formSessoes_buscarVotos(idFicha) {
  try {
    const ss        = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetVot  = ss.getSheetByName('tabVotos');
    if (!sheetVot) throw new Error("Aba 'tabVotos' não encontrada.");

    const mapa  = getMapaColunas(sheetVot);
    const dados = sheetVot.getDataRange().getValues();
    dados.shift();

    const chaveUrl = _encontrarChave(mapa, ['url voto', 'urlvoto', 'url_voto']);

    const votos = dados
      .filter(v => String(v[mapa['idfichavotacao'] - 1]).trim() === String(idFicha).trim())
      .map(v => ({
        id:             String(v[mapa['id'] - 1] || ''),
        idfichavotacao: String(v[mapa['idfichavotacao'] - 1] || ''),
        idprocesso:     String(v[mapa['idprocesso'] - 1] || ''),
        tipovoto:       String(v[mapa['tipovoto'] - 1] || ''),
        relator:        String(v[mapa['relator'] - 1] || ''),
        voto:           String(v[mapa['voto'] - 1] || ''),
        resultado:      String(v[mapa['resultado'] - 1] || ''),
        urlvoto:        chaveUrl ? String(v[mapa[chaveUrl] - 1] || '') : ''
      }));

    return { sucesso: true, votos: votos };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva um novo voto ou atualiza um existente.
 * @param {Object} obj  { id?, idfichavotacao, idprocesso, tipovoto, relator, voto, resultado }
 */
function formSessoes_salvarVoto(obj) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabVotos');
    if (!sheet) throw new Error("Aba 'tabVotos' não encontrada.");

    const mapa    = getMapaColunas(sheet);
    const dados   = sheet.getDataRange().getValues();
    const numCols = Object.keys(mapa).length;

    const chaveUrl = _encontrarChave(mapa, ['url voto', 'urlvoto', 'url_voto']);

    if (!obj.id) obj.id = novoIdTimeStamp();

    let linhaAlvo = -1;
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][mapa['id'] - 1]).trim() === String(obj.id).trim()) {
        linhaAlvo = i + 1;
        break;
      }
    }

    const linha = linhaAlvo !== -1
      ? dados[linhaAlvo - 1].slice()
      : new Array(numCols).fill('');

    linha[mapa['id'] - 1]             = obj.id;
    linha[mapa['idfichavotacao'] - 1] = obj.idfichavotacao || '';
    linha[mapa['idprocesso'] - 1]     = obj.idprocesso || '';
    linha[mapa['tipovoto'] - 1]       = obj.tipovoto || '';
    linha[mapa['relator'] - 1]        = obj.relator || '';
    linha[mapa['voto'] - 1]           = obj.voto || '';
    linha[mapa['resultado'] - 1]      = obj.resultado || '';
    if (chaveUrl) linha[mapa[chaveUrl] - 1] = obj.urlvoto || '';

    if (linhaAlvo !== -1) {
      sheet.getRange(linhaAlvo, 1, 1, linha.length).setValues([linha]);
    } else {
      sheet.appendRow(linha);
    }

    return { sucesso: true, id: obj.id };
  } catch (e) {
    throw new Error('Erro ao salvar voto: ' + e.message);
  }
}

/**
 * Exclui um voto pelo Id.
 */
function formSessoes_excluirVoto(id) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName('tabVotos');
    if (!sheet) throw new Error("Aba 'tabVotos' não encontrada.");

    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return { sucesso: true };
      }
    }
    throw new Error('Voto não encontrado.');
  } catch (e) {
    throw new Error(e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD DE RELATÓRIO EM PDF
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Faz upload de um PDF para a pasta "Relatórios" no mesmo diretório do documento.
 * Salva a URL no campo URL Voto da tabVotos.
 *
 * @param {string} idVoto     Id do registro em tabVotos que receberá a URL
 * @param {string} base64Data Conteúdo do arquivo em Base64
 * @param {string} fileName   Nome original do arquivo
 */
function formSessoes_uploadRelatorio(idVoto, base64Data, fileName) {
  try {
    // Localiza a pasta pai do documento
    const doc = DocumentApp.getActiveDocument();
    const docFile = DriveApp.getFileById(doc.getId());
    const parentIter = docFile.getParents();
    if (!parentIter.hasNext()) throw new Error('Não foi possível localizar a pasta do documento.');
    const parentFolder = parentIter.next();

    // Localiza ou cria a pasta "Relatórios"
    let relFolder;
    const folderIter = parentFolder.getFoldersByName('Relatórios');
    relFolder = folderIter.hasNext() ? folderIter.next() : parentFolder.createFolder('Relatórios');

    // Cria o arquivo
    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, 'application/pdf', fileName);
    const file  = relFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();

    // Atualiza a URL no registro de voto correspondente
    if (idVoto) {
      const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
      const sheet = ss.getSheetByName('tabVotos');
      if (sheet) {
        const mapa    = getMapaColunas(sheet);
        const chaveUrl = _encontrarChave(mapa, ['url voto', 'urlvoto', 'url_voto']);
        if (chaveUrl) {
          const dados = sheet.getDataRange().getValues();
          for (let i = 1; i < dados.length; i++) {
            if (String(dados[i][mapa['id'] - 1]).trim() === String(idVoto).trim()) {
              sheet.getRange(i + 1, mapa[chaveUrl]).setValue(fileUrl);
              break;
            }
          }
        }
      }
    }

    return { sucesso: true, url: fileUrl, nome: fileName };

  } catch (e) {
    throw new Error('Erro no upload do relatório: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNÇÕES AUXILIARES INTERNAS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Procura a chave do mapa em uma lista de possíveis nomes (case-insensitive, sem acentos opcionais).
 */
function _encontrarChave(mapa, candidatos) {
  for (const cand of candidatos) {
    if (mapa[cand] !== undefined) return cand;
    // Tenta sem acentos
    const semAcento = cand.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const chaveEncontrada = Object.keys(mapa).find(k => {
      const kNorm = k.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      return kNorm === semAcento;
    });
    if (chaveEncontrada) return chaveEncontrada;
  }
  return null;
}