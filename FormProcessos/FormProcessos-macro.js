/**
 * @fileoverview Backend para Gestão de Processos SDP — OAB/GO
 * Funções: CRUD processos, pauta, histórico, alteração de status, importação em massa.
 */

function formProcessos_abrirModal() {
  const template = HtmlService.createTemplateFromFile('FormProcessos-layout');
  const html = template.evaluate()
    .setTitle('SDP - Gestão de Processos')
    .setWidth(1200)
    .setHeight(850);
  DocumentApp.getUi().showModalDialog(html, ' ');
}

function PainelLateral_abrirProcessos() {
  formProcessos_abrirModal();
  return true;
}

function sdp_abrirFormProcessos() {
  formProcessos_abrirModal();
}

/**
 * Payload inicial: processos (cruzados com último evento do histórico),
 * sessões para o modal de pauta e membros para autocomplete.
 */
function formProcessos_buscarDadosCompletos() {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);

    // ── tabProcessos ──────────────────────────────────────────────────────────
    const sheetProc = ss.getSheetByName("tabProcessos");
    if (!sheetProc) throw new Error("Aba 'tabProcessos' não encontrada.");
    const mapaProc  = getMapaColunas(sheetProc);
    const dadosProc = sheetProc.getDataRange().getValues();
    dadosProc.shift();

    // ── tabHistorico: último evento por processo ───────────────────────────────
    const sheetHist = ss.getSheetByName("tabHistorico");
    const ultimoEventoMap = {};
    if (sheetHist) {
      const mapaHist  = getMapaColunas(sheetHist);
      const dadosHist = sheetHist.getDataRange().getValues();
      dadosHist.shift();
      const chaveDesc = mapaHist['descrição'] !== undefined ? 'descrição' : 'descricão';

      dadosHist.forEach(h => {
        const idProc  = String(h[mapaHist['idprocesso'] - 1]);
        const dataRaw = h[mapaHist['datahora'] - 1];
        const desc    = h[mapaHist[chaveDesc]  - 1];
        if (idProc && dataRaw) {
          const dObj = new Date(dataRaw);
          if (!isNaN(dObj.getTime())) {
            if (!ultimoEventoMap[idProc] || dObj > ultimoEventoMap[idProc].data) {
              ultimoEventoMap[idProc] = { data: dObj, descricao: desc || "" };
            }
          }
        }
      });
    }

    // ── tabSessoes: lista para o modal de pauta ───────────────────────────────
    const sheetSess = ss.getSheetByName("tabSessoes");
    let listaSessoes = [];
    if (sheetSess) {
      const mapaSess  = getMapaColunas(sheetSess);
      const dadosSess = sheetSess.getDataRange().getValues();
      dadosSess.shift();
      const chaveOrgao = Object.keys(mapaSess).find(k => k.includes('rg'));
      const chaveLocal = Object.keys(mapaSess).find(k => k.includes('local'));

      listaSessoes = dadosSess.map(s => {
        const rawDate = s[mapaSess['datasessao'] - 1];
        const dObj    = rawDate ? new Date(rawDate) : null;
        const dataOk  = dObj && !isNaN(dObj.getTime());
        return {
          id:       s[mapaSess['id'] - 1],
          data:     dataOk ? Utilities.formatDate(dObj, "GMT-3", "dd/MM/yyyy") : "Data não informada",
          dataSort: dataOk ? dObj.getTime() : 0,
          orgao:    chaveOrgao ? (s[mapaSess[chaveOrgao] - 1] || "—") : "—",
          local:    chaveLocal ? (s[mapaSess[chaveLocal]  - 1] || "") : ""
        };
      }).sort((a, b) => b.dataSort - a.dataSort);
    }

    // ── tabMembros: autocomplete do campo Procurador ──────────────────────────
    const sheetMembros = ss.getSheetByName("tabMembros");
    const membrosAuto  = {};
    if (sheetMembros) {
      const mapaMembros  = getMapaColunas(sheetMembros);
      const dadosMembros = sheetMembros.getDataRange().getValues();
      dadosMembros.shift();
      dadosMembros.forEach(m => {
        const nome = m[mapaMembros['nome'] - 1];
        if (nome) membrosAuto[nome.toString().trim()] = null;
      });
    }

    // ── Monta lista final de processos ────────────────────────────────────────
    const listaFinal = dadosProc.map(linha => {
      const id = String(linha[mapaProc['id'] - 1] || "");
      let dataFmt = "---", descFmt = "Sem registros";
      if (ultimoEventoMap[id]) {
        dataFmt = Utilities.formatDate(ultimoEventoMap[id].data, "GMT-3", "dd/MM/yyyy");
        descFmt = ultimoEventoMap[id].descricao;
      }
      return {
        id:         id,
        processo:   String(linha[mapaProc['processo']   - 1] || "S/N"),
        requerente: String(linha[mapaProc['requerente'] - 1] || ""),
        requerido:  String(linha[mapaProc['requerido']  - 1] || ""),
        procurador: String(linha[mapaProc['procurador'] - 1] || ""),
        local:      mapaProc['local da ocorrência'] ? String(linha[mapaProc['local da ocorrência'] - 1] || "") : "",
        status:     mapaProc['status'] ? String(linha[mapaProc['status'] - 1] || "") : "",
        resumo:     mapaProc['resumo'] ? String(linha[mapaProc['resumo'] - 1] || "") : "",
        ementa:     mapaProc['ementa'] ? String(linha[mapaProc['ementa'] - 1] || "") : "",
        provas:     mapaProc['provas'] ? String(linha[mapaProc['provas'] - 1] || "") : "",
        ultimaData: dataFmt,
        ultimaDesc: descFmt
      };
    });

    return { sucesso: true, dados: listaFinal, sessoes: listaSessoes, membros: membrosAuto };

  } catch (e) {
    console.error("Erro em formProcessos_buscarDadosCompletos: " + e.message);
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Salva ou atualiza um processo na tabProcessos.
 */
function formProcessos_salvarRegistro(obj) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName("tabProcessos");
    if (!sheet) throw new Error("Aba 'tabProcessos' não encontrada.");

    const dados   = sheet.getDataRange().getValues();
    const colunas = dados[0].map(c => c.toLowerCase().trim());

    let rowIndex = -1;
    if (obj.id) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]).trim() === String(obj.id).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (!obj.id || rowIndex === -1) obj.id = novoIdTimeStamp();

    const linhaParaSalvar = colunas.map(col => {
      if (col === 'id') return obj.id;
      if (col === 'local da ocorrência') return obj['local'] || "";
      return obj[col] !== undefined ? obj[col] : "";
    });

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, linhaParaSalvar.length).setValues([linhaParaSalvar]);
    } else {
      sheet.appendRow(linhaParaSalvar);
    }

    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro ao salvar processo: " + e.message);
  }
}

/**
 * Exclui um processo da tabProcessos pelo Id.
 * Registros dependentes (Fichas, Votos, Histórico) são preservados.
 */
function formProcessos_excluirRegistro(id) {
  try {
    const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName("tabProcessos");
    if (!sheet) throw new Error("Aba 'tabProcessos' não encontrada.");

    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return { sucesso: true };
      }
    }
    throw new Error("Processo não encontrado para exclusão.");
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * Busca todos os registros do histórico de um processo, ordenados do mais recente.
 * @param {string} idProcesso Id Base36 do processo.
 */
function formProcessos_buscarHistorico(idProcesso) {
  try {
    const ss        = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetHist = ss.getSheetByName("tabHistorico");
    if (!sheetHist) throw new Error("Aba 'tabHistorico' não encontrada.");

    const mapaHist  = getMapaColunas(sheetHist);
    const dadosHist = sheetHist.getDataRange().getValues();
    dadosHist.shift();

    const chaveDesc = mapaHist['descrição'] !== undefined ? 'descrição' : 'descricão';

    const registros = dadosHist
      .filter(h => String(h[mapaHist['idprocesso'] - 1]).trim() === String(idProcesso).trim())
      .map(h => {
        const rawDate = h[mapaHist['datahora'] - 1];
        const dObj    = rawDate ? new Date(rawDate) : null;
        const dataOk  = dObj && !isNaN(dObj.getTime());
        return {
          data:      dataOk ? Utilities.formatDate(dObj, "GMT-3", "dd/MM/yyyy") : "—",
          dataSort:  dataOk ? dObj.getTime() : 0,
          tipo:      String(h[mapaHist['tipo'] - 1] || ""),
          descricao: String(h[mapaHist[chaveDesc] - 1] || "")
        };
      })
      .sort((a, b) => b.dataSort - a.dataSort);

    return { sucesso: true, registros };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Altera o status de um processo e registra o evento no histórico.
 * @param {string} idProcesso  Id Base36 do processo.
 * @param {string} novoStatus  "Concluso" ou "Na secretaria".
 * @param {string} dataISO     Data da mudança no formato 'yyyy-mm-dd'.
 */
function formProcessos_alterarStatus(idProcesso, novoStatus, dataISO) {
  try {
    const ss        = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetProc = ss.getSheetByName("tabProcessos");
    if (!sheetProc) throw new Error("Aba 'tabProcessos' não encontrada.");

    const mapaProc  = getMapaColunas(sheetProc);
    const dadosProc = sheetProc.getDataRange().getValues();
    const colStatus = mapaProc['status'] - 1;

    let linhaAlvo = -1;
    for (let i = 1; i < dadosProc.length; i++) {
      if (String(dadosProc[i][mapaProc['id'] - 1]).trim() === String(idProcesso).trim()) {
        linhaAlvo = i + 1;
        break;
      }
    }
    if (linhaAlvo === -1) throw new Error("Processo não encontrado.");

    // Atualiza a coluna Status
    sheetProc.getRange(linhaAlvo, colStatus + 1).setValue(novoStatus);

    // Registra no histórico
    _registrarHistorico(ss, idProcesso, "Andamento", novoStatus, dataISO);

    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro ao alterar status: " + e.message);
  }
}

/**
 * Importação em massa de processos a partir de texto estruturado.
 * Formato esperado por linha: Processo | Data | Hora | Requerente
 * 
 * Regras:
 * - Se o número do processo NÃO existir na tabela → cria novo registro com status "Concluso"
 * - Se o número do processo JÁ existir → apenas muda status para "Concluso" e registra histórico
 * Em ambos os casos, adiciona registro em tabHistorico:
 * Tipo: "Andamento" | Descrição: "Concluso ao órgão deliberativo"
 *
 * @param {Array} itens Lista de objetos { processo, data, requerente } parseados no frontend.
 */
function formProcessos_importarEmMassa(itens) {
  try {
    const ss        = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetProc = ss.getSheetByName("tabProcessos");
    if (!sheetProc) throw new Error("Aba 'tabProcessos' não encontrada.");

    const mapaProc  = getMapaColunas(sheetProc);
    const dadosProc = sheetProc.getDataRange().getValues();
    const colunas   = dadosProc[0].map(c => c.toLowerCase().trim());
    const colStatus = mapaProc['status'] - 1;
    const colId     = mapaProc['id']     - 1;
    const colNum    = mapaProc['processo']   - 1;

    let criados = 0, atualizados = 0;

    itens.forEach(item => {
      const numeroLimpo = String(item.processo || "").trim();
      if (!numeroLimpo) return;

      // Busca processo existente pelo número
      let linhaExistente = -1;
      let idExistente    = "";
      for (let i = 1; i < dadosProc.length; i++) {
        if (String(dadosProc[i][colNum]).trim() === numeroLimpo) {
          linhaExistente = i + 1;
          idExistente    = String(dadosProc[i][colId]);
          break;
        }
      }

      if (linhaExistente !== -1) {
        // Processo existe: apenas atualiza status
        sheetProc.getRange(linhaExistente, colStatus + 1).setValue("Concluso");
        _registrarHistorico(ss, idExistente, "Andamento", "Concluso ao órgão deliberativo", item.data);
        atualizados++;

      } else {
        // Processo novo: cria linha completa
        const novoId    = novoIdTimeStamp();
        Utilities.sleep(5); // Garante unicidade do Id
        const novaLinha = colunas.map(col => {
          if (col === 'id')         return novoId;
          if (col === 'processo')   return numeroLimpo;
          if (col === 'requerente') return item.requerente || "";
          if (col === 'status')     return "Concluso";
          return "";
        });
        sheetProc.appendRow(novaLinha);
        _registrarHistorico(ss, novoId, "Andamento", "Concluso ao órgão deliberativo", item.data);
        criados++;
      }
    });

    return {
      sucesso: true,
      mensagem: `Importação concluída: ${criados} processo(s) criado(s), ${atualizados} atualizado(s).`
    };

  } catch (e) {
    throw new Error("Erro na importação em massa: " + e.message);
  }
}

/**
 * Adiciona um processo à pauta de uma sessão existente (tabFichas).
 */
function formProcessos_adicionarAPauta(idProcesso, idSessao) {
  try {
    const ss          = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetFichas = ss.getSheetByName("tabFichas");
    if (!sheetFichas) throw new Error("Aba 'tabFichas' não encontrada.");

    const mapa      = getMapaColunas(sheetFichas);
    const dados     = sheetFichas.getDataRange().getValues();
    const colSessao = mapa['idsessao']   - 1;
    const colProc   = mapa['idprocesso'] - 1;

    for (let i = 1; i < dados.length; i++) {
      if (
        String(dados[i][colSessao]).trim() === String(idSessao).trim()  &&
        String(dados[i][colProc]).trim()   === String(idProcesso).trim()
      ) {
        throw new Error("Este processo já está na pauta desta sessão.");
      }
    }

    const numCols   = Object.keys(mapa).length;
    const novaLinha = new Array(numCols).fill("");
    novaLinha[mapa['id'] - 1]         = novoIdTimeStamp();
    novaLinha[colSessao]              = idSessao;
    novaLinha[colProc]                = idProcesso;
    novaLinha[mapa['expediente'] - 1] = "Aguardando relato";

    sheetFichas.appendRow(novaLinha);
    return { sucesso: true, mensagem: "Processo adicionado à pauta com sucesso." };

  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * Cria nova sessão (Órgão + Data + Sala/Local opcional) e inclui o processo na pauta.
 */
function formProcessos_criarSessaoEPautar(idProcesso, orgao, dataISO, local) {
  try {
    const ss           = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheetSessoes = ss.getSheetByName("tabSessoes");
    if (!sheetSessoes) throw new Error("Aba 'tabSessoes' não encontrada.");

    const mapa      = getMapaColunas(sheetSessoes);
    const numCols   = Object.keys(mapa).length;
    const novaLinha = new Array(numCols).fill("");

    const novoIdSessao = novoIdTimeStamp();
    novaLinha[mapa['id'] - 1] = novoIdSessao;

    if (dataISO) {
      const p = dataISO.split('-');
      if (p.length === 3) {
        novaLinha[mapa['datasessao'] - 1] =
          new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 12, 0, 0);
      }
    }

    const chaveOrgao = Object.keys(mapa).find(k => k.includes('rg'));
    const chaveLocal = Object.keys(mapa).find(k => k.includes('local'));
    if (chaveOrgao) novaLinha[mapa[chaveOrgao] - 1] = orgao;
    if (chaveLocal && local) novaLinha[mapa[chaveLocal] - 1] = local;

    sheetSessoes.appendRow(novaLinha);
    Utilities.sleep(250);

    formProcessos_adicionarAPauta(idProcesso, novoIdSessao);
    return { sucesso: true, mensagem: "Sessão criada e processo incluído na pauta." };

  } catch (e) {
    throw new Error("Erro ao criar sessão: " + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Funções auxiliares internas
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Insere um registro na tabHistorico.
 * @param {Spreadsheet} ss          Instância da planilha já aberta.
 * @param {string}      idProcesso  Id do processo.
 * @param {string}      tipo        Ex: "Andamento", "Ato processual".
 * @param {string}      descricao   Texto do evento.
 * @param {string}      dataISO     Data no formato 'yyyy-mm-dd' (pode ser "dd/mm/yyyy" para retrocompatibilidade).
 */
function _registrarHistorico(ss, idProcesso, tipo, descricao, dataISO) {
  const sheetHist = ss.getSheetByName("tabHistorico");
  if (!sheetHist) return;

  const mapaHist = getMapaColunas(sheetHist);
  const numCols  = Object.keys(mapaHist).length;
  const linha    = new Array(numCols).fill("");

  const chaveDesc = mapaHist['descrição'] !== undefined ? 'descrição' : 'descricão';

  linha[mapaHist['id']          - 1] = novoIdTimeStamp();
  linha[mapaHist['idprocesso']  - 1] = idProcesso;
  linha[mapaHist['tipo']        - 1] = tipo;
  linha[mapaHist[chaveDesc]     - 1] = descricao;

  // Converte data: aceita 'yyyy-mm-dd' ou 'dd/mm/yyyy'
  if (dataISO) {
    let dObj;
    if (dataISO.includes('-')) {
      const p = dataISO.split('-');
      dObj = p.length === 3
        ? new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 12, 0, 0)
        : null;
    } else if (dataISO.includes('/')) {
      const p = dataISO.split('/');
      dObj = p.length === 3
        ? new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]), 12, 0, 0)
        : null;
    }
    if (dObj && !isNaN(dObj.getTime())) {
      linha[mapaHist['datahora'] - 1] = dObj;
    }
  }

  sheetHist.appendRow(linha);
}