/**
 * @fileoverview Backend para o formulário de Gestão de Membros
 * Gerencia a persistência na aba 'tabMembros'
 */


/**
 * Função para abrir o formulário principal (pode ser chamada via menu ou barra lateral)
 */
function abrirFormMembros() {
  const template = HtmlService.createTemplateFromFile('FormMembros-layout');
  const html = template.evaluate()
    .setTitle('Gestão de Membros')
    .setWidth(900)
    .setHeight(600);
  
  DocumentApp.getUi().showModalDialog(html, 'Membros cadastrados e empossados');
}

/**
 * Busca membros e cargos únicos em uma única leitura de planilha
 */
function buscarMembrosMacro() {
  try {
    // CORREÇÃO 1: Usar PLANILHA_DADOS_ID (conforme seu Startup.gs)
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID); 
    const sheet = ss.getSheetByName("tabMembros");
    
    if (!sheet) throw new Error("Aba 'tabMembros' não encontrada.");

    // CORREÇÃO 2: Passar o objeto 'sheet' para a função de mapeamento
    const mapa = getMapaColunas(sheet); 
    const dadosPlena = sheet.getDataRange().getValues();
    
    const listaMembros = [];
    const cargosSet = new Set();

    if (dadosPlena.length > 1) {
      const cabecalhos = dadosPlena[0];
      
      // CORREÇÃO 3: Buscar o índice da coluna usando o mapa (que está em lowercase)
      const colunaCargoIndex = mapa["cargo"] - 1; 

      for (let i = 1; i < dadosPlena.length; i++) {
        let linha = {};
        cabecalhos.forEach((cab, index) => {
          linha[cab] = dadosPlena[i][index];
        });
        listaMembros.push(linha);

        const valorCargo = dadosPlena[i][colunaCargoIndex];
        if (valorCargo) {
          cargosSet.add(valorCargo.toString().trim());
        }
      }
    }

    const objCargos = {};
    cargosSet.forEach(c => objCargos[c] = null);

    return {
      membros: listaMembros,
      cargos: objCargos
    };

  } catch (e) {
    // Log para depuração no console do Google Cloud
    console.error("Erro em buscarMembrosMacro: " + e.message);
    throw new Error("Erro ao buscar dados: " + e.message);
  }
}

/**
 * Salva ou Atualiza um membro
 * @param {Object} dados Objeto vindo do frontend
 */
function salvarMembroMacro(dados) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    let sheet = ss.getSheetByName("tabMembros");
    
    // No backend (salvarMembroMacro), a lista de colunas oficiais deve ser:
    const colunasOficiais = ["Id", "Nome", "Email", "Gênero", "Cargo"];

    if (!sheet) {
      sheet = ss.insertSheet("tabMembros");
      sheet.getRange(1, 1, 1, colunasOficiais.length).setValues([colunasOficiais]);
    }

    const mapa = getMapaColunas(sheet);
    const fullData = sheet.getDataRange().getValues();
    
    const buscarNoObjeto = (chavePlanilha) => {
      const alvo = chavePlanilha.toLowerCase().trim();
      const realKey = Object.keys(dados).find(k => k.toLowerCase().trim() === alvo);
      return realKey ? dados[realKey] : undefined;
    };

    let linhaAlvo = -1;
    const idEnviado = buscarNoObjeto("Id");

    if (idEnviado) {
      const idColKey = Object.keys(mapa).find(k => k.toLowerCase().trim() === "id");
      if (idColKey) {
        const idColIndex = mapa[idColKey] - 1;
        for (let i = 1; i < fullData.length; i++) {
          if (String(fullData[i][idColIndex]).trim() === String(idEnviado).trim()) {
            linhaAlvo = i + 1;
            break;
          }
        }
      }
    }

    const valoresLinha = new Array(Object.keys(mapa).length);
    const dadosAntigos = linhaAlvo !== -1 ? fullData[linhaAlvo - 1] : [];

    Object.keys(mapa).forEach(nomeColuna => {
      let val = buscarNoObjeto(nomeColuna);
      let colIdx = mapa[nomeColuna] - 1;

      if (linhaAlvo !== -1 && val === undefined) {
        valoresLinha[colIdx] = dadosAntigos[colIdx];
        return;
      }

      if (nomeColuna.toLowerCase().includes("data") && val) {
        val = new Date(val);
      }
      
      valoresLinha[colIdx] = val || "";
    });

    if (linhaAlvo !== -1) {
      sheet.getRange(linhaAlvo, 1, 1, valoresLinha.length).setValues([valoresLinha]);
    } else {
      sheet.appendRow(valoresLinha);
    }

    return { sucesso: true };
  } catch (e) {
    throw new Error("Erro ao salvar no banco de dados: " + e.message);
  }
}

/**
 * Exclui um membro pelo ID
 */
function excluirMembroMacro(id) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sheet = ss.getSheetByName("tabMembros");
    if (!sheet) throw new Error("Aba 'tabMembros' não encontrada.");

    const data = sheet.getDataRange().getValues();
    const mapa = getMapaColunas(sheet);
    
    const idColKey = Object.keys(mapa).find(k => k.toLowerCase().trim() === "id");
    if (!idColKey) throw new Error("Coluna de ID não mapeada na planilha.");
    
    const idColIndex = mapa[idColKey] - 1;
    let registroExcluido = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        registroExcluido = true;
        break; 
      }
    }

    if (!registroExcluido) {
      throw new Error("Registro com ID " + id + " não encontrado para exclusão.");
    }

    return true;
  } catch (e) {
    throw new Error("Erro na exclusão: " + e.message);
  }
}

/**
 * Função para ser chamada pelo Painel Lateral
 * IMPORTANTE: Esta função é executada no contexto do servidor quando chamada
 * via google.script.run do painel lateral, então tem acesso à UI
 */
function PainelLateral_abrirMembros() {
  // Fecha o painel lateral antes de abrir o modal (opcional)
  // SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(''));
  
  abrirFormMembros();
  return true;
}