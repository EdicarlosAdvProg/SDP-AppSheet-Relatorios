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

/**
 * Sincroniza a lista de membros com o site da OAB-GO.
 * Grava os dados na planilha de banco de dados do sistema.
 */
function sincronizaMembrosSiteOAB() {
  const url = "https://www.oabgo.org.br/comissao/sistema-de-defesa-das-prerrogativas-sdp/";
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const sheet = ss.getSheetByName("tabMembros");

  let pageContent;
  try {
    pageContent = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
  } catch (e) {
    throw new Error("Erro de conexão com o site da OAB.");
  }

  const regexAba = /<div\s+role="tabpanel"\s+class="tab-pane"\s+id="aba1">([\s\S]*?)<\/div>/i;
  const matchAba = pageContent.match(regexAba);
  if (!matchAba) throw new Error("Conteúdo não encontrado.");

  const htmlLista = matchAba[1];
  const regexParagrafos = /<p>([\s\S]*?)<\/p>/gi;
  let parágrafo;

  // Usaremos um objeto para guardar Arrays de cargos: { "Nome": ["Cargo1", "Cargo2"] }
  const dicionarioMembros = {};
  let cargoAtual = "Membro";

  const cargoKeywords = [
    "vice", "presidente", "secretário", "secretários", "secretária", "secretaria",
    "coordenador", "coordenadora", "procurador", "procuradora", "órgão", "membro",
    "membros", "diretor", "diretora", "conselheiro", "conselheira", "regional",
    "comissão", "representante", "deliberativo", "sistema", "defesa"
  ];

  while ((parágrafo = regexParagrafos.exec(htmlLista)) !== null) {
    let conteudo = parágrafo[1];
    const matchCargo = conteudo.match(/<strong>(.*?)<\/strong>/i);

    if (matchCargo) {
      cargoAtual = matchCargo[1].replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim();
      conteudo = conteudo.replace(/<strong>.*?<\/strong>/i, '');
    }

    const nomesNoParagrafo = conteudo
      .split(/<br\s*\/?>|\n/i)
      .map(n => n.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim())
      .filter(n => {
        if (n.length < 4) return false;
        return !cargoKeywords.some(k => new RegExp(`\\b${k}\\b`, 'i').test(n));
      });

    nomesNoParagrafo.forEach(nome => {
      const nomeLimpo = nome.split('-')[0].trim();
      if (nomeLimpo) {
        if (!dicionarioMembros[nomeLimpo]) {
          dicionarioMembros[nomeLimpo] = [];
        }
        // Adiciona o cargo se ele ainda não estiver na lista deste membro
        if (dicionarioMembros[nomeLimpo].indexOf(cargoAtual) === -1) {
          dicionarioMembros[nomeLimpo].push(cargoAtual);
        }
      }
    });
  }

  // --- Lógica de Formatação dos Cargos e Gravação ---
  const mapa = getMapaColunas(sheet);
  const colId = mapa["id"];
  const colNome = mapa["nome"];
  const colCargo = mapa["cargo"];
  const ultimaCol = sheet.getLastColumn();

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, ultimaCol).clearContent();
    if (lastRow > 2) sheet.deleteRows(3, lastRow - 2);
  }

  const listaFinalNomes = Object.keys(dicionarioMembros).sort((a, b) => a.localeCompare(b));

  if (listaFinalNomes.length > 0) {
    let idAtual = "";
    const matrixFinal = [];

    listaFinalNomes.forEach(nome => {
      idAtual = gerarProximoIdIncremental(idAtual);

      // FORMATAÇÃO GRAMATICAL: "cargo1, cargo2 e cargo3"
      const listaCargos = dicionarioMembros[nome];
      let cargosFormatados = "";

      if (listaCargos.length === 1) {
        cargosFormatados = listaCargos[0];
      } else if (listaCargos.length > 1) {
        // Pega todos menos o último e junta com vírgula
        const ultItem = listaCargos.pop();
        cargosFormatados = listaCargos.join(", ") + " e " + ultItem;
      }

      let linha = new Array(ultimaCol).fill("");
      linha[colId - 1] = idAtual;
      linha[colNome - 1] = nome;
      linha[colCargo - 1] = cargosFormatados;
      matrixFinal.push(linha);
    });

    sheet.getRange(2, 1, matrixFinal.length, ultimaCol).setValues(matrixFinal);
    return "Sucesso! " + listaFinalNomes.length + " membros processados.";
  }
  return "Nenhum dado processado.";
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

function novoIdTimeStamp() {
  const ano = new Date().getFullYear() - 2025;
  const ms = Date.now() % 46656000000;
  return ano.toString(36) + ms.toString(36).padStart(5, '0');
}

function salvarGenerosEmMassa(listaNovosGeneros) {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const sheet = ss.getSheetByName("tabMembros");
  const dados = sheet.getDataRange().getValues();
  const mapa = getMapaColunas(sheet);

  const colId = mapa["id"] - 1;
  const colGen = mapa["gênero"] - 1;

  // Cria um objeto para busca rápida dos novos gêneros
  const dePara = {};
  listaNovosGeneros.forEach(item => dePara[item.id] = item.genero);

  // Percorre a planilha e atualiza apenas os IDs que estão na lista
  for (let i = 1; i < dados.length; i++) {
    const idAtual = dados[i][colId];
    if (dePara[idAtual]) {
      sheet.getRange(i + 1, colGen + 1).setValue(dePara[idAtual]);
    }
  }

  return "Gêneros atualizados com sucesso!";
}