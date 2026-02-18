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
 * 1. FUNÇÃO PRINCIPAL PARA ATUALIZAR A TABELA DE MEMBROS CONFORME O SISTE DA OAB
 */
function atualizarTabelaMembros() {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  
  // A. Arquiva dados atuais e limpa a tabela principal
  arquivarELimparMembros(ss);
  
  // B. Extrai dados do site usando sua lógica original intacta
  const dicionarioMembrosBruto = extrairDadosSiteOAB();
  
  // C. Enriquece os nomes extraídos com E-mail e Gênero do Arquivo
  const membrosEnriquecidos = enriquecerDadosComArquivo(ss, dicionarioMembrosBruto);
  
  // D. Gravação Final na tabMembros
  const sheetMembros = ss.getSheetByName("tabMembros");
  const mapa = getMapaColunas(sheetMembros);
  const ultimaCol = sheetMembros.getLastColumn();
  
  if (membrosEnriquecidos.length > 0) {
    let idAtual = "";
    const matrixFinal = membrosEnriquecidos.map(m => {
      idAtual = gerarProximoIdIncremental(idAtual);
      let linha = new Array(ultimaCol).fill("");
      
      linha[mapa["id"] - 1] = idAtual;
      linha[mapa["nome"] - 1] = m.nome;
      linha[mapa["cargo"] - 1] = m.cargo;
      linha[mapa["email"] - 1] = m.email;
      linha[mapa["gênero"] - 1] = m.genero;
      
      return linha;
    });

    sheetMembros.getRange(2, 1, matrixFinal.length, ultimaCol).setValues(matrixFinal);
    return "Sucesso! " + membrosEnriquecidos.length + " membros processados e sincronizados.";
  }
  return "Nenhum dado processado.";
}

/**
 * 2. EXTRAÇÃO DOS NOMES E CARGOS DO SITE DA OAB
 */
function extrairDadosSiteOAB() {
  const url = "https://www.oabgo.org.br/comissao/sistema-de-defesa-das-prerrogativas-sdp/";
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
  const dicionarioMembros = {};
  let cargoAtual = "Membro";

  // LISTA COMPLETA RESTAURADA
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
        if (!dicionarioMembros[nomeLimpo]) dicionarioMembros[nomeLimpo] = [];
        if (dicionarioMembros[nomeLimpo].indexOf(cargoAtual) === -1) {
          dicionarioMembros[nomeLimpo].push(cargoAtual);
        }
      }
    });
  }
  return dicionarioMembros;
}

/**
 * 3. ENRIQUECIMENTO DOS DADOS COM E-MAIL E GÊNERO PREVIAMENTE CADASTRADO
 */
function enriquecerDadosComArquivo(ss, dicionarioBruto) {
  const sheetArq = ss.getSheetByName("tabMembrosArquivo");
  const mapaArq = getMapaColunas(sheetArq);
  const valoresArq = sheetArq.getDataRange().getValues();
  
  const cacheArquivo = {};
  for (let i = 1; i < valoresArq.length; i++) {
    const nomeArq = valoresArq[i][mapaArq["nome"] - 1];
    cacheArquivo[nomeArq] = {
      email: valoresArq[i][mapaArq["email"] - 1] || "",
      genero: valoresArq[i][mapaArq["gênero"] - 1] || ""
    };
  }

  const nomesExtraidos = Object.keys(dicionarioBruto).sort((a, b) => a.localeCompare(b));

  return nomesExtraidos.map(nome => {
    const memoria = cacheArquivo[nome] || { email: "", genero: "" };
    const genero = memoria.genero;
    const listaCargosBrutos = dicionarioBruto[nome];

    // Aplicar a flexão de gênero em cada cargo da lista
    const listaCargosFlexionados = listaCargosBrutos.map(cargo => flexionarCargo(cargo, genero));

    let cargosFormatados = "";
    if (listaCargosFlexionados.length === 1) {
      cargosFormatados = listaCargosFlexionados[0];
    } else if (listaCargosFlexionados.length > 1) {
      const copiaCargos = [...listaCargosFlexionados];
      const ultItem = copiaCargos.pop();
      cargosFormatados = copiaCargos.join(", ") + " e " + ultItem;
    }

    return {
      nome: nome,
      cargo: cargosFormatados,
      email: memoria.email,
      genero: genero
    };
  });
}

/**
 * Função Auxiliar para tratar a gramática dos cargos com flexão composta
 */
function flexionarCargo(cargo, genero) {
  let novoCargo = cargo;

  // 1. Tratamento de Plurais e Padronização para o Singular Masculino (Base)
  // Remove o plural independentemente do gênero para padronizar
  novoCargo = novoCargo.replace(/Secretários-Gerais Executivos/gi, "Secretário-Geral Executivo");
  novoCargo = novoCargo.replace(/Vice-Presidentes/gi, "Vice-Presidente");
  
  // Se não houver gênero definido, paramos por aqui (fica no singular masculino)
  if (!genero || genero === "") return novoCargo;

  // 2. Flexão para o Feminino
  if (genero === "Feminino") {
    // Substituições simples
    novoCargo = novoCargo.replace(/\bMembro\b/g, "Membra");
    novoCargo = novoCargo.replace(/\bCoordenador\b/g, "Coordenadora");
    novoCargo = novoCargo.replace(/\bProcurador\b/g, "Procuradora");
    novoCargo = novoCargo.replace(/\bDiretor\b/g, "Diretora");
    novoCargo = novoCargo.replace(/\bConselheiro\b/g, "Conselheira");

    // Substituição Composta: Secretário e Executivo
    // A ordem importa: primeiro trocamos o radical para garantir a concordância
    if (novoCargo.includes("Secretário")) {
      novoCargo = novoCargo.replace(/\bSecretário\b/g, "Secretária");
      
      // Se houver a palavra Executivo acompanhando, flexiona também
      if (novoCargo.includes("Executivo")) {
        novoCargo = novoCargo.replace(/\bExecutivo\b/g, "Executiva");
      }
    }
  }

  return novoCargo;
}

/**
 * 4. ARQUIVAMENTO DOS DADOS ENRIQUECIDOS E EXCLUSÃO DOS DADOS DE TABMEMBROS
 */
function arquivarELimparMembros(ss) {
  const sheetMembros = ss.getSheetByName("tabMembros");
  const sheetArq = ss.getSheetByName("tabMembrosArquivo");
  
  const mapaM = getMapaColunas(sheetMembros);
  const mapaA = getMapaColunas(sheetArq);
  
  const dadosMembros = sheetMembros.getDataRange().getValues();
  if (dadosMembros.length < 2) return;

  const dadosArq = sheetArq.getDataRange().getValues();
  const nomesNoArquivo = dadosArq.map(r => r[mapaA["nome"] - 1]);

  for (let i = 1; i < dadosMembros.length; i++) {
    const nomeM = dadosMembros[i][mapaM["nome"] - 1];
    const emailM = dadosMembros[i][mapaM["email"] - 1];
    const generoM = dadosMembros[i][mapaM["gênero"] - 1];
    
    if (!nomeM) continue;

    const indexNoArq = nomesNoArquivo.indexOf(nomeM);

    if (indexNoArq !== -1) {
      // Atualiza Email e Gênero na linha encontrada (index + 1)
      sheetArq.getRange(indexNoArq + 1, mapaA["email"]).setValue(emailM);
      sheetArq.getRange(indexNoArq + 1, mapaA["gênero"]).setValue(generoM);
    } else {
      // Adiciona novo registro com ID de arquivo exclusivo
      const novoIdArq = "ARQ-" + novoIdTimeStamp();
      const novaLinha = [];
      novaLinha[mapaA["id"] - 1] = novoIdArq;
      novaLinha[mapaA["nome"] - 1] = nomeM;
      novaLinha[mapaA["email"] - 1] = emailM;
      novaLinha[mapaA["gênero"] - 1] = generoM;
      sheetArq.appendRow(novaLinha);
    }
  }

  // Limpa a tabela principal
  const lastRow = sheetMembros.getLastRow();
  if (lastRow >= 2) {
    sheetMembros.getRange(2, 1, lastRow - 1, sheetMembros.getLastColumn()).clearContent();
    if (lastRow > 2) sheetMembros.deleteRows(3, lastRow - 2);
  }
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