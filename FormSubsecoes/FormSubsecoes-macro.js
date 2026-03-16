/**
 * @fileoverview Robô de Sincronização Unificado (Versão Final)
 * Local: Subsecoes-macro.gs
 */

function subsecoes_sincronizarComSite() {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const shSub = ss.getSheetByName("tabSubsecoes");
  const shCid = ss.getSheetByName("tabCidades");

  // 1. Limpeza das tabelas (preservando o cabeçalho)
  if (shSub.getLastRow() > 1) shSub.getRange(2, 1, shSub.getLastRow() - 1, shSub.getLastColumn()).clearContent();
  if (shCid.getLastRow() > 1) shCid.getRange(2, 1, shCid.getLastRow() - 1, shCid.getLastColumn()).clearContent();

  const urlIndex = "https://www.oabgo.org.br/subsecoes/";
  const htmlIndex = UrlFetchApp.fetch(urlIndex).getContentText();
  
  const regexBloco = /<aside class="item">([\s\S]*?)<\/aside>/g;
  let match;
  
  const listaSubsecoes = [];
  const listaCidades = [];
  let ultimoIdSub = ""; 
  let ultimoIdCid = "";

  // --- INSERÇÃO DE GOIÂNIA (ESTRUTURA DE SUBSEÇÃO E CIDADE) ---
  // 1. Adiciona Goiânia na tabSubsecoes para manter a integridade da modelagem
  listaSubsecoes.push([
    "SUB-SEDE",              // Id fixo
    "GOIÂNIA",               // Nome
    "Frederico Manoel Sousa Álvares", // Presidente (Coordenador)
    "oabnet@oabgo.org.br",    // Email
    "(62) 3238-2000",         // Telefone
    "Todas (Sede)"           // Região
  ]);

  // 2. Adiciona Goiânia na tabCidades vinculada ao ID SUB-SEDE
  ultimoIdCid = gerarProximoIdIncremental(ultimoIdCid);
  listaCidades.push([
    "CID-" + ultimoIdCid,
    "SUB-SEDE",
    "GOIÂNIA"
  ]);

  // --- LOOP DAS SUBSEÇÕES VINDAS DO SITE ---
  while ((match = regexBloco.exec(htmlIndex)) !== null) {
    const bloco = match[1];

    const nomeSubRaw = subsecoes_limparTexto(/<h3>(.*?)<\/h3>/.exec(bloco));
    const nomeSubLimpo = nomeSubRaw.toUpperCase().replace("OAB - ", "").trim();
    
    // Pula se for Goiânia no site (para não duplicar a que inserimos manualmente)
    if (nomeSubLimpo === "GOIÂNIA") continue;

    let telefone = "";
    const trechoTelefone = /<i class="fas fa-phone"><\/i>([\s\S]*?)class="btn">Saiba mais/i.exec(bloco);
    if (trechoTelefone) {
      const telMatch = /\(\d{2}\)\s\d{4,5}-\d{4}/.exec(trechoTelefone[1]);
      if (telMatch) telefone = telMatch[0];
    }

    const cfEmailMatch = /data-cfemail="(.*?)"/.exec(bloco);
    const email = cfEmailMatch ? subsecoes_decodificarEmail(cfEmailMatch[1]) : "";
    const linkDetalhe = /href="(https:\/\/www.oabgo.org.br\/subsecao\/.*?\/)"/.exec(bloco);

    if (linkDetalhe) {
      const regiaoSubsecao = buscarRegiaoNoArquivo(nomeSubLimpo);
      
      ultimoIdSub = gerarProximoIdIncremental(ultimoIdSub);
      const idSubAtual = "SUB-" + ultimoIdSub;
      
      const detalhes = subsecoes_capturarDetalhes(linkDetalhe[1]);
      
      // tabSubsecoes: Id, Nome, Presidente, Email, Telefone, Regiao
      listaSubsecoes.push([
        idSubAtual,
        nomeSubLimpo,
        detalhes.presidente,
        email,
        telefone,
        regiaoSubsecao
      ]);

      // tabCidades: Id, IdSubsecao, Nome (Simplificada sem a 4ª coluna)
      detalhes.cidades.forEach(cidade => {
        ultimoIdCid = gerarProximoIdIncremental(ultimoIdCid);
        listaCidades.push([
          "CID-" + ultimoIdCid,
          idSubAtual,
          cidade.toUpperCase()
        ]);
      });
    }
  }

  // Gravação Final (Batch Update)
  if (listaSubsecoes.length > 0) shSub.getRange(2, 1, listaSubsecoes.length, 6).setValues(listaSubsecoes);
  if (listaCidades.length > 0) shCid.getRange(2, 1, listaCidades.length, 3).setValues(listaCidades);

  return `Sincronização concluída! Goiânia inserida como Sede e subseções mapeadas por região.`;
}

/**
 * Busca Região na aba tabSubsecoesArquivo
 */
function buscarRegiaoNoArquivo(nomeSubsecao) {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const shArq = ss.getSheetByName("tabSubsecoesArquivo");
  if (!shArq) return "SEDE / NÃO MAPEADA";

  const dados = shArq.getDataRange().getValues();
  const busca = nomeSubsecao.toUpperCase().trim();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().toUpperCase().trim() === busca) {
      return dados[i][1];
    }
  }
  return "SEDE / NÃO MAPEADA";
}

/**
 * Captura Presidente e Cidades (Nível 2)
 */
function subsecoes_capturarDetalhes(url) {
  try {
    const html = UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getContentText();
    const presMatch = /<strong>Presidente<\/strong>:(.*?)<br>/i.exec(html);
    const presidente = presMatch ? subsecoes_limparTexto([null, presMatch[1]]) : "Não informado";

    const cidades = [];
    const abaCidadesMatch = /id="aba4"[\s\S]*?<ul>([\s\S]*?)<\/ul>/i.exec(html);
    if (abaCidadesMatch) {
      const liRegex = /<li>(.*?)<\/li>/gi;
      let liMatch;
      while ((liMatch = liRegex.exec(abaCidadesMatch[1])) !== null) {
        cidades.push(subsecoes_limparTexto([null, liMatch[1]]));
      }
    }
    return { presidente, cidades };
  } catch (e) {
    return { presidente: "Erro", cidades: [] };
  }
}

/**
 * Decodifica Cloudflare Email
 */
function subsecoes_decodificarEmail(encodedString) {
  let email = "";
  let k = parseInt(encodedString.substr(0, 2), 16);
  for (let n = 2; n < encodedString.length; n += 2) {
    let i = parseInt(encodedString.substr(n, 2), 16) ^ k;
    email += String.fromCharCode(i);
  }
  return email;
}

/**
 * Limpeza de texto HTML
 */
function subsecoes_limparTexto(matchArray) {
  if (!matchArray || !matchArray[1]) return "";
  return matchArray[1]
    .replace(/<\/?[^>]+(>|$)/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}