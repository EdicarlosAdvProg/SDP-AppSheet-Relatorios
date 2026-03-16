/**
 * @fileoverview Lógica de Backend para Gestão de Subseções
 * Bloco: FormSubsecoes-macro.gs
 */

/**
 * Abre o formulário de Gestão de Subseções como diálogo modal
 */
function abrirFormSubsecoes() {
  const html = HtmlService.createTemplateFromFile('FormSubsecoes-layout')
      .evaluate()
      .setTitle('Gestão de Subseções - OAB/GO')
      .setWidth(1200)
      .setHeight(800);
  DocumentApp.getUi().showModalDialog(html, ' ');
}

// ---------------------------------------------------------------------------
// CARREGAMENTO DE DADOS
// ---------------------------------------------------------------------------

/**
 * Busca todas as subseções enriquecidas com procurador responsável e
 * lista de cidades, em uma única chamada ao back-end.
 *
 * Relações:
 *   tabSubsecoes.Região  →  tabProcuradores.Região  (N:1)
 *   tabCidades.IdSubsecao →  tabSubsecoes.Id          (N:1)
 *
 * @returns {Array<Object>} Lista ordenada alfabeticamente com o shape:
 *   { id, nome, presidente, email, telefone, regiao,
 *     procurador: { id, nome, cargo, email, telefone } | null,
 *     cidades: string[] }
 */
function subsecoes_buscarTodas() {
  const ss = SpreadsheetApp.openById(PLANILHA_DADOS_ID);

  // --- Leitura das três tabelas ---
  const dadosSub  = ss.getSheetByName('tabSubsecoes').getDataRange().getValues();
  const dadosProc = ss.getSheetByName('tabProcuradores').getDataRange().getValues();
  const dadosCid  = ss.getSheetByName('tabCidades').getDataRange().getValues();

  dadosSub.shift();   // remove cabeçalho
  dadosProc.shift();
  dadosCid.shift();

  // --- Mapa: REGIÃO_NORMALIZADA → objeto procurador ---
  // Ex: "TODAS (SEDE)" → { id, nome, cargo, email, telefone }
  const mapaProcuradores = {};
  dadosProc.forEach(function (p) {
    const chave = String(p[5] || '').trim().toUpperCase();
    mapaProcuradores[chave] = {
      id:       p[0],
      nome:     p[1],
      email:    p[2],
      telefone: p[3],
      cargo:    p[4]
    };
  });

  // --- Mapa: idSubsecao → [ 'CIDADE A', 'CIDADE B', ... ] ---
  const mapaCidades = {};
  dadosCid.forEach(function (c) {
    const idSub = c[1];
    if (!mapaCidades[idSub]) mapaCidades[idSub] = [];
    mapaCidades[idSub].push(String(c[2] || '').trim());
  });

  // --- Montagem da lista enriquecida ---
  const fallbackProc = mapaProcuradores['TODAS (SEDE)'] || null;

  const lista = dadosSub.map(function (linha) {
    const regiaoNorm = String(linha[5] || '').trim().toUpperCase();
    const idSub      = linha[0];

    // Procurador: tenta pela região exata; se não achar, usa o da Sede
    const procurador = mapaProcuradores[regiaoNorm] || fallbackProc;

    return {
      id:          idSub,
      nome:        String(linha[1] || '').trim(),
      presidente:  String(linha[2] || '').trim(),
      email:       String(linha[3] || '').trim(),
      telefone:    String(linha[4] || '').trim(),
      regiao:      String(linha[5] || '').trim(),
      procurador:  procurador,
      cidades:     (mapaCidades[idSub] || []).sort(function (a, b) {
                     return a.localeCompare(b, 'pt-BR');
                   })
    };
  });

  return lista.sort(function (a, b) {
    return a.nome.localeCompare(b.nome, 'pt-BR');
  });
}

// ---------------------------------------------------------------------------
// AÇÕES DO MENU
// ---------------------------------------------------------------------------

/**
 * Dispara o robô de sincronização e retorna resultado para o front-end
 * @returns {{ sucesso: boolean, msg: string }}
 */
function subsecoes_executarSincronizacao() {
  try {
    const resultado = subsecoes_sincronizarComSite();
    return { sucesso: true, msg: resultado };
  } catch (e) {
    return { sucesso: false, msg: 'Erro na sincronização: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// ROBÔ DE SINCRONIZAÇÃO COM O SITE DA OAB
// ---------------------------------------------------------------------------

/**
 * Sincroniza tabSubsecoes e tabCidades com o site oficial da OAB-GO
 */
function subsecoes_sincronizarComSite() {
  const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const shSub = ss.getSheetByName('tabSubsecoes');
  const shCid = ss.getSheetByName('tabCidades');

  if (shSub.getLastRow() > 1) shSub.getRange(2, 1, shSub.getLastRow() - 1, shSub.getLastColumn()).clearContent();
  if (shCid.getLastRow() > 1) shCid.getRange(2, 1, shCid.getLastRow() - 1, shCid.getLastColumn()).clearContent();

  const urlIndex   = 'https://www.oabgo.org.br/subsecoes/';
  const htmlIndex  = UrlFetchApp.fetch(urlIndex).getContentText();
  const regexBloco = /<aside class="item">([\s\S]*?)<\/aside>/g;

  const listaSubsecoes = [];
  const listaCidades   = [];
  let ultimoIdSub = '';
  let ultimoIdCid = '';

  // Sede (Goiânia) inserida manualmente
  listaSubsecoes.push(['SUB-SEDE', 'GOIÂNIA', '', 'oabnet@oabgo.org.br', '(62) 3238-2000', 'Todas (Sede)']);
  ultimoIdCid = gerarProximoIdIncremental(ultimoIdCid);
  listaCidades.push(['CID-' + ultimoIdCid, 'SUB-SEDE', 'GOIÂNIA']);

  let match;
  while ((match = regexBloco.exec(htmlIndex)) !== null) {
    const bloco        = match[1];
    const nomeSubRaw   = subsecoes_limparTexto(/<h3>(.*?)<\/h3>/.exec(bloco));
    const nomeSubLimpo = nomeSubRaw.toUpperCase().replace('OAB - ', '').trim();

    if (nomeSubLimpo === 'GOIÂNIA') continue;

    let telefone = '';
    const trechoTel = /<i class="fas fa-phone"><\/i>([\s\S]*?)class="btn">Saiba mais/i.exec(bloco);
    if (trechoTel) {
      const telMatch = /\(\d{2}\)\s\d{4,5}-\d{4}/.exec(trechoTel[1]);
      if (telMatch) telefone = telMatch[0];
    }

    const cfEmailMatch = /data-cfemail="(.*?)"/.exec(bloco);
    const email        = cfEmailMatch ? subsecoes_decodificarEmail(cfEmailMatch[1]) : '';
    const linkDetalhe  = /href="(https:\/\/www.oabgo.org.br\/subsecao\/.*?\/)"/.exec(bloco);

    if (linkDetalhe) {
      const regiaoSubsecao = buscarRegiaoNoArquivo(nomeSubLimpo);
      ultimoIdSub = gerarProximoIdIncremental(ultimoIdSub);
      const idSubAtual = 'SUB-' + ultimoIdSub;
      const detalhes   = subsecoes_capturarDetalhes(linkDetalhe[1]);

      listaSubsecoes.push([idSubAtual, nomeSubLimpo, detalhes.presidente, email, telefone, regiaoSubsecao]);

      detalhes.cidades.forEach(function (cidade) {
        ultimoIdCid = gerarProximoIdIncremental(ultimoIdCid);
        listaCidades.push(['CID-' + ultimoIdCid, idSubAtual, cidade.toUpperCase()]);
      });
    }
  }

  if (listaSubsecoes.length > 0) shSub.getRange(2, 1, listaSubsecoes.length, 6).setValues(listaSubsecoes);
  if (listaCidades.length  > 0) shCid.getRange(2, 1, listaCidades.length,   3).setValues(listaCidades);

  return 'Sincronização concluída. ' + listaSubsecoes.length + ' subseção(ões) registrada(s).';
}

function buscarRegiaoNoArquivo(nomeSubsecao) {
  const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const shArq = ss.getSheetByName('tabSubsecoesArquivo');
  if (!shArq) return 'Todas (Sede)';
  const dados = shArq.getDataRange().getValues();
  const busca = nomeSubsecao.toUpperCase().trim();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().toUpperCase().trim() === busca) return dados[i][1];
  }
  return 'Todas (Sede)';
}

function subsecoes_capturarDetalhes(url) {
  try {
    const html       = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
    const presMatch  = /<strong>Presidente<\/strong>:(.*?)<br>/i.exec(html);
    const presidente = presMatch ? subsecoes_limparTexto([null, presMatch[1]]) : 'Não informado';
    const cidades    = [];
    const abaMatch   = /id="aba4"[\s\S]*?<ul>([\s\S]*?)<\/ul>/i.exec(html);
    if (abaMatch) {
      const liRegex = /<li>(.*?)<\/li>/gi;
      let liMatch;
      while ((liMatch = liRegex.exec(abaMatch[1])) !== null) {
        cidades.push(subsecoes_limparTexto([null, liMatch[1]]));
      }
    }
    return { presidente, cidades };
  } catch (e) {
    return { presidente: 'Erro', cidades: [] };
  }
}

function subsecoes_decodificarEmail(encodedString) {
  let email = '';
  const k   = parseInt(encodedString.substr(0, 2), 16);
  for (let n = 2; n < encodedString.length; n += 2) {
    email += String.fromCharCode(parseInt(encodedString.substr(n, 2), 16) ^ k);
  }
  return email;
}

function subsecoes_limparTexto(matchArray) {
  if (!matchArray || !matchArray[1]) return '';
  return matchArray[1]
    .replace(/<\/?[^>]+(>|$)/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}