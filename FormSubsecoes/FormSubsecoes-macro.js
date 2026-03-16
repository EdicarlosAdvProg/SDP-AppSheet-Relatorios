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
// EDIÇÃO DE SUBSEÇÃO
// ---------------------------------------------------------------------------

/**
 * Atualiza os campos editáveis de uma subseção na tabSubsecoes.
 * Apenas colunas 2–6 (Subseção, Presidente, Email, Telefone, Região)
 * são gravadas. A coluna 1 (Id) nunca é alterada.
 *
 * @param {{ id, nome, presidente, email, telefone, regiao }} dados
 * @returns {{ sucesso: boolean, msg: string }}
 */
function subsecoes_salvarEdicao(dados) {
  try {
    if (!dados || !dados.id) throw new Error('ID da subseção não informado.');

    const ss  = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sh  = ss.getSheetByName('tabSubsecoes');
    const all = sh.getDataRange().getValues();

    // Localiza a linha pelo Id (coluna 1, índice 0)
    let linhaAlvo = -1;
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]).trim() === String(dados.id).trim()) {
        linhaAlvo = i + 1; // +1 porque getValues é 0-indexed mas Sheets é 1-indexed
        break;
      }
    }

    if (linhaAlvo === -1) throw new Error('Subseção não encontrada: ' + dados.id);

    // Grava colunas 2–6 (B até F) em uma única operação
    sh.getRange(linhaAlvo, 2, 1, 5).setValues([[
      String(dados.nome       || '').trim(),
      String(dados.presidente || '').trim(),
      String(dados.email      || '').trim(),
      String(dados.telefone   || '').trim(),
      String(dados.regiao     || '').trim()
    ]]);

    SpreadsheetApp.flush();
    return { sucesso: true, msg: 'Subseção "' + dados.nome + '" atualizada com sucesso.' };

  } catch (e) {
    return { sucesso: false, msg: 'Erro ao salvar: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// CRUD DE PROCURADORES
// ---------------------------------------------------------------------------

/**
 * Retorna todos os procuradores da tabProcuradores.
 * Shape: { id, nome, email, telefone, cargo, regiao }
 * @returns {Array<Object>}
 */
function procuradores_buscarTodos() {
  const ss    = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const dados = ss.getSheetByName('tabProcuradores').getDataRange().getValues();
  dados.shift(); // remove cabeçalho
  return dados.map(function (r) {
    return {
      id:       String(r[0] || '').trim(),
      nome:     String(r[1] || '').trim(),
      email:    String(r[2] || '').trim(),
      telefone: String(r[3] || '').trim(),
      cargo:    String(r[4] || '').trim(),
      regiao:   String(r[5] || '').trim()
    };
  });
}

/**
 * Atualiza os campos de um procurador na tabProcuradores.
 * Apenas colunas 2–6 (Nome, Email, Telefone, Cargo, Região) são alteradas.
 * @param {{ id, nome, email, telefone, cargo, regiao }} dados
 * @returns {{ sucesso: boolean, msg: string }}
 */
function procuradores_salvarEdicao(dados) {
  try {
    if (!dados || !dados.id) throw new Error('ID do procurador não informado.');

    const ss  = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sh  = ss.getSheetByName('tabProcuradores');
    const all = sh.getDataRange().getValues();

    let linhaAlvo = -1;
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]).trim() === String(dados.id).trim()) {
        linhaAlvo = i + 1;
        break;
      }
    }

    if (linhaAlvo === -1) throw new Error('Procurador não encontrado: ' + dados.id);

    sh.getRange(linhaAlvo, 2, 1, 5).setValues([[
      String(dados.nome     || '').trim(),
      String(dados.email    || '').trim(),
      String(dados.telefone || '').trim(),
      String(dados.cargo    || '').trim(),
      String(dados.regiao   || '').trim()
    ]]);

    SpreadsheetApp.flush();
    return { sucesso: true, msg: 'Procurador "' + dados.nome + '" atualizado com sucesso.' };
  } catch (e) {
    return { sucesso: false, msg: 'Erro ao salvar: ' + e.message };
  }
}

/**
 * Remove um procurador da tabProcuradores pelo Id.
 * @param {string} id
 * @returns {{ sucesso: boolean, msg: string }}
 */
function procuradores_excluir(id) {
  try {
    if (!id) throw new Error('ID não informado.');

    const ss  = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sh  = ss.getSheetByName('tabProcuradores');
    const all = sh.getDataRange().getValues();

    let linhaAlvo = -1;
    let nomeProc  = '';
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]).trim() === String(id).trim()) {
        linhaAlvo = i + 1;
        nomeProc  = String(all[i][1]).trim();
        break;
      }
    }

    if (linhaAlvo === -1) throw new Error('Procurador não encontrado: ' + id);

    sh.deleteRow(linhaAlvo);
    SpreadsheetApp.flush();
    return { sucesso: true, msg: 'Procurador "' + nomeProc + '" removido com sucesso.' };
  } catch (e) {
    return { sucesso: false, msg: 'Erro ao excluir: ' + e.message };
  }
}

/**
 * Cria um novo procurador na tabProcuradores.
 * O Id é gerado incrementalmente a partir do maior Id existente.
 * @param {{ nome, email, telefone, cargo, regiao }} dados
 * @returns {{ sucesso: boolean, msg: string, id: string }}
 */
function procuradores_criar(dados) {
  try {
    if (!dados || !dados.nome || !dados.nome.trim()) throw new Error('Nome obrigatório.');

    const ss  = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sh  = ss.getSheetByName('tabProcuradores');
    const all = sh.getDataRange().getValues();

    // Encontra o maior número sequencial existente nos IDs "PROC-XX"
    let maxNum = 0;
    for (let i = 1; i < all.length; i++) {
      const m = /^PROC-(\d+)$/i.exec(String(all[i][0]).trim());
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
    const novoId = 'PROC-' + String(maxNum + 1).padStart(2, '0');

    sh.appendRow([
      novoId,
      String(dados.nome     || '').trim(),
      String(dados.email    || '').trim(),
      String(dados.telefone || '').trim(),
      String(dados.cargo    || '').trim(),
      String(dados.regiao   || '').trim()
    ]);

    SpreadsheetApp.flush();
    return { sucesso: true, msg: 'Procurador "' + dados.nome.trim() + '" cadastrado.', id: novoId };
  } catch (e) {
    return { sucesso: false, msg: 'Erro ao cadastrar: ' + e.message, id: '' };
  }
}


 /* Fonte: valores distintos de tabProcuradores.Região + 'Todas (Sede)'.
 * @returns {string[]}
 */
function subsecoes_listarRegioes() {
  const ss      = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
  const dados   = ss.getSheetByName('tabProcuradores').getDataRange().getValues();
  dados.shift();
  const regioes = dados
    .map(function (r) { return String(r[5] || '').trim(); })
    .filter(function (v) { return v !== ''; });
  // Garante que "Todas (Sede)" apareça primeiro e sem duplicatas
  const unicas  = ['Todas (Sede)'].concat(
    regioes.filter(function (v) { return v.toUpperCase() !== 'TODAS (SEDE)'; })
  );
  return unicas;
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

/**
 * Adiciona uma cidade à tabCidades vinculada a uma subseção.
 * @param {string} idSubsecao
 * @param {string} nomeCidade
 * @returns {{ sucesso: boolean, msg: string }}
 */
function cidades_adicionar(idSubsecao, nomeCidade) {
  try {
    if (!idSubsecao) throw new Error('ID da subseção não informado.');
    const nome = String(nomeCidade || '').trim().toUpperCase();
    if (!nome) throw new Error('Nome da cidade é obrigatório.');

    const ss  = SpreadsheetApp.openById(PLANILHA_DADOS_ID);
    const sh  = ss.getSheetByName('tabCidades');
    const all = sh.getDataRange().getValues();

    // Verifica duplicata na subseção
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][1]).trim() === idSubsecao &&
          String(all[i][2]).trim().toUpperCase() === nome) {
        throw new Error('"' + nome + '" já está cadastrada nesta subseção.');
      }
    }

    // Gera próximo ID incremental
    let maxNum = 0;
    for (let i = 1; i < all.length; i++) {
      const m = /^CID-(.+)$/.exec(String(all[i][0]).trim());
      if (m) maxNum++;
    }
    const novoId = 'CID-' + gerarProximoIdIncremental(String(maxNum));

    sh.appendRow([novoId, idSubsecao, nome]);
    SpreadsheetApp.flush();

    return { sucesso: true, msg: 'Cidade "' + nome + '" adicionada com sucesso.' };
  } catch (e) {
    return { sucesso: false, msg: e.message };
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