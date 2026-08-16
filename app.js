
const ExcelJS = require("exceljs");

const BASE_URL = "https://pncp.gov.br/api/consulta";

const CONFIG = {
  dataInicial: "20250101",
  dataFinal: "20250131",
  codigoModalidadeContratacao: 6, 
  uf: null, // 
  tamanhoPagina: 50,
  maxPaginas: 3,
  atrasoEntreChamadasMs: 300,
};


function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
async function fetchJson(url) {
  const resposta = await fetch(url, {
    method: "GET",
    headers: { Accept: "application/json" },
  });

  if (resposta.status === 204) {
    return null;
  }

  if (!resposta.ok) {
    throw new Error(`HTTP ${resposta.status} em ${url}`);
  }

  return resposta.json();
}

async function buscarContratacoes() {
  const contratacoes = [];
  let pagina = 1;

  while (pagina <= CONFIG.maxPaginas) {
    const params = new URLSearchParams({
      dataInicial: CONFIG.dataInicial,
      dataFinal: CONFIG.dataFinal,
      codigoModalidadeContratacao: String(CONFIG.codigoModalidadeContratacao),
      pagina: String(pagina),
      tamanhoPagina: String(CONFIG.tamanhoPagina),
    });
    if (CONFIG.uf) params.set("uf", CONFIG.uf);

    const url = `${BASE_URL}/v1/contratacoes/publicacao?${params.toString()}`;
    console.log(`[contratações] página ${pagina} -> ${url}`);

    let json;
    try {
      json = await fetchJson(url);
    } catch (erro) {
      console.error(`  Falha ao buscar página ${pagina}: ${erro.message}`);
      break;
    }

    if (!json || !Array.isArray(json.data) || json.data.length === 0) {
      console.log("  Sem mais registros. Encerrando paginação.");
      break;
    }

    contratacoes.push(...json.data);

    const totalPaginas = json.totalPaginas ?? pagina;
    if (pagina >= totalPaginas) break;

    pagina += 1;
    await delay(CONFIG.atrasoEntreChamadasMs);
  }

  return contratacoes;
}

async function buscarItensDaContratacao(cnpj, ano, sequencial) {
  const url = `${BASE_URL}/v1/orgaos/${cnpj}/compras/${ano}/${sequencial}/itens`;
  try {
    const json = await fetchJson(url);
    return Array.isArray(json) ? json : json?.data ?? [];
  } catch (erro) {
    console.warn(`  [itens] falha em ${url}: ${erro.message}`);
    return [];
  }
}

async function buscarResultadosDoItem(cnpj, ano, sequencial, numeroItem) {
  const url = `${BASE_URL}/v1/orgaos/${cnpj}/compras/${ano}/${sequencial}/itens/${numeroItem}/resultados`;
  try {
    const json = await fetchJson(url);
    if (!json) return [];
    return Array.isArray(json) ? json : json?.data ?? [];
  } catch (erro) {
    console.warn(`  [resultados] falha em ${url}: ${erro.message}`);
    return [];
  }
}

// extrair dados

async function extrairDados() {
  const linhasItens = [];
  const linhasSemResultado = [];

  const contratacoes = await buscarContratacoes();
  console.log(`\nTotal de contratações encontradas: ${contratacoes.length}\n`);

  for (const contratacao of contratacoes) {
    const cnpj = contratacao.orgaoEntidade?.cnpj;
    const ano = contratacao.anoCompra;
    const sequencial = contratacao.sequencialCompra;
    const numeroControlePNCP = contratacao.numeroControlePNCP ?? null;
    const objeto = contratacao.objetoCompra ?? null;
    const orgao = contratacao.orgaoEntidade?.razaoSocial ?? null;

    if (!cnpj || !ano || !sequencial) {
      console.warn(
        `  Contratação sem cnpj/ano/sequencial completos, pulando itens: ${numeroControlePNCP}`
      );
      continue;
    }

    await delay(CONFIG.atrasoEntreChamadasMs);
    const itens = await buscarItensDaContratacao(cnpj, ano, sequencial);

    for (const item of itens) {
      const numeroItem = item.numeroItem;
      if (numeroItem === undefined) continue;

      await delay(CONFIG.atrasoEntreChamadasMs);
      const resultados = await buscarResultadosDoItem(
        cnpj,
        ano,
        sequencial,
        numeroItem
      );

      if (resultados.length === 0) {
        linhasSemResultado.push({
          numeroControlePNCP,
          orgao,
          numeroItem,
          descricaoItem: item.descricao ?? null,
          motivo: "Sem resultado/adjudicação publicado na API para este item",
        });
        continue;
      }

      for (const resultado of resultados) {
        linhasItens.push({
          processo: numeroControlePNCP ?? null,
          lote: item.numeroItem ?? null,
          item: item.descricao ?? null,
          valorProposta:
            resultado.valorTotalHomologado ??
            resultado.valorTotal ??
            resultado.valorProposta ??
            null,
          valorUnitario:
            resultado.valorUnitarioHomologado ??
            resultado.valorUnitario ??
            null,
          valorNegociado:
            resultado.valorTotalHomologado ??
            resultado.valorTotal ??
            resultado.valorNegociado ??
            null,
          marcaFabricante:
            resultado.marca ??
            resultado.fabricante ??
            resultado.marcaFabricante ??
            null,
          modeloVersao:
            resultado.modeloVersao ??
            resultado.descricaoModelo ??
            resultado.modelo ??
            null,
        });
      }
    }
  }

  return { linhasItens, linhasSemResultado };
}

// excel
async function exportarParaExcel({ linhasItens, linhasSemResultado }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "extrairPropostasPNCP.js";
  workbook.created = new Date();
 
  const abaItens = workbook.addWorksheet("Itens Adjudicados");
  abaItens.columns = [
    { header: "Processo", key: "processo", width: 30 },
    { header: "Lote", key: "lote", width: 10 },
    { header: "Item", key: "item", width: 40 },
    { header: "Valor Proposta", key: "valorProposta", width: 16 },
    { header: "Valor Unitário", key: "valorUnitario", width: 16 },
    { header: "Valor Negociado", key: "valorNegociado", width: 16 },
    { header: "Marca/Fabricante", key: "marcaFabricante", width: 20 },
    { header: "Modelo/Versão", key: "modeloVersao", width: 20 },
  ];
  abaItens.addRows(linhasItens);
  abaItens.getRow(1).font = { bold: true };
  abaItens.getColumn("valorProposta").numFmt = '"R$" #,##0.00';
  abaItens.getColumn("valorUnitario").numFmt = '"R$" #,##0.00';
  abaItens.getColumn("valorNegociado").numFmt = '"R$" #,##0.00';
 
  const abaSemResultado = workbook.addWorksheet("Sem Resultado Publicado");
  abaSemResultado.columns = [
    { header: "Nº Controle PNCP", key: "numeroControlePNCP", width: 30 },
    { header: "Órgão", key: "orgao", width: 40 },
    { header: "Nº Item", key: "numeroItem", width: 10 },
    { header: "Descrição Item", key: "descricaoItem", width: 40 },
    { header: "Motivo", key: "motivo", width: 45 },
  ];
  abaSemResultado.addRows(linhasSemResultado);
  abaSemResultado.getRow(1).font = { bold: true };
 
  const caminhoArquivo = "./propostas_pncp.xlsx";
  await workbook.xlsx.writeFile(caminhoArquivo);
  console.log(`\nArquivo Excel gerado: ${caminhoArquivo}`);
  console.log(`  - Itens adjudicados: ${linhasItens.length}`);
  console.log(`  - Itens sem resultado publicado: ${linhasSemResultado.length}`);
}


// ---------------------- EXECUÇÃO --------------------------------------------

(async () => {
  try {
    const { linhasItens, linhasSemResultado } = await extrairDados();

    console.log("\n===== ITENS ADJUDICADOS =====");
    console.log(JSON.stringify(linhasItens, null, 2));

    console.log("\n===== SEM RESULTADO PUBLICADO =====");
    console.log(JSON.stringify(linhasSemResultado, null, 2));

    // console.log(`\nResumo: ${linhasItens.length} itens com resultado, ${linhasSemResultado.length} sem resultado.`);
  } catch (erro) {
    console.error("Erro fatal na extração:", erro);
    process.exit(1);
  }
})();