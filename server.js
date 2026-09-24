const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// path.join(__dirname, 'public') funciona independente de onde o processo foi iniciado —
// importante quando o app roda empacotado (Electron), onde o diretório de trabalho muda.
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Formata nomes de tabelas e colunas conforme o banco de dados.
// IMPORTANTE: nunca altera o conteúdo do nome, apenas embrulha no
// delimitador correto — os nomes chegam do usuário já tratados.
function formatarIdentifier(nome, banco) {
  const limpo = String(nome).trim();
  if (banco === 'sqlserver') return `[${limpo}]`;
  if (banco === 'postgres') return `"${limpo}"`;
  return `\`${limpo}\``; // MySQL / MariaDB
}

// Formata um objeto Date (vindo de célula de data real do Excel) em texto SQL.
// Usa os getters UTC porque o SheetJS constrói o Date a partir do serial do
// Excel usando Date.UTC(...) internamente — ler em horário local causaria
// um deslocamento de fuso horário no dia/hora.
function formatarDataSQL(date) {
  const pad = n => String(n).padStart(2, '0');
  const ano = date.getUTCFullYear();
  const mes = pad(date.getUTCMonth() + 1);
  const dia = pad(date.getUTCDate());
  const hora = pad(date.getUTCHours());
  const min = pad(date.getUTCMinutes());
  const seg = pad(date.getUTCSeconds());

  const temHora = date.getUTCHours() !== 0 || date.getUTCMinutes() !== 0 || date.getUTCSeconds() !== 0;
  const dataStr = `${ano}-${mes}-${dia}`;
  return temHora ? `${dataStr} ${hora}:${min}:${seg}` : dataStr;
}

// Classifica um tipo de coluna REAL (vindo do banco de destino, ex: "int", "varchar",
// "datetime") em 3 categorias amplas — é só isso que a gente precisa pra decidir como
// formatar o valor: com aspas, como número puro, ou como data.
function categoriaDoTipoReal(tipoReal) {
  if (!tipoReal) return null;
  const t = String(tipoReal).toLowerCase();
  if (/date|time/.test(t)) return 'data';
  if (/char|text|nchar|nvarchar|string|uuid|uniqueidentifier/.test(t)) return 'texto';
  if (/int|decimal|numeric|float|real|money|bit|double|bigint|smallint|tinyint/.test(t)) return 'numero';
  return null; // tipo não reconhecido — cai na heurística antiga
}

// Lê a linha colada pelo usuário (formato "coluna:tipo;coluna:tipo;...", exatamente o
// que os scripts de consulta do front devolvem) e monta um mapa { coluna: tipoReal }.
function parseEstruturaReal(texto) {
  const mapa = {};
  if (!texto || typeof texto !== 'string') return mapa;

  texto.trim().split(';').forEach(par => {
    const idx = par.indexOf(':');
    if (idx === -1) return;
    const coluna = par.slice(0, idx).trim();
    const tipo = par.slice(idx + 1).trim();
    if (coluna && tipo) mapa[coluna] = tipo;
  });

  return mapa;
}

// Detecta se uma string com zero à esquerda é provavelmente um CÓDIGO
// (NCM, CEST, código de barras, etc.) e não um número — nesses casos,
// tratar como número apagaria os zeros à esquerda e corromperia o dado.
function pareceCodigoComZeroEsquerda(strVal) {
  return /^-?0\d/.test(strVal.trim());
}

// Diz se uma string pode virar um valor numérico puro no SQL com segurança
// (sem perder zeros à esquerda importantes).
function pareceNumeroSeguro(strVal) {
  const limpo = strVal.trim();
  if (limpo === '' || pareceCodigoComZeroEsquerda(limpo)) return false;
  const num = Number(limpo.replace(',', '.'));
  return !isNaN(num) && isFinite(num);
}

// Formata valores para inserção/atualização SQL.
// Se "tipoReal" for informado (veio da linha colada pelo usuário, com o tipo de
// verdade da coluna no banco de destino), ele manda — sem adivinhação nenhuma.
// Sem ele, cai na heurística de sempre (deduzir pelo formato do valor na planilha).
function formatarValor(val, banco, tipoReal) {
  if (val === null || val === undefined || String(val).trim() === '') {
    return 'NULL';
  }

  const categoria = categoriaDoTipoReal(tipoReal);

  // Banco disse que a coluna é TEXTO: nunca converte pra número, mesmo que pareça um.
  // Isso protege exatamente os casos como NCM/CEST/código de barras guardados como varchar.
  if (categoria === 'texto') {
    const bruto = val instanceof Date ? formatarDataSQL(val) : String(val);
    let strVal = bruto;
    if (banco === 'mysql') strVal = strVal.replace(/\\/g, '\\\\');
    strVal = strVal.replace(/'/g, "''");
    const prefixo = banco === 'sqlserver' ? 'N' : '';
    return `${prefixo}'${strVal}'`;
  }

  // Banco disse que a coluna é NUMÉRICA: sempre número puro, sem aspas.
  if (categoria === 'numero') {
    if (typeof val === 'number') return val;
    const num = Number(String(val).trim().replace(',', '.'));
    return isNaN(num) ? 'NULL' : num;
  }

  // Célula que o Excel já reconhece como data real (cellDates: true na leitura)
  if (val instanceof Date) {
    return `'${formatarDataSQL(val)}'`;
  }

  if (typeof val === 'number') {
    return val;
  }

  const strValOriginal = String(val);

  // Sem tipo real conhecido: heurística de sempre — número "de verdade"
  // (ex: "0.00000", "123", "45,90") vira valor numérico puro; strings com
  // zero à esquerda (prováveis códigos) ficam de fora dessa regra.
  if (pareceNumeroSeguro(strValOriginal)) {
    return Number(strValOriginal.trim().replace(',', '.'));
  }

  let strVal = strValOriginal;

  // MySQL trata backslash como caractere de escape dentro de strings
  if (banco === 'mysql') {
    strVal = strVal.replace(/\\/g, '\\\\');
  }
  strVal = strVal.replace(/'/g, "''");

  // N'' no SQL Server garante que acentuação/caracteres especiais não sejam corrompidos
  const prefixo = banco === 'sqlserver' ? 'N' : '';
  return `${prefixo}'${strVal}'`;
}

// Infere o tipo de coluna para o modo 'Criar do Zero'
function inferirTipoColuna(valores, banco) {
  let temDado = false;
  let temData = false;     // algum valor é uma data real (célula formatada como data no Excel)
  let temNaoData = false;  // algum valor NÃO é uma data real
  let temHora = false;

  let ehNumero = true;
  let ehInt = true;

  for (const val of valores) {
    if (val === null || val === undefined || String(val).trim() === '') continue;
    temDado = true;

    if (val instanceof Date) {
      temData = true;
      if (val.getUTCHours() !== 0 || val.getUTCMinutes() !== 0 || val.getUTCSeconds() !== 0) {
        temHora = true;
      }
      continue;
    }

    temNaoData = true;

    if (typeof val === 'number') {
      if (!Number.isInteger(val)) ehInt = false;
    } else if (typeof val === 'string') {
      const bruto = val.trim();
      if (pareceCodigoComZeroEsquerda(bruto)) {
        // Ex: "007", "0012345" — provavelmente um código, não um número de verdade
        ehNumero = false;
      } else {
        const num = Number(bruto.replace(',', '.'));
        if (isNaN(num)) {
          ehNumero = false;
        } else if (!Number.isInteger(num) || /[.,]/.test(bruto)) {
          // "/[.,]/" pega casos como "0.00000": matematicamente é inteiro (zero),
          // mas o formato original indica que a coluna é decimal (ex: custo/preço).
          ehInt = false;
        }
      }
    } else {
      ehNumero = false;
    }
  }

  if (!temDado) return 'VARCHAR(255)';

  // Só assume DATE/DATETIME quando TODOS os valores da coluna são datas reais.
  // Se aparecer qualquer texto/número junto, não arrisca — trata como texto,
  // já que planilhas de bancos diferentes podem guardar "data" como string.
  if (temData && !temNaoData) {
    if (banco === 'sqlserver') return temHora ? 'DATETIME2' : 'DATE';
    if (banco === 'postgres') return temHora ? 'TIMESTAMP' : 'DATE';
    return temHora ? 'DATETIME' : 'DATE'; // MySQL / MariaDB
  }

  if (temNaoData && !temData && ehNumero) {
    if (ehInt) return banco === 'postgres' ? 'INTEGER' : 'INT';
    return banco === 'postgres' ? 'NUMERIC(15,2)' : 'DECIMAL(15,2)';
  }

  return banco === 'sqlserver' ? 'VARCHAR(MAX)' : 'TEXT';
}

// Normaliza o modo vindo da interface
function normalizarModo(modo) {
  if (!modo) return 'inserir';
  const m = String(modo).toLowerCase().trim();
  if (m.includes('criar')) return 'criar';
  if (m.includes('substituir')) return 'substituir';
  return 'inserir';
}

// Gerador do script SQL — devolve um array de "blocos" (até TAMANHO_BLOCO linhas cada),
// prontos pra exibir e copiar separadamente, em vez de um único script gigante.
function gerarSqlScript(linhas, nomeTabela, bancoDados, modoMigracao, colunasPermitidas, colunaChave, mapaTiposReais, mapeamentoColunas) {
  if (!linhas || linhas.length === 0) {
    throw new Error('A planilha está vazia ou não possui dados válidos.');
  }

  const tiposReais = mapaTiposReais || {};
  const mapeamento = mapeamentoColunas || {};

  // Nome que a coluna vai ter no DESTINO (banco). Se o usuário não mapeou essa
  // coluna explicitamente, assume que o nome é o mesmo da planilha.
  const colunaDestino = (colOrigem) => mapeamento[colOrigem] || colOrigem;

  const modo = normalizarModo(modoMigracao);
  const tabelaFormatada = formatarIdentifier(nomeTabela, bancoDados);

  // Filtra as colunas se o utilizador tiver especificado uma seleção
  let colunas = Object.keys(linhas[0]);
  if (colunasPermitidas && Array.isArray(colunasPermitidas) && colunasPermitidas.length > 0) {
    colunas = colunas.filter(col => colunasPermitidas.includes(col));
  }

  if (colunas.length === 0) {
    throw new Error('Nenhuma coluna foi selecionada para gerar o script.');
  }

  // Permite escolher a coluna chave; se não vier, assume a primeira coluna selecionada.
  // colunaPk é o nome na PLANILHA (usado pra ler o valor); colunaPkDestino é o nome no BANCO.
  const colunaPk = (colunaChave && colunas.includes(colunaChave)) ? colunaChave : colunas[0];
  const colunaPkDestino = colunaDestino(colunaPk);
  const TAMANHO_BLOCO = 500;

  // ---------------- CABEÇALHO (comentários + SET NOCOUNT) ----------------
  let cabecalho = `-- Script de Migração SQL\n`;
  cabecalho += `-- Tabela: ${tabelaFormatada}\n`;
  cabecalho += `-- Banco Alvo: ${bancoDados.toUpperCase()}\n`;
  cabecalho += `-- Modo Selecionado: ${modo.toUpperCase()}\n`;
  cabecalho += `-- Coluna Chave (PK): ${colunaPk}${colunaPkDestino !== colunaPk ? ` → ${colunaPkDestino}` : ''}\n`;
  cabecalho += `-- Colunas Incluídas: ${colunas.join(', ')}\n`;
  const colunasRenomeadas = colunas.filter(col => colunaDestino(col) !== col);
  if (colunasRenomeadas.length > 0) {
    cabecalho += `-- Mapeamento (Planilha → Banco): ${colunasRenomeadas.map(col => `${col}→${colunaDestino(col)}`).join(', ')}\n`;
  }
  cabecalho += `-- Total de Registros: ${linhas.length}\n`;
  cabecalho += `-- OBS: cada linha roda como comando independente — se alguma der erro,\n`;
  cabecalho += `-- o banco registra o erro daquela linha e segue para as próximas.\n\n`;
  if (bancoDados === 'sqlserver') {
    cabecalho += `SET NOCOUNT ON;\n\n`;
  }

  // ---------------- PREPARAÇÃO (CREATE / DROP / TRUNCATE / IDENTITY_INSERT ON) ----------------
  let preparacao = '';

  if (modo === 'criar') {
    if (bancoDados === 'sqlserver') {
      preparacao += `IF OBJECT_ID('${nomeTabela}', 'U') IS NOT NULL DROP TABLE ${tabelaFormatada};\n`;
    } else {
      preparacao += `DROP TABLE IF EXISTS ${tabelaFormatada};\n`;
    }

    preparacao += `CREATE TABLE ${tabelaFormatada} (\n`;
    const defsColunas = colunas.map(col => {
      const colFormatada = formatarIdentifier(colunaDestino(col), bancoDados);
      const valoresColuna = linhas.map(l => l[col]);
      const tipo = inferirTipoColuna(valoresColuna, bancoDados);
      return `  ${colFormatada} ${tipo}`;
    });
    preparacao += defsColunas.join(',\n') + `\n);\n\n`;
  }

  if (modo === 'substituir') {
    if (bancoDados === 'sqlserver') {
      preparacao += `-- Desativa temporariamente as Foreign Keys para permitir a limpeza sem erros\n`;
      preparacao += `EXEC sp_msforeachtable 'ALTER TABLE ? NOCHECK CONSTRAINT ALL';\n\n`;
      preparacao += `DELETE FROM ${tabelaFormatada};\n`;
      preparacao += `-- Só reseta o contador de identidade se a tabela REALMENTE tiver uma coluna IDENTITY\n`;
      preparacao += `IF OBJECTPROPERTY(OBJECT_ID('${nomeTabela}'), 'TableHasIdentity') = 1\n`;
      preparacao += `  DBCC CHECKIDENT ('${nomeTabela}', RESEED, 0);\n\n`;
    } else {
      preparacao += `TRUNCATE TABLE ${tabelaFormatada};\n\n`;
    }
  }

  const podeTerIdentity = bancoDados === 'sqlserver' && modo !== 'criar';
  if (podeTerIdentity) {
    preparacao += `-- Habilita inserção explícita na coluna de ID, caso a tabela possua propriedade IDENTITY\n`;
    preparacao += `IF OBJECTPROPERTY(OBJECT_ID('${nomeTabela}'), 'TableHasIdentity') = 1\n`;
    preparacao += `  SET IDENTITY_INSERT ${tabelaFormatada} ON;\n\n`;
  }

  // ---------------- FECHAMENTO (IDENTITY_INSERT OFF / reabilita FKs) ----------------
  let fechamento = '';
  if (bancoDados === 'sqlserver') {
    if (podeTerIdentity) {
      fechamento += `-- Restaura o padrão de IDENTITY_INSERT, caso tenha sido habilitado acima\n`;
      fechamento += `IF OBJECTPROPERTY(OBJECT_ID('${nomeTabela}'), 'TableHasIdentity') = 1\n`;
      fechamento += `  SET IDENTITY_INSERT ${tabelaFormatada} OFF;\n`;
    }
    if (modo === 'substituir') {
      fechamento += `\n-- Reabilita a checagem de Foreign Keys\n`;
      fechamento += `EXEC sp_msforeachtable 'ALTER TABLE ? WITH CHECK CHECK CONSTRAINT ALL';\n`;
    }
  }

  // ---------------- UM COMANDO SQL POR LINHA DA PLANILHA ----------------
  const colunasFormatadas = colunas.map(col => formatarIdentifier(colunaDestino(col), bancoDados)).join(', ');

  const comandosPorLinha = linhas.map(linha => {
    const pkValor = linha[colunaPk];
    const pkValorFormatado = formatarValor(pkValor, bancoDados, tiposReais[colunaPkDestino]);
    const valoresInsert = colunas.map(col => formatarValor(linha[col], bancoDados, tiposReais[colunaDestino(col)])).join(', ');

    if (modo === 'inserir' && bancoDados === 'sqlserver') {
      const camposUpdate = colunas
        .filter(col => col !== colunaPk)
        .map(col => `${formatarIdentifier(colunaDestino(col), bancoDados)} = ${formatarValor(linha[col], bancoDados, tiposReais[colunaDestino(col)])}`)
        .join(', ');

      let cmd = `IF EXISTS (SELECT 1 FROM ${tabelaFormatada} WHERE ${formatarIdentifier(colunaPkDestino, bancoDados)} = ${pkValorFormatado})\n`;
      cmd += camposUpdate
        ? `  UPDATE ${tabelaFormatada} SET ${camposUpdate} WHERE ${formatarIdentifier(colunaPkDestino, bancoDados)} = ${pkValorFormatado};\n`
        : `  BEGIN /* nada além da chave para atualizar */ END\n`;
      cmd += `ELSE\n`;
      cmd += `  INSERT INTO ${tabelaFormatada} (${colunasFormatadas}) VALUES (${valoresInsert});`;
      return cmd;
    }

    if (modo === 'inserir' && bancoDados === 'postgres') {
      const colunaPkFormatada = formatarIdentifier(colunaPkDestino, bancoDados);
      const updates = colunas
        .filter(col => col !== colunaPk)
        .map(col => `${formatarIdentifier(colunaDestino(col), bancoDados)} = EXCLUDED.${formatarIdentifier(colunaDestino(col), bancoDados)}`)
        .join(', ');

      let cmd = `INSERT INTO ${tabelaFormatada} (${colunasFormatadas}) VALUES (${valoresInsert})\n`;
      cmd += updates
        ? `  ON CONFLICT (${colunaPkFormatada}) DO UPDATE SET ${updates};`
        : `  ON CONFLICT (${colunaPkFormatada}) DO NOTHING;`;
      return cmd;
    }

    if (modo === 'inserir') {
      const colunaPkFormatada = formatarIdentifier(colunaPkDestino, bancoDados);
      const updates = colunas
        .filter(col => col !== colunaPk)
        .map(col => `${formatarIdentifier(colunaDestino(col), bancoDados)} = VALUES(${formatarIdentifier(colunaDestino(col), bancoDados)})`)
        .join(', ');

      let cmd = `INSERT INTO ${tabelaFormatada} (${colunasFormatadas}) VALUES (${valoresInsert})\n`;
      cmd += updates
        ? `  ON DUPLICATE KEY UPDATE ${updates};`
        : `  ON DUPLICATE KEY UPDATE ${colunaPkFormatada} = ${colunaPkFormatada};`;
      return cmd;
    }

    // MODOS 'CRIAR' OU 'SUBSTITUIR'
    return `INSERT INTO ${tabelaFormatada} (${colunasFormatadas}) VALUES (${valoresInsert});`;
  });

  // ---------------- MONTAGEM DOS BLOCOS ----------------
  const blocos = [];
  const totalBlocos = Math.max(1, Math.ceil(comandosPorLinha.length / TAMANHO_BLOCO));

  for (let i = 0; i < totalBlocos; i++) {
    const inicio = i * TAMANHO_BLOCO;
    const fim = Math.min(inicio + TAMANHO_BLOCO, comandosPorLinha.length);
    const comandosDoBloco = comandosPorLinha.slice(inicio, fim).join('\n\n');

    let sqlDoBloco = '';
    if (i === 0) {
      sqlDoBloco += cabecalho + preparacao;
      sqlDoBloco += totalBlocos > 1
        ? `-- Processamento de Registros — bloco 1 de ${totalBlocos} (linhas ${inicio + 1} a ${fim})\n`
        : `-- Processamento de Registros (${comandosPorLinha.length} linhas)\n`;
    } else {
      sqlDoBloco += `-- Continuação — bloco ${i + 1} de ${totalBlocos} (linhas ${inicio + 1} a ${fim})\n`;
    }

    sqlDoBloco += comandosDoBloco;

    if (i === totalBlocos - 1 && fechamento) {
      sqlDoBloco += `\n\n${fechamento}`;
    }

    blocos.push({
      titulo: totalBlocos > 1
        ? `Bloco ${i + 1} de ${totalBlocos} (linhas ${inicio + 1}–${fim})`
        : `Script Completo (${comandosPorLinha.length} linhas)`,
      sql: sqlDoBloco
    });
  }

  const sqlCompleto = blocos.map(b => b.sql).join('\n\n');

  return { blocos, sqlCompleto, totalBlocos };
}

// -------------------------------------------------------------
// ENDPOINTS DA API
// -------------------------------------------------------------

// 1. ENDPOINT PARA LER AS COLUNAS DA PLANILHA
app.post('/api/ler-colunas', upload.single('planilha'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ erro: 'Nenhum arquivo de planilha foi fornecido.' });
    }

    // cellDates: true faz o SheetJS entregar células de data real como objetos Date,
    // em vez do número serial do Excel — só usado depois em gerar-sql.
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const primeiraAba = workbook.SheetNames[0];
    const aba = workbook.Sheets[primeiraAba];
    const dadosJson = XLSX.utils.sheet_to_json(aba, { defval: '' });

    if (!dadosJson || dadosJson.length === 0) {
      return res.status(400).json({ erro: 'A planilha informada não possui linhas com dados.' });
    }

    const colunas = Object.keys(dadosJson[0]);

    return res.json({
      sucesso: true,
      totalLinhas: dadosJson.length,
      colunas: colunas
    });

  } catch (err) {
    console.error('Erro ao ler colunas:', err);
    return res.status(500).json({ erro: err.message || 'Erro interno ao ler planilha.' });
  }
});

// 2. ENDPOINT PARA GERAR O SCRIPT SQL
app.post('/api/gerar-sql', upload.single('planilha'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ erro: 'Nenhum arquivo de planilha foi fornecido.' });
    }

    const { nomeTabela, bancoDados, modoMigracao, colunasSelecionadas, colunaChave, estruturaReal, mapeamentoColunas } = req.body;

    let colunasArray = [];
    if (colunasSelecionadas) {
      colunasArray = typeof colunasSelecionadas === 'string'
        ? JSON.parse(colunasSelecionadas)
        : colunasSelecionadas;
    }

    let mapaColunas = {};
    if (mapeamentoColunas) {
      mapaColunas = typeof mapeamentoColunas === 'string'
        ? JSON.parse(mapeamentoColunas)
        : mapeamentoColunas;
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const primeiraAba = workbook.SheetNames[0];
    const aba = workbook.Sheets[primeiraAba];

    const dadosJson = XLSX.utils.sheet_to_json(aba, { defval: '' });

    if (!dadosJson || dadosJson.length === 0) {
      return res.status(400).json({ erro: 'A planilha informada não possui linhas com dados.' });
    }

    const mapaTiposReais = parseEstruturaReal(estruturaReal);

    const resultado = gerarSqlScript(
      dadosJson,
      nomeTabela || 'minha_tabela',
      bancoDados || 'sqlserver',
      modoMigracao || 'inserir',
      colunasArray,
      colunaChave,
      mapaTiposReais,
      mapaColunas
    );

    return res.json({
      sucesso: true,
      totalLinhas: dadosJson.length,
      blocos: resultado.blocos,
      sqlCompleto: resultado.sqlCompleto
    });

  } catch (err) {
    console.error('Erro no processamento:', err);
    return res.status(500).json({ erro: err.message || 'Erro interno ao gerar script SQL.' });
  }
});

// Sobe o servidor na porta desejada; se estiver ocupada, tenta a próxima
// automaticamente (até um limite de tentativas), em vez de simplesmente falhar.
// Retorna uma Promise que resolve com a porta em que ele REALMENTE subiu.
function iniciarServidor(portaDesejada, tentativasRestantes = 10) {
  return new Promise((resolve, reject) => {
    function tentar(porta, restam) {
      const servidor = app.listen(porta, () => {
        console.log(`🚀 Servidor rodando em http://localhost:${porta}`);
        resolve({ servidor, porta });
      });

      servidor.on('error', (err) => {
        if (err.code === 'EADDRINUSE' && restam > 0) {
          console.warn(`⚠️  Porta ${porta} ocupada, tentando a porta ${porta + 1}...`);
          tentar(porta + 1, restam - 1);
        } else if (err.code === 'EADDRINUSE') {
          reject(new Error(`Não foi possível encontrar uma porta livre a partir de ${portaDesejada}.`));
        } else {
          reject(err);
        }
      });
    }

    tentar(portaDesejada, tentativasRestantes);
  });
}

module.exports = { app, iniciarServidor };

// Quando rodado diretamente (node server.js / npm start), sobe sozinho na porta padrão.
// Quando importado pelo main.js do Electron, quem decide quando/como subir é o main.js.
if (require.main === module) {
  const PORT = Number(process.env.PORT) || 3000;
  iniciarServidor(PORT).catch((err) => {
    console.error('❌ Não foi possível iniciar o servidor:', err.message);
  });
}