/**
 * Implantação do Web App no Apps Script
 * -------------------------------------
 * 1. Abra o Google Drive, crie uma nova planilha com as abas "Entradas" e "Saídas" (ou "Saidas").
 * 2. No menu Extensões > Apps Script, crie um novo projeto e cole este arquivo como Code.gs.
 * 3. Crie um novo arquivo HTML chamado "index" e cole o conteúdo do arquivo index.html.
 * 4. Em "Implantar" > "Implantações" > "Nova implantação", escolha "Aplicativo da web",
 *    defina "Executar como" como "Você" e "Quem tem acesso" como "Qualquer pessoa com o link".
 * 5. Clique em "Implantar" para gerar a URL pública do aplicativo.
 */

const APP_CONFIG = {
  timezone: Session.getScriptTimeZone() || 'America/Sao_Paulo',
  headerRow: 4,
  firstDataRow: 5,
  headers: ['Data', 'Descrição', 'Valor'],
  sheetNames: {
    entradas: 'Entradas',
    saidas: 'Saídas'
  },
  numberFormat: 'R$ #,##0.00',
  dateFormat: 'dd/MM/yyyy',
  sessionPrefix: 'sessao.',
  sessionDurationHours: 12,
  usuarios: {
    bruno: { senha: 'Cesar177*', perfil: 'completo' },
    alexandre: { senha: 'plus123', perfil: 'leitura' }
  }
};

const CURRENCY_FORMATTER = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Plus Podcast – Finanças');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function login(credentials) {
  const { usuario, senha } = credentials || {};
  if (!usuario || !senha) {
    throw new Error('Informe usuário e senha.');
  }
  const normalizedUser = usuario.toLowerCase().trim();
  const registro = APP_CONFIG.usuarios[normalizedUser];
  if (!registro || registro.senha !== senha) {
    Utilities.sleep(400);
    throw new Error('Usuário ou senha inválidos.');
  }
  const token = Utilities.getUuid();
  const expiresAt = Date.now() + APP_CONFIG.sessionDurationHours * 60 * 60 * 1000;
  const payload = { token, usuario: normalizedUser, perfil: registro.perfil, expiresAt };
  saveSessionPayload(normalizedUser, payload);
  return {
    sucesso: true,
    mensagem: 'Login realizado com sucesso.',
    token,
    usuario: normalizedUser,
    perfil: registro.perfil
  };
}

function validarSessao(token) {
  const session = readSession(token);
  if (!session) {
    return { valido: false };
  }
  return {
    valido: true,
    usuario: session.usuario,
    perfil: session.perfil,
    expiraEm: session.expiresAt
  };
}

function logout(token) {
  const session = readSession(token);
  if (session) {
    removeSession(session.usuario);
  }
  return { sucesso: true };
}

function getDashboardData(request) {
  const { token, mes } = request || {};
  const session = assertSession(token, { needsWrite: false });
  const range = mes ? getMonthRangeFromKey(mes) : getMonthRangeFromDate(new Date());
  const dados = getAllLancamentos();
  const totais = dados.reduce((acc, item) => {
    if (item.tipo === 'entrada') {
      acc.totalEntradas += item.valor;
    } else {
      acc.totalSaidas += item.valor;
    }
    if (item.data >= range.inicio && item.data <= range.fim) {
      if (item.tipo === 'entrada') {
        acc.entradasMes += item.valor;
      } else {
        acc.saidasMes += item.valor;
      }
    }
    return acc;
  }, { totalEntradas: 0, totalSaidas: 0, entradasMes: 0, saidasMes: 0 });
  const formatter = getCurrencyFormatter();
  return {
    usuario: session.usuario,
    perfil: session.perfil,
    mesSelecionado: formatMonthLabel(range.inicio),
    saldoAtual: totais.totalEntradas - totais.totalSaidas,
    totalEntradas: totais.entradasMes,
    totalSaidas: totais.saidasMes,
    formatado: {
      saldoAtual: formatter(totais.totalEntradas - totais.totalSaidas),
      totalEntradas: formatter(totais.entradasMes),
      totalSaidas: formatter(totais.saidasMes)
    }
  };
}

function getMesesDisponiveis(request) {
  const { token } = request || {};
  assertSession(token, { needsWrite: false });
  const dados = getAllLancamentos();
  const meses = {};
  dados.forEach(item => {
    const chave = getMonthKey(item.data);
    if (!meses[chave]) {
      meses[chave] = formatMonthLabel(item.data);
    }
  });
  const lista = Object.keys(meses)
    .sort((a, b) => b.localeCompare(a))
    .map(key => ({ chave: key, rotulo: meses[key] }));
  return lista;
}

function getLancamentosRecentes(request) {
  const { token, filtros = {} } = request || {};
  const session = assertSession(token, { needsWrite: false });
  const todos = getAllLancamentos();
  const limite = Math.max(1, Number(filtros.limite) || 50);
  let intervalo = buildDateInterval(filtros);
  let filtrados = applyDateFilter(todos, intervalo);
  let fallbackAplicado = false;
  if (filtrados.length === 0) {
    filtrados = todos
      .slice()
      .sort((a, b) => b.data - a.data)
      .slice(0, limite);
    fallbackAplicado = true;
  }
  const formatter = getCurrencyFormatter();
  const resposta = filtrados
    .sort((a, b) => b.data - a.data)
    .slice(0, limite)
    .map(item => ({
      tipo: item.tipo,
      descricao: item.descricao,
      dataISO: Utilities.formatDate(item.data, APP_CONFIG.timezone, 'yyyy-MM-dd'),
      dataFormatada: Utilities.formatDate(item.data, APP_CONFIG.timezone, APP_CONFIG.dateFormat),
      valor: item.valor,
      valorFormatado: formatter(item.valor),
      linha: item.linha,
      aba: item.aba
    }));
  return {
    usuario: session.usuario,
    perfil: session.perfil,
    filtrosAplicados: intervalo.descricao,
    fallbackAplicado,
    lancamentos: resposta
  };
}

function addLancamento(request) {
  const { token, lancamento } = request || {};
  const session = assertSession(token, { needsWrite: true });
  if (!lancamento) {
    throw new Error('Envie os dados do lançamento.');
  }
  const tipoNormalizado = normalizeTipo(lancamento.tipo);
  if (!tipoNormalizado) {
    throw new Error('Tipo de lançamento inválido.');
  }
  const data = parseISODate(lancamento.data);
  if (!data) {
    throw new Error('Data inválida.');
  }
  const descricao = String(lancamento.descricao || '').trim();
  if (!descricao) {
    throw new Error('Descrição é obrigatória.');
  }
  const valor = normalizeValor(lancamento.valor);
  if (valor === null) {
    throw new Error('Valor inválido.');
  }
  const sheet = getSheetByTipo(tipoNormalizado);
  const targetRow = Math.max(sheet.getLastRow() + 1, APP_CONFIG.firstDataRow);
  sheet.getRange(targetRow, 1, 1, 3).setValues([[data, descricao, valor]]);
  sheet.getRange(targetRow, 1).setNumberFormat(APP_CONFIG.dateFormat);
  sheet.getRange(targetRow, 3).setNumberFormat(APP_CONFIG.numberFormat);
  return {
    sucesso: true,
    mensagem: 'Lançamento salvo com sucesso.'
  };
}

function deleteLancamentos(request) {
  const { token, linhas } = request || {};
  const session = assertSession(token, { needsWrite: true });
  if (session.usuario !== 'bruno') {
    throw new Error('Somente Bruno pode excluir lançamentos.');
  }
  if (!Array.isArray(linhas) || linhas.length === 0) {
    throw new Error('Selecione ao menos um lançamento para excluir.');
  }
  const agrupados = linhas.reduce((acc, item) => {
    const tipo = normalizeTipo(item.tipo);
    if (!tipo) {
      return acc;
    }
    if (!acc[tipo]) {
      acc[tipo] = [];
    }
    const linha = Number(item.linha);
    if (linha >= APP_CONFIG.firstDataRow) {
      acc[tipo].push(linha);
    }
    return acc;
  }, {});
  if (Object.keys(agrupados).length === 0) {
    throw new Error('Não encontramos linhas válidas para exclusão.');
  }
  const planilha = getSpreadsheet();
  Object.keys(agrupados).forEach(tipo => {
    const sheet = getSheetByTipo(tipo, planilha);
    const linhasParaExcluir = agrupados[tipo].sort((a, b) => b - a);
    linhasParaExcluir.forEach(linha => {
      if (linha >= APP_CONFIG.firstDataRow && linha <= sheet.getLastRow()) {
        sheet.deleteRow(linha);
      }
    });
  });
  return {
    sucesso: true,
    mensagem: 'Lançamentos excluídos com sucesso.'
  };
}

// --- Sessões ---

function saveSessionPayload(usuario, payload) {
  PropertiesService.getScriptProperties().setProperty(APP_CONFIG.sessionPrefix + usuario, JSON.stringify(payload));
}

function readSession(token) {
  if (!token) {
    return null;
  }
  const props = PropertiesService.getScriptProperties();
  const entries = props.getProperties();
  const chave = Object.keys(entries).find(key => {
    if (!key.startsWith(APP_CONFIG.sessionPrefix)) {
      return false;
    }
    try {
      const dados = JSON.parse(entries[key]);
      return dados.token === token;
    } catch (e) {
      return false;
    }
  });
  if (!chave) {
    return null;
  }
  try {
    const dados = JSON.parse(entries[chave]);
    if (dados.expiresAt && Date.now() > dados.expiresAt) {
      PropertiesService.getScriptProperties().deleteProperty(chave);
      return null;
    }
    return dados;
  } catch (error) {
    PropertiesService.getScriptProperties().deleteProperty(chave);
    return null;
  }
}

function removeSession(usuario) {
  PropertiesService.getScriptProperties().deleteProperty(APP_CONFIG.sessionPrefix + usuario);
}

function assertSession(token, options) {
  const session = readSession(token);
  if (!session) {
    throw new Error('Sessão expirada. Faça login novamente.');
  }
  if (options && options.needsWrite && APP_CONFIG.usuarios[session.usuario].perfil === 'leitura') {
    throw new Error('Seu perfil permite apenas leitura.');
  }
  return session;
}

// --- Manipulação de planilhas ---

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByTipo(tipo, spreadsheet) {
  const ss = spreadsheet || getSpreadsheet();
  const nomePreferencial = tipo === 'entrada' ? APP_CONFIG.sheetNames.entradas : APP_CONFIG.sheetNames.saidas;
  const nomesAlternativos = tipo === 'entrada'
    ? [APP_CONFIG.sheetNames.entradas, 'Entrada']
    : [APP_CONFIG.sheetNames.saidas, 'Saidas', 'Saída'];
  let sheet = nomesAlternativos.reduce((acc, nome) => acc || ss.getSheetByName(nome), null);
  if (!sheet) {
    sheet = ss.insertSheet(nomePreferencial);
  }
  prepareSheetStructure(sheet);
  return sheet;
}

function prepareSheetStructure(sheet) {
  const headerRow = APP_CONFIG.headerRow;
  const headersRange = sheet.getRange(headerRow, 1, 1, APP_CONFIG.headers.length);
  const currentValues = headersRange.getValues()[0];
  let needsUpdate = false;
  APP_CONFIG.headers.forEach((header, index) => {
    if (currentValues[index] !== header) {
      needsUpdate = true;
    }
  });
  if (needsUpdate) {
    headersRange.setValues([APP_CONFIG.headers]);
    headersRange.setFontWeight('bold');
    sheet.getRange(headerRow, 1, 1, APP_CONFIG.headers.length).setBackground('#f5f5f5');
  }
  sheet.getRange(APP_CONFIG.firstDataRow, 1, Math.max(1, sheet.getMaxRows() - APP_CONFIG.firstDataRow + 1), 1)
    .setNumberFormat(APP_CONFIG.dateFormat);
  sheet.getRange(APP_CONFIG.firstDataRow, 3, Math.max(1, sheet.getMaxRows() - APP_CONFIG.firstDataRow + 1), 1)
    .setNumberFormat(APP_CONFIG.numberFormat);
}

function getAllLancamentos() {
  const planilha = getSpreadsheet();
  const entradas = readSheetData(getSheetByTipo('entrada', planilha), 'entrada');
  const saidas = readSheetData(getSheetByTipo('saida', planilha), 'saida');
  return entradas.concat(saidas);
}

function readSheetData(sheet, tipo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG.firstDataRow) {
    return [];
  }
  const range = sheet.getRange(APP_CONFIG.firstDataRow, 1, lastRow - APP_CONFIG.firstDataRow + 1, APP_CONFIG.headers.length);
  const values = range.getValues();
  const result = [];
  values.forEach((row, index) => {
    const data = normalizeDateValue(row[0]);
    const descricao = String(row[1] || '').trim();
    const valor = normalizeValor(row[2]);
    if (!data || !descricao || valor === null) {
      return;
    }
    result.push({
      tipo,
      data,
      descricao,
      valor,
      linha: APP_CONFIG.firstDataRow + index,
      aba: sheet.getName()
    });
  });
  return result;
}

// --- Filtros e datas ---

function buildDateInterval(filtros) {
  filtros = filtros || {};
  if (filtros.mes) {
    const intervaloMes = getMonthRangeFromKey(filtros.mes);
    return {
      inicio: intervaloMes.inicio,
      fim: intervaloMes.fim,
      descricao: `Mês ${intervaloMes.rotulo}`
    };
  }
  let inicio = filtros.dataInicial ? parseISODate(filtros.dataInicial) : null;
  let fim = filtros.dataFinal ? parseISODate(filtros.dataFinal) : null;
  if (!inicio && !fim) {
    fim = endOfDay(new Date());
    inicio = startOfDay(new Date(fim.getTime() - 29 * 24 * 60 * 60 * 1000));
    return { inicio, fim, descricao: 'Últimos 30 dias' };
  }
  if (inicio && !fim) {
    fim = endOfDay(new Date());
  }
  if (fim && !inicio) {
    inicio = startOfDay(new Date(fim.getTime() - 29 * 24 * 60 * 60 * 1000));
  }
  return {
    inicio: startOfDay(inicio),
    fim: endOfDay(fim),
    descricao: `${Utilities.formatDate(inicio, APP_CONFIG.timezone, APP_CONFIG.dateFormat)} a ${Utilities.formatDate(fim, APP_CONFIG.timezone, APP_CONFIG.dateFormat)}`
  };
}

function applyDateFilter(registros, intervalo) {
  return registros.filter(item => item.data >= intervalo.inicio && item.data <= intervalo.fim);
}

function getMonthRangeFromKey(key) {
  if (!key || !/^\d{4}-\d{2}$/.test(key)) {
    return getMonthRangeFromDate(new Date());
  }
  const [ano, mes] = key.split('-').map(Number);
  const inicio = startOfDay(new Date(ano, mes - 1, 1));
  return getMonthRangeFromDate(inicio);
}

function getMonthRangeFromDate(date) {
  const inicio = startOfDay(new Date(date.getFullYear(), date.getMonth(), 1));
  const fim = endOfDay(new Date(date.getFullYear(), date.getMonth() + 1, 0));
  return { inicio, fim, rotulo: formatMonthLabel(inicio) };
}

function startOfDay(date) {
  const nova = new Date(date);
  nova.setHours(0, 0, 0, 0);
  return nova;
}

function endOfDay(date) {
  const nova = new Date(date);
  nova.setHours(23, 59, 59, 999);
  return nova;
}

const MONTH_LABELS = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

function formatMonthLabel(date) {
  const mes = MONTH_LABELS[date.getMonth()] || '';
  return `${mes.charAt(0).toUpperCase()}${mes.slice(1)} de ${date.getFullYear()}`;
}

function getMonthKey(date) {
  return Utilities.formatDate(date, APP_CONFIG.timezone, 'yyyy-MM');
}

// --- Conversões ---

function normalizeTipo(tipo) {
  if (!tipo) return null;
  const valor = tipo.toString().toLowerCase();
  if (valor.startsWith('e')) return 'entrada';
  if (valor.startsWith('s')) return 'saida';
  return null;
}

function parseISODate(valor) {
  if (valor instanceof Date) {
    return valor;
  }
  if (typeof valor !== 'string') {
    return null;
  }
  const partes = valor.split('-').map(Number);
  if (partes.length !== 3) {
    return null;
  }
  const data = new Date(partes[0], partes[1] - 1, partes[2]);
  return isNaN(data.getTime()) ? null : startOfDay(data);
}

function normalizeDateValue(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return startOfDay(valor);
  }
  if (typeof valor === 'string' && valor.includes('/')) {
    const [dia, mes, ano] = valor.split('/').map(Number);
    const data = new Date(ano, mes - 1, dia);
    return isNaN(data.getTime()) ? null : startOfDay(data);
  }
  if (typeof valor === 'string' && valor.includes('-')) {
    return parseISODate(valor);
  }
  return null;
}

function normalizeValor(valor) {
  if (typeof valor === 'number' && !isNaN(valor)) {
    return Number(valor);
  }
  if (typeof valor === 'string') {
    const limpo = valor.replace(/[^0-9,-]/g, '').replace(',', '.');
    const numero = parseFloat(limpo);
    return isNaN(numero) ? null : numero;
  }
  return null;
}

function getCurrencyFormatter() {
  return valor => CURRENCY_FORMATTER.format(typeof valor === 'number' ? valor : 0);
}
