/**
 * ============================================================
 * SISTEMA DE MESADA FAMILIAR - BACKEND (Google Apps Script)
 * ============================================================
 * INSTRUÇÕES DE INSTALAÇÃO:
 * 1. Crie uma planilha nova no Google Sheets.
 * 2. Vá em Extensões > Apps Script.
 * 3. Cole TODO este código no editor (substitua o conteúdo padrão).
 * 4. Execute a função `setupPlanilha` uma vez (Executar > setupPlanilha).
 *    - Na primeira execução, autorize as permissões solicitadas.
 * 5. Clique em "Implantar" > "Nova implantação".
 *    - Tipo: "Aplicativo da Web"
 *    - Executar como: "Eu"
 *    - Quem pode acessar: "Qualquer pessoa"
 * 6. Copie a URL gerada (terminada em /exec) e cole no arquivo
 *    index.html, na constante API_URL.
 * ============================================================
 */

// ============================================================
// CONFIGURAÇÃO GERAL
// ============================================================

const SHEET_USUARIOS = 'Usuarios';
const SHEET_CONFIG = 'Configuracoes';
const SHEET_LANCAMENTOS = 'Lancamentos';
const SHEET_HISTORICO = 'HistoricoRecorrencias';

// ============================================================
// SETUP INICIAL - Cria as abas e dados padrão
// ============================================================
function setupPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Aba Usuarios ----
  let shUsers = ss.getSheetByName(SHEET_USUARIOS);
  if (!shUsers) shUsers = ss.insertSheet(SHEET_USUARIOS);
  shUsers.clear();
  shUsers.appendRow(['id', 'usuario', 'senha', 'perfil', 'criadoEm']);
  shUsers.appendRow([1, 'admin', 'admin', 'Admin', new Date().toISOString()]);
  shUsers.appendRow([2, 'Mariana', 'Mari', 'Restrito', new Date().toISOString()]);

  // ---- Aba Configuracoes ----
  let shConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!shConfig) shConfig = ss.insertSheet(SHEET_CONFIG);
  shConfig.clear();
  shConfig.appendRow(['chave', 'valor']);
  shConfig.appendRow(['valorMesada', '50']);
  shConfig.appendRow(['tipoRecorrencia', 'semanal']); // 'semanal' ou 'mensal'
  shConfig.appendRow(['diaSemana', '6']); // 0=Domingo ... 6=Sábado (Apps Script: 0-6)
  shConfig.appendRow(['diaMes', '1']); // 1 a 28
  shConfig.appendRow(['taxaJuros', '1']); // em %
  shConfig.appendRow(['saldoAtual', '0']);
  shConfig.appendRow(['ultimoFechamento', '']); // ISO date do último ciclo processado

  // ---- Aba Lancamentos ----
  let shLanc = ss.getSheetByName(SHEET_LANCAMENTOS);
  if (!shLanc) shLanc = ss.insertSheet(SHEET_LANCAMENTOS);
  shLanc.clear();
  shLanc.appendRow(['id', 'data', 'tipo', 'descricao', 'valor', 'saldoApos', 'usuario']);

  // ---- Aba HistoricoRecorrencias ----
  let shHist = ss.getSheetByName(SHEET_HISTORICO);
  if (!shHist) shHist = ss.insertSheet(SHEET_HISTORICO);
  shHist.clear();
  shHist.appendRow(['id', 'dataProcessamento', 'tipo', 'valorDeposito', 'valorJuros', 'saldoAntes', 'saldoDepois']);

  // Remove a aba padrão "Página1"/"Sheet1" se existir e estiver vazia
  const def1 = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (def1 && ss.getSheets().length > 1) {
    ss.deleteSheet(def1);
  }

  Logger.log('Setup concluído com sucesso!');
}

// ============================================================
// UTILITÁRIOS DE PLANILHA
// ============================================================
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheetName) {
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) val = val.toISOString();
        obj[h] = val;
      });
      return obj;
    });
}

function getConfigMap() {
  const arr = sheetToObjects(SHEET_CONFIG);
  const map = {};
  arr.forEach(item => { map[item.chave] = item.valor; });
  return map;
}

function setConfigValue(chave, valor) {
  const sh = getSheet(SHEET_CONFIG);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === chave) {
      sh.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  sh.appendRow([chave, valor]);
}

function getNextId(sheetName) {
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0], 10);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

// Formata data para "YYYY-MM-DD" baseado no fuso de São Paulo
function formatDateISO(date) {
  return Utilities.formatDate(date, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function todayDate() {
  const now = new Date();
  // Normaliza para meia-noite no fuso de SP, representado como Date UTC
  const isoStr = formatDateISO(now);
  return new Date(isoStr + 'T00:00:00');
}

// ============================================================
// LANÇAMENTOS - Criar e atualizar saldo
// ============================================================
function addLancamento(tipo, descricao, valor, usuario, dataCustom) {
  const config = getConfigMap();
  let saldoAtual = parseFloat(config.saldoAtual) || 0;

  let novoSaldo;
  if (tipo === 'Saque') {
    novoSaldo = saldoAtual - Math.abs(valor);
  } else {
    // Deposito ou Juros
    novoSaldo = saldoAtual + Math.abs(valor);
  }

  const sh = getSheet(SHEET_LANCAMENTOS);
  const id = getNextId(SHEET_LANCAMENTOS);
  const dataLanc = dataCustom ? dataCustom : new Date();

  const valorRegistrado = (tipo === 'Saque') ? -Math.abs(valor) : Math.abs(valor);

  sh.appendRow([
    id,
    dataLanc.toISOString(),
    tipo,
    descricao,
    valorRegistrado,
    novoSaldo,
    usuario || ''
  ]);

  setConfigValue('saldoAtual', novoSaldo);

  return { id, novoSaldo };
}

// ============================================================
// PROCESSAMENTO DE RECORRÊNCIAS (Mesada + Juros)
// ============================================================
/**
 * Verifica se ciclos de mesada recorrente foram "perdidos" desde o
 * último fechamento e os processa, aplicando depósito + juros
 * sobre o saldo existente no momento de CADA fechamento.
 */
function processarRecorrencias() {
  const config = getConfigMap();
  const valorMesada = parseFloat(config.valorMesada) || 0;
  const taxaJuros = parseFloat(config.taxaJuros) || 0;
  const tipoRecorrencia = config.tipoRecorrencia || 'semanal';
  const diaSemana = parseInt(config.diaSemana, 10); // 0-6 (Dom-Sab)
  const diaMes = parseInt(config.diaMes, 10); // 1-28

  let ultimoFechamento = config.ultimoFechamento
    ? new Date(config.ultimoFechamento + 'T00:00:00')
    : null;

  const hoje = todayDate();

  // Se nunca houve fechamento, define o "marco zero" como hoje
  // (não processa retroativamente desde o início dos tempos)
  if (!ultimoFechamento) {
    setConfigValue('ultimoFechamento', formatDateISO(hoje));
    return { processados: 0 };
  }

  let processados = 0;
  let cursor = new Date(ultimoFechamento.getTime());
  let safetyCounter = 0; // evita loop infinito

  while (safetyCounter < 1000) {
    safetyCounter++;
    const proximaData = calcularProximaDataFechamento(cursor, tipoRecorrencia, diaSemana, diaMes);

    if (proximaData.getTime() > hoje.getTime()) {
      break; // ainda não chegou a próxima data de fechamento
    }

    // Processa o fechamento na "proximaData"
    processarUmFechamento(proximaData, valorMesada, taxaJuros);
    processados++;
    cursor = proximaData;
    setConfigValue('ultimoFechamento', formatDateISO(proximaData));
  }

  return { processados };
}

/**
 * Calcula a próxima data de fechamento APÓS a data 'partir'.
 */
function calcularProximaDataFechamento(partir, tipoRecorrencia, diaSemana, diaMes) {
  const d = new Date(partir.getTime());

  if (tipoRecorrencia === 'semanal') {
    // Avança dia a dia até encontrar o próximo diaSemana após 'partir'
    do {
      d.setDate(d.getDate() + 1);
    } while (d.getDay() !== diaSemana);
    return d;
  } else {
    // Mensal: próximo mês no dia 'diaMes'
    let nextMonth = d.getMonth() + 1;
    let nextYear = d.getFullYear();
    if (nextMonth > 11) { nextMonth = 0; nextYear++; }
    const diaAjustado = Math.min(diaMes, diasNoMes(nextYear, nextMonth));
    return new Date(nextYear, nextMonth, diaAjustado);
  }
}

function diasNoMes(ano, mesIndex) {
  return new Date(ano, mesIndex + 1, 0).getDate();
}

/**
 * Aplica juros sobre o saldo ATUAL (antes do depósito) e depois
 * lança o depósito da mesada, na data informada.
 */
function processarUmFechamento(dataFechamento, valorMesada, taxaJuros) {
  const config = getConfigMap();
  const saldoAntes = parseFloat(config.saldoAtual) || 0;

  // 1. Calcula e aplica juros sobre o saldo existente
  const valorJuros = Math.round((saldoAntes * (taxaJuros / 100)) * 100) / 100;
  if (valorJuros > 0) {
    addLancamento('Juros', 'Juros sobre saldo (ciclo recorrente)', valorJuros, 'sistema', dataFechamento);
  }

  // 2. Aplica o depósito recorrente da mesada
  if (valorMesada > 0) {
    addLancamento('Deposito', 'Mesada recorrente', valorMesada, 'sistema', dataFechamento);
  }

  const configDepois = getConfigMap();
  const saldoDepois = parseFloat(configDepois.saldoAtual) || 0;

  // 3. Registra no histórico
  const sh = getSheet(SHEET_HISTORICO);
  const id = getNextId(SHEET_HISTORICO);
  sh.appendRow([
    id,
    dataFechamento.toISOString(),
    'recorrencia',
    valorMesada,
    valorJuros,
    saldoAntes,
    saldoDepois
  ]);
}

// ============================================================
// ROTEAMENTO DA API (doGet / doPost)
// ============================================================
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let action = '';
  let params = {};

  try {
    if (e.parameter && e.parameter.action) {
      action = e.parameter.action;
      params = e.parameter;
    }
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      action = body.action || action;
      params = body;
    }

    let result;
    switch (action) {
      case 'login':
        result = apiLogin(params);
        break;
      case 'getDados':
        result = apiGetDados(params);
        break;
      case 'registrarSaque':
        result = apiRegistrarSaque(params);
        break;
      case 'aporteIndividual':
        result = apiAporteIndividual(params);
        break;
      case 'salvarConfiguracoes':
        result = apiSalvarConfiguracoes(params);
        break;
      case 'listarUsuarios':
        result = apiListarUsuarios(params);
        break;
      case 'salvarUsuario':
        result = apiSalvarUsuario(params);
        break;
      case 'excluirUsuario':
        result = apiExcluirUsuario(params);
        break;
      case 'exportarDados':
        result = apiExportarDados(params);
        break;
      case 'importarDados':
        result = apiImportarDados(params);
        break;
      default:
        result = { sucesso: false, erro: 'Ação não reconhecida: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ sucesso: false, erro: err.message, stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ENDPOINTS DA API
// ============================================================

// ---- LOGIN ----
function apiLogin(params) {
  const usuarios = sheetToObjects(SHEET_USUARIOS);
  const user = usuarios.find(
    u => String(u.usuario).toLowerCase() === String(params.usuario).toLowerCase()
      && String(u.senha) === String(params.senha)
  );

  if (!user) {
    return { sucesso: false, erro: 'Usuário ou senha inválidos.' };
  }

  return {
    sucesso: true,
    usuario: {
      id: user.id,
      usuario: user.usuario,
      perfil: user.perfil
    }
  };
}

// ---- DADOS GERAIS (saldo, lançamentos, config) ----
function apiGetDados(params) {
  // Processa recorrências pendentes antes de retornar os dados
  processarRecorrencias();

  const config = getConfigMap();
  const lancamentos = sheetToObjects(SHEET_LANCAMENTOS)
    .sort((a, b) => new Date(b.data) - new Date(a.data));

  let usuarios = [];
  if (params.incluirUsuarios === 'true' || params.incluirUsuarios === true) {
    usuarios = sheetToObjects(SHEET_USUARIOS).map(u => ({
      id: u.id, usuario: u.usuario, perfil: u.perfil, senha: u.senha
    }));
  }

  return {
    sucesso: true,
    saldoAtual: parseFloat(config.saldoAtual) || 0,
    configuracoes: {
      valorMesada: parseFloat(config.valorMesada) || 0,
      tipoRecorrencia: config.tipoRecorrencia || 'semanal',
      diaSemana: parseInt(config.diaSemana, 10),
      diaMes: parseInt(config.diaMes, 10),
      taxaJuros: parseFloat(config.taxaJuros) || 0,
      ultimoFechamento: config.ultimoFechamento || ''
    },
    lancamentos: lancamentos,
    usuarios: usuarios
  };
}

// ---- SAQUE ----
function apiRegistrarSaque(params) {
  const valor = parseFloat(params.valor);
  const comentario = params.comentario;
  const usuario = params.usuario || 'desconhecido';

  if (isNaN(valor) || valor <= 0) {
    return { sucesso: false, erro: 'Valor de saque inválido.' };
  }
  if (!comentario || comentario.trim() === '') {
    return { sucesso: false, erro: 'Comentário/justificativa é obrigatório.' };
  }

  const config = getConfigMap();
  const saldoAtual = parseFloat(config.saldoAtual) || 0;

  if (valor > saldoAtual) {
    return { sucesso: false, erro: 'Saldo insuficiente para este saque.' };
  }

  const resultado = addLancamento('Saque', comentario, valor, usuario);

  return { sucesso: true, novoSaldo: resultado.novoSaldo };
}

// ---- APORTE INDIVIDUAL (Admin) ----
function apiAporteIndividual(params) {
  const valor = parseFloat(params.valor);
  const descricao = params.descricao || 'Aporte individual';
  const usuario = params.usuario || 'admin';

  if (isNaN(valor) || valor <= 0) {
    return { sucesso: false, erro: 'Valor de aporte inválido.' };
  }

  const resultado = addLancamento('Deposito', descricao, valor, usuario);
  return { sucesso: true, novoSaldo: resultado.novoSaldo };
}

// ---- SALVAR CONFIGURAÇÕES (Admin) ----
function apiSalvarConfiguracoes(params) {
  if (params.valorMesada !== undefined) setConfigValue('valorMesada', parseFloat(params.valorMesada));
  if (params.tipoRecorrencia !== undefined) setConfigValue('tipoRecorrencia', params.tipoRecorrencia);
  if (params.diaSemana !== undefined) setConfigValue('diaSemana', parseInt(params.diaSemana, 10));
  if (params.diaMes !== undefined) setConfigValue('diaMes', parseInt(params.diaMes, 10));
  if (params.taxaJuros !== undefined) setConfigValue('taxaJuros', parseFloat(params.taxaJuros));

  return { sucesso: true };
}

// ---- USUÁRIOS (Admin) ----
function apiListarUsuarios() {
  const usuarios = sheetToObjects(SHEET_USUARIOS).map(u => ({
    id: u.id, usuario: u.usuario, perfil: u.perfil, senha: u.senha
  }));
  return { sucesso: true, usuarios: usuarios };
}

function apiSalvarUsuario(params) {
  const sh = getSheet(SHEET_USUARIOS);
  const data = sh.getDataRange().getValues();

  const id = params.id ? parseInt(params.id, 10) : null;
  const usuario = params.usuario;
  const senha = params.senha;
  const perfil = params.perfil;

  if (!usuario || !senha || !perfil) {
    return { sucesso: false, erro: 'Campos obrigatórios: usuario, senha, perfil.' };
  }

  if (id) {
    // Edição
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) === id) {
        sh.getRange(i + 1, 2).setValue(usuario);
        sh.getRange(i + 1, 3).setValue(senha);
        sh.getRange(i + 1, 4).setValue(perfil);
        return { sucesso: true, id: id };
      }
    }
    return { sucesso: false, erro: 'Usuário não encontrado.' };
  } else {
    // Criação - verifica duplicidade
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === usuario.toLowerCase()) {
        return { sucesso: false, erro: 'Já existe um usuário com este nome.' };
      }
    }
    const newId = getNextId(SHEET_USUARIOS);
    sh.appendRow([newId, usuario, senha, perfil, new Date().toISOString()]);
    return { sucesso: true, id: newId };
  }
}

function apiExcluirUsuario(params) {
  const id = parseInt(params.id, 10);
  const sh = getSheet(SHEET_USUARIOS);
  const data = sh.getDataRange().getValues();

  // Impede excluir o último admin
  const admins = data.slice(1).filter(row => row[3] === 'Admin');

  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0], 10) === id) {
      if (data[i][3] === 'Admin' && admins.length <= 1) {
        return { sucesso: false, erro: 'Não é possível excluir o único administrador.' };
      }
      sh.deleteRow(i + 1);
      return { sucesso: true };
    }
  }
  return { sucesso: false, erro: 'Usuário não encontrado.' };
}

// ---- EXPORTAR DADOS (Backup) ----
function apiExportarDados() {
  const dados = {
    versao: 1,
    exportadoEm: new Date().toISOString(),
    usuarios: sheetToObjects(SHEET_USUARIOS),
    configuracoes: sheetToObjects(SHEET_CONFIG),
    lancamentos: sheetToObjects(SHEET_LANCAMENTOS),
    historicoRecorrencias: sheetToObjects(SHEET_HISTORICO)
  };
  return { sucesso: true, dados: dados };
}

// ---- IMPORTAR DADOS (Backup) ----
function apiImportarDados(params) {
  let dados;
  try {
    dados = typeof params.dados === 'string' ? JSON.parse(params.dados) : params.dados;
  } catch (err) {
    return { sucesso: false, erro: 'JSON inválido: ' + err.message };
  }

  if (!dados || dados.versao === undefined) {
    return { sucesso: false, erro: 'Arquivo de backup inválido (campo "versao" ausente).' };
  }

  // ---- Usuarios ----
  if (Array.isArray(dados.usuarios)) {
    const sh = getSheet(SHEET_USUARIOS);
    sh.clear();
    sh.appendRow(['id', 'usuario', 'senha', 'perfil', 'criadoEm']);
    dados.usuarios.forEach(u => {
      sh.appendRow([u.id, u.usuario, u.senha, u.perfil, u.criadoEm || '']);
    });
  }

  // ---- Configuracoes ----
  if (Array.isArray(dados.configuracoes)) {
    const sh = getSheet(SHEET_CONFIG);
    sh.clear();
    sh.appendRow(['chave', 'valor']);
    dados.configuracoes.forEach(c => {
      sh.appendRow([c.chave, c.valor]);
    });
  }

  // ---- Lancamentos ----
  if (Array.isArray(dados.lancamentos)) {
    const sh = getSheet(SHEET_LANCAMENTOS);
    sh.clear();
    sh.appendRow(['id', 'data', 'tipo', 'descricao', 'valor', 'saldoApos', 'usuario']);
    dados.lancamentos.forEach(l => {
      sh.appendRow([l.id, l.data, l.tipo, l.descricao, l.valor, l.saldoApos, l.usuario || '']);
    });
  }

  // ---- HistoricoRecorrencias ----
  if (Array.isArray(dados.historicoRecorrencias)) {
    const sh = getSheet(SHEET_HISTORICO);
    sh.clear();
    sh.appendRow(['id', 'dataProcessamento', 'tipo', 'valorDeposito', 'valorJuros', 'saldoAntes', 'saldoDepois']);
    dados.historicoRecorrencias.forEach(h => {
      sh.appendRow([h.id, h.dataProcessamento, h.tipo, h.valorDeposito, h.valorJuros, h.saldoAntes, h.saldoDepois]);
    });
  }

  return { sucesso: true };
}
