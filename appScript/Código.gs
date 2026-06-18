// ============================================================
// VIEIRA MÓVEIS SOB MEDIDA — Sistema de Gestão
// Google Apps Script | Arquivo principal: Code.gs
// ============================================================

// ─── CONFIGURAÇÕES GLOBAIS ───────────────────────────────────
var CONFIG = { // a variavel config inicia uma lista de informações imutaveis, pois nao é um let
  EMAIL_PROPRIETARIO:"EMAIL@gmail.com",   // ← ALTERE AQUI PARA O EMAIL CORRETO
  DIAS_ALERTA_PRAZO:3, // Define um numero de dias que é contado para disparar um alarm na planilha
  NOME_EMPRESA:"Vieira Móveis Sob Medida" // Autoexplicativo: nome da empresa associada
};

// Nomes das abas (não altere)
var ABAS = { // a variavel ABAS cria uma lista de opções linkadas no google sheets
  PEDIDOS:"Pedidos", // onde define "pedidos" como pedidos
  CLIENTES:"Clientes", // clientes como clientes
  ESTOQUE:"Estoque", // estoque como estoque
  MOVIMENTOS:"Mov_Estoque", // movimentações como mov_estoque
  DASHBOARD:"Dashboard"//dashboard como dashboard
};


// ─── INICIALIZAÇÃO ───────────────────────────────────────────

/** Cria o menu personalizado ao abrir a planilha */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Vieira Móveis")
    .addItem("Novo Pedido","abrirFormPedido")
    .addItem("Novo Cliente","abrirFormCliente")
    .addSeparator()
    .addItem("Entrada de Material","abrirFormEntradaEstoque")
    .addItem("Saída de Material","abrirFormSaidaEstoque")
    .addSeparator()
    .addItem("Buscar Cliente","abrirBuscaCliente")
    .addSeparator()
    .addItem("Verificar Alertas Agora","verificarTodosAlertas")
    .addItem("Atualizar Dashboard","atualizarDashboard")
    .addSeparator()
    .addItem("Configuração Inicial","configuracaoInicial")
    .addToUi();
}

/** Roda automaticamente ao editar qualquer célula */
function onEdit(e) {
  var aba = e.range.getSheet().getName();
  if (aba === ABAS.PEDIDOS || aba === ABAS.ESTOQUE) {
    atualizarDashboard();
  }
}

/** Cria todas as abas e cabeçalhos na primeira execução */
function configuracaoInicial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var resposta = ui.alert(
    "Configuração Inicial",
    "Isso irá criar todas as abas e estrutura do sistema.\n" +
    "Dados existentes serão preservados. Deseja continuar?",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  _criarAba(ss, ABAS.PEDIDOS, [
    "Nº Pedido", "Código Cliente", "Nome Cliente",
    "Descrição do Projeto", "Data Entrada", "Prazo Entrega",
    "Data Montagem", "Local Instalação", "Valor Total (R$)",
    "Valor Recebido (R$)", "Saldo (R$)", "Responsável",
    "Status", "Observações", "ID Evento Calendar"
  ]);

  _criarAba(ss, ABAS.CLIENTES, [
    "Código", "Nome Completo", "Telefone/WhatsApp", "E-mail",
    "Endereço", "Cidade", "Data Cadastro",
    "Total de Pedidos", "Observações"
  ]);

  _criarAba(ss, ABAS.ESTOQUE, [
    "Código", "Descrição", "Unidade", "Qtd. Mínima",
    "Qtd. Atual", "Último Fornecedor", "Preço Unitário (R$)",
    "Última Entrada", "Última Saída"
  ]);

  _criarAba(ss, ABAS.MOVIMENTOS, [
    "Data/Hora", "Tipo", "Código Material", "Descrição Material",
    "Quantidade", "Pedido Vinculado", "Responsável", "Observações"
  ]);

  _criarAba(ss, ABAS.DASHBOARD, []);
  atualizarDashboard();
  _configurarGatilhos();

  ui.alert("✅ Configuração concluída!\n\nO sistema está pronto para uso.\n" +
           "Acesse o menu '🪵 Vieira Móveis' para começar.");
}

function _criarAba(ss, nome, cabecalhos) {
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
  }
  if (cabecalhos.length > 0 && aba.getLastRow() === 0) {
    var range = aba.getRange(1, 1, 1, cabecalhos.length);
    range.setValues([cabecalhos]);
    range.setFontWeight("bold");
    range.setBackground("#2d5016");
    range.setFontColor("#ffffff");
    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, cabecalhos.length);
  }
}

function _configurarGatilhos() {
  // Remove gatilhos antigos para evitar duplicatas
  var gatilhos = ScriptApp.getProjectTriggers();
  gatilhos.forEach(function(g) { ScriptApp.deleteTrigger(g); });

  // Alerta diário às 07:00
  ScriptApp.newTrigger("verificarTodosAlertas")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  Logger.log("Gatilhos configurados com sucesso.");
}


// ─── MÓDULO 1: PEDIDOS ───────────────────────────────────────

function abrirFormPedido() {
  var html = HtmlService.createHtmlOutputFromFile("FormPedido")
    .setTitle("Novo Pedido")
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Chamado pelo formulário HTML para salvar o pedido */
function salvarPedido(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABAS.PEDIDOS);

  var numPedido = _proximoNumeroPedido(aba);
  var dataHoje  = new Date();
  var saldo     = (parseFloat(dados.valorTotal) || 0) -
                  (parseFloat(dados.valorRecebido) || 0);

  // Busca nome do cliente pelo código
  var nomeCliente = _buscarNomeCliente(dados.codigoCliente);

  var linha = [
    numPedido,
    dados.codigoCliente,
    nomeCliente,
    dados.descricao,
    dataHoje,
    dados.prazoEntrega  ? new Date(dados.prazoEntrega)  : "",
    dados.dataMontagem  ? new Date(dados.dataMontagem)  : "",
    dados.localInstalacao,
    parseFloat(dados.valorTotal)    || 0,
    parseFloat(dados.valorRecebido) || 0,
    saldo,
    dados.responsavel,
    "Aprovado",
    dados.observacoes,
    ""  // ID Calendar (preenchido depois)
  ];

  aba.appendRow(linha);
  _formatarUltimaLinhaPedido(aba);

  // Cria evento no Calendar se tiver data de montagem
  if (dados.dataMontagem) {
    var idEvento = criarEventoCalendar(
      numPedido, nomeCliente, dados.descricao,
      dados.dataMontagem, dados.localInstalacao
    );
    // Registra o ID do evento na coluna O
    var ultimaLinha = aba.getLastRow();
    aba.getRange(ultimaLinha, 15).setValue(idEvento);
  }

  // Atualiza contador de pedidos do cliente
  _incrementarPedidosCliente(dados.codigoCliente);
  atualizarDashboard();

  return "✅ Pedido #" + numPedido + " cadastrado com sucesso!";
}

function _proximoNumeroPedido(aba) {
  var ultLinha = aba.getLastRow();
  if (ultLinha <= 1) return "PED-001";
  var valores = aba.getRange(2, 1, ultLinha - 1, 1).getValues();
  var nums = valores
    .map(function(r) { return parseInt((r[0] + "").replace("PED-", "")) || 0; })
    .filter(function(n) { return !isNaN(n); });
  var proximo = nums.length > 0 ? Math.max.apply(null, nums) + 1 : 1;
  return "PED-" + String(proximo).padStart(3, "0");
}

function _formatarUltimaLinhaPedido(aba) {
  var linha = aba.getLastRow();
  // Formata colunas de data
  aba.getRange(linha, 5).setNumberFormat("dd/MM/yyyy");
  aba.getRange(linha, 6).setNumberFormat("dd/MM/yyyy");
  aba.getRange(linha, 7).setNumberFormat("dd/MM/yyyy");
  // Formata colunas de valor
  aba.getRange(linha, 9, 1, 3).setNumberFormat("R$ #,##0.00");
  // Cor por status
  _colorirLinhaPorStatus(aba, linha);
}

function _colorirLinhaPorStatus(aba, linha) {
  var status = aba.getRange(linha, 13).getValue();
  var cores = {
    "Orçamento":        "#fff9c4",
    "Aprovado":         "#e3f2fd",
    "Em Produção":      "#fff3e0",
    "Aguardando Entrega": "#f3e5f5",
    "Concluído":        "#e8f5e9"
  };
  var cor = cores[status] || "#ffffff";
  aba.getRange(linha, 1, 1, 15).setBackground(cor);
}

/** Retorna lista de pedidos para o frontend */
function listarPedidos() {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.PEDIDOS);
  if (aba.getLastRow() <= 1) return [];
  var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getValues();
  return dados.map(function(r, i) {
    return {
      linha:        i + 2,
      numPedido:    r[0],
      cliente:      r[2],
      descricao:    r[3],
      prazo:        r[5] ? Utilities.formatDate(new Date(r[5]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "—",
      valorTotal:   r[8],
      saldo:        r[10],
      responsavel:  r[11],
      status:       r[12]
    };
  });
}

/** Atualiza o status de um pedido */
function atualizarStatusPedido(linha, novoStatus) {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.PEDIDOS);
  aba.getRange(linha, 13).setValue(novoStatus);
  _colorirLinhaPorStatus(aba, linha);
  atualizarDashboard();
  return "Status atualizado para: " + novoStatus;
}


// ─── MÓDULO 2: CLIENTES ──────────────────────────────────────

function abrirFormCliente() {
  var html = HtmlService.createHtmlOutputFromFile("FormCliente")
    .setTitle("Novo Cliente")
    .setWidth(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

function salvarCliente(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABAS.CLIENTES);

  var codigo = _proximoCodigoCliente(aba);
  var linha = [
    codigo,
    dados.nomeCompleto,
    dados.telefone,
    dados.email,
    dados.endereco,
    dados.cidade,
    new Date(),
    0,
    dados.observacoes
  ];

  aba.appendRow(linha);
  var ultLinha = aba.getLastRow();
  aba.getRange(ultLinha, 7).setNumberFormat("dd/MM/yyyy");

  return "✅ Cliente " + codigo + " – " + dados.nomeCompleto + " cadastrado!";
}

function _proximoCodigoCliente(aba) {
  var ultLinha = aba.getLastRow();
  if (ultLinha <= 1) return "CLI-001";
  var valores = aba.getRange(2, 1, ultLinha - 1, 1).getValues();
  var nums = valores
    .map(function(r) { return parseInt((r[0] + "").replace("CLI-", "")) || 0; })
    .filter(function(n) { return !isNaN(n); });
  var proximo = nums.length > 0 ? Math.max.apply(null, nums) + 1 : 1;
  return "CLI-" + String(proximo).padStart(3, "0");
}

function _buscarNomeCliente(codigo) {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.CLIENTES);
  if (aba.getLastRow() <= 1) return codigo;
  var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] === codigo) return dados[i][1];
  }
  return codigo;
}

function _incrementarPedidosCliente(codigo) {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.CLIENTES);
  if (aba.getLastRow() <= 1) return;
  var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] === codigo) {
      var linha   = i + 2;
      var atual   = parseInt(dados[i][7]) || 0;
      aba.getRange(linha, 8).setValue(atual + 1);
      break;
    }
  }
}

/** Retorna lista de clientes para dropdowns */
function listarClientes() {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.CLIENTES);
  if (aba.getLastRow() <= 1) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues()
    .map(function(r) {
      return { codigo: r[0], nome: r[1], telefone: r[2] };
    });
}

function abrirBuscaCliente() {
  var html = HtmlService.createHtmlOutputFromFile("BuscaCliente")
    .setTitle("Buscar Cliente")
    .setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}

function buscarCliente(termo) {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.CLIENTES);
  if (aba.getLastRow() <= 1) return [];
  var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
  termo = termo.toLowerCase();
  var resultado = dados.filter(function(r) {
    return (r[0] + "").toLowerCase().includes(termo) ||
           (r[1] + "").toLowerCase().includes(termo) ||
           (r[2] + "").toLowerCase().includes(termo);
  });
  // Busca pedidos do cliente
  return resultado.map(function(r) {
    var pedidos = _pedidosDoCliente(r[0]);
    return {
      codigo:       r[0], nome:      r[1], telefone:  r[2],
      email:        r[3], endereco:  r[4], cidade:    r[5],
      cadastro:     r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "—",
      totalPedidos: r[7], observacoes: r[8],
      pedidos:      pedidos
    };
  });
}

function _pedidosDoCliente(codigoCliente) {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.PEDIDOS);
  if (aba.getLastRow() <= 1) return [];
  var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 13).getValues();
  return dados
    .filter(function(r) { return r[1] === codigoCliente; })
    .map(function(r) {
      return {
        num: r[0], descricao: r[3],
        prazo: r[5] ? Utilities.formatDate(new Date(r[5]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "—",
        valor: r[8], status: r[12]
      };
    });
}


// ─── MÓDULO 3: ESTOQUE ───────────────────────────────────────

function abrirFormEntradaEstoque() {
  _abrirFormMovimento("entrada");
}

function abrirFormSaidaEstoque() {
  _abrirFormMovimento("saida");
}

function _abrirFormMovimento(tipo) {
  var html = HtmlService.createHtmlOutputFromFile("FormEstoque")
    .setTitle(tipo === "entrada" ? "📦 Entrada de Material" : "📤 Saída de Material")
    .setWidth(460);
  // Passa o tipo via URL para o HTML
  html = HtmlService.createHtmlOutput(
    HtmlService.createHtmlOutputFromFile("FormEstoque")
      .getContent()
      .replace("{{TIPO}}", tipo)
  ).setTitle(tipo === "entrada" ? "📦 Entrada de Material" : "📤 Saída de Material")
   .setWidth(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

function listarMateriais() {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.ESTOQUE);
  if (aba.getLastRow() <= 1) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 5).getValues()
    .map(function(r) {
      return {
        codigo: r[0], descricao: r[1], unidade: r[2],
        minimo: r[3], atual: r[4]
      };
    });
}

function salvarNovoMaterial(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABAS.ESTOQUE);

  var codigo = _proximoCodigoMaterial(aba);
  aba.appendRow([
    codigo, dados.descricao, dados.unidade,
    parseFloat(dados.minimo) || 0,
    parseFloat(dados.inicial) || 0,
    dados.fornecedor,
    parseFloat(dados.preco) || 0,
    dados.inicial > 0 ? new Date() : "",
    ""
  ]);
  return "✅ Material " + codigo + " – " + dados.descricao + " cadastrado!";
}

function _proximoCodigoMaterial(aba) {
  var ultLinha = aba.getLastRow();
  if (ultLinha <= 1) return "MAT-001";
  var valores = aba.getRange(2, 1, ultLinha - 1, 1).getValues();
  var nums = valores
    .map(function(r) { return parseInt((r[0] + "").replace("MAT-", "")) || 0; })
    .filter(function(n) { return !isNaN(n); });
  var proximo = nums.length > 0 ? Math.max.apply(null, nums) + 1 : 1;
  return "MAT-" + String(proximo).padStart(3, "0");
}

function registrarMovimento(dados) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var abaEst    = ss.getSheetByName(ABAS.ESTOQUE);
  var abaMov    = ss.getSheetByName(ABAS.MOVIMENTOS);
  var tipo      = dados.tipo;         // "entrada" ou "saida"
  var codigo    = dados.codigo;
  var qtd       = parseFloat(dados.quantidade) || 0;

  // Encontra e atualiza o material no estoque
  var dados_est = abaEst.getRange(2, 1, abaEst.getLastRow() - 1, 9).getValues();
  var linhaEncontrada = -1;
  var descricao = "";

  for (var i = 0; i < dados_est.length; i++) {
    if (dados_est[i][0] === codigo) {
      linhaEncontrada = i + 2;
      descricao       = dados_est[i][1];
      var atual       = parseFloat(dados_est[i][4]) || 0;

      if (tipo === "saida" && qtd > atual) {
        return "❌ Quantidade insuficiente. Disponível: " + atual +
               " " + dados_est[i][2];
      }

      var novaQtd = tipo === "entrada" ? atual + qtd : atual - qtd;
      abaEst.getRange(linhaEncontrada, 5).setValue(novaQtd);

      if (tipo === "entrada") {
        abaEst.getRange(linhaEncontrada, 8).setValue(new Date());
        if (dados.fornecedor) {
          abaEst.getRange(linhaEncontrada, 6).setValue(dados.fornecedor);
        }
      } else {
        abaEst.getRange(linhaEncontrada, 9).setValue(new Date());
      }
      break;
    }
  }

  if (linhaEncontrada === -1) return "❌ Material não encontrado.";

  // Registra no log de movimentos
  abaMov.appendRow([
    new Date(), tipo === "entrada" ? "ENTRADA" : "SAÍDA",
    codigo, descricao, qtd,
    dados.pedidoVinculado || "—",
    dados.responsavel,
    dados.observacoes
  ]);
  abaMov.getRange(abaMov.getLastRow(), 1)
        .setNumberFormat("dd/MM/yyyy HH:mm");

  atualizarDashboard();
  return "✅ " + (tipo === "entrada" ? "Entrada" : "Saída") + " de " +
         qtd + " unidade(s) de " + descricao + " registrada!";
}


// ─── MÓDULO 4: GOOGLE CALENDAR ───────────────────────────────

function criarEventoCalendar(numPedido, cliente, projeto, data, local) {
  try {
    var calendario = CalendarApp.getDefaultCalendar();
    var inicio     = new Date(data);
    inicio.setHours(8, 0, 0, 0);
    var fim = new Date(inicio);
    fim.setHours(12, 0, 0, 0);

    var evento = calendario.createEvent(
      "🪵 Montagem/Entrega – " + cliente + " | " + numPedido,
      inicio,
      fim,
      {
        description: "Projeto: " + projeto +
                     "\nLocal: " + (local || "—") +
                     "\nCliente: " + cliente +
                     "\n\nGerado pelo sistema Vieira Móveis.",
        location: local || ""
      }
    );

    return evento.getId();
  } catch (e) {
    Logger.log("Erro ao criar evento: " + e.message);
    return "";
  }
}


// ─── MÓDULO 5: ALERTAS AUTOMÁTICOS ───────────────────────────

function verificarTodosAlertas() {
  verificarPrazos();
  verificarEstoqueMinimo();
}

function verificarPrazos() {
  var aba    = SpreadsheetApp.getActiveSpreadsheet()
                             .getSheetByName(ABAS.PEDIDOS);
  if (aba.getLastRow() <= 1) return;

  var dados  = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getValues();
  var hoje   = new Date();
  hoje.setHours(0, 0, 0, 0);
  var alertas = [];

  dados.forEach(function(r) {
    var numPedido = r[0];
    var cliente   = r[2];
    var projeto   = r[3];
    var prazo     = r[5];
    var status    = r[12];

    if (!prazo || status === "Concluído") return;

    var dataPrazo = new Date(prazo);
    dataPrazo.setHours(0, 0, 0, 0);
    var diasRestantes = Math.floor((dataPrazo - hoje) / 86400000);

    if (diasRestantes <= CONFIG.DIAS_ALERTA_PRAZO && diasRestantes >= 0) {
      alertas.push({
        num: numPedido, cliente: cliente, projeto: projeto,
        dias: diasRestantes, status: status
      });
    }
  });

  if (alertas.length === 0) return;

  var corpo = "Olá!\n\nOs pedidos abaixo estão com prazo próximo:\n\n";
  alertas.forEach(function(a) {
    corpo += "━━━━━━━━━━━━━━━━━━━━━━━━\n";
    corpo += "📋 Pedido: " + a.num + "\n";
    corpo += "👤 Cliente: " + a.cliente + "\n";
    corpo += "🪵 Projeto: " + a.projeto + "\n";
    corpo += "⏰ Dias restantes: " + a.dias +
             (a.dias === 0 ? " (HOJE!)" : "") + "\n";
    corpo += "📌 Status atual: " + a.status + "\n\n";
  });
  corpo += "Acesse a planilha para verificar e atualizar os status.\n\n";
  corpo += "— Sistema " + CONFIG.NOME_EMPRESA;

  MailApp.sendEmail({
    to:      CONFIG.EMAIL_PROPRIETARIO,
    subject: "⚠️ " + alertas.length + " pedido(s) com prazo próximo – " +
             CONFIG.NOME_EMPRESA,
    body:    corpo
  });

  Logger.log("Alerta de prazo enviado: " + alertas.length + " pedido(s).");
}

function verificarEstoqueMinimo() {
  var aba = SpreadsheetApp.getActiveSpreadsheet()
                          .getSheetByName(ABAS.ESTOQUE);
  if (aba.getLastRow() <= 1) return;

  var dados   = aba.getRange(2, 1, aba.getLastRow() - 1, 5).getValues();
  var criticos = [];

  dados.forEach(function(r) {
    var codigo     = r[0];
    var descricao  = r[1];
    var unidade    = r[2];
    var minimo     = parseFloat(r[3]) || 0;
    var disponivel = parseFloat(r[4]) || 0;

    if (disponivel <= minimo) {
      criticos.push({
        codigo: codigo, descricao: descricao, unidade: unidade,
        disponivel: disponivel, minimo: minimo
      });
    }
  });

  if (criticos.length === 0) return;

  var corpo = "Olá!\n\nOs seguintes materiais estão no nível crítico de estoque:\n\n";
  criticos.forEach(function(m) {
    corpo += "━━━━━━━━━━━━━━━━━━━━━━━━\n";
    corpo += "📦 " + m.codigo + " – " + m.descricao + "\n";
    corpo += "   Disponível: " + m.disponivel + " " + m.unidade + "\n";
    corpo += "   Mínimo:     " + m.minimo     + " " + m.unidade + "\n\n";
  });
  corpo += "Providencie a reposição para evitar interrupção na produção.\n\n";
  corpo += "— Sistema " + CONFIG.NOME_EMPRESA;

  MailApp.sendEmail({
    to:      CONFIG.EMAIL_PROPRIETARIO,
    subject: "⚠️ Estoque crítico – " + criticos.length +
             " material(is) abaixo do mínimo – " + CONFIG.NOME_EMPRESA,
    body:    corpo
  });

  Logger.log("Alerta de estoque enviado: " + criticos.length + " material(is).");
}


// ─── MÓDULO 6: DASHBOARD ─────────────────────────────────────

function atualizarDashboard() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABAS.DASHBOARD);
  if (!aba) return;

  aba.clearContents();

  var abaPed = ss.getSheetByName(ABAS.PEDIDOS);
  var abaEst = ss.getSheetByName(ABAS.ESTOQUE);

  var dadosPed = abaPed.getLastRow() > 1
    ? abaPed.getRange(2, 1, abaPed.getLastRow() - 1, 13).getValues()
    : [];

  var dadosEst = abaEst.getLastRow() > 1
    ? abaEst.getRange(2, 1, abaEst.getLastRow() - 1, 5).getValues()
    : [];

  // Contadores de status
  var statusCount = {};
  var faturamentoTotal = 0;
  var saldoPendente    = 0;
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var prazoProximo = 0;
  var emAtraso     = 0;

  dadosPed.forEach(function(r) {
    var status = r[12] || "Sem status";
    statusCount[status] = (statusCount[status] || 0) + 1;
    faturamentoTotal += parseFloat(r[8]) || 0;
    if (status !== "Concluído") {
      saldoPendente += parseFloat(r[10]) || 0;
      if (r[5]) {
        var prazo = new Date(r[5]);
        prazo.setHours(0, 0, 0, 0);
        var dias = Math.floor((prazo - hoje) / 86400000);
        if (dias >= 0 && dias <= 3) prazoProximo++;
        if (dias < 0)               emAtraso++;
      }
    }
  });

  var estoqueCritico = dadosEst.filter(function(r) {
    return (parseFloat(r[4]) || 0) <= (parseFloat(r[3]) || 0);
  }).length;

  // Escreve o dashboard
  var dados = [
    ["🪵 VIEIRA MÓVEIS SOB MEDIDA – PAINEL DE ACOMPANHAMENTO", "", ""],
    ["Atualizado em: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"), "", ""],
    ["", "", ""],
    ["📋 PEDIDOS POR STATUS", "", ""],
    ["Status", "Quantidade", ""],
  ];

  Object.keys(statusCount).forEach(function(s) {
    dados.push([s, statusCount[s], ""]);
  });

  dados.push(["TOTAL", dadosPed.length, ""]);
  dados.push(["", "", ""]);
  dados.push(["💰 FINANCEIRO", "", ""]);
  dados.push(["Faturamento total (todos os pedidos)", "R$ " + faturamentoTotal.toFixed(2), ""]);
  dados.push(["Saldo a receber (pedidos em aberto)",  "R$ " + saldoPendente.toFixed(2), ""]);
  dados.push(["", "", ""]);
  dados.push(["⏰ ALERTAS", "", ""]);
  dados.push(["Pedidos com prazo nos próximos 3 dias", prazoProximo, ""]);
  dados.push(["Pedidos em atraso",                     emAtraso,     ""]);
  dados.push(["Materiais em estoque crítico",           estoqueCritico, ""]);

  aba.getRange(1, 1, dados.length, 3).setValues(dados);

  // Formatação visual
  aba.getRange(1, 1, 1, 3).setBackground("#2d5016").setFontColor("#ffffff")
     .setFontWeight("bold").setFontSize(13);
  aba.getRange(4, 1, 1, 3).setBackground("#558b2f").setFontColor("#ffffff")
     .setFontWeight("bold");
  aba.getRange(9 + Object.keys(statusCount).length, 1, 1, 3)
     .setBackground("#558b2f").setFontColor("#ffffff").setFontWeight("bold");

  aba.autoResizeColumns(1, 3);
  Logger.log("Dashboard atualizado.");
}
