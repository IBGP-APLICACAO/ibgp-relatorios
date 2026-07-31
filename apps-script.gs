// ════════════════════════════════════════════════════════════════
//  IBGP — Apps Script (versão completa)
//  Cole em: script.google.com → projeto vinculado à planilha
//  Depois: Implantar → Nova implantação → Aplicativo da Web
//  Executar como: Eu | Acesso: Qualquer pessoa
// ════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    salvarConcurso(ss, data);
    salvarLocais(ss, data);
    salvarResultados(ss, data);
    salvarHorarios(ss, data);
    salvarTestemunhas(ss, data);
    salvarOcorrencias(ss, data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "erro", msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helpers ──────────────────────────────────────────────────────
function getOrCreateSheet(ss, nome, cabecalho) {
  var ws = ss.getSheetByName(nome);
  if (!ws) {
    ws = ss.insertSheet(nome);
    ws.appendRow(cabecalho);
  }
  return ws;
}

function deletarLinhasEscola(ws, escola) {
  var ultima = ws.getLastRow();
  if (ultima < 2) return;
  var col = ws.getRange(2, 1, ultima - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {
    if (col[i][0] === escola) ws.deleteRow(i + 2);
  }
}

function atualizarOuInserir(ws, escola, novaLinha) {
  var ultima = ws.getLastRow();
  if (ultima > 1) {
    var col = ws.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      if (col[i][0] === escola) {
        ws.getRange(i + 2, 1, 1, novaLinha.length).setValues([novaLinha]);
        return;
      }
    }
  }
  ws.appendRow(novaLinha);
}

// ── Aba CONCURSO ─────────────────────────────────────────────────
function salvarConcurso(ss, data) {
  if (ss.getSheetByName("CONCURSO")) return; // já existe, não sobrescreve
  var ws = ss.insertSheet("CONCURSO");
  ws.appendRow(["CAMPO", "VALOR"]);
  [
    ["Título do Concurso",        data.concurso  || ""],
    ["Edital",                    data.edital    || ""],
    ["Data da Prova",             data.data      || ""],
    ["Coordenador Pedagógico",    ""],
    ["Coordenador Logístico",     ""],
    ["Hr Reunião — Início",       "05h50"],
    ["Hr Reunião — Fim",          "06h45"],
    ["Hr Abertura do Portão",     "07h00"],
    ["Hr Distribuição de Provas", "07h30"],
    ["Hr Fechamento do Portão",   "08h00"],
    ["Hr Início da Prova",        "08h30"],
  ].forEach(function(c){ ws.appendRow(c); });
}

// ── Aba LOCAIS ────────────────────────────────────────────────────
function salvarLocais(ss, data) {
  var ws = getOrCreateSheet(ss, "LOCAIS", [
    "Local / Escola", "Coordenador Local", "Nº Salas", "Turno",
    "Houve apoio?", "Nome do apoio", "Observações gerais", "Enviado em"
  ]);

  var turnos = [];
  if (data.manha) turnos.push("MANHÃ");
  if (data.tarde) turnos.push("TARDE");
  var turno = turnos.join(" e ") || "";

  atualizarOuInserir(ws, data.escola, [
    data.escola, data.coord_local, data.salas, turno,
    data.houve_apoio, data.nome_apoio, data.obs, data.enviado_em
  ]);
}

// ── Aba RESULTADOS ────────────────────────────────────────────────
function salvarResultados(ss, data) {
  var ws = getOrCreateSheet(ss, "RESULTADOS", [
    "Local / Escola", "Sala", "Turno", "Cargo",
    "Inscritos", "Presentes", "Ausentes", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  ["manha", "tarde"].forEach(function(t) {
    if (!data[t]) return;
    data[t].estatistica.forEach(function(r) {
      ws.appendRow([data.escola, r.sala, r.turno, r.cargo, r.inscr, r.pres, r.aus, data.enviado_em]);
    });
  });
}

// ── Aba HORÁRIOS ──────────────────────────────────────────────────
function salvarHorarios(ss, data) {
  var ws = getOrCreateSheet(ss, "HORÁRIOS", [
    "Local / Escola", "Turno",
    "Chegada da equipe", "Abertura do portão", "Fechamento do portão",
    "Distribuição dos pacotes", "Início da aplicação", "Encerramento",
    "Fechamento do malote", "Abertura do malote", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    var key = par[0], label = par[1];
    if (!data[key]) return;
    var h = data[key].horarios;
    ws.appendRow([
      data.escola, label,
      h.chegada, h.ab_portao, h.fch_portao,
      h.distrib, h.inicio, h.encerr,
      h.fch_malote, h.ab_malote, data.enviado_em
    ]);
  });
}

// ── Aba TESTEMUNHAS ───────────────────────────────────────────────
function salvarTestemunhas(ss, data) {
  var ws = getOrCreateSheet(ss, "TESTEMUNHAS", [
    "Local / Escola", "Turno", "Tipo",
    "Testemunha 1 — Nome", "Testemunha 1 — CPF",
    "Testemunha 2 — Nome", "Testemunha 2 — CPF", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    var key = par[0], label = par[1];
    if (!data[key]) return;
    var t = data[key].testemunhas;

    ws.appendRow([
      data.escola, label, "Abertura do Portão",
      t.portao[0].nome, t.portao[0].cpf,
      t.portao[1].nome, t.portao[1].cpf, data.enviado_em
    ]);
    ws.appendRow([
      data.escola, label, "Abertura do Malote",
      t.malote[0].nome, t.malote[0].cpf,
      t.malote[1].nome, t.malote[1].cpf, data.enviado_em
    ]);
  });
}

// ── Aba OCORRÊNCIAS ───────────────────────────────────────────────
function salvarOcorrencias(ss, data) {
  var ws = getOrCreateSheet(ss, "OCORRÊNCIAS", [
    "Local / Escola", "Turno",
    "Ocorrências em sala", "Descrição",
    "Prova condicional", "Descrição",
    "Toque de celular", "Descrição",
    "Declaração de comparecimento", "Descrição",
    "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    var key = par[0], label = par[1];
    if (!data[key]) return;
    var o = data[key].ocorrencias;
    ws.appendRow([
      data.escola, label,
      o.sala.resp, o.sala.desc,
      o.cond.resp, o.cond.desc,
      o.toque.resp, o.toque.desc,
      o.decl.resp, o.decl.desc,
      data.enviado_em
    ]);
  });
}
