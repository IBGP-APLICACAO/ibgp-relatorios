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
    salvarCandidatos(ss, data);
    salvarFotos(data);

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
    "Inscritos", "Presentes", "Ausentes", "Inclusões", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  ["manha", "tarde"].forEach(function(t) {
    if (!data[t]) return;
    data[t].estatistica.forEach(function(r) {
      ws.appendRow([data.escola, r.sala, r.turno, r.cargo, r.inscr, r.pres, r.aus, r.incl || 0, data.enviado_em]);
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
    "Testemunha 2 — Nome", "Testemunha 2 — CPF",
    "Testemunha 3 — Nome", "Testemunha 3 — CPF", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    var key = par[0], label = par[1];
    if (!data[key]) return;
    var t = data[key].testemunhas;

    ws.appendRow([
      data.escola, label, "Abertura do Malote",
      t.ab_malote[0].nome, t.ab_malote[0].cpf,
      t.ab_malote[1].nome, t.ab_malote[1].cpf,
      t.ab_malote[2].nome, t.ab_malote[2].cpf, data.enviado_em
    ]);
    ws.appendRow([
      data.escola, label, "Fechamento do Malote",
      t.fch_malote[0].nome, t.fch_malote[0].cpf,
      t.fch_malote[1].nome, t.fch_malote[1].cpf,
      t.fch_malote[2].nome, t.fch_malote[2].cpf, data.enviado_em
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

// ── Fotos → Google Drive ──────────────────────────────────────────
function salvarFotos(data) {
  if (!data.fotos || data.fotos.length === 0) return;

  // Cria estrutura: Minha Unidade / IBGP_RELATORIOS / <NomeEscola>
  var root = DriveApp.getRootFolder();

  var ibgpFolder;
  var ibgpIt = root.getFoldersByName("IBGP_RELATORIOS");
  ibgpFolder = ibgpIt.hasNext() ? ibgpIt.next() : root.createFolder("IBGP_RELATORIOS");

  var nomeEscola = data.escola.replace(/[\/\\:*?"<>|]/g, "_");
  var escolaFolder;
  var escolaIt = ibgpFolder.getFoldersByName(nomeEscola);
  escolaFolder = escolaIt.hasNext() ? escolaIt.next() : ibgpFolder.createFolder(nomeEscola);

  data.fotos.forEach(function(foto, idx) {
    try {
      var bytes = Utilities.base64Decode(foto.b64);
      var cat   = (foto.categoria || "Foto").replace(/[\/\\:*?"<>|]/g, "_");
      var nome  = cat + "_" + (idx + 1) + "_" + foto.nome;
      var blob  = Utilities.newBlob(bytes, "image/jpeg", nome);
      escolaFolder.createFile(blob);
    } catch(e) { /* ignora foto com erro */ }
  });
}

// ── Aba CANDIDATOS ────────────────────────────────────────────────
function salvarCandidatos(ss, data) {
  var ws = getOrCreateSheet(ss, "CANDIDATOS", [
    "Local / Escola", "Turno", "Tipo", "Nº Inscrição", "Candidato", "Sala", "Enviado em"
  ]);

  deletarLinhasEscola(ws, data.escola);

  // Apoio (apenas nome)
  if (data.apoio_entries && data.apoio_entries.length > 0) {
    data.apoio_entries.forEach(function(nome) {
      ws.appendRow([data.escola, "—", "Apoio", "", nome, "", data.enviado_em]);
    });
  }

  // Ocorrências por turno
  var tipos = [
    {key: "sala",  label: "Ocorrências em sala"},
    {key: "cond",  label: "Prova condicional"},
    {key: "toque", label: "Toque de celular"},
    {key: "decl",  label: "Declaração de comparecimento"},
  ];

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    var tkey = par[0], tlabel = par[1];
    if (!data[tkey]) return;
    tipos.forEach(function(tipo) {
      var ocorr = data[tkey].ocorrencias[tipo.key];
      if (!ocorr || !ocorr.candidatos) return;
      ocorr.candidatos.forEach(function(c) {
        ws.appendRow([data.escola, tlabel, tipo.label, c.inscricao, c.candidato, c.sala, data.enviado_em]);
      });
    });
  });
}
