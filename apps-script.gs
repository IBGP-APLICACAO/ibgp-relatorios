// ════════════════════════════════════════════════════════════════
//  IBGP — Apps Script para receber dados do formulário web
//  Cole este código em: script.google.com → Novo projeto
//  Depois: Implantar → Novo Implantação → Aplicativo da Web
// ════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    salvarConcurso(ss, data);
    salvarLocais(ss, data);
    salvarResultados(ss, data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "erro", msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Aba CONCURSO (info geral, salva apenas uma vez) ───────────────
function salvarConcurso(ss, data) {
  var ws = ss.getSheetByName("CONCURSO");
  if (!ws) {
    ws = ss.insertSheet("CONCURSO");
    ws.appendRow(["CAMPO", "VALOR"]);
    var campos = [
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
    ];
    campos.forEach(function(c) { ws.appendRow(c); });
  }
}

// ── Aba LOCAIS (uma linha por escola) ─────────────────────────────
function salvarLocais(ss, data) {
  var ws = ss.getSheetByName("LOCAIS");
  if (!ws) {
    ws = ss.insertSheet("LOCAIS");
    ws.appendRow([
      "Local / Escola", "Coordenador Local", "Nº Salas", "Turno",
      "Andamento Geral", "Toque de Celular", "Ocorrências"
    ]);
  }

  // Calcula turno a partir dos resultados
  var hasManha = data.resultados.some(function(r){ return r.turno === "Manhã"; });
  var haTarde  = data.resultados.some(function(r){ return r.turno === "Tarde"; });
  var turno    = (hasManha && haTarde) ? "AMBOS" : hasManha ? "MANHÃ" : "TARDE";

  var novaLinha = [
    data.escola, data.coord_local, data.salas, turno,
    data.andamento, data.toque_celular, data.ocorrencias
  ];

  // Atualiza linha existente ou insere nova
  var ultima = ws.getLastRow();
  if (ultima > 1) {
    var colA = ws.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      if (colA[i][0] === data.escola) {
        ws.getRange(i + 2, 1, 1, novaLinha.length).setValues([novaLinha]);
        return;
      }
    }
  }
  ws.appendRow(novaLinha);
}

// ── Aba RESULTADOS (uma linha por sala/cargo) ─────────────────────
function salvarResultados(ss, data) {
  var ws = ss.getSheetByName("RESULTADOS");
  if (!ws) {
    ws = ss.insertSheet("RESULTADOS");
    ws.appendRow([
      "Local / Escola", "Sala", "Turno", "Cargo",
      "Inscritos", "Presentes", "Ausentes", "Eliminados", "Enviado em"
    ]);
  }

  // Remove linhas anteriores desta escola (evita duplicatas em reenvio)
  var ultima = ws.getLastRow();
  if (ultima > 1) {
    var colA = ws.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] === data.escola) {
        ws.deleteRow(i + 2);
      }
    }
  }

  // Insere novas linhas
  data.resultados.forEach(function(r) {
    ws.appendRow([
      data.escola, r.sala, r.turno, r.cargo,
      r.inscr, r.pres, r.aus, r.elim, data.enviado_em
    ]);
  });
}
