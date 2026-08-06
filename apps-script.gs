// ════════════════════════════════════════════════════════════════
//  IBGP — Apps Script (versão completa)
//  Cole em: script.google.com → projeto vinculado à planilha
//  Depois: Implantar → Nova implantação → Aplicativo da Web
//  Executar como: Eu | Acesso: Qualquer pessoa
// ════════════════════════════════════════════════════════════════

// E-mail que recebe os relatórios de cada escola
var EMAIL_DESTINO = "natalia.matias@ibgp.org.br";

// ────────────────────────────────────────────────────────────────
//  doPost — recebe formulário do coordenador
// ────────────────────────────────────────────────────────────────
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
    var linkFotos = salvarFotos(data);

    enviarEmailRelatorio(data, linkFotos);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "erro", msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ────────────────────────────────────────────────────────────────
//  doGet — retorna dados consolidados para o Relatório Geral
// ────────────────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";

  if (action === "relatorio") {
    var dados = coletarDadosRelatorio();
    return ContentService
      .createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────────
//  Helper — formata valor de hora do Sheets como "HH:mm"
// ────────────────────────────────────────────────────────────────
function horaStr(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(v);
}

// ────────────────────────────────────────────────────────────────
//  Coleta dados de todas as abas para o Relatório Geral
// ────────────────────────────────────────────────────────────────
function coletarDadosRelatorio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultado = { concurso: {}, escolas: [] };

  // CONCURSO
  var wsConcurso = ss.getSheetByName("CONCURSO");
  if (wsConcurso && wsConcurso.getLastRow() > 1) {
    var rows = wsConcurso.getRange(2, 1, wsConcurso.getLastRow() - 1, 2).getValues();
    rows.forEach(function(r) { resultado.concurso[r[0]] = r[1]; });
  }

  // LOCAIS — monta lista de escolas
  var wsLocais = ss.getSheetByName("LOCAIS");
  var escolas = {};
  if (wsLocais && wsLocais.getLastRow() > 1) {
    var lRows = wsLocais.getRange(2, 1, wsLocais.getLastRow() - 1, 9).getValues();
    lRows.forEach(function(r) {
      if (!r[0]) return;
      escolas[r[0]] = {
        nome: r[0], coord_local: r[1], salas: r[2], turno: r[3],
        houve_apoio: r[4], nome_apoio: r[5], obs: r[6], enviado_em: r[7],
        foto_url: r[8] || "",
        horarios: [], resultados: [], testemunhas: [], ocorrencias: [], candidatos: []
      };
    });
  }

  // HORÁRIOS
  var wsHor = ss.getSheetByName("HORÁRIOS");
  if (wsHor && wsHor.getLastRow() > 1) {
    wsHor.getRange(2, 1, wsHor.getLastRow() - 1, 11).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].horarios.push({
        turno: r[1],
        chegada:   horaStr(r[2]),  ab_portao: horaStr(r[3]),
        fch_portao: horaStr(r[4]), distrib:   horaStr(r[5]),
        inicio:    horaStr(r[6]),  encerr:    horaStr(r[7]),
        fch_malote: horaStr(r[8]), ab_malote: horaStr(r[9])
      });
    });
  }

  // RESULTADOS
  var wsRes = ss.getSheetByName("RESULTADOS");
  if (wsRes && wsRes.getLastRow() > 1) {
    wsRes.getRange(2, 1, wsRes.getLastRow() - 1, 9).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      if (!r[3]) return; // ignora linhas sem cargo
      escolas[r[0]].resultados.push({
        sala: r[1], turno: r[2], cargo: String(r[3]),
        inscr: r[4], pres: r[5], aus: r[6], incl: r[7]
      });
    });
  }

  // TESTEMUNHAS
  var wsTest = ss.getSheetByName("TESTEMUNHAS");
  if (wsTest && wsTest.getLastRow() > 1) {
    wsTest.getRange(2, 1, wsTest.getLastRow() - 1, 10).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].testemunhas.push({
        turno: r[1], tipo: r[2],
        t1_nome: r[3], t1_cpf: r[4],
        t2_nome: r[5], t2_cpf: r[6],
        t3_nome: r[7], t3_cpf: r[8]
      });
    });
  }

  // OCORRÊNCIAS
  var wsOcorr = ss.getSheetByName("OCORRÊNCIAS");
  if (wsOcorr && wsOcorr.getLastRow() > 1) {
    wsOcorr.getRange(2, 1, wsOcorr.getLastRow() - 1, 7).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].ocorrencias.push({
        turno: r[1],
        sala: r[2], cond: r[3], toque: r[4], decl: r[5]
      });
    });
  }

  // CANDIDATOS
  var wsCand = ss.getSheetByName("CANDIDATOS");
  if (wsCand && wsCand.getLastRow() > 1) {
    wsCand.getRange(2, 1, wsCand.getLastRow() - 1, 7).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].candidatos.push({
        turno: r[1], tipo: r[2], inscricao: r[3], candidato: r[4], sala: r[5]
      });
    });
  }

  resultado.escolas = Object.values(escolas);
  return resultado;
}

// ────────────────────────────────────────────────────────────────
//  E-mail — envia relatório do coordenador após cada envio
// ────────────────────────────────────────────────────────────────
function enviarEmailRelatorio(data, linkFotos) {
  try {
    var assunto = "Relatório de Aplicação — " + data.escola + " — " + data.data;
    var html = montarHtmlEmail(data, linkFotos);
    MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: html });
  } catch(e) { /* não bloqueia o fluxo principal */ }
}

function montarHtmlEmail(data, linkFotos) {
  var s = [];
  var cor = { manha: "#1F3A6E", tarde: "#CF3432", header: "#393939" };

  s.push('<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;background:#F2F2F2;padding:20px;">');
  s.push('<div style="background:#393939;padding:16px 24px;border-radius:8px 8px 0 0;">');
  s.push('<h1 style="color:#FFF;font-size:16px;margin:0;">IBGP — Relatório de Aplicação de Prova</h1>');
  s.push('<p style="color:#BBB;font-size:13px;margin:4px 0 0;">Enviado automaticamente pelo sistema</p>');
  s.push('</div>');
  s.push('<div style="height:4px;background:#CF3432;margin-bottom:16px;"></div>');

  // Identificação
  s.push(card("Identificação", cor.header, [
    row2col("Escola", data.escola, "Concurso", data.concurso || ""),
    row2col("Coordenador Local", data.coord_local, "Data da Prova", data.data || ""),
    row2col("Houve Apoio?", data.houve_apoio, "Nome do Apoio", data.nome_apoio || "—"),
    row2col("Enviado em", data.enviado_em, "", "")
  ]));

  // Turnos
  [["manha", "Manhã", cor.manha], ["tarde", "Tarde", cor.tarde]].forEach(function(par) {
    var key = par[0], label = par[1], c = par[2];
    if (!data[key]) return;
    var d = data[key];
    var h = d.horarios;

    // Horários
    s.push(card("Turno " + label + " — Horários", c, [
      row2col("Chegada da equipe", h.chegada, "Abertura do portão", h.ab_portao),
      row2col("Abertura do malote", h.ab_malote, "Distribuição dos pacotes", h.distrib),
      row2col("Fechamento do portão", h.fch_portao, "Início da aplicação", h.inicio),
      row2col("Encerramento", h.encerr, "Fechamento do malote", h.fch_malote)
    ]));

    // Testemunhas
    var t = d.testemunhas;
    s.push(card("Turno " + label + " — Testemunhas Abertura do Malote", c, [
      row2col("Testemunha 1", t.ab_malote[0].nome + " — " + t.ab_malote[0].cpf, "Testemunha 2", t.ab_malote[1].nome + " — " + t.ab_malote[1].cpf),
      row2col("Testemunha 3", t.ab_malote[2].nome + " — " + t.ab_malote[2].cpf, "", "")
    ]));
    s.push(card("Turno " + label + " — Testemunhas Fechamento do Malote", c, [
      row2col("Testemunha 1", t.fch_malote[0].nome + " — " + t.fch_malote[0].cpf, "Testemunha 2", t.fch_malote[1].nome + " — " + t.fch_malote[1].cpf),
      row2col("Testemunha 3", t.fch_malote[2].nome + " — " + t.fch_malote[2].cpf, "", "")
    ]));

    // Estatística
    var tabela = '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;">';
    tabela += '<tr style="background:' + c + ';color:#FFF;"><th style="padding:7px 8px;text-align:left;">Cargo</th><th style="padding:7px 8px;">Inscritos</th><th style="padding:7px 8px;">Presentes</th><th style="padding:7px 8px;">Ausentes</th><th style="padding:7px 8px;">Inclusões</th></tr>';
    var totI=0,totP=0,totA=0,totIn=0;
    d.estatistica.forEach(function(r,i) {
      totI+=r.inscr; totP+=r.pres; totA+=r.aus; totIn+=r.incl||0;
      tabela += '<tr style="background:' + (i%2===0?'#FFF':'#F4F4F4') + ';">';
      tabela += '<td style="padding:6px 8px;">' + r.sala + ' — ' + r.cargo.replace(/^01\s+/,'') + '</td>';
      tabela += '<td style="text-align:center;padding:6px 8px;">' + r.inscr + '</td>';
      tabela += '<td style="text-align:center;padding:6px 8px;">' + r.pres + '</td>';
      tabela += '<td style="text-align:center;padding:6px 8px;">' + r.aus + '</td>';
      tabela += '<td style="text-align:center;padding:6px 8px;">' + (r.incl||0) + '</td></tr>';
    });
    tabela += '<tr style="background:#EEE;font-weight:700;"><td style="padding:7px 8px;">TOTAL</td>';
    tabela += '<td style="text-align:center;padding:7px 8px;">' + totI + '</td>';
    tabela += '<td style="text-align:center;padding:7px 8px;">' + totP + '</td>';
    tabela += '<td style="text-align:center;padding:7px 8px;">' + totA + '</td>';
    tabela += '<td style="text-align:center;padding:7px 8px;">' + totIn + '</td></tr>';
    tabela += '</table>';
    s.push(card("Turno " + label + " — Estatística por Cargo", c, [tabela], true));

    // Ocorrências
    var o = d.ocorrencias;
    var linhasOcorr = [
      row2col("Ocorrências em sala", o.sala.resp, "Prova condicional", o.cond.resp),
      row2col("Toque de celular", o.toque.resp, "Declaração de comparecimento", o.decl.resp)
    ];
    ["sala","cond","toque","decl"].forEach(function(tipo) {
      var cands = o[tipo].candidatos || [];
      if(cands.length > 0) {
        var lista = cands.map(function(c){ return c.candidato + (c.inscricao?" ("+c.inscricao+")":"") + (c.sala?" — Sala "+c.sala:"") + (c.desc?" — "+c.desc:""); }).join("<br>");
        linhasOcorr.push('<div style="padding:6px 0;font-size:13px;color:#555;">' + lista + '</div>');
      }
    });
    s.push(card("Turno " + label + " — Ocorrências", c, linhasOcorr, true));
  });

  // Observações
  if (data.obs) {
    s.push(card("Observações Gerais", cor.header, [
      '<p style="font-size:13px;color:#393939;margin:0;">' + data.obs + '</p>'
    ], true));
  }

  // Fotos
  if (linkFotos) {
    s.push(card("Registro Fotográfico", cor.header, [
      '<p style="font-size:13px;"><a href="' + linkFotos + '" style="color:#CF3432;">Ver fotos no Google Drive →</a></p>'
    ], true));
  }

  s.push('</div>');
  return s.join('');
}

function card(titulo, cor, linhas, raw) {
  var html = '<div style="background:#FFF;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.07);margin-bottom:16px;overflow:hidden;">';
  html += '<div style="background:' + cor + ';padding:10px 16px;">';
  html += '<h3 style="color:#FFF;font-size:13px;margin:0;text-transform:uppercase;letter-spacing:.4px;">' + titulo + '</h3></div>';
  html += '<div style="padding:16px;">';
  if (raw) {
    linhas.forEach(function(l){ html += l; });
  } else {
    linhas.forEach(function(l){ html += l; });
  }
  html += '</div></div>';
  return html;
}

function row2col(l1, v1, l2, v2) {
  var cell = function(l, v) {
    if (!l) return '<td style="width:50%;"></td>';
    return '<td style="width:50%;padding:4px 8px 4px 0;vertical-align:top;">' +
      '<span style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.4px;">' + l + '</span>' +
      '<br><span style="font-size:13px;color:#393939;">' + (v || '—') + '</span></td>';
  };
  return '<table style="width:100%;margin-bottom:10px;"><tr>' + cell(l1,v1) + cell(l2,v2) + '</tr></table>';
}

// ════════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════════
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

// ════════════════════════════════════════════════════════════════
//  Abas de dados
// ════════════════════════════════════════════════════════════════

function salvarConcurso(ss, data) {
  if (ss.getSheetByName("CONCURSO")) return;
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

function salvarLocais(ss, data) {
  var ws = getOrCreateSheet(ss, "LOCAIS", [
    "Local / Escola", "Coordenador Local", "Nº Salas", "Turno",
    "Houve apoio?", "Nome do apoio", "Observações gerais", "Enviado em"
  ]);
  var turnos = [];
  if (data.manha) turnos.push("MANHÃ");
  if (data.tarde) turnos.push("TARDE");
  atualizarOuInserir(ws, data.escola, [
    data.escola, data.coord_local, data.salas, turnos.join(" e ") || "",
    data.houve_apoio, data.nome_apoio, data.obs, data.enviado_em
  ]);
}

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

function salvarHorarios(ss, data) {
  var ws = getOrCreateSheet(ss, "HORÁRIOS", [
    "Local / Escola", "Turno",
    "Chegada da equipe", "Abertura do portão", "Fechamento do portão",
    "Distribuição dos pacotes", "Início da aplicação", "Encerramento",
    "Fechamento do malote", "Abertura do malote", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);
  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    if (!data[par[0]]) return;
    var h = data[par[0]].horarios;
    ws.appendRow([
      data.escola, par[1],
      h.chegada, h.ab_portao, h.fch_portao,
      h.distrib, h.inicio, h.encerr,
      h.fch_malote, h.ab_malote, data.enviado_em
    ]);
  });
}

function salvarTestemunhas(ss, data) {
  var ws = getOrCreateSheet(ss, "TESTEMUNHAS", [
    "Local / Escola", "Turno", "Tipo",
    "Testemunha 1 — Nome", "Testemunha 1 — CPF",
    "Testemunha 2 — Nome", "Testemunha 2 — CPF",
    "Testemunha 3 — Nome", "Testemunha 3 — CPF", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);
  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    if (!data[par[0]]) return;
    var t = data[par[0]].testemunhas;
    ws.appendRow([
      data.escola, par[1], "Abertura do Malote",
      t.ab_malote[0].nome, t.ab_malote[0].cpf,
      t.ab_malote[1].nome, t.ab_malote[1].cpf,
      t.ab_malote[2].nome, t.ab_malote[2].cpf, data.enviado_em
    ]);
    ws.appendRow([
      data.escola, par[1], "Fechamento do Malote",
      t.fch_malote[0].nome, t.fch_malote[0].cpf,
      t.fch_malote[1].nome, t.fch_malote[1].cpf,
      t.fch_malote[2].nome, t.fch_malote[2].cpf, data.enviado_em
    ]);
  });
}

function salvarOcorrencias(ss, data) {
  var ws = getOrCreateSheet(ss, "OCORRÊNCIAS", [
    "Local / Escola", "Turno",
    "Ocorrências em sala", "Prova condicional",
    "Toque de celular", "Declaração de comparecimento", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);
  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    if (!data[par[0]]) return;
    var o = data[par[0]].ocorrencias;
    ws.appendRow([
      data.escola, par[1],
      o.sala.resp, o.cond.resp, o.toque.resp, o.decl.resp,
      data.enviado_em
    ]);
  });
}

function salvarCandidatos(ss, data) {
  var ws = getOrCreateSheet(ss, "CANDIDATOS", [
    "Local / Escola", "Turno", "Tipo", "Nº Inscrição", "Candidato", "Sala", "Descrição", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);

  var tipos = [
    {key:"sala",  label:"Ocorrências em sala"},
    {key:"cond",  label:"Prova condicional"},
    {key:"toque", label:"Toque de celular"},
    {key:"decl",  label:"Declaração de comparecimento"},
  ];

  [["manha","Manhã"],["tarde","Tarde"]].forEach(function(par) {
    if (!data[par[0]]) return;
    tipos.forEach(function(tipo) {
      var ocorr = data[par[0]].ocorrencias[tipo.key];
      if (!ocorr || !ocorr.candidatos) return;
      ocorr.candidatos.forEach(function(c) {
        ws.appendRow([data.escola, par[1], tipo.label, c.inscricao, c.candidato, c.sala, c.desc || "", data.enviado_em]);
      });
    });
  });
}

function salvarFotos(data) {
  if (!data.fotos || data.fotos.length === 0) return null;

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
    } catch(e) {}
  });

  // Retorna link público para a pasta
  escolaFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = "https://drive.google.com/drive/folders/" + escolaFolder.getId();

  // Salva o link na aba LOCAIS (coluna 9) para uso no Relatório Geral
  try {
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var wsL = ss2.getSheetByName("LOCAIS");
    if (wsL && wsL.getLastRow() > 1) {
      var nomes = wsL.getRange(2, 1, wsL.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < nomes.length; i++) {
        if (String(nomes[i][0]).trim() === String(data.escola).trim()) {
          wsL.getRange(i + 2, 9).setValue(link);
          break;
        }
      }
    }
  } catch(e2) {}

  return link;
}
