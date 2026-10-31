// ════════════════════════════════════════════════════════════════
//  IBGP — Apps Script (versão dinâmica — suporta N turnos)
//  Cole em: script.google.com → projeto vinculado à planilha
//  Depois: Implantar → Nova implantação → Aplicativo da Web
//  Executar como: Eu | Acesso: Qualquer pessoa
//  ⚠️  OBRIGATÓRIO: Configurações (engrenagem) → ativar "Ambiente V8"
//      Sem V8, Object.values() não funciona e o relatório não gera.
// ════════════════════════════════════════════════════════════════

var EMAIL_DESTINO = "natalia.matias@ibgp.org.br";

// ────────────────────────────────────────────────────────────────
//  Helper — retorna array de nomes de turnos ativos no payload
// ────────────────────────────────────────────────────────────────
function getTurnosAtivos(data) {
  var t = data.turnos || {};
  return Object.keys(t).filter(function(k) { return t[k] !== null; });
}

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

  // ── Lookup de candidato por CPF ─────────────────────────────────
  if (action === "buscarCandidato") {
    var cpf = String((e.parameter && e.parameter.cpf) || "").replace(/\D/g, "");
    if (!cpf) {
      return ContentService
        .createTextOutput(JSON.stringify({ erro: "CPF não informado" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var wsAl = ss.getSheetByName("DADOS ALOCAÇÃO");
    if (!wsAl || wsAl.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ erro: "Aba DADOS ALOCAÇÃO não encontrada" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // Colunas: 1=Índice 2=INSCRIÇÃO 3=CANDIDATO 4=CPF 8=CARGO 10=TURNO 11=SALA 14=ESCOLA
    var dados = wsAl.getRange(2, 1, wsAl.getLastRow() - 1, 14).getValues();
    for (var i = 0; i < dados.length; i++) {
      var cpfPlanilha = String(dados[i][3]).replace(/\D/g, "").padStart(11, "0");
      if (cpfPlanilha === cpf) {
        return ContentService
          .createTextOutput(JSON.stringify({
            nome:      String(dados[i][2]),
            inscricao: String(dados[i][1]),
            cargo:     String(dados[i][7]),
            turno:     String(dados[i][9]),
            sala:      String(dados[i][10]),
            escola:    String(dados[i][13])
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ erro: "Candidato não encontrado" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────────
//  Helper — formata hora do Sheets como "HH:mm"
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

  var wsConcurso = ss.getSheetByName("CONCURSO");
  if (wsConcurso && wsConcurso.getLastRow() > 1) {
    wsConcurso.getRange(2, 1, wsConcurso.getLastRow() - 1, 2).getValues()
      .forEach(function(r) { resultado.concurso[r[0]] = r[1]; });
  }

  var wsLocais = ss.getSheetByName("LOCAIS");
  var escolas = {};
  if (wsLocais && wsLocais.getLastRow() > 1) {
    wsLocais.getRange(2, 1, wsLocais.getLastRow() - 1, 10).getValues()
      .forEach(function(r) {
        if (!r[0]) return;
        var membStr = r[7] ? String(r[7]) : "";
        escolas[r[0]] = {
          nome: r[0], coord_local: r[1], salas: r[2], turno: r[3],
          houve_apoio: r[4], nome_apoio: r[5],
          houve_comissao: r[6], membros_comissao: membStr,
          // comissao como array (esperado pelo admin.html)
          comissao: membStr ? membStr.split("; ").filter(Boolean) : [],
          obs: r[8], enviado_em: r[9], foto_url: "",
          horarios: [], resultados: [], testemunhas: [],
          ocorrencias: [], candidatos: [], correcoes: [], fotos: []
        };
      });
  }

  // Foto URL (coluna 11)
  if (wsLocais && wsLocais.getLastRow() > 1) {
    wsLocais.getRange(2, 1, wsLocais.getLastRow() - 1, 11).getValues()
      .forEach(function(r) {
        if (escolas[r[0]]) escolas[r[0]].foto_url = r[10] || "";
      });
  }

  var wsHor = ss.getSheetByName("HORÁRIOS");
  if (wsHor && wsHor.getLastRow() > 1) {
    wsHor.getRange(2, 1, wsHor.getLastRow() - 1, 11).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].horarios.push({
        turno: r[1],
        chegada: horaStr(r[2]), ab_portao: horaStr(r[3]),
        fch_portao: horaStr(r[4]), distrib: horaStr(r[5]),
        inicio: horaStr(r[6]), encerr: horaStr(r[7]),
        fch_malote: horaStr(r[8]), ab_malote: horaStr(r[9])
      });
    });
  }

  var wsRes = ss.getSheetByName("RESULTADOS");
  if (wsRes && wsRes.getLastRow() > 1) {
    wsRes.getRange(2, 1, wsRes.getLastRow() - 1, 9).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      if (!r[3]) return;
      escolas[r[0]].resultados.push({
        sala: r[1] || "", turno: r[2] || "",
        cargo: String(r[3]),
        inscr: r[4], pres: r[5], aus: r[6], incl: r[7]
      });
    });
  }

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

  var wsOcorr = ss.getSheetByName("OCORRÊNCIAS");
  if (wsOcorr && wsOcorr.getLastRow() > 1) {
    wsOcorr.getRange(2, 1, wsOcorr.getLastRow() - 1, 7).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].ocorrencias.push({
        turno: r[1] || "",
        sala: r[2], cond: r[3], toque: r[4], decl: r[5]
      });
    });
  }

  var wsCand = ss.getSheetByName("CANDIDATOS");
  if (wsCand && wsCand.getLastRow() > 1) {
    wsCand.getRange(2, 1, wsCand.getLastRow() - 1, 8).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      escolas[r[0]].candidatos.push({
        turno: r[1], tipo: r[2], inscricao: r[3],
        candidato: r[4], sala: r[5], desc: r[6] || ""
      });
    });
  }

  // Fotos individuais (IDs para embutir no relatório)
  var wsFotos = ss.getSheetByName("FOTOS");
  if (wsFotos && wsFotos.getLastRow() > 1) {
    wsFotos.getRange(2, 1, wsFotos.getLastRow() - 1, 6).getValues().forEach(function(r) {
      if (!escolas[r[0]]) return;
      var fileId = r[3] ? String(r[3]).trim() : "";
      if (!fileId) return;
      escolas[r[0]].fotos.push({
        categoria: r[1] || "Foto",
        turno:     r[2] || "",
        fileId:    fileId,
        thumbUrl:  r[4] || ("https://drive.google.com/thumbnail?id=" + fileId + "&sz=w600")
      });
    });
  }

  // Extrai correcoes e ampliacao dos candidatos (subtipos especiais)
  Object.values(escolas).forEach(function(esc) {
    // Correções de dados: tipo="Correção de Dados", desc="tipoCampo → valorCorreto"
    esc.correcoes = esc.candidatos
      .filter(function(c) { return c.tipo === "Correção de Dados"; })
      .map(function(c) {
        var partes = (c.desc || "").split(" → ");
        return {
          turno: c.turno,
          inscricao: c.inscricao,
          candidato: c.candidato,
          tipo: partes[0] || "",
          correcao: partes[1] || c.desc
        };
      });
    // Remove correções e ampliações da lista de candidatos (são seções separadas)
    esc.candidatos = esc.candidatos.filter(function(c) {
      return c.tipo !== "Correção de Dados" && c.tipo !== "Ampliação de Sala";
    });
  });

  resultado.escolas = Object.values(escolas);
  return resultado;
}

// ────────────────────────────────────────────────────────────────
//  E-mail
// ────────────────────────────────────────────────────────────────
function enviarEmailRelatorio(data, linkFotos) {
  try {
    var assunto = "Relatório de Aplicação — " + data.escola + " — " + data.data;
    MailApp.sendEmail({ to: EMAIL_DESTINO, subject: assunto, htmlBody: montarHtmlEmail(data, linkFotos) });
  } catch(e) {}
}

function montarHtmlEmail(data, linkFotos) {
  var s = [];
  var CORES = ["#1F3A6E","#CF3432","#6A0DAD","#1E6B3B"];
  var header = "#393939";
  var turnosAtivos = getTurnosAtivos(data);

  s.push('<div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;background:#F2F2F2;padding:20px;">');
  s.push('<div style="background:#393939;padding:16px 24px;border-radius:8px 8px 0 0;">');
  s.push('<h1 style="color:#FFF;font-size:16px;margin:0;">IBGP — Relatório de Aplicação de Prova</h1>');
  s.push('<p style="color:#BBB;font-size:13px;margin:4px 0 0;">Enviado automaticamente pelo sistema</p>');
  s.push('</div><div style="height:4px;background:#CF3432;margin-bottom:16px;"></div>');

  // Identificação
  var memb = (data.membros_comissao||[]).join(", ")||"—";
  s.push(card("Identificação", header, [
    row2col("Escola", data.escola, "Concurso", data.concurso||""),
    row2col("Coordenador Local", data.coord_local, "Data da Prova", data.data||""),
    row2col("Houve Apoio?", data.houve_apoio, "Nome do Apoio", data.nome_apoio||"—"),
    row2col("Houve Comissão?", data.houve_comissao, "Membros", memb),
    row2col("Enviado em", data.enviado_em, "", "")
  ]));

  // Turnos
  turnosAtivos.forEach(function(nomeTurno, idx) {
    var cor = CORES[idx % CORES.length];
    var d = data.turnos[nomeTurno];
    var h = d.horarios || {};

    s.push(card("Turno " + nomeTurno + " — Horários", cor, [
      row2col("Chegada da equipe", h.chegada, "Abertura do portão", h.ab_portao),
      row2col("Abertura do malote", h.ab_malote, "Distribuição dos pacotes", h.distrib),
      row2col("Fechamento do portão", h.fch_portao, "Início da aplicação", h.inicio),
      row2col("Encerramento", h.encerr, "Fechamento do malote", h.fch_malote)
    ]));

    var t = d.testemunhas || {};
    var abm = t.ab_malote || [{},{},{}];
    var fchm = t.fch_malote || [{},{},{}];
    s.push(card("Turno " + nomeTurno + " — Testemunhas Abertura do Malote", cor, [
      row2col("Testemunha 1", (abm[0].nome||"—")+" — "+(abm[0].cpf||"—"), "Testemunha 2", (abm[1].nome||"—")+" — "+(abm[1].cpf||"—")),
      row2col("Testemunha 3", (abm[2].nome||"—")+" — "+(abm[2].cpf||"—"), "", "")
    ]));
    s.push(card("Turno " + nomeTurno + " — Testemunhas Fechamento do Malote", cor, [
      row2col("Testemunha 1", (fchm[0].nome||"—")+" — "+(fchm[0].cpf||"—"), "Testemunha 2", (fchm[1].nome||"—")+" — "+(fchm[1].cpf||"—")),
      row2col("Testemunha 3", (fchm[2].nome||"—")+" — "+(fchm[2].cpf||"—"), "", "")
    ]));

    // Estatística
    var tabela = '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;">';
    tabela += '<tr style="background:'+cor+';color:#FFF;"><th style="padding:7px 8px;text-align:left;">Sala — Cargo</th><th style="padding:7px 8px;">Inscritos</th><th style="padding:7px 8px;">Presentes</th><th style="padding:7px 8px;">Ausentes</th><th style="padding:7px 8px;">Inclusões</th></tr>';
    var totI=0,totP=0,totA=0,totIn=0;
    (d.estatistica||[]).forEach(function(r,i){
      totI+=Number(r.inscr)||0; totP+=Number(r.pres)||0; totA+=Number(r.aus)||0; totIn+=Number(r.incl)||0;
      tabela+='<tr style="background:'+(i%2===0?'#FFF':'#F4F4F4')+'">';
      tabela+='<td style="padding:6px 8px;">'+(r.sala||"")+(r.sala?" — ":"")+r.cargo.replace(/^01\s+/,'')+'</td>';
      tabela+='<td style="text-align:center;padding:6px 8px;">'+r.inscr+'</td>';
      tabela+='<td style="text-align:center;padding:6px 8px;">'+r.pres+'</td>';
      tabela+='<td style="text-align:center;padding:6px 8px;">'+r.aus+'</td>';
      tabela+='<td style="text-align:center;padding:6px 8px;">'+(r.incl||0)+'</td></tr>';
    });
    tabela+='<tr style="background:#EEE;font-weight:700;"><td style="padding:7px 8px;">TOTAL</td>';
    tabela+='<td style="text-align:center;padding:7px 8px;">'+totI+'</td><td style="text-align:center;padding:7px 8px;">'+totP+'</td>';
    tabela+='<td style="text-align:center;padding:7px 8px;">'+totA+'</td><td style="text-align:center;padding:7px 8px;">'+totIn+'</td></tr></table>';
    s.push(card("Turno " + nomeTurno + " — Estatística por Cargo", cor, [tabela], true));

    // Ampliação
    if (d.ampliacao && d.ampliacao.length > 0) {
      var ampHtml = d.ampliacao.map(function(a){
        return '<div style="margin-bottom:8px;"><strong>'+a.ata+'</strong> — '+a.qtd+' candidatos<br><span style="font-size:12px;color:#555;">'+a.justificativa+'</span></div>';
      }).join('');
      s.push(card("Turno " + nomeTurno + " — Ampliação de Sala", cor, [ampHtml], true));
    }

    // Ocorrências
    var o = d.ocorrencias || {};
    var linhasOcorr = [
      row2col("Ocorrências em sala", (o.sala&&o.sala.resp)||"Não", "Prova condicional", (o.cond&&o.cond.resp)||"Não"),
      row2col("Toque de celular", (o.toque&&o.toque.resp)||"Não", "Declaração de comparecimento", (o.decl&&o.decl.resp)||"Não")
    ];
    ["cond","toque","decl"].forEach(function(tipo){
      var cands = (o[tipo]&&o[tipo].candidatos)||[];
      if (cands.length > 0) {
        var lista = cands.map(function(c){
          return c.candidato+(c.inscricao?" ("+c.inscricao+")":"")+(c.sala?" — Sala "+c.sala:"");
        }).join("<br>");
        linhasOcorr.push('<div style="padding:6px 0;font-size:13px;color:#555;">'+lista+'</div>');
      }
    });
    if (o.sala && o.sala.tipos && o.sala.tipos.length > 0) {
      var tiposHtml = o.sala.tipos.map(function(t2){
        return '<li>'+t2.tipo+(t2.candidato?" — "+t2.candidato:"")+(t2.sala?" (Sala "+t2.sala+")":"")+'</li>';
      }).join('');
      linhasOcorr.push('<ul style="margin:4px 0 0 16px;font-size:13px;color:#555;">'+tiposHtml+'</ul>');
    }
    s.push(card("Turno " + nomeTurno + " — Ocorrências", cor, linhasOcorr, true));

    // Correção de dados
    if (d.correcao && d.correcao.length > 0) {
      var corrHtml = '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
      corrHtml += '<tr style="background:'+cor+';color:#FFF;"><th style="padding:6px;">Inscrição</th><th style="padding:6px;">Candidato</th><th style="padding:6px;">Tipo</th><th style="padding:6px;">Dado Correto</th></tr>';
      d.correcao.forEach(function(r,i){
        corrHtml += '<tr style="background:'+(i%2===0?'#FFF':'#F4F4F4')+'"><td style="padding:5px 8px;">'+r.inscricao+'</td><td style="padding:5px 8px;">'+r.candidato+'</td><td style="padding:5px 8px;">'+r.tipo+'</td><td style="padding:5px 8px;">'+r.correcao+'</td></tr>';
      });
      corrHtml += '</table>';
      s.push(card("Turno " + nomeTurno + " — Correção de Dados Cadastrais", cor, [corrHtml], true));
    }
  });

  if (data.obs) {
    s.push(card("Observações Gerais", header, ['<p style="font-size:13px;color:#393939;margin:0;">'+data.obs+'</p>'], true));
  }
  if (linkFotos) {
    s.push(card("Registro Fotográfico", header, ['<p style="font-size:13px;"><a href="'+linkFotos+'" style="color:#CF3432;">Ver fotos no Google Drive →</a></p>'], true));
  }
  s.push('</div>');
  return s.join('');
}

function card(titulo, cor, linhas, raw) {
  var html = '<div style="background:#FFF;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.07);margin-bottom:16px;overflow:hidden;">';
  html += '<div style="background:'+cor+';padding:10px 16px;"><h3 style="color:#FFF;font-size:13px;margin:0;text-transform:uppercase;letter-spacing:.4px;">'+titulo+'</h3></div>';
  html += '<div style="padding:16px;">';
  linhas.forEach(function(l){ html += l; });
  html += '</div></div>';
  return html;
}

function row2col(l1, v1, l2, v2) {
  var cell = function(l, v) {
    if (!l) return '<td style="width:50%;"></td>';
    return '<td style="width:50%;padding:4px 8px 4px 0;vertical-align:top;">'
      +'<span style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.4px;">'+l+'</span>'
      +'<br><span style="font-size:13px;color:#393939;">'+(v||'—')+'</span></td>';
  };
  return '<table style="width:100%;margin-bottom:10px;"><tr>'+cell(l1,v1)+cell(l2,v2)+'</tr></table>';
}

// ════════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, nome, cabecalho) {
  var ws = ss.getSheetByName(nome);
  if (!ws) { ws = ss.insertSheet(nome); ws.appendRow(cabecalho); }
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
//  Abas de dados — versão dinâmica (data.turnos)
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
    "Local / Escola", "Coordenador Local", "Nº Salas", "Turno(s)",
    "Houve apoio?", "Nome do apoio",
    "Houve comissão?", "Membros da comissão",
    "Observações gerais", "Enviado em", "Foto URL"
  ]);
  var turnosAtivos = getTurnosAtivos(data);
  var memb = (data.membros_comissao || []).map(function(m) {
    return typeof m === "object" && m !== null ? (m.nome || "") : String(m);
  }).filter(Boolean).join("; ");
  atualizarOuInserir(ws, data.escola, [
    data.escola, data.coord_local, data.salas, turnosAtivos.join(" e "),
    data.houve_apoio, data.nome_apoio,
    data.houve_comissao, memb,
    data.obs, data.enviado_em, ""
  ]);
}

function salvarResultados(ss, data) {
  var ws = getOrCreateSheet(ss, "RESULTADOS", [
    "Local / Escola", "Sala", "Turno", "Cargo",
    "Inscritos", "Presentes", "Ausentes", "Inclusões", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);
  getTurnosAtivos(data).forEach(function(nomeTurno) {
    var d = data.turnos[nomeTurno];
    (d.estatistica||[]).forEach(function(r) {
      ws.appendRow([
        data.escola,
        r.sala || "",
        nomeTurno,
        r.cargo,
        r.inscr, r.pres, r.aus, r.incl || 0,
        data.enviado_em
      ]);
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
  getTurnosAtivos(data).forEach(function(nomeTurno) {
    var h = (data.turnos[nomeTurno].horarios) || {};
    ws.appendRow([
      data.escola, nomeTurno,
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
  getTurnosAtivos(data).forEach(function(nomeTurno) {
    var t = (data.turnos[nomeTurno].testemunhas) || {};
    var abm = t.ab_malote  || [{nome:"",cpf:""},{nome:"",cpf:""},{nome:"",cpf:""}];
    var fch = t.fch_malote || [{nome:"",cpf:""},{nome:"",cpf:""},{nome:"",cpf:""}];
    ws.appendRow([
      data.escola, nomeTurno, "Abertura do Malote",
      abm[0].nome, abm[0].cpf, abm[1].nome, abm[1].cpf, abm[2].nome, abm[2].cpf,
      data.enviado_em
    ]);
    ws.appendRow([
      data.escola, nomeTurno, "Fechamento do Malote",
      fch[0].nome, fch[0].cpf, fch[1].nome, fch[1].cpf, fch[2].nome, fch[2].cpf,
      data.enviado_em
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
  getTurnosAtivos(data).forEach(function(nomeTurno) {
    var o = (data.turnos[nomeTurno].ocorrencias) || {};
    ws.appendRow([
      data.escola, nomeTurno,
      (o.sala && o.sala.resp) || "Não",
      (o.cond && o.cond.resp) || "Não",
      (o.toque && o.toque.resp) || "Não",
      (o.decl && o.decl.resp) || "Não",
      data.enviado_em
    ]);
  });
}

function salvarCandidatos(ss, data) {
  var ws = getOrCreateSheet(ss, "CANDIDATOS", [
    "Local / Escola", "Turno", "Tipo", "Nº Inscrição", "Candidato", "Sala", "Descrição", "Enviado em"
  ]);
  deletarLinhasEscola(ws, data.escola);

  getTurnosAtivos(data).forEach(function(nomeTurno) {
    var o = (data.turnos[nomeTurno].ocorrencias) || {};
    // Ocorrências em sala — tipos detalhados
    ((o.sala && o.sala.tipos) || []).forEach(function(t) {
      ws.appendRow([
        data.escola, nomeTurno, "Ocorrências em sala",
        t.inscricao || "", t.candidato || "", t.sala || "",
        t.tipo + (t.candidato ? " — " + t.candidato : "") + (t.sala ? " (Sala " + t.sala + ")" : ""), data.enviado_em
      ]);
    });
    // Prova condicional, Toque, Declaração — candidatos
    [
      {key:"cond",  label:"Prova condicional"},
      {key:"toque", label:"Toque de celular"},
      {key:"decl",  label:"Declaração de comparecimento"},
    ].forEach(function(tipo) {
      var cands = (o[tipo.key] && o[tipo.key].candidatos) || [];
      cands.forEach(function(c) {
        ws.appendRow([
          data.escola, nomeTurno, tipo.label,
          c.inscricao || "", c.candidato || "", c.sala || "", c.desc || "",
          data.enviado_em
        ]);
      });
    });
    // Correção de dados
    (data.turnos[nomeTurno].correcao || []).forEach(function(c) {
      ws.appendRow([
        data.escola, nomeTurno, "Correção de Dados",
        c.inscricao || "", c.candidato || "", "", c.tipo+" → "+c.correcao,
        data.enviado_em
      ]);
    });
    // Ampliação de sala
    (data.turnos[nomeTurno].ampliacao || []).forEach(function(a) {
      ws.appendRow([
        data.escola, nomeTurno, "Ampliação de Sala",
        "", "", "", a.ata+" — "+a.qtd+" cand. — "+a.justificativa,
        data.enviado_em
      ]);
    });
  });
}

function salvarFotos(data) {
  if (!data.fotos || data.fotos.length === 0) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var root = DriveApp.getRootFolder();
  var ibgpIt = root.getFoldersByName("IBGP_RELATORIOS");
  var ibgpFolder = ibgpIt.hasNext() ? ibgpIt.next() : root.createFolder("IBGP_RELATORIOS");

  var nomeEscola = data.escola.replace(/[\/\\:*?"<>|]/g, "_");
  var escolaIt = ibgpFolder.getFoldersByName(nomeEscola);
  var escolaFolder = escolaIt.hasNext() ? escolaIt.next() : ibgpFolder.createFolder(nomeEscola);

  // Aba FOTOS — salva ID individual de cada foto para embutir no relatório
  var wsFotos = getOrCreateSheet(ss, "FOTOS", [
    "Local / Escola", "Categoria", "Turno", "File ID", "URL Thumbnail", "Enviado em"
  ]);
  deletarLinhasEscola(wsFotos, data.escola);

  data.fotos.forEach(function(foto, idx) {
    try {
      var bytes = Utilities.base64Decode(foto.b64);
      var cat   = (foto.categoria || "Foto").replace(/[\/\\:*?"<>|]/g, "_");
      var nome  = cat + "_" + (idx + 1) + "_" + (foto.nome || "foto.jpg");
      var file  = escolaFolder.createFile(Utilities.newBlob(bytes, "image/jpeg", nome));

      // Torna cada arquivo individualmente público
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      var fileId    = file.getId();
      var thumbUrl  = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w600";

      wsFotos.appendRow([
        data.escola,
        foto.categoria || "Foto",
        foto.turno || "",
        fileId,
        thumbUrl,
        data.enviado_em
      ]);
    } catch(e) {}
  });

  escolaFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = "https://drive.google.com/drive/folders/" + escolaFolder.getId();

  // Salva link da pasta na aba LOCAIS (coluna 11)
  try {
    var wsL = ss.getSheetByName("LOCAIS");
    if (wsL && wsL.getLastRow() > 1) {
      var nomes = wsL.getRange(2, 1, wsL.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < nomes.length; i++) {
        if (String(nomes[i][0]).trim() === String(data.escola).trim()) {
          wsL.getRange(i + 2, 11).setValue(link);
          break;
        }
      }
    }
  } catch(e) {}

  return link;
}
