# IBGP — Sistema de Relatórios de Aplicação de Prova

Gerador automático de formulários individuais e relatório geral consolidado para concursos públicos.

---

## Estrutura do projeto

```
ibgp-relatorios/
├── planilha_ibgp.xlsx       ← Preencher antes de cada concurso
├── gerar_formularios.py     ← Gera os PDFs individuais por escola
├── gerar_relatorio_geral.py ← Gera o relatório consolidado (PDF + Word)
├── logo/
│   └── IBGP.png             ← Logo do IBGP (não alterar)
├── saida/                   ← Arquivos gerados (não versionar)
├── requirements.txt
└── package.json
```

---

## Instalação (primeira vez)

**Pré-requisitos:** Python 3.10+ e Node.js 18+

```bash
pip install -r requirements.txt
npm install
```

---

## Fluxo de uso

### ETAPA 1 — Antes da prova: preencher a planilha

Abra `planilha_ibgp.xlsx` e preencha as 3 abas:

| Aba | O que preencher |
|-----|-----------------|
| **CONCURSO** | Título, edital, data, coordenadores e horários |
| **LOCAIS** | Uma linha por escola: nome, coord. local, nº salas, turno, observações |
| **FORMULÁRIOS** | Mesmas escolas com dados do concurso (usada pelo gerador de PDFs) |

A aba **ALOCAÇÃO POR SALA** (opcional) pré-preenche os cargos nos formulários se disponível.

### ETAPA 2 — Gerar os formulários individuais

```bash
python gerar_formularios.py planilha_ibgp.xlsx
```

Será gerado um ZIP com um PDF por escola na mesma pasta da planilha.
Distribuir os PDFs para cada coordenador local preenchê-los no dia da prova.

### ETAPA 3 — Após a prova: lançar os resultados

Quando **todos os formulários forem devolvidos**, preencha a aba **RESULTADOS** da planilha com os dados de cada sala:

- Local / Escola
- Sala
- Turno (Manhã ou Tarde)
- Cargo
- Inscritos, Presentes, Ausentes, Eliminados

### ETAPA 4 — Gerar o Relatório Geral

```bash
python gerar_relatorio_geral.py planilha_ibgp.xlsx
```

Serão gerados dois arquivos na mesma pasta da planilha:
- `relatorio_geral_aplicacao.pdf`
- `relatorio_geral_aplicacao.docx`

---

## Logo

Coloque o arquivo `IBGP.png` dentro da pasta `logo/`.
Os scripts buscam automaticamente o logo nessa pasta.

---

## Observações

- A pasta `saida/` e os arquivos gerados (`.pdf`, `.docx`) não são versionados no Git.
- Os dados de exemplo embutidos nos scripts servem apenas para testes.
- Para gerar com dados reais, sempre passe a planilha como argumento.
