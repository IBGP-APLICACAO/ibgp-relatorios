# Como configurar o sistema (passo a passo)

## ETAPA 1 — Ativar o site no GitHub Pages

1. No repositório do GitHub, clique em **Configurações**
2. No menu lateral, clique em **Pages**
3. Em "Origem", selecione **Implantar de um branch**
4. Selecione o branch **main** e a pasta **/ (root)**
5. Clique em **Salvar**

Após alguns minutos, o site estará disponível em:
`https://ibgp-aplicacao.github.io/ibgp-relatorios/`

---

## ETAPA 2 — Criar a planilha Google Sheets

1. Acesse **sheets.google.com** e crie uma planilha em branco
2. Dê um nome a ela, ex: `IBGP — Coleta Alto Rio Doce`
3. Deixe a planilha aberta (usaremos ela na próxima etapa)

---

## ETAPA 3 — Configurar o Google Apps Script

1. Na planilha do Google Sheets, clique em **Extensões → Apps Script**
2. Apague o código que aparecer e cole o conteúdo do arquivo `apps-script.gs`
3. Clique em **Salvar** (ícone de disquete)
4. Clique em **Implantar → Nova implantação**
5. Em "Tipo", selecione **Aplicativo da Web**
6. Configure assim:
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
7. Clique em **Implantar**
8. Autorize o acesso quando solicitado (clique em "Avançado" → "Acessar ibgp...")
9. **Copie a URL** que aparecer (começa com `https://script.google.com/macros/s/...`)

---

## ETAPA 4 — Conectar o site ao Apps Script

1. Abra o arquivo `index.html` (pode fazer pelo GitHub mesmo, clicando no arquivo e depois no lápis ✏️)
2. Encontre esta linha:
   ```
   const SCRIPT_URL = "COLE_AQUI_A_URL_DO_APPS_SCRIPT";
   ```
3. Substitua `COLE_AQUI_A_URL_DO_APPS_SCRIPT` pela URL copiada no passo anterior
4. Clique em **Confirmar alterações**

---

## ETAPA 5 — Testar

1. Acesse o link do GitHub Pages
2. Selecione uma escola, preencha os dados e clique em **Enviar Dados**
3. Abra a planilha do Google Sheets — as abas **RESULTADOS** e **LOCAIS** devem aparecer com os dados

---

## Após todos os coordenadores enviarem

1. Abra a planilha Google Sheets
2. Clique em **Arquivo → Baixar → Microsoft Excel (.xlsx)**
3. Salve como `planilha_ibgp.xlsx` (substituindo o arquivo do projeto)
4. Também preencha a aba **CONCURSO** com os coordenadores e horários
5. Rode o gerador do relatório:
   ```
   python gerar_relatorio_geral.py planilha_ibgp.xlsx
   ```

---

## Para um novo concurso

Edite o arquivo `index.html` no GitHub e atualize o objeto `CONCURSO` no início do script com os dados do novo concurso (título, escolas, salas, cargos e inscritos previstos).
