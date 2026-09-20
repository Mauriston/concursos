# Documentação — Sistema de IS de Ingresso (JRS/HNRe)

Documentação funcional e técnica do Apps Script vinculado à planilha
"TEMPLATE CONCURSOS", que automatiza a gestão da Inspeção de Saúde (IS)
de Ingresso de um concurso da Marinha do Brasil, da mensagem
administrativa inicial (SIGAD-MB) até a minuta da mensagem final com os
resultados.

- Planilha: `TEMPLATE CONCURSOS` (Google Sheets)
- Projeto Apps Script: `TEMPLATE CONCURSOS`
- Arquivos do projeto: `Código.gs`, `appsscript.json`, `Modal.html`,
  `Alerta.html`, `UploadMensagem.html`, `AgendamentoIS.html`

---

## 1. Visão geral do fluxo

```
1. Upload da mensagem administrativa (PDF, SIGAD-MB)
   → extrai candidatos, período de agendamento e dados da mensagem
   → grava em "candidatos", "candidatosDataBase", "mensagens" e "agendamentos"

2. Configuração do agendamento das IS
   → usuário escolhe quantidade de IS/dia e dias da semana
   → script distribui os candidatos pendentes pelas datas disponíveis
   → grava em "candidatosDataBase" (dataAgendamento) e na tabela "principal" (aba Principal)
   → gera a minuta da MENSAGEM DE AGENDAMENTO DA IS

3. Condução das IS (uso corrente da tabela Principal)
   → a tabela "principal" funciona como CRUD front-end de "candidatosDataBase"
   → usuário edita Status, Observações e Nº TIS por candidato
   → script sincroniza automaticamente com "candidatosDataBase"
   → Status = INAPTO dispara confirmação e oferece gerar o Termo de
     Cientificação de Recurso individual

4. Mensagem final de resultados
   → script bloqueia a geração se houver candidato não finalizado
   → agrupa os candidatos de "candidatosDataBase" por status
   → gera a minuta da MENSAGEM DE RESULTADOS DA IS
```

---

## 2. Estrutura das abas da planilha

### 2.1 `mensagens`
Registro de cada mensagem administrativa (SIGAD-MB) processada.

| Coluna | Conteúdo |
|---|---|
| A `dataHora` | ID único da mensagem (ex.: `P141721Z/JUL/2026`) |
| B `fileUrl` | Link do PDF salvo no Drive (pasta "Mensagens") |
| C `proposito` | `"Apresentação e IS"` ou `"Outros"` (detectado automaticamente) |
| D `remetente` | Campo "De:" do cabeçalho SIGAD |
| E `destinatario` | Campo "Para:" |
| F `informacao` | Campo "Info:" |
| G `assunto` | Campo "Assunto:" (usado para extrair o nome do concurso, ex.: `CPAEAM/2026`) |
| H `corpoTexto` | Corpo do campo "Texto:" da mensagem, já normalizado |

Novas mensagens são inseridas sempre na linha 2 (mais recente no topo).
A `dataHora` da mensagem cujo `proposito = "Apresentação e IS"` é usada
como identificador da mensagem inicial em ambas as minutas geradas.

### 2.2 `candidatos`
Lista mestra de candidatos do concurso, ordenada alfabeticamente pelo nome.

| Coluna | Conteúdo |
|---|---|
| A `id` | Matrícula do candidato (ex.: `108842-0`) |
| B `candidato` | Nome completo |

### 2.3 `candidatosDataBase`
Base de dados operacional de cada candidato — é a fonte de verdade do
sistema. A tabela `principal` (aba Principal) funciona como um
front-end de edição para parte destas colunas.

| Col. | Nome | Tipo | Preenchida por |
|---|---|---|---|
| A | `id` | texto | matrícula, espelha `candidatos` |
| B | `dataAgendamento` | data | agendamento da IS |
| C | `reagendamento` | — | uso manual (não gerenciada pelo script) |
| D | `status` | texto | `APTO` / `INAPTO` / `FALTOU` / `INSUF DOCUMENTAL` / `Pendente` / vazio |
| E | `observacoes` | texto | espelho da coluna "Observações" da tabela Principal |
| F | `finalizado` | ☑️ booleano | `TRUE` quando um status com laudo é selecionado |
| G | `recurso` | ☑️ booleano | `TRUE` quando o Termo de Recurso é gerado |
| H | `dataLaudo` | data | data do laudo (hoje, no momento em que o Status é definido) |
| I | `Laudo` | texto | texto do laudo correspondente ao Status (ver §4.3) |
| J | `nº TIS` | texto | espelho da coluna "Nº TIS" da tabela Principal |
| K | `termoRecursoUrl` | texto (URL) | link do PDF do Termo de Recurso gerado |

### 2.4 `agendamentos`
Calendário de datas úteis disponíveis para agendamento das IS.

| Coluna | Conteúdo |
|---|---|
| A `agendamentos` (data) | Data útil (seg–sex, exclui feriados nacionais) |
| B `diaDaSemana` | Nome do dia (Segunda...Sexta) |
| C `ativa` | ☑️ `TRUE` se o dia da semana foi selecionado na última configuração de agendamento |

Gerada automaticamente a partir do período "JRS" extraído da mensagem
inicial (ex.: `03AGO a 14SET2026 (JRS)`), preservando `ativa` de datas
já existentes ao reprocessar.

### 2.5 `Principal` — tabela estruturada `principal`
Tela de trabalho do dia a dia da JRS. É uma **tabela estruturada** do
Google Sheets (não uma faixa fixa), com o cabeçalho localizado
dinamicamente pelo script (procura as colunas "Matricula" e
"Candidato").

| Coluna | Conteúdo | Sincroniza com `candidatosDataBase` |
|---|---|---|
| `Data` | Data do dia (mesclada) + dia da semana, preenchida automaticamente ao confirmar um agendamento | — |
| `dataAgendamento` | Data "crua" do agendamento (mesmo valor usado internamente) | ← gravada na confirmação do agendamento |
| `Matricula` | ID do candidato — **chave de correspondência** com `candidatosDataBase`/`candidatos` | — |
| `Candidato` | Nome do candidato | — |
| `Status` | Lista suspensa: `APTO`, `INAPTO`, `FALTOU`, `INSUF DOCUMENTAL`, `Pendente` | ✅ editável → sincroniza |
| `Observações` | Texto livre | ✅ editável → sincroniza |
| `Nº TIS` | Texto livre | ✅ editável → sincroniza |
| `Data laudo` | Preenchida automaticamente quando um Status com laudo é selecionado | somente leitura (espelho de `dataLaudo`) |

Cada bloco de linhas de uma mesma data de agendamento recebe fundo
alternado (branco/cinza-claro) e bordas de destaque na primeira e
última linha do bloco.

### 2.6 `Listas por Conclusões` (legado)
Aba usada pelo fluxo antigo, anterior à tabela `principal` dinâmica e a
`candidatosDataBase`. Ainda é referenciada por `processarGeracaoRecursos`
(ver §6 — Itens legados).

---

## 3. Arquivos do projeto Apps Script

| Arquivo | Papel |
|---|---|
| `Código.gs` | Toda a lógica do sistema |
| `appsscript.json` | Manifesto: fuso `America/Recife`, serviços avançados Drive v3, Docs v1, Sheets v4 |
| `UploadMensagem.html` | Modal de upload do PDF da mensagem administrativa |
| `AgendamentoIS.html` | Modal de configuração e confirmação do agendamento das IS |
| `Modal.html` | Modal genérico de exibição de minuta (usado tanto para a minuta de agendamento quanto para a minuta de resultados), com botão "Copiar Minuta" |
| `Alerta.html` | Modal genérico de alerta/confirmação (usado em vários fluxos: confirmação inicial, geração de recursos em lote, alerta final) |

---

## 4. Fluxo detalhado

### 4.1 Upload e extração da mensagem inicial
Menu **✏️ Termos → 📄 Registrar Mensagem (PDF)**.

1. `iniciarUploadMensagem()` abre `UploadMensagem.html`.
2. O PDF selecionado é enviado em base64 para `processarMensagemPDFUpload()`,
   que salva o arquivo na subpasta "Mensagens" (mesma pasta da planilha)
   e chama `processarArquivoMensagem()`.
3. `processarArquivoMensagem()`:
   - Converte o PDF para Google Docs via serviço avançado Drive (com
     OCR em português) só para extrair o texto (`extrairTextoPdf`), e
     descarta a cópia temporária.
   - Remove ruído de paginação (marca d'água "HNRe - 02.2", "Página X
     de Y") com `limparRuidoPaginacao()`.
   - Extrai o cabeçalho SIGAD-MB (Data-Hora, De, Para, Info, Assunto,
     Texto) com `extrairCabecalhoMensagem()`.
   - Extrai a lista de candidatos (`extrairCandidatos()`), tolerante ao
     layout em colunas do PDF (corta cada item no próximo código de
     matrícula `000000-0`, não na quebra de linha).
   - Reconstrói o bloco de candidatos no texto da mensagem já
     formatado (`normalizarTextoComCandidatos()`), para ficar
     consistente com o padrão de pontuação do resto do sistema.
   - Grava os candidatos novos em `candidatos` (`gravarExaminee()`,
     reordena tudo alfabeticamente) e sincroniza `candidatosDataBase`
     na mesma ordem (`gravarExamineeDataBase()`), preservando dados já
     existentes de candidatos repetidos.
   - Registra a mensagem em `mensagens` (`gravarMensagem()`).
   - Extrai o período "JRS" do texto (`extrairPeriodoJRS()`, padrão
     `03AGO a 14SET2026 (JRS)`), calcula os dias úteis excluindo
     feriados nacionais fixos e móveis (`calcularDiasUteis()`,
     `obterFeriadosNacionais()`, cálculo da Páscoa por
     Meeus/Jones/Butcher) e grava em `agendamentos`
     (`gravarDatasAgendamento()`).
4. Ao final, oferece ir direto para o agendamento
   (`abrirModalAgendamento()`).

### 4.2 Configuração do agendamento das IS
Acionado a partir do resultado do upload, ou reaberto manualmente
chamando `abrirModalAgendamento()`.

1. `abrirModalAgendamento()` lista os candidatos ainda sem
   `dataAgendamento` em `candidatosDataBase`
   (`listarCandidatosPendentesAgendamento()`) e monta o resumo (total
   pendente, período, dias úteis disponíveis) para `AgendamentoIS.html`.
2. O usuário informa quantas IS por dia e quais dias da semana usar.
   `verificarViabilidadeAgendamento()`:
   - marca `ativa=TRUE/FALSE` em `agendamentos` conforme os dias da
     semana escolhidos (`ativarDatasPorDiaSemana()`);
   - calcula se a capacidade (`datas × qtde/dia`) é suficiente; se não
     for, sugere aumentar a quantidade por dia ou os dias da semana;
   - se for viável, distribui os candidatos pendentes pelas datas
     disponíveis, em ordem cronológica, preenchendo cada data até o
     limite antes de passar para a próxima
     (`distribuirCandidatosNasDatas()`).
3. Ao confirmar (`confirmarAgendamentos()`):
   - grava a data de cada candidato em `candidatosDataBase.dataAgendamento`
     (`gravarDatasAgendamentoCandidatos()`);
   - anexa os candidatos à tabela `principal` (aba Principal), sem
     sobrescrever confirmações anteriores
     (`preencherAbaPrincipal()` → `proximaLinhaVaziaTabelaPrincipal()` →
     `garantirCapacidadeTabelaPrincipal()`, que expande a tabela
     estruturada via Sheets API v4 quando faltam linhas);
   - formata a coluna "Data" (mescla data + dia da semana, ou célula
     única quando há 1 candidato no dia) e pinta o bloco de linhas de
     cada dia com fundo alternado e bordas de destaque
     (`aplicarFormatacaoColunaData()`);
   - gera e exibe a minuta da **MENSAGEM DE AGENDAMENTO DA IS**
     (`gerarTextoMinutaAgendamento()`) no `Modal.html`, com o título
     "MENSAGEM DE AGENDAMENTOS DA IS - Minuta gerada". Ao fechar esse
     modal **não** aparece o alerta de "Processo Concluído" (é
     exclusivo do fluxo de conclusão, §4.4).

Formato da minuta de agendamento:
```
{Data-Hora da mensagem inicial}, PTC:

ALFA - As IS de Ingresso dos Candidatos a {concurso} estão agendadas conforme:

UNO - {data} às 7h30:
- {matricula} {nome};
...
- {matricula} {nome}; e
- {matricula} {nome}.

DOIS - {próxima data} às 7h30:
...

BRAVO - CFM o item 3.1.2 da DGPM-406 (9ª Revisão), Os candidatos que não
comparecerem... terão suas IS concluídas e assinadas tempestivamente com
laudos, respectivamente, de "faltou" ou "Insuficiência Documental Médica" BT
```

### 4.3 Sincronização Principal ↔ candidatosDataBase (uso corrente)
A tabela Principal passa a ser a tela de trabalho: o usuário edita
Status, Observações e Nº TIS diretamente nela, e o script propaga para
`candidatosDataBase`.

**Ativação (uma única vez por planilha):** menu **✏️ Termos → ⚙️
Ativar automação da tabela Principal** (`instalarGatilhoOnEditPrincipal()`).
Isso é necessário porque a sincronização precisa exibir alertas
(`SpreadsheetApp.getUi()`), o que um gatilho `onEdit` **simples** não
tem permissão para fazer — só um gatilho **instalável**, criado com
autorização plena do usuário via um item de menu. `onOpen()` (gatilho
simples) não pode criar esse gatilho sozinho.

O gatilho instalável (`aoEditarPrincipalInstalavel()`) só age sobre
edições de **uma única célula**, dentro das linhas de dados da tabela
`principal`, localizando o candidato pela `Matricula` (coluna que casa
com `candidatosDataBase.id`):

- **Nº TIS** (`sincronizarNumTISPrincipal`) → grava/limpa
  `candidatosDataBase.nº TIS` (coluna J).
- **Observações** (`sincronizarObservacoesPrincipal`) → grava/limpa
  `candidatosDataBase.observacoes` (coluna E).
- **Status** (`processarEdicaoStatusPrincipal`):
  - célula limpa → limpa `status`, `finalizado` (desmarca), `dataLaudo`
    e `Laudo` em `candidatosDataBase`, e "Data laudo" na Principal;
  - `Pendente` → grava `status = "Pendente"`, mas **não** marca
    `finalizado` nem grava `Laudo`/`dataLaudo` (mesmo comportamento de
    limpeza dos demais campos que a célula vazia);
  - `APTO` / `FALTOU` / `INSUF DOCUMENTAL` → grava `status`, marca
    `finalizado = TRUE`, grava `dataLaudo` = hoje e `Laudo` conforme
    o mapa abaixo, e replica a data em "Data laudo" na Principal;
  - `INAPTO` → antes de gravar, pergunta
    *"Registrar {Candidato} como inapto hoje ({data})?"* (Sim/Não). Se
    "Não", a célula volta ao valor anterior (`e.oldValue`) e nada é
    gravado. Se "Sim", grava normalmente (como acima) e, em seguida,
    pergunta *"Gerar o termo de Cientificação de Recurso para
    {Candidato}?"*; se "Sim", chama `gerarTermoRecursoIndividual(id)`.

Mapa Status → Laudo (`MAPA_LAUDO_POR_STATUS`):

| Status | Laudo |
|---|---|
| `APTO` | Apto para Ingresso |
| `INAPTO` | Inapto para Ingresso |
| `FALTOU` | IS não concluída por não comparecimento |
| `INSUF DOCUMENTAL` | IS não concluída por Insuficiência Documental Médica |

### 4.4 Termo de Cientificação de Recurso — individual
`gerarTermoRecursoIndividual(id)` (disparado pelo fluxo INAPTO acima):

1. Busca o nome em `candidatos` e a `dataLaudo` em `candidatosDataBase`
   pelo `id`.
2. Reaproveita um PDF já existente na subpasta (`Termo Recurso
   {candidato}.pdf`, dentro da pasta fixa de Termos), ou gera um novo a
   partir do template do Google Docs (substitui `{{Candidato}}`,
   `{{Data Laudo}}`, `{{DATA_HOJE}}`) e exporta como PDF.
3. Marca `candidatosDataBase.recurso = TRUE` e grava a URL do PDF em
   `termoRecursoUrl`.
4. Mostra um alerta com o link do termo gerado.

### 4.5 Mensagem final de resultados da IS
Menu **✏️ Termos → 🛑 Cientificação de Recurso** aciona
`iniciarGeracaoRecursos()`, que pergunta se deseja "verificar os
candidatos com interposição de recurso e gerar os Termos individuais" —
**este é o fluxo legado em lote**, ver §6. A minuta de resultados
propriamente dita é gerada por `abrirModal()`, chamada a partir do
alerta inicial de confirmação de dados (`iniciarProcesso()` → "Sim,
gerar minuta").

`abrirModal()`:

1. **Bloqueio de segurança**: `listarCandidatosNaoFinalizados()`
   verifica se todo candidato em `candidatosDataBase` tem
   `finalizado = TRUE`. Se houver algum pendente, mostra um alerta
   listando quem falta finalizar e **interrompe** — a minuta só é
   gerada com todos finalizados.
2. Busca a Data-Hora da mensagem inicial em `mensagens`
   (`obterDadosMensagemInicial()`) e o nome do concurso a partir do
   Assunto (`extrairNomeConcurso()`).
3. Agrupa os candidatos de `candidatosDataBase` por `status`
   (`obterCandidatosPorStatus()`): APTO, INAPTO, FALTOU, INSUF
   DOCUMENTAL, e a lista de quem tem `recurso = TRUE` (com a data do
   `dataLaudo`).
4. Monta e exibe a minuta em `Modal.html`, título "Minuta Gerada". Ao
   fechar, **aparece** o alerta "Processo Concluído" orientando colar
   no SIGAD e tramitar para o Supervisor JRS/HNRe (`mostrarAlertaFinal()`)
   — diferente do fluxo de agendamento (§4.2).

Formato da minuta de resultados:
```
{Data-Hora da mensagem inicial}, PTC que JRS/HNRe concluiu em {hoje} as
IS dos {total} candidatos APS FIM Ingresso no {concurso} CFM os
resultados abaixo relacionados:

ALFA - Candidatos considerados "Aptos para Ingresso" (total: N):
BRAVO - Candidatos considerados "Inaptos para Ingresso" (total: N):
CHARLIE - Candidatos com IS não concluídas por não comparecimento (total: N):
DELTA - Candidatos com IS não concluídas por Insuficiência Documental Médica (total: N):
ECHO - Candidatos que interpuseram recurso... (total: N): BT
```

---

## 5. Automações e gatilhos

| Gatilho | Tipo | Escopo | Função |
|---|---|---|---|
| Abertura da planilha | simples (`onOpen`) | global | Cria o menu "✏️ Termos" |
| Edição de célula | simples (`onEdit`) | aba "Principal", coluna 8, linha ≥ 11 | Legado — grava a data de hoje quando um valor é digitado (ver §6) |
| Edição de célula | **instalável** (`aoEditarPrincipalInstalavel`, precisa ser ativado 1x pelo menu) | tabela `principal` (aba Principal), colunas Status/Observações/Nº TIS | Sincroniza com `candidatosDataBase`, dispara confirmações e o Termo de Recurso individual |

### Itens de menu (✏️ Termos)
- 🛑 Cientificação de Recurso → geração de Termos em lote (legado, §6)
- 📄 Registrar Mensagem (PDF) → início do fluxo (§4.1)
- ⚙️ Ativar automação da tabela Principal → instala o gatilho de
  sincronização (uma vez só, §4.3)

---

## 6. Itens legados / atenção

- **`onEdit(e)` simples** (linha ~18): grava a data de hoje quando a
  coluna 8 é editada em linha ≥ 11 da aba "Principal". Essa referência
  por índice fixo de coluna/linha é anterior à tabela estruturada
  `principal` (que tem cabeçalho e colunas localizados
  dinamicamente); pode não corresponder mais ao layout atual da
  planilha. Vale revisar/remover se não estiver mais em uso.
- **`processarGeracaoRecursos()`** (geração de Termos de Recurso **em
  lote**, acionada pelo menu "🛑 Cientificação de Recurso"): lê a aba
  "Principal" em posições fixas (`A12:K`, coluna A = Data, D =
  Candidato, **G = Recurso** com valor literal "Sim", J = Data Laudo).
  A tabela `principal` atual **não tem** uma coluna "Recurso" nessa
  posição (o layout real é `Data, dataAgendamento, Matricula,
  Candidato, Status, Observações, Nº TIS, Data laudo`) — esse fluxo em
  lote está desalinhado com a estrutura atual e provavelmente não
  encontra candidatos para gerar. O fluxo válido e testado atualmente
  é o **individual**, disparado ao selecionar Status = INAPTO na
  Principal (§4.4, `gerarTermoRecursoIndividual`). Recomenda-se decidir
  entre atualizar `processarGeracaoRecursos()` para a estrutura atual
  ou removê-lo do menu.
- **Aba "Listas por Conclusões"**: usada apenas por
  `processarGeracaoRecursos()` (legado acima). A minuta de resultados
  (§4.5) não depende mais dela desde a migração para
  `candidatosDataBase`/`candidatos`.
- **Coluna `reagendamento`** de `candidatosDataBase`: existe na
  estrutura mas não é lida nem escrita por nenhuma função atual do
  script (uso manual).

---

## 7. Convenções auxiliares usadas nas minutas

- **Pontuação de listas** (`aplicarPontuacao`): item do meio termina
  em `;`, penúltimo item em `; e`, último item em `.` (ou sem
  pontuação, no modo "eco", usado na lista de recursos que já termina
  em `BT`).
- **Numeração de itens** (`numeroItemLista` / `numeroCardinalExtenso`):
  por extenso em maiúsculas, com `UNO` para o item 1 (convenção naval,
  evita ambiguidade com o artigo "um").
- **Datas**:
  - `formatarDataSimples`: `DD/MM/AAAA`.
  - `formatarDataMilitar` / `formatarDataDDMMMAAAA`: `DDMMMAAAA` (ex.:
    `14AGO2026`).
  - Datas de agendamento são sempre construídas ao meio-dia (`12:00`)
    em vez de meia-noite, para evitar que o deslocamento de fuso entre
    o runtime V8 (UTC) e o fuso da planilha (`America/Recife`) jogue a
    data para o dia anterior.
- **Feriados**: fixos (Confraternização Universal, Tiradentes, Dia do
  Trabalho, Independência, N. Sra. Aparecida, Finados, Proclamação da
  República, Consciência Negra, Natal) + móveis calculados a partir da
  Páscoa (Carnaval, Quarta de Cinzas, Sexta-feira Santa, Corpus
  Christi).
