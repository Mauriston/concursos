# Documentação — Sistema de IS de Ingresso (JRS/HNRe)
## Planilha "TEMPLATE CONCURSOS" — Apps Script `Código.gs`

Documentação funcional e técnica completa do projeto Apps Script
**container-bound** (vinculado à planilha) "TEMPLATE CONCURSOS", que
automatiza a gestão da Inspeção de Saúde (IS) de Ingresso de um
concurso da Marinha do Brasil, desde a mensagem administrativa inicial
(SIGAD-MB) até a minuta da mensagem final com os resultados — incluindo
a leitura/extração de dados de PDFs, o agendamento das IS, a
sincronização em tempo real da tabela de trabalho da JRS e a geração de
Termos de Cientificação de Recurso em PDF.

> Este projeto é distinto do backend web standalone `CodeConcursos.gs`
> (repositório `DoencasEPareceresJRS`), que expõe parte destes mesmos
> dados via API HTTP para a tela "Planilhas de Controle" do app web da
> JRS. Este documento cobre exclusivamente o projeto Apps Script
> vinculado à própria planilha "TEMPLATE CONCURSOS" (uso direto dentro
> do Google Sheets, por menu).

- Planilha: `TEMPLATE CONCURSOS` (Google Sheets)
- Projeto Apps Script: `TEMPLATE CONCURSOS` (container-bound)
- Arquivos do projeto: `Código.gs`, `appsscript.json`, `Modal.html`,
  `Alerta.html`, `UploadMensagem.html`, `AgendamentoIS.html`
- Fuso horário do projeto: `America/Recife` (ver `appsscript.json`)

---

## 0. Identificadores fixos usados pelo script

O script **não** referencia a planilha por ID fixo — ele sempre usa
`SpreadsheetApp.getActiveSpreadsheet()`, por ser um projeto
container-bound (roda vinculado à própria planilha aberta). Os únicos
identificadores fixos gravados diretamente no código são os do Drive
usados para o Termo de Cientificação de Recurso:

| Identificador | Valor | Uso |
|---|---|---|
| ID da pasta de destino dos Termos | `1_dJV8HP1WFXa5lSV-p0V0N22_YXWIRDa` | Pasta do Drive onde os PDFs `Termo Recurso {candidato}.pdf` são salvos/reaproveitados (`gerarTermoRecursoIndividual`, linha ~1762 de `Código.gs`) |
| ID do template Google Docs do Termo | `1CpgsInQSHnx_ji6NBfAiKmRmbfczKYW-M0QO4LZllgc` | Documento-modelo copiado e preenchido a cada novo Termo gerado (`gerarTermoRecursoIndividual`, linha ~1772) |

**Placeholders substituídos no template do Termo** (via
`body.replaceText(...)`, com escape de chaves como regex):

| Placeholder no Doc | Substituído por |
|---|---|
| `{{Candidato}}` | Nome completo do candidato (aba `candidatos`) |
| `{{Data Laudo}}` | `candidatosDataBase.dataLaudo` formatada `DD/MM/AAAA` (`formatarDataSimples`) |
| `{{DATA_HOJE}}` | Data de hoje, formatada `DD/MM/AAAA` |

O nome do concurso usado na mensagem de sucesso do Termo é lido da
célula **`F3` da aba `Principal`** (`abaPrincipal.getRange('F3')`),
removendo o prefixo `"CONCURSO "` se presente — é a **única** leitura
de célula fixa (não dinâmica) que resta no script, fora do trigger
legado `onEdit`.

A pasta "Mensagens" (PDFs das mensagens administrativas enviadas via
upload) **não** tem ID fixo: é localizada/criada dinamicamente por
`obterPastaMensagens()`, como subpasta com o nome exato `Mensagens`
dentro da mesma pasta do Drive onde está a planilha.

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
     Cientificação de Recurso individual (PDF)

4. Mensagem final de resultados
   → script bloqueia a geração se houver candidato não finalizado
   → agrupa os candidatos de "candidatosDataBase" por status
   → gera a minuta da MENSAGEM DE RESULTADOS DA IS
```

Diagrama de dependência entre abas:

```
mensagens ──────────┐
                     ├─→ obterDadosMensagemInicial() → Data-Hora + Assunto usados nas 2 minutas
agendamentos ←── extrairPeriodoJRS()/calcularDiasUteis() (a partir do texto da mensagem)
     │
     ├─→ listarDatasDisponiveis() / ativarDatasPorDiaSemana()
     │        │
     │        ▼
candidatos ──→ candidatosDataBase ──→ Principal (tabela "principal")
  (mestre)      (fonte de verdade      (tela de trabalho; sincroniza
                 operacional)           de volta via onEdit instalável)
                     │
                     ├─→ obterCandidatosPorStatus() → minuta de resultados
                     └─→ gerarTermoRecursoIndividual() → PDF do Termo de Recurso
```

---

## 2. Estrutura das abas da planilha

### 2.1 `mensagens`
Registro de cada mensagem administrativa (SIGAD-MB) processada.

| Col. | Nome | Conteúdo |
|---|---|---|
| A | `dataHora` | ID único da mensagem (ex.: `P141721Z/JUL/2026`) |
| B | `fileUrl` | Link do PDF salvo no Drive (pasta "Mensagens") |
| C | `proposito` | `"Apresentação e IS"` ou `"Outros"` (detectado automaticamente) |
| D | `remetente` | Campo "De:" do cabeçalho SIGAD |
| E | `destinatario` | Campo "Para:" |
| F | `informacao` | Campo "Info:" |
| G | `assunto` | Campo "Assunto:" (usado para extrair o nome do concurso, ex.: `CPAEAM/2026`) |
| H | `corpoTexto` | Corpo do campo "Texto:" da mensagem, já normalizado/reconstruído |

Novas mensagens são inseridas sempre na linha 2 (mais recente no
topo, via `aba.insertRowBefore(2)`). A `dataHora` da mensagem cujo
`proposito = "Apresentação e IS"` é usada como identificador da
mensagem inicial em ambas as minutas geradas (função
`obterDadosMensagemInicial`). É proibido registrar duas mensagens com
a mesma `dataHora` (`gravarMensagem` lança erro se o ID já existir).

### 2.2 `candidatos`
Lista mestra de candidatos do concurso, ordenada alfabeticamente pelo
nome (`localeCompare` pt-BR, `sensitivity: 'base'`, ignora
maiúsculas/acentos no critério de ordenação).

| Col. | Nome | Conteúdo |
|---|---|---|
| A | `id` | Matrícula do candidato (ex.: `108842-0`) |
| B | `candidato` | Nome completo |

Toda vez que uma mensagem é processada com candidatos novos, a aba
inteira é **reescrita do zero** (limpa e regravada) na nova ordem
alfabética (`gravarExaminee`).

### 2.3 `candidatosDataBase`
Base de dados operacional de cada candidato — é a **fonte de verdade**
do sistema (11 colunas). A tabela `principal` (aba Principal) funciona
como front-end de edição para as colunas Status/Observações/Nº TIS.

| Col. | Índice | Nome | Tipo | Preenchida por |
|---|---|---|---|---|
| A | 1 | `id` | texto | matrícula, espelha `candidatos` |
| B | 2 | `dataAgendamento` | data | agendamento da IS (`gravarDatasAgendamentoCandidatos`) |
| C | 3 | `reagendamento` | — | **não gerenciada pelo script** (uso manual) |
| D | 4 | `status` | texto | `APTO` / `INAPTO` / `FALTOU` / `INSUF DOCUMENTAL` / `Pendente` / vazio |
| E | 5 | `observacoes` | texto | espelho da coluna "Observações" da tabela Principal |
| F | 6 | `finalizado` | ☑️ booleano | `TRUE` somente quando um status com laudo (APTO/INAPTO/FALTOU/INSUF DOCUMENTAL) está gravado **e** a coluna "Nº TIS" da tabela Principal está preenchida para o candidato (ver `candidatoEstaFinalizado`, §4.3) |
| G | 7 | `recurso` | ☑️ booleano | `TRUE` quando o Termo de Recurso é gerado |
| H | 8 | `dataLaudo` | data | data do laudo (hoje, ao meio-dia, no momento em que o Status é definido) |
| I | 9 | `Laudo` | texto | texto do laudo correspondente ao Status (ver `MAPA_LAUDO_POR_STATUS`, §4.3) |
| J | 10 | `nº TIS` | texto | espelho da coluna "Nº TIS" da tabela Principal |
| K | 11 | `termoRecursoUrl` | texto (URL) | link do PDF do Termo de Recurso gerado |

A constante `NUM_COLUNAS = 11` (em `gravarExamineeDataBase`) precisa
ser atualizada junto caso uma coluna seja adicionada/removida nesta
aba.

### 2.4 `agendamentos`
Calendário de datas úteis disponíveis para agendamento das IS.

| Col. | Nome | Conteúdo |
|---|---|---|
| A | `agendamentos` (data) | Data útil (seg–sex, exclui feriados nacionais), sempre gravada ao meio-dia (12:00) |
| B | `diaDaSemana` | Nome do dia (`Segunda`...`Sexta`) |
| C | `ativa` | ☑️ `TRUE` se o dia da semana foi selecionado na última configuração de agendamento (`ativarDatasPorDiaSemana`) |

Gerada automaticamente a partir do período "JRS" extraído da mensagem
inicial (ex.: `03AGO a 14SET2026 (JRS)`, via `extrairPeriodoJRS` +
`calcularDiasUteis`), preservando o valor de `ativa` de datas já
existentes ao reprocessar (`gravarDatasAgendamento`) — nunca duplica
datas, apenas soma as novas ao conjunto já presente.

### 2.5 `Principal` — tabela estruturada `principal`
Tela de trabalho do dia a dia da JRS. É uma **tabela estruturada** do
Google Sheets (Table, não uma faixa fixa de células), com o cabeçalho
localizado **dinamicamente** pelo script (`localizarTabelaPrincipal`,
varre até 20 linhas × 20 colunas procurando as colunas obrigatórias
"Matricula" e "Candidato"; as demais colunas são opcionais).

| Coluna (nome no cabeçalho) | Conteúdo | Sincroniza com `candidatosDataBase` |
|---|---|---|
| `Data` | Data do dia (mesclada) + dia da semana, preenchida automaticamente ao confirmar um agendamento | — |
| `dataAgendamento` | Data "crua" do agendamento (mesmo valor usado internamente) | ← gravada na confirmação do agendamento |
| `Matricula` | ID do candidato — **chave de correspondência** com `candidatosDataBase`/`candidatos` (coluna obrigatória) | — |
| `Candidato` | Nome do candidato (coluna obrigatória) | — |
| `Status` | Lista suspensa: `APTO`, `INAPTO`, `FALTOU`, `INSUF DOCUMENTAL`, `Pendente` | ✅ editável → sincroniza (gatilho instalável, §4.3) |
| `Observações` | Texto livre | ✅ editável → sincroniza |
| `Nº TIS` | Texto livre (busca por regex `/tis/i` no cabeçalho, não pelo nome exato) | ✅ editável → sincroniza |
| `Data laudo` | Preenchida automaticamente quando um Status com laudo é selecionado | somente leitura (espelho de `dataLaudo`) |

Notas de implementação:
- A busca de colunas ignora acentos/caixa/espaços
  (`normalizarTexto`), então variações como "Matrícula" ou
  "matricula " são reconhecidas igualmente.
- Se `colData`/`colDataAgendamento` não forem encontradas, o script
  usa posições padrão relativas a `colMatricula` (fallback, não
  interrompe a operação).
- Se `colStatus`/`colObservacoes`/`colNumTIS`/`colDataLaudo` não forem
  encontradas, retornam `-1` e o gatilho de sincronização
  simplesmente ignora aquela coluna (nenhum erro).
- Cada bloco de linhas de uma mesma data de agendamento recebe fundo
  alternado (branco `#ffffff` / cinza-claro `#f6f6f6`, contínuo a
  partir do último bloco já existente na aba) e bordas de destaque
  (`SOLID_MEDIUM`, cor `#434343`) na primeira e última linha do bloco
  (`aplicarFormatacaoColunaData`).
- Ao preencher novos agendamentos, o script **nunca sobrescreve**
  linhas já preenchidas: sempre anexa a partir da primeira linha sem
  `Matricula` (`proximaLinhaVaziaTabelaPrincipal`), expandindo a
  tabela estruturada via Sheets API v4 quando necessário
  (`garantirCapacidadeTabelaPrincipal`).
- **Célula `F3`**: única referência fixa que resta no script — lida
  por `gerarTermoRecursoIndividual` como "nome do concurso" para a
  mensagem de sucesso do Termo (fora da tabela estruturada
  propriamente dita).

### 2.6 `Listas por Conclusões` (legado)
Aba usada pelo fluxo antigo, anterior à tabela `principal` dinâmica e a
`candidatosDataBase`. **Não é referenciada por nenhuma função do
script atual** (ver §7 — Itens legados/removidos).

---

## 3. Arquivos do projeto Apps Script

| Arquivo | Papel |
|---|---|
| `Código.gs` | Toda a lógica do sistema (1800 linhas) |
| `appsscript.json` | Manifesto do projeto (ver conteúdo abaixo) |
| `UploadMensagem.html` | Modal de upload do PDF da mensagem administrativa |
| `AgendamentoIS.html` | Modal de configuração e confirmação do agendamento das IS |
| `Modal.html` | Modal genérico de exibição de minuta (agendamento e resultados), com botão "Copiar Minuta" |
| `Alerta.html` | Modal genérico de alerta/confirmação (confirmação inicial, avisos/erros, alerta final) |

### 3.1 `appsscript.json` (manifesto)
```json
{
  "timeZone": "America/Recife",
  "dependencies": {
    "enabledAdvancedServices": [
      { "userSymbol": "Drive",  "version": "v3", "serviceId": "drive"  },
      { "userSymbol": "Docs",   "version": "v1", "serviceId": "docs"   },
      { "userSymbol": "Sheets", "version": "v4", "serviceId": "sheets" }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```
Os três serviços avançados são indispensáveis:
- **Drive v3** (`Drive.Files.create`): conversão de PDF → Google Docs
  com OCR em `extrairTextoPdf`.
- **Docs v1**: não é chamado diretamente por símbolo próprio no
  código atual (o corpo do documento é manipulado via `DocumentApp`,
  serviço básico), mas fica habilitado no manifesto para uso do
  template do Termo.
- **Sheets v4** (`Sheets.Spreadsheets.get` / `.batchUpdate`):
  necessário para ler os metadados da tabela estruturada `principal`
  (limites/`tableId`) e estendê-la programaticamente em
  `garantirCapacidadeTabelaPrincipal` — a classe padrão
  `SpreadsheetApp` não expõe operações sobre `Table` (recurso mais
  recente do Sheets).

### 3.2 `UploadMensagem.html`
Modal de upload (550×480). Fluxo de tela: `tela-upload` (dropzone +
`<input type=file accept="application/pdf,.pdf">`, com suporte a
arrastar-e-soltar) → `tela-resultado` (mensagem de processamento →
sucesso/erro).

Chamadas `google.script.run`:
- `processarMensagemPDFUpload(base64Data, nomeArquivo, mimeType)` —
  envia o arquivo lido em base64 pelo `FileReader` do navegador.
- Em caso de sucesso, o botão de ação vira "Ir ao agendamento" e
  chama `abrirModalAgendamento()` (fecha este modal e abre
  `AgendamentoIS.html` na sequência).

### 3.3 `AgendamentoIS.html`
Modal de configuração de agendamento (600×650). Recebe o **contexto
inicial já calculado no servidor** via template scriptlet
(`<?!= JSON.stringify(contexto) ?>` — variável `contexto` setada em
`abrirModalAgendamento()`): total de candidatos pendentes, período e
dias úteis disponíveis.

Fluxo de tela: `tela-formulario` (quantidade por dia + checkboxes de
dias da semana, todos marcados por padrão) → `tela-viavel` (prévia da
distribuição) **ou** `tela-inviavel` (mensagem de erro com sugestão de
ajuste).

Chamadas `google.script.run`:
- `verificarViabilidadeAgendamento(quantidade, dias)` — botão
  "Verificar datas".
- `confirmarAgendamentos(quantidade, dias)` — botão "Confirmar
  agendamentos e minutar MSG" (só aparece após uma verificação
  viável).

### 3.4 `Modal.html`
Modal genérico de exibição de minuta (750×800), reutilizado tanto
para a **minuta de agendamento** quanto para a **minuta de
resultados**. Recebe três variáveis do servidor via template
scriptlet: `textoFinal` (corpo da minuta), `titulo` e
`fecharComAlerta` (booleano).

- Botão "Copiar Minuta": usa um `<textarea>` oculto fora da tela +
  `document.execCommand('copy')` (compatibilidade ampla dentro do
  iframe sandboxed do Apps Script HtmlService).
- Botão "Fechar": se `fecharComAlerta === true`, chama
  `mostrarAlertaFinal()` antes de fechar (exclusivo do fluxo de
  conclusão/resultados, §4.5); caso contrário fecha direto (fluxo de
  agendamento, §4.2).

### 3.5 `Alerta.html`
Modal genérico de alerta/confirmação (500×400). Recebe `titulo`,
`mensagem` (HTML, via `<?!= mensagem ?>`) e `tipo` do servidor.

- `tipo === 'confirmacao'`: botões "Não" (fecha) / "Sim, gerar
  minuta" (chama `abrirModal()` no servidor — único uso concreto
  deste tipo hoje, na confirmação inicial de `iniciarProcesso()`,
  §4.5).
- Qualquer outro valor de `tipo` (`'final'`, ou outro): botão único
  "OK, Fechar".

---

## 4. Fluxo detalhado

### 4.1 Upload e extração da mensagem inicial
Menu **✏️ Termos → 📄 Registrar Mensagem (PDF)**.

1. `iniciarUploadMensagem()` abre `UploadMensagem.html`.
2. O PDF selecionado é enviado em base64 para
   `processarMensagemPDFUpload(base64Data, nomeArquivo, mimeType)`,
   que decodifica o base64 (`Utilities.base64Decode`), monta um
   `Blob` e salva o arquivo na subpasta "Mensagens" (mesma pasta da
   planilha, criada se não existir — `obterPastaMensagens()`), e
   chama `processarArquivoMensagem(arquivoPdf)`.
3. `processarArquivoMensagem()` executa, em ordem:
   1. **Extração de texto com OCR**: `extrairTextoPdf(arquivoPdf)`
      converte o PDF para uma cópia temporária em Google Docs via
      Drive API v3 (`Drive.Files.create(recurso, blob, {ocrLanguage:
      'pt'})`, `mimeType: MimeType.GOOGLE_DOCS`), lê o texto pelo
      `DocumentApp` e **descarta** a cópia temporária
      (`setTrashed(true)`) no `finally`.
   2. **Limpeza de ruído**: `limparRuidoPaginacao(texto)` remove a
      marca d'água repetida `HNRe - 02.2` (regex
      `/HNRe\s*-?\s*0?2\.2/gi`) e os rodapés `Página X de Y` (regex
      `/P[áa]gina\s+\d+\s+de\s+\d+/gi`), que aparecem como texto real
      embutido nas quebras de página (não apenas elementos visuais) e
      ficam intercalados no meio do corpo da mensagem e da lista de
      candidatos. Também normaliza espaços/tabs repetidos e colapsa
      3+ quebras de linha em 2.
   3. **Extração do cabeçalho SIGAD-MB**:
      `extrairCabecalhoMensagem(texto)`, com os seguintes padrões:
      | Campo | Regex |
      |---|---|
      | `dataHora` | `/Data-Hora\s*[\r\n]+\s*([^\r\n]+)/i` |
      | `sender` (De:) | `/\bDe:\s*([^\r\n]+)/i` |
      | `recipient` (Para:) | `/\bPara:\s*([^\r\n]+)/i` |
      | `info` (Info:) | `/\bInfo:\s*([^\r\n]+)/i` |
      | `subject` (Assunto:) | `/\bAssunto:\s*([^\r\n]+)/i` |
      | `corpoTexto` (Texto:) | `/\bTexto:\s*([\s\S]*?)(?:\r?\n\s*Tr[âa]mite:\|\r?\n\s*Prazo para Transmiss\|$)/i` |

      O `purpose` (propósito) da mensagem é determinado testando se o
      corpo contém a frase `candidatos abaixo relacionados`
      (`/candidatos\s+abaixo\s+relacionados/i`); se sim,
      `"Apresentação e IS"`, senão `"Outros"`. Se a `dataHora` não for
      encontrada, o processamento é abortado com erro ("Não foi
      possível localizar o código Data-Hora...").
   4. **Extração da lista de candidatos**: `extrairCandidatos(texto)`
      isola primeiro o trecho entre `candidatos abaixo relacionados`
      e o próximo marcador `DOIS -` (regex de busca de posição:
      `/candidatos\s+abaixo\s+relacionados/i` e `/\bDOIS\s*[-–—]/i`);
      dentro desse trecho, itera com
      `/(\d{6}-\d)\s+([\s\S]+?)(?=\d{6}-\d|$)/g` — **cada item para
      no próximo código de matrícula encontrado, nunca na quebra de
      linha**, porque a lista costuma vir diagramada em 2+ colunas no
      PDF e mais de um candidato pode cair na mesma linha física do
      texto extraído. O nome de cada candidato é normalizado (espaços
      colapsados, hífen/`; e`/`;`/`.` finais removidos).
   5. **Reconstrução do bloco de candidatos**:
      `normalizarTextoComCandidatos(texto, candidatos)` substitui o
      trecho original (potencialmente com candidatos colados na mesma
      linha) pela lista já corretamente separada, reaproveitando
      `aplicarPontuacao()` para manter o padrão `;` / `; e` / `.` —
      preserva intocado o resto do texto (introdução e itens
      `DOIS`/`TRÊS`/`QUATRO` seguintes).
   6. **Gravação dos candidatos**: se `candidatos.length > 0`,
      `gravarExaminee(ss, candidatos)` funde os novos com os já
      existentes em `candidatos` (dedup por `id`), reordena tudo
      alfabeticamente e reescreve a aba; em seguida
      `gravarExamineeDataBase(ss, listaOrdenada)` reescreve a coluna
      `id` de `candidatosDataBase` na mesma ordem, **preservando** os
      dados das demais colunas de candidatos já existentes (e
      preservando, ao final, qualquer `id` com dados que não esteja
      mais na lista de `candidatos`, em vez de descartar
      silenciosamente).
   7. **Registro da mensagem**: `gravarMensagem(ss, dadosMsg,
      arquivoPdf.getUrl())` insere na linha 2 de `mensagens` (lança
      erro se já existir uma mensagem com a mesma `dataHora`).
   8. **Extração do período/agendamento**: `extrairPeriodoJRS(texto)`
      busca o padrão `03AGO a 14SET2026 (JRS)` (regex completa:
      `/(\d{1,2})\s*([A-ZÇ]{3})\s*(\d{4})?\s*a\s*(\d{1,2})\s*([A-ZÇ]{3})\s*(\d{4})\s*\(\s*JRS\s*\)/i`
      — o ano do início é opcional, assume-se o mesmo ano do fim
      quando ausente); se encontrado, `calcularDiasUteis(inicio,
      fim)` gera a lista de dias úteis excluindo sábados/domingos e
      feriados nacionais (fixos + móveis calculados a partir da
      Páscoa, ver §8), e `gravarDatasAgendamento(ss, diasUteis)` grava
      em `agendamentos` (sem duplicar, preservando `ativa` de datas
      já presentes).
4. Ao final, retorna um resumo em HTML (candidatos identificados,
   novos inseridos, período de agendamento) e oferece ir direto para
   o agendamento (`abrirModalAgendamento()`, chamado pelo próprio
   botão "Ir ao agendamento" do modal de resultado).

### 4.2 Configuração do agendamento das IS
Acionado a partir do resultado do upload, ou reaberto manualmente
chamando `abrirModalAgendamento()`.

1. `abrirModalAgendamento()` lista os candidatos ainda sem
   `dataAgendamento` em `candidatosDataBase`
   (`listarCandidatosPendentesAgendamento()`) e monta o resumo (total
   pendente, período, dias úteis disponíveis a partir de
   `listarDatasDisponiveis()`) para `AgendamentoIS.html`. Se não
   houver pendentes, mostra apenas um aviso e não abre o modal.
2. O usuário informa quantas IS por dia e quais dias da semana usar.
   `verificarViabilidadeAgendamento(quantidadePorDia,
   diasSemanaSelecionados)`:
   - marca `ativa=TRUE/FALSE` em `agendamentos` conforme os dias da
     semana escolhidos (`ativarDatasPorDiaSemana()`);
   - calcula se a capacidade (`datas × qtde/dia`) é suficiente; se não
     for, sugere aumentar a quantidade por dia (`Math.ceil(pendentes /
     datas.length)`) ou os dias da semana (`Math.ceil(pendentes /
     quantidadePorDia)` dias necessários);
   - se for viável, distribui os candidatos pendentes pelas datas
     disponíveis, em ordem cronológica, preenchendo cada data até o
     limite antes de passar para a próxima
     (`distribuirCandidatosNasDatas()`) — algoritmo guloso e
     determinístico (mesma entrada → mesma distribuição, o que
     permite recalcular sem persistir estado intermediário).
3. Ao confirmar (`confirmarAgendamentos(quantidadePorDia,
   diasSemanaSelecionados)`):
   - **recalcula** a mesma distribuição (não reaproveita nenhum
     estado do passo anterior — os parâmetros são reenviados do
     modal);
   - grava a data de cada candidato em
     `candidatosDataBase.dataAgendamento`
     (`gravarDatasAgendamentoCandidatos()`);
   - anexa os candidatos à tabela `principal` (aba Principal), sem
     sobrescrever confirmações anteriores (`preencherAbaPrincipal()` →
     `proximaLinhaVaziaTabelaPrincipal()` →
     `garantirCapacidadeTabelaPrincipal()`, que expande a tabela
     estruturada via Sheets API v4 quando faltam linhas);
   - formata a coluna "Data" (mescla data + dia da semana, ou célula
     única quando há 1 candidato no dia) e pinta o bloco de linhas de
     cada dia com fundo alternado e bordas de destaque
     (`aplicarFormatacaoColunaData()`);
   - gera e exibe a minuta da **MENSAGEM DE AGENDAMENTO DA IS**
     (`gerarTextoMinutaAgendamento()`) no `Modal.html`, com o título
     "MENSAGEM DE AGENDAMENTOS DA IS - Minuta gerada" e
     `fecharComAlerta = false` — ao fechar esse modal **não** aparece
     o alerta de "Processo Concluído" (é exclusivo do fluxo de
     conclusão, §4.5).

Formato da minuta de agendamento (`gerarTextoMinutaAgendamento`):
```
{Data-Hora da mensagem inicial}, PTC:

ALFA - As IS de Ingresso dos Candidatos a {concurso} estão agendadas conforme:

UNO - {data DDMMMAAAA} às 7h30:
- {matricula} {nome};
...
- {matricula} {nome}; e
- {matricula} {nome}.

DOIS - {próxima data} às 7h30:
...

BRAVO - CFM o item 3.1.2 da DGPM-406 (9ª Revisão), Os candidatos que não
comparecerem das respectivas datas de agendamentos de suas IS ou não
apresentarem a totalidade dos exames previstos no edital do certame da data
agendada, terão suas IS concluídas e assinadas tempestivamente com laudos,
respectivamente, de "faltou" ou "Insuficiência Documental Médica" BT
```
`nomeConcurso` vem de `extrairNomeConcurso(subject)` — regex
`/([A-ZÇ]{2,10}\/\d{4})/` sobre o Assunto da mensagem inicial (ex.:
`CPAEAM/2026`); se não casar, usa o Assunto bruto como fallback.

### 4.3 Sincronização Principal ↔ candidatosDataBase (uso corrente)
A tabela Principal passa a ser a tela de trabalho: o usuário edita
Status, Observações e Nº TIS diretamente nela, e o script propaga
essas edições em tempo real para `candidatosDataBase`.

**Ativação (uma única vez por planilha):** menu **✏️ Termos → ⚙️
Ativar automação da tabela Principal**
(`instalarGatilhoOnEditPrincipal()`). Isso é necessário porque a
sincronização precisa exibir alertas (`SpreadsheetApp.getUi()`), o que
um gatilho `onEdit` **simples** não tem permissão para fazer — só um
gatilho **instalável**, criado com autorização plena do usuário via um
item de menu. `onOpen()` (gatilho simples) não pode criar esse gatilho
sozinho. A função verifica antes se já existe um trigger idêntico
(`ScriptApp.getProjectTriggers()`) para não duplicar.

O gatilho instalável (`aoEditarPrincipalInstalavel()`) só age sobre
edições de **uma única célula** (`e.range.getNumRows() === 1 &&
e.range.getNumColumns() === 1`), dentro das linhas de dados da tabela
`principal` (abaixo do cabeçalho localizado dinamicamente), localizando
o candidato pela `Matricula` (coluna que casa com
`candidatosDataBase.id`, via `localizarLinhaCandidatosDataBasePorId`):

- **Nº TIS** (`sincronizarNumTISPrincipal`) → grava/limpa
  `candidatosDataBase.nº TIS` (coluna J / índice 10) **e recalcula
  `finalizado`** (coluna F): lê o `status` já gravado em
  `candidatosDataBase` (coluna D) e chama `candidatoEstaFinalizado(status,
  novoValorTIS)` — se o status for um laudo válido e o Nº TIS que acabou
  de ser gravado não estiver vazio, marca `finalizado = TRUE`; caso
  contrário (Nº TIS limpo, ou status ainda sem laudo), desmarca. Ou seja,
  informar o Nº TIS **depois** de já ter um Status com laudo é o que
  efetivamente marca a caixa "finalizado"; apagar o Nº TIS depois
  desmarca de novo.
- **Observações** (`sincronizarObservacoesPrincipal`) → grava/limpa
  `candidatosDataBase.observacoes` (coluna E / índice 5).
- **Status** (`processarEdicaoStatusPrincipal`):
  - célula limpa (`novoValor` vazio) → limpa `status` (coluna D),
    `finalizado` (coluna F, desmarca), `dataLaudo` (coluna H) e
    `Laudo` (coluna I) em `candidatosDataBase`, e "Data laudo" na
    Principal;
  - `PENDENTE` (comparação `.toUpperCase()`) → grava `status =
    "Pendente"` em `candidatosDataBase`, mas **não** marca
    `finalizado` nem grava `Laudo`/`dataLaudo` (mesmo comportamento de
    limpeza dos demais campos que a célula vazia);
  - `APTO` / `INAPTO` / `FALTOU` / `INSUF DOCUMENTAL` (deve existir
    em `MAPA_LAUDO_POR_STATUS`, senão a edição é ignorada
    silenciosamente) → grava `status`, grava `dataLaudo` = hoje (às
    12:00, para evitar problema de fuso) e `Laudo` conforme o mapa
    abaixo, e replica a data em "Data laudo" na Principal. **`finalizado`
    só é marcado `TRUE` nesse momento se a coluna "Nº TIS" da Principal
    já estiver preenchida** para aquele candidato
    (`candidatoEstaFinalizado(novoValor, numTisAtual)`, lendo o valor
    atual da célula "Nº TIS" via `estrutura.colNumTIS`); se o Nº TIS
    ainda não foi informado, `finalizado` permanece/fica `FALSE` mesmo
    com o Status já definido, até que o Nº TIS seja preenchido (o que
    aciona o recálculo descrito acima em `sincronizarNumTISPrincipal`);
  - **especificamente para `INAPTO`**: antes de gravar, pergunta *"Registrar
    {Candidato} como inapto hoje ({data})?"* (Sim/Não, via
    `ui.alert(...)`). Se "Não", a célula volta ao valor anterior
    (`e.oldValue`) e **nada é gravado** em `candidatosDataBase`. Se
    "Sim", grava normalmente (como no item acima) e, em seguida,
    pergunta *"Gerar o termo de Cientificação de Recurso para
    {Candidato}?"*; se "Sim", chama
    `gerarTermoRecursoIndividual(id)` (§4.4). A geração do Termo **não**
    depende do Nº TIS estar preenchido — apenas a marcação de
    `finalizado`.

Mapa Status → Laudo (`MAPA_LAUDO_POR_STATUS`):

| Status | Laudo |
|---|---|
| `APTO` | Apto para Ingresso |
| `INAPTO` | Inapto para Ingresso |
| `FALTOU` | IS não concluída por não comparecimento |
| `INSUF DOCUMENTAL` | IS não concluída por Insuficiência Documental Médica |

**Regra de `finalizado`** (`candidatoEstaFinalizado(statusValor,
numTisValor)`): retorna `TRUE` apenas quando `statusValor` (maiúsculo,
aparado) existe em `MAPA_LAUDO_POR_STATUS` **e** `numTisValor` (aparado)
não é vazio. É a única função que decide o valor de `finalizado`, chamada
tanto pela sincronização de Status quanto pela de Nº TIS, para que o
resultado seja o mesmo qualquer que seja a ordem em que o usuário
preencha as duas colunas na Principal. Como a minuta de resultados
(§4.5) exige que todo candidato esteja `finalizado` antes de gerar,
isso implica, na prática, que **o Nº TIS de cada candidato precisa
estar preenchido** antes que a minuta de resultados possa ser gerada.

### 4.4 Termo de Cientificação de Recurso — individual (PDF)
`gerarTermoRecursoIndividual(id)` (disparado apenas pelo fluxo INAPTO
descrito acima — não existe mais nenhum caminho de geração em lote,
ver §7):

1. Busca o nome do candidato em `candidatos` (por `id`); se não
   encontrado, alerta e interrompe sem gerar nada.
2. Localiza a linha do candidato em `candidatosDataBase`
   (`localizarLinhaCandidatosDataBasePorId`) e lê `dataLaudo` (coluna
   H), formatada `DD/MM/AAAA`.
3. Lê o "nome do concurso" da célula `Principal!F3` (removendo o
   prefixo `CONCURSO ` se presente) — usado só na mensagem final ao
   usuário, não no PDF em si.
4. **Reaproveitamento**: procura na pasta de Termos
   (`1_dJV8HP1WFXa5lSV-p0V0N22_YXWIRDa`) um arquivo já existente com o
   nome exato `Termo Recurso {candidato}.pdf`
   (`subPasta.getFilesByName(...)`); se existir, reaproveita esse PDF
   em vez de gerar de novo.
5. **Geração** (só quando não há arquivo prévio):
   - copia o template Google Docs
     (`1CpgsInQSHnx_ji6NBfAiKmRmbfczKYW-M0QO4LZllgc`,
     `makeCopy("Temp_Recurso_{candidato}")`);
   - abre a cópia com `DocumentApp`, substitui os placeholders
     `{{Candidato}}`, `{{Data Laudo}}` e `{{DATA_HOJE}}` no corpo
     (`body.replaceText`, com as chaves escapadas para regex:
     `"\\{\\{Candidato\\}\\}"` etc.) e salva
     (`docAberto.saveAndClose()`);
   - exporta a cópia como PDF (`docCopia.getAs("application/pdf")`),
     nomeia o blob e salva o arquivo na pasta de Termos
     (`subPasta.createFile(pdfBlob)`);
   - descarta (lixeira) a cópia temporária do Google Docs
     (`docCopia.setTrashed(true)`) — só o PDF final permanece.
6. Marca `candidatosDataBase.recurso = TRUE` (coluna G) e grava a URL
   do PDF em `termoRecursoUrl` (coluna K).
7. Mostra um alerta genérico com o link do termo gerado (aberto em
   nova aba).

### 4.5 Mensagem final de resultados da IS
Menu **✏️ Termos → ✅ Gerar Minuta de Resultados da IS** aciona
`iniciarProcesso()`, que exibe (via `Alerta.html`, `tipo =
'confirmacao'`) o alerta de confirmação: *"Os dados de Status de cada
candidato (aba 'Principal' / candidatosDataBase) foram devidamente
checados com os dados do SINAIS (candidato a candidato) pelo
Supervisor?"*. Ao confirmar ("Sim, gerar minuta"), chama
`abrirModal()`.

`abrirModal()`:

1. **Bloqueio de segurança**: `listarCandidatosNaoFinalizados(ss)`
   verifica se todo candidato em `candidatosDataBase` tem `finalizado
   = TRUE`. Se houver algum pendente, mostra um alerta listando
   `- {id}  {nome}` de quem falta finalizar e **interrompe a função**
   — a minuta só é gerada com todos finalizados. Como `finalizado`
   exige Status com laudo **e** Nº TIS preenchido (§4.3), este bloqueio
   também barra a geração da minuta enquanto houver candidato com
   Status definido mas sem Nº TIS informado na Principal.
2. Busca a Data-Hora da mensagem inicial em `mensagens`
   (`obterDadosMensagemInicial()`, procura a linha com `proposito =
   "Apresentação e IS"`) e o nome do concurso a partir do Assunto
   (`extrairNomeConcurso()`).
3. Agrupa os candidatos de `candidatosDataBase` por `status`
   (`obterCandidatosPorStatus()`, lendo as 11 colunas de cada linha):
   `APTO`, `INAPTO`, `FALTOU`, `INSUF DOCUMENTAL` (status vazio ou
   `Pendente` **não entram em nenhum grupo**), e a lista de quem tem
   `recurso = TRUE` (com a `dataLaudo` formatada em estilo militar,
   `DDMMMAAAA`). Também retorna o `total` de candidatos cadastrados
   (independente do status).
4. Monta e exibe a minuta em `Modal.html` (750×800), título "Minuta
   Gerada", `fecharComAlerta = true`. Ao fechar, **aparece** o alerta
   "Processo Concluído" orientando colar no SIGAD e tramitar para o
   Supervisor JRS/HNRe (`mostrarAlertaFinal()`) — diferente do fluxo
   de agendamento (§4.2, que fecha direto sem esse alerta).

Formato da minuta de resultados (`abrirModal`):
```
{Data-Hora da mensagem inicial}, PTC que JRS/HNRe concluiu em {hoje, DDMMMAAAA} as
IS dos {total} candidatos APS FIM Ingresso no {concurso} CFM os
resultados abaixo relacionados:

ALFA - Candidatos considerados "Aptos para Ingresso" (total: N):
- {id}  {nome};
...

BRAVO - Candidatos considerados "Inaptos para Ingresso" (total: N):
...

CHARLIE - Candidatos com IS não concluídas por não comparecimento (total: N):
...

DELTA - Candidatos com IS não concluídas por Insuficiência Documental Médica (total: N):
...

ECHO - Candidatos que interpuseram recurso junto à JSD/COM3ºDN através da
assinatura do Termo de Reconhecimento de Recurso (total: N):
- Em {dataLaudo DDMMMAAAA}: {id}  {nome}
... BT
```
Cada bloco `ALFA`...`DELTA` usa `aplicarPontuacao(lista, false)`
(pontuação `;`/`; e`/`.`); o bloco `ECHO` usa o modo "eco"
(`aplicarPontuacao(lista, true)`, sem pontuação no último item, pois a
linha já termina em `BT`).

---

## 5. Referência de funções (`Código.gs`)

Catálogo completo das funções do script, agrupadas por área. Nomes
entre parênteses indicam a origem da chamada quando não é uma chamada
interna comum (menu, `google.script.run` a partir de um HTML, ou
gatilho).

### 5.1 Menu, UI e alertas genéricos
| Função | Descrição |
|---|---|
| `onOpen()` | Gatilho simples — cria o menu "✏️ Termos" ao abrir a planilha |
| `onEdit(e)` | Gatilho simples **legado** — ver §7 |
| `iniciarProcesso()` | (menu) Abre `Alerta.html` de confirmação antes da minuta de resultados |
| `mostrarAlertaFinal()` | Abre `Alerta.html` tipo `final`, mensagem "Processo Concluído" |
| `mostrarAlertaGenerico(titulo, mensagem)` | Abre `Alerta.html` tipo `final` com título/mensagem arbitrários |

### 5.2 Upload e extração da mensagem PDF
| Função | Descrição |
|---|---|
| `iniciarUploadMensagem()` | (menu) Abre `UploadMensagem.html` |
| `processarMensagemPDFUpload(base64Data, nomeArquivo, mimeType)` | (`google.script.run`) Decodifica o PDF, salva na pasta "Mensagens", chama `processarArquivoMensagem` |
| `obterPastaMensagens()` | Localiza/cria a subpasta "Mensagens" ao lado da planilha |
| `processarArquivoMensagem(arquivoPdf)` | Orquestra toda a extração e gravação (§4.1) |
| `extrairTextoPdf(arquivoPdf)` | OCR via conversão temporária para Google Docs (Drive API v3) |
| `limparRuidoPaginacao(texto)` | Remove marca d'água/rodapés de paginação |
| `extrairCabecalhoMensagem(texto)` | Extrai Data-Hora/De/Para/Info/Assunto/Texto/propósito |
| `extrairCandidatos(texto)` | Extrai a lista `{id, nome}` de candidatos do corpo da mensagem |
| `normalizarTextoComCandidatos(texto, candidatos)` | Reconstrói o bloco de candidatos no texto já formatado |
| `gravarExaminee(ss, candidatos)` | Funde e reordena `candidatos` (mestre); retorna `{novos, listaOrdenada}` |
| `lerParesIdNome(aba)` | Utilitário: lê pares `{id, nome}` das colunas A/B de uma aba |
| `gravarExamineeDataBase(ss, listaOrdenada)` | Reordena `candidatosDataBase` preservando dados existentes |
| `gravarMensagem(ss, dadosMsg, urlArquivo)` | Insere o registro da mensagem em `mensagens` (linha 2) |
| `coletarIdsExistentes(aba)` | Utilitário: mapa `{id: true}` da coluna A de uma aba |

### 5.3 Datas de agendamento (período JRS, dias úteis, feriados)
| Função | Descrição |
|---|---|
| `extrairPeriodoJRS(texto)` | Extrai `{inicio, fim}` do padrão `DDMMM a DDMMMAAAA (JRS)` |
| `calcularPascoa(ano)` | Data da Páscoa pelo algoritmo Meeus/Jones/Butcher |
| `obterFeriadosNacionais(ano)` | Lista de feriados fixos + móveis (a partir da Páscoa) de um ano |
| `formatarChaveData(d)` | Formata `Date` como chave `AAAA-MM-DD` (independente de fuso) |
| `calcularDiasUteis(dataInicial, dataFinal)` | Dias úteis (seg-sex, exclui feriados) no intervalo |
| `gravarDatasAgendamento(ss, diasUteis)` | Grava/mescla datas em `agendamentos`, sem duplicar |

### 5.4 Agendamento das IS
| Função | Descrição |
|---|---|
| `abrirModalAgendamento()` | (menu / pós-upload) Abre `AgendamentoIS.html` com o contexto inicial |
| `listarCandidatosPendentesAgendamento(ss)` | Candidatos de `candidatosDataBase` sem `dataAgendamento` |
| `listarDatasDisponiveis(ss)` | Datas de `agendamentos`, ordenadas |
| `ativarDatasPorDiaSemana(ss, dias)` | Marca `ativa` conforme os dias da semana escolhidos |
| `verificarViabilidadeAgendamento(qtdePorDia, dias)` | (`google.script.run`) Verifica capacidade e monta a prévia |
| `distribuirCandidatosNasDatas(candidatos, datas, qtdePorDia)` | Algoritmo guloso de distribuição |
| `confirmarAgendamentos(qtdePorDia, dias)` | (`google.script.run`) Confirma, grava e gera a minuta de agendamento |
| `gravarDatasAgendamentoCandidatos(ss, agendamento)` | Grava `dataAgendamento` de cada candidato |
| `preencherAbaPrincipal(ss, agendamento)` | Anexa os candidatos agendados à tabela `principal` |
| `proximaLinhaVaziaTabelaPrincipal(aba, estrutura)` | Primeira linha sem `Matricula` na tabela |
| `garantirCapacidadeTabelaPrincipal(ss, aba, linhasNecessarias)` | Expande a grade e a tabela estruturada (Sheets API v4) |
| `localizarTabelaPrincipal(aba)` | Localiza dinamicamente cabeçalho/colunas da tabela `principal` |
| `normalizarTexto(texto)` | Utilitário: minúsculas, sem acentos, sem espaços (comparação de cabeçalhos) |
| `obterUltimaColunaTabela(aba, linhaCabecalho)` | Última coluna com cabeçalho preenchido |
| `aplicarFormatacaoColunaData(aba, estrutura, agendamento, linhaInicio)` | Mescla "Data" + fundo alternado + bordas por bloco de dia |
| `gerarTextoMinutaAgendamento(ss, agendamento)` | Monta o texto da minuta de agendamento |
| `obterDadosMensagemInicial(ss)` | Localiza a mensagem `"Apresentação e IS"` mais recente |
| `extrairNomeConcurso(subject)` | Extrai `SIGLA/AAAA` do Assunto da mensagem |
| `formatarDataDDMMMAAAA(data)` | Formata `DDMMMAAAA` (ex.: `14AGO2026`) |
| `numeroCardinalExtenso(n)` | Número por extenso em pt-BR maiúsculo (1–999) |
| `numeroItemLista(n)` | Marcador de item de lista naval (`UNO` para 1, extenso para os demais) |

### 5.5 Sincronização Principal → candidatosDataBase e Termo de Recurso
| Função | Descrição |
|---|---|
| `instalarGatilhoOnEditPrincipal()` | (menu) Cria o gatilho instalável `aoEditarPrincipalInstalavel`, se ainda não existir |
| `aoEditarPrincipalInstalavel(e)` | Gatilho instalável — roteia a edição para o sincronizador certo |
| `candidatoEstaFinalizado(statusValor, numTisValor)` | Única regra que decide `finalizado`: `TRUE` apenas se o status tiver laudo válido e o Nº TIS não estiver vazio |
| `localizarLinhaCandidatosDataBasePorId(aba, id)` | Localiza a linha de um candidato em `candidatosDataBase` |
| `sincronizarNumTISPrincipal(aba, estrutura, linha, novoValor)` | Sincroniza Nº TIS |
| `sincronizarObservacoesPrincipal(aba, estrutura, linha, novoValor)` | Sincroniza Observações |
| `processarEdicaoStatusPrincipal(aba, estrutura, linha, e)` | Sincroniza Status (com confirmações via `ui.alert`) |
| `gerarTermoRecursoIndividual(id)` | Gera/reaproveita o PDF do Termo de Cientificação de Recurso |

### 5.6 Mensagem final de resultados
| Função | Descrição |
|---|---|
| `abrirModal()` | (via `Alerta.html`) Bloqueia se houver pendente, monta e exibe a minuta de resultados |
| `aplicarPontuacao(lista, isEcho)` | Pontuação de listas (`;` / `; e` / `.`, ou "eco" sem pontuação final) |
| `obterCandidatosPorStatus(ss)` | Agrupa candidatos por status + lista de recursos + total |
| `listarCandidatosNaoFinalizados(ss)` | Candidatos com `finalizado = FALSE` |
| `formatarDataMilitar(dataOrig)` | Formata `DDMMMAAAA`, aceita `Date` ou string `DD/MM/AAAA` |
| `formatarDataSimples(dataOrig)` | Formata `DD/MM/AAAA`, aceita `Date` ou string |

---

## 6. Automações e gatilhos

| Gatilho | Tipo | Escopo | Função |
|---|---|---|---|
| Abertura da planilha | simples (`onOpen`) | global | Cria o menu "✏️ Termos" |
| Edição de célula | simples (`onEdit`) | aba "Principal", coluna 8, linha ≥ 11 | Legado — grava a data de hoje quando um valor é digitado (ver §7) |
| Edição de célula | **instalável** (`aoEditarPrincipalInstalavel`, precisa ser ativado 1x pelo menu) | tabela `principal` (aba Principal), colunas Status/Observações/Nº TIS | Sincroniza com `candidatosDataBase`, dispara confirmações e o Termo de Recurso individual |

### Itens de menu (✏️ Termos)
- ✅ Gerar Minuta de Resultados da IS → fluxo de conclusão/minuta final (§4.5)
- 📄 Registrar Mensagem (PDF) → início do fluxo (§4.1)
- ⚙️ Ativar automação da tabela Principal → instala o gatilho de
  sincronização (uma vez só, §4.3)

---

## 7. Itens legados / atenção

- **`onEdit(e)` simples** (linha ~18 de `Código.gs`): grava a data de
  hoje quando a coluna **8** é editada em linha **≥ 11** da aba
  "Principal" (referência por índice fixo de coluna/linha, não pelo
  cabeçalho). Essa referência é anterior à tabela estruturada
  `principal` (que tem cabeçalho e colunas localizados
  dinamicamente); **pode não corresponder mais ao layout atual da
  planilha**. Convive com o gatilho instalável sem conflito direto
  (agem em condições diferentes), mas é redundante com "Data laudo"
  sincronizada pelo gatilho instalável. Vale revisar/remover se não
  estiver mais em uso real.
- **Coluna `reagendamento`** de `candidatosDataBase` (coluna C):
  existe na estrutura mas não é lida nem escrita por nenhuma função
  atual do script (uso manual apenas).
- **Célula `Principal!F3`**: única leitura de célula fixa (fora da
  tabela estruturada) que resta no script, usada só como texto de
  exibição na mensagem de sucesso do Termo de Recurso — não afeta a
  lógica de negócio, mas depende de um layout fixo daquela célula
  específica.

### Removidos (histórico, PRs #13–#18 deste repositório)
- **Geração de Termos de Recurso em lote** (`iniciarGeracaoRecursos()`
  / `processarGeracaoRecursos()`, antigo item de menu "🛑
  Cientificação de Recurso"): lia a aba "Principal" em posições fixas
  (`A12:K`, coluna G = "Recurso" com valor literal "Sim", J = Data
  Laudo) que não correspondiam mais à tabela `principal` atual
  (layout real: `Data, dataAgendamento, Matricula, Candidato, Status,
  Observações, Nº TIS, Data laudo`, sem coluna "Recurso"). Removido do
  script e do menu — o fluxo válido e único hoje é o **individual**,
  disparado ao selecionar Status = INAPTO na Principal (§4.4,
  `gerarTermoRecursoIndividual`). O ramo `recurso_confirmacao` e a
  função `avancarRecursos()` também foram removidos de `Alerta.html`
  por ficarem órfãos.
- **Aba "Listas por Conclusões"**: só era usada pelo fluxo em lote
  removido acima. Não é mais referenciada por nenhuma função do
  script — a minuta de resultados (§4.5) já usava
  `candidatosDataBase`/`candidatos` desde a migração anterior.
- **API Web (`doGet`/`doPost`)**: chegou a ser adicionada
  temporariamente a este mesmo projeto (PR #19) para uso pelo app web
  da JRS, mas foi **revertida** — o backend web vive hoje em um
  projeto Apps Script **standalone** separado
  (`CodeConcursos.gs`, repositório `DoencasEPareceresJRS`), que
  duplica a lógica necessária usando `SpreadsheetApp.openById()` (sem
  UI), em vez de acoplar uma API HTTP a este projeto container-bound.

---

## 8. Convenções auxiliares usadas nas minutas

- **Pontuação de listas** (`aplicarPontuacao`): item do meio termina
  em `;`, penúltimo item em `; e`, último item em `.` (ou sem
  pontuação, no modo "eco", usado na lista de recursos que já termina
  em `BT`).
- **Numeração de itens** (`numeroItemLista` / `numeroCardinalExtenso`):
  por extenso em maiúsculas, com `UNO` para o item 1 (convenção naval,
  evita ambiguidade com o artigo "um"). Suporta de 1 a 999.
- **Datas**:
  - `formatarDataSimples`: `DD/MM/AAAA`.
  - `formatarDataMilitar` / `formatarDataDDMMMAAAA`: `DDMMMAAAA` (ex.:
    `14AGO2026`).
  - Datas de agendamento são sempre construídas ao meio-dia (`12:00`)
    em vez de meia-noite, para evitar que o deslocamento de fuso entre
    o runtime V8 (UTC) e o fuso da planilha (`America/Recife`) jogue a
    data para o dia anterior.
- **Feriados** (`obterFeriadosNacionais`): fixos (01/01
  Confraternização Universal, 21/04 Tiradentes, 01/05 Dia do
  Trabalho, 07/09 Independência, 12/10 N. Sra. Aparecida, 02/11
  Finados, 15/11 Proclamação da República, 20/11 Consciência Negra,
  25/12 Natal) + móveis calculados a partir da Páscoa
  (`calcularPascoa`, algoritmo Meeus/Jones/Butcher): Carnaval
  (segunda e terça, Páscoa − 48/−47 dias), Quarta-feira de Cinzas
  (Páscoa − 46), Sexta-feira Santa (Páscoa − 2), Domingo de Páscoa, e
  Corpus Christi (Páscoa + 60).
- **Normalização de texto para cabeçalhos** (`normalizarTexto`):
  minúsculas, sem acentuação (`normalize('NFD')` + remoção de
  diacríticos) e sem espaços — usada para localizar dinamicamente as
  colunas da tabela `principal` de forma tolerante a variações de
  grafia/capitalização.
