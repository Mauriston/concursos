/**
 * Cria o menu personalizado assim que a planilha for aberta.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✏️ Termos')
    .addItem('🛑 Cientificação de Recurso', 'iniciarGeracaoRecursos')
    .addSeparator()
    .addItem('📄 Registrar Mensagem (PDF)', 'iniciarUploadMensagem')
    .addSeparator()
    .addItem('⚙️ Ativar automação da tabela Principal', 'instalarGatilhoOnEditPrincipal')
    .addToUi();
}

/**
 * Função nativa do Google Sheets que é acionada ao editar uma célula.
 */
function onEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  
  if (sheet.getName() === "Principal") {
    var colEditada = e.range.getColumn();
    var linhaEditada = e.range.getRow();
    
    if (colEditada === 8 && linhaEditada >= 11) {
      var dateCell = sheet.getRange(linhaEditada, 10); 
      if (e.value && !dateCell.getValue()) {
        dateCell.setValue(new Date());
      }
    }
  }
}

/**
 * Dispara o alerta HTML inicial com identidade visual.
 */
function iniciarProcesso() {
  var htmlTemplate = HtmlService.createTemplateFromFile('Alerta');
  htmlTemplate.titulo = 'Confirmação de Dados';
  htmlTemplate.mensagem = "Os dados da aba 'Listas por Conclusões' foram devidamente checados com os dados do SINAIS (candidato a candidato) pelo Supervisor?";
  htmlTemplate.tipo = 'confirmacao';
  
  var htmlOutput = htmlTemplate.evaluate()
    .setWidth(500)
    .setHeight(400)
    .setTitle('Inspeção de Saúde - Marinha do Brasil');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function aplicarPontuacao(lista, isEcho) {
  if (!lista || lista.length === 0) return [];
  var listaFormatada = [];
  for (var i = 0; i < lista.length; i++) {
    var item = lista[i];
    if (i === lista.length - 1) {
      listaFormatada.push(isEcho ? item : item + ".");
    } else if (i === lista.length - 2) {
      listaFormatada.push(item + "; e");
    } else {
      listaFormatada.push(item + ";");
    }
  }
  return listaFormatada;
}

function abrirModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaPrincipal = ss.getSheetByName('Principal');
  var abaListas = ss.getSheetByName('Listas por Conclusões');
  
  var nomeConcursoBruto = abaPrincipal.getRange('F3').getValue() || "NÃO INFORMADO";
  var nomeConcurso = String(nomeConcursoBruto).replace(/CONCURSO\s+/i, '');
  
  var dadosCandidatos = abaPrincipal.getRange('D12:D').getValues();
  var qtdeCandidatos = 0;
  for (var c = 0; c < dadosCandidatos.length; c++) {
    if (String(dadosCandidatos[c][0]).trim() !== "") qtdeCandidatos++;
  }
  
  var dadosListas = abaListas.getDataRange().getValues();
  var aptos = [], inaptos = [], faltosos = [], idm = [], recursos = [];
  
  for (var i = 1; i < dadosListas.length; i++) {
    var linha = dadosListas[i];
    if (linha[0]) aptos.push("- " + linha[0] + "  " + linha[1]);
    if (linha[3]) inaptos.push("- " + linha[3] + "  " + linha[4]);
    if (linha[6]) faltosos.push("- " + linha[6] + "  " + linha[7]);
    if (linha[9]) idm.push("- " + linha[9] + "  " + linha[10]);
    if (linha[12]) {
      var dataBruta = linha[14];
      var dataFormatada = formatarDataMilitar(dataBruta);
      recursos.push("- Em " + dataFormatada + ": " + linha[12] + "  " + linha[13]);
    }
  }

  aptos = aplicarPontuacao(aptos, false);
  inaptos = aplicarPontuacao(inaptos, false);
  faltosos = aplicarPontuacao(faltosos, false);
  idm = aplicarPontuacao(idm, false);
  recursos = aplicarPontuacao(recursos, true);

  var texto = [];
  
  texto.push("R-000000Z/MMM/AAAA, PTC que JRS/HNRe concluiu em " + formatarDataMilitar(new Date()) + " as IS dos " + qtdeCandidatos + " candidatos APS FIM Ingresso no " + nomeConcurso + " CFM os resultados abaixo relacionados:");
  texto.push(""); 
  
  texto.push('ALFA - Candidatos considerados "Aptos para Ingresso" (total: ' + aptos.length + '):');
  if (aptos.length > 0) texto.push(aptos.join('\n'));
  texto.push(""); 
  
  texto.push('BRAVO - Candidatos considerados "Inaptos para Ingresso" (total: ' + inaptos.length + '):');
  if (inaptos.length > 0) texto.push(inaptos.join('\n'));
  texto.push(""); 
  
  texto.push('CHARLIE - Candidatos com IS não concluídas por não comparecimento (total: ' + faltosos.length + '):');
  if (faltosos.length > 0) texto.push(faltosos.join('\n'));
  texto.push(""); 
  
  texto.push('DELTA - Candidatos com IS não concluídas por Insuficiência Documental Médica (total: ' + idm.length + '):');
  if (idm.length > 0) texto.push(idm.join('\n'));
  texto.push(""); 
  
  texto.push('ECHO - Candidatos que interpuseram recurso junto à JSD/COM3ºDN através da assinatura do Termo de Reconhecimento de Recurso (total: ' + recursos.length + '):');
  if (recursos.length > 0) {
    texto.push(recursos.join('\n') + " BT"); 
  } else {
    texto.push("(total: 0) BT"); 
  }

  var textoCompleto = texto.join('\n');
  var htmlTemplate = HtmlService.createTemplateFromFile('Modal');
  htmlTemplate.textoFinal = textoCompleto;
  htmlTemplate.titulo = 'Minuta Gerada';
  htmlTemplate.fecharComAlerta = true;

  var htmlOutput = htmlTemplate.evaluate()
    .setWidth(750)
    .setHeight(800)
    .setTitle('Inspeção de Saúde - Marinha do Brasil');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function formatarDataMilitar(dataOrig) {
  if (!dataOrig) return "DATA NÃO INFORMADA";
  var d;
  if (Object.prototype.toString.call(dataOrig) === "[object Date]") {
    d = dataOrig;
  } else {
    var partes = String(dataOrig).split('/');
    if (partes.length === 3) d = new Date(partes[2], partes[1] - 1, partes[0]);
    else d = new Date(dataOrig); 
  }
  if (isNaN(d.getTime())) return dataOrig; 

  var dia = String(d.getDate()).padStart(2, '0');
  var meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
  var mes = meses[d.getMonth()];
  var ano = d.getFullYear();
  return dia + mes + ano;
}

function mostrarAlertaFinal() {
  var htmlTemplate = HtmlService.createTemplateFromFile('Alerta');
  htmlTemplate.titulo = 'Processo Concluído';
  htmlTemplate.mensagem = 'A minuta está na sua área de transferência.<br><br>Cole no SIGAD e tramite para o Supervisor JRS/HNRe para que ele execute o 2º filtro de verificação de dados junto ao SINAIS.';
  htmlTemplate.tipo = 'final';
  
  var htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(400).setTitle('Inspeção de Saúde');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function mostrarAlertaGenerico(titulo, mensagem) {
  var htmlTemplate = HtmlService.createTemplateFromFile('Alerta');
  htmlTemplate.titulo = titulo;
  htmlTemplate.mensagem = mensagem;
  htmlTemplate.tipo = 'final';
  
  var htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(400).setTitle('Inspeção de Saúde');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}


/**
 * Função auxiliar para formatar a data normal (ex: 25/03/2026) para os Termos
 */
function formatarDataSimples(dataOrig) {
  if (!dataOrig) return "___/___/_____";
  var d;
  if (Object.prototype.toString.call(dataOrig) === "[object Date]") {
    d = dataOrig;
  } else {
    var partes = String(dataOrig).split('/');
    if (partes.length === 3) d = new Date(partes[2], partes[1] - 1, partes[0]);
    else d = new Date(dataOrig); 
  }
  if (isNaN(d.getTime())) return dataOrig;

  var dia = String(d.getDate()).padStart(2, '0');
  var mes = String(d.getMonth() + 1).padStart(2, '0');
  var ano = d.getFullYear();
  return dia + "/" + mes + "/" + ano;
}


// =========================================================================
// NOVA FUNCIONALIDADE: TERMOS DE RECONHECIMENTO DE RECURSO (ARQUIVOS INDIVIDUAIS)
// =========================================================================

/**
 * 1. Exibe o modal perguntando se deseja avançar.
 */
function iniciarGeracaoRecursos() {
  var htmlTemplate = HtmlService.createTemplateFromFile('Alerta');
  htmlTemplate.titulo = 'Gerar Termos de Recurso';
  htmlTemplate.mensagem = 'Deseja verificar os candidatos com interposição de recurso e gerar os Termos individuais?<br><br>Ficheiros já existentes na pasta não serão duplicados.';
  htmlTemplate.tipo = 'recurso_confirmacao';
  
  var htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(400).setTitle('Inspeção de Saúde');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * 2. Processa os recursos na aba Principal, verifica duplicações e gera os PDFs.
 */
function processarGeracaoRecursos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var abaPrincipal = ss.getSheetByName('Principal');
    
    // Extrai e formata o nome do concurso
    var nomeConcursoBruto = abaPrincipal.getRange('F3').getValue() || "NÃO INFORMADO";
    var nomeConcurso = String(nomeConcursoBruto).replace(/CONCURSO\s+/i, '');

    // Cria Pasta TERMOS deste concurso
    var idPastaPai = '1_dJV8HP1WFXa5lSV-p0V0N22_YXWIRDa';
    var subPasta = DriveApp.getFolderById(idPastaPai);

    // Leitura dos dados da aba Principal (A12:K)
    // A=0(Data), D=3(Candidato), G=6(Recurso), J=9(Data Laudo)
    var ultimaLinha = abaPrincipal.getLastRow();
    if (ultimaLinha < 12) {
      mostrarAlertaGenerico("Aviso", "Não há dados de candidatos na aba Principal.");
      return;
    }
    
    var dados = abaPrincipal.getRange('A12:K' + ultimaLinha).getValues();
    
    var idTemplate = '1CpgsInQSHnx_ji6NBfAiKmRmbfczKYW-M0QO4LZllgc';
    var dataHojeFormatada = formatarDataSimples(new Date());
    
    var quantidadeGerados = 0;
    var quantidadeIgnorados = 0;

    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];
      var temRecurso = String(linha[6]).trim(); // Coluna G
      
      if (temRecurso.toLowerCase() === 'sim') {
        var candidato = String(linha[3]).trim(); // Coluna D
        var dataLaudo = formatarDataSimples(linha[9]); // Coluna J
        
        var nomeArquivoPdf = "Termo Recurso " + candidato + ".pdf";
        
        // VERIFICAÇÃO DE DUPLICIDADE: Checa se já existe arquivo com este nome na subpasta
        var arquivosExistentes = subPasta.getFilesByName(nomeArquivoPdf);
        if (arquivosExistentes.hasNext()) {
          quantidadeIgnorados++;
          continue; // Já existe, ignora e vai para o próximo
        }

        // Se não existir, faz a cópia e preenche os dados
        var docCopia = DriveApp.getFileById(idTemplate).makeCopy("Temp_Recurso_" + candidato);
        var docAberto = DocumentApp.openById(docCopia.getId());
        var body = docAberto.getBody();
        
        body.replaceText("\\{\\{Candidato\\}\\}", candidato);
        body.replaceText("\\{\\{Data Laudo\\}\\}", dataLaudo);
        body.replaceText("\\{\\{DATA_HOJE\\}\\}", dataHojeFormatada);
        
        docAberto.saveAndClose();
        
        // Converte para PDF e grava na subpasta
        var pdfBlob = docCopia.getAs("application/pdf");
        pdfBlob.setName(nomeArquivoPdf);
        subPasta.createFile(pdfBlob);
        
        // Remove o arquivo temporário (Google Doc)
        docCopia.setTrashed(true);
        
        quantidadeGerados++;
      }
    }

    var linkSubpasta = subPasta.getUrl();
    var mensagemSucesso = '';

    if (quantidadeGerados === 0 && quantidadeIgnorados === 0) {
      mensagemSucesso = "Nenhum candidato com <b>Recurso='Sim'</b> foi encontrado na aba Principal.";
    } else {
      mensagemSucesso = 'Verificação e Processamento Concluídos!<br><br>' +
                        '<b>Novos termos gerados:</b> ' + quantidadeGerados + '<br>' +
                        '<b>Termos já existentes (ignorados):</b> ' + quantidadeIgnorados + '<br><br>' +
                        '<b>Pasta de Destino:</b> <a href="' + linkSubpasta + '" target="_blank">' + nomeConcurso + '</a><br><br>' +
                        '<i>Clique no link acima para abrir a pasta com os ficheiros.</i>';
    }

    mostrarAlertaGenerico("Processo de Recursos Concluído", mensagemSucesso);

  } catch (erro) {
    mostrarAlertaGenerico("Erro", "Ocorreu um erro ao processar os recursos: <br><br>" + erro.message);
  }
}


// =========================================================================
// LEITURA DE MENSAGEM ADMINISTRATIVA (PDF) E EXTRAÇÃO DE DADOS
// =========================================================================

/**
 * 1. Abre o modal de upload direto do PDF da mensagem administrativa (SIGAD-MB).
 */
function iniciarUploadMensagem() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('UploadMensagem')
    .setWidth(550)
    .setHeight(480)
    .setTitle('Inspeção de Saúde - Marinha do Brasil');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * 2. Recebe o PDF enviado pelo modal (conteúdo em base64), salva na pasta
 *    "Mensagens" (na mesma pasta da planilha) e dispara o processamento.
 *    Chamada via google.script.run a partir de UploadMensagem.html.
 */
function processarMensagemPDFUpload(base64Data, nomeArquivo, mimeType) {
  if (!base64Data) {
    throw new Error('Nenhum arquivo foi recebido.');
  }

  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType || 'application/pdf', nomeArquivo || ('mensagem_' + new Date().getTime() + '.pdf'));

  var pastaMensagens = obterPastaMensagens();
  var arquivoPdf = pastaMensagens.createFile(blob);

  return processarArquivoMensagem(arquivoPdf);
}

/**
 * Retorna (criando se necessário) a subpasta "Mensagens" na mesma pasta
 * onde está a planilha, para guardar os PDFs enviados via upload.
 */
function obterPastaMensagens() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var arquivoPlanilha = DriveApp.getFileById(ss.getId());
  var pais = arquivoPlanilha.getParents();
  var pastaPai = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();

  var subPastas = pastaPai.getFoldersByName('Mensagens');
  if (subPastas.hasNext()) return subPastas.next();

  return pastaPai.createFolder('Mensagens');
}

/**
 * Núcleo do processamento: recebe o arquivo PDF já salvo no Drive, extrai
 * o texto, o cabeçalho, os candidatos e o período de agendamento da JRS, e
 * grava tudo nas abas correspondentes. Retorna uma mensagem HTML de resultado.
 */
function processarArquivoMensagem(arquivoPdf) {
  var textoMensagem = limparRuidoPaginacao(extrairTextoPdf(arquivoPdf));
  var dadosMsg = extrairCabecalhoMensagem(textoMensagem);

  if (!dadosMsg.dataHora) {
    throw new Error('Não foi possível localizar o código Data-Hora (ID único) da mensagem no PDF.');
  }

  var candidatos = extrairCandidatos(dadosMsg.texto || textoMensagem);
  dadosMsg.texto = normalizarTextoComCandidatos(dadosMsg.texto, candidatos);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var novosExaminee = 0;
  if (candidatos.length) {
    var resultadoExaminee = gravarExaminee(ss, candidatos);
    novosExaminee = resultadoExaminee.novos;
    gravarExamineeDataBase(ss, resultadoExaminee.listaOrdenada);
  }
  gravarMensagem(ss, dadosMsg, arquivoPdf.getUrl());

  var periodoJRS = extrairPeriodoJRS(dadosMsg.texto || textoMensagem);
  var infoPeriodo;
  if (periodoJRS) {
    var diasUteis = calcularDiasUteis(periodoJRS.inicio, periodoJRS.fim);
    gravarDatasAgendamento(ss, diasUteis);
    infoPeriodo = formatarDataSimples(periodoJRS.inicio) + ' à ' + formatarDataSimples(periodoJRS.fim) +
      ' (' + diasUteis.length + ' dias úteis).';
  } else {
    infoPeriodo = 'não identificado no texto da mensagem.';
  }

  var avisoCandidatos = candidatos.length === 0
    ? '<br><br><i>Nenhum candidato foi identificado no texto da mensagem.</i>'
    : '';

  return 'Mensagem <b>' + dadosMsg.dataHora + '</b> processada com sucesso!<br><br>' +
    '<b>Candidatos identificados na MSG:</b> ' + candidatos.length + '<br>' +
    '<b>Novos candidatos inseridos na planilha:</b> ' + novosExaminee + '<br>' +
    '<b>Período de agendamento identificado:</b> ' + infoPeriodo +
    avisoCandidatos;
}

/**
 * Converte o PDF para Google Docs (com OCR) via serviço avançado Drive
 * apenas para extrair o texto, e depois descarta a cópia temporária.
 * Requer o serviço avançado "Drive" (API v3) habilitado no appsscript.json.
 */
function extrairTextoPdf(arquivoPdf) {
  var blob = arquivoPdf.getBlob();
  var recurso = {
    name: 'OCR_TEMP_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_DOCS
  };

  var arquivoConvertido = Drive.Files.create(recurso, blob, { ocrLanguage: 'pt' });

  try {
    var doc = DocumentApp.openById(arquivoConvertido.id);
    return doc.getBody().getText();
  } finally {
    DriveApp.getFileById(arquivoConvertido.id).setTrashed(true);
  }
}

/**
 * Remove ruído de paginação do texto extraído do PDF: a marca d'água
 * repetida "HNRe - 02.2" e os rodapés "Página X de Y", que aparecem como
 * texto real embutido nas quebras de página (não apenas elementos visuais)
 * e acabam intercalados no meio do corpo da mensagem e da lista de
 * candidatos. Também normaliza espaços/quebras de linha resultantes.
 */
function limparRuidoPaginacao(texto) {
  var limpo = texto
    .replace(/HNRe\s*-?\s*0?2\.2/gi, ' ')
    .replace(/P[áa]gina\s+\d+\s+de\s+\d+/gi, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .replace(/\n{3,}/g, '\n\n');

  return limpo.trim();
}

/**
 * Extrai os campos do cabeçalho da mensagem SIGAD-MB (Data-Hora, De, Para,
 * Info, Assunto e o corpo do Texto).
 */
function extrairCabecalhoMensagem(texto) {
  var extrair = function(padrao) {
    var m = texto.match(padrao);
    return m ? m[1].trim() : '';
  };

  var dataHora = extrair(/Data-Hora\s*[\r\n]+\s*([^\r\n]+)/i);
  var sender = extrair(/\bDe:\s*([^\r\n]+)/i);
  var recipient = extrair(/\bPara:\s*([^\r\n]+)/i);
  var info = extrair(/\bInfo:\s*([^\r\n]+)/i);
  var subject = extrair(/\bAssunto:\s*([^\r\n]+)/i);

  var mTexto = texto.match(/\bTexto:\s*([\s\S]*?)(?:\r?\n\s*Tr[âa]mite:|\r?\n\s*Prazo para Transmiss|$)/i);
  var corpoTexto = mTexto ? mTexto[1].trim() : '';

  var purpose = /candidatos\s+abaixo\s+relacionados/i.test(corpoTexto)
    ? 'Apresentação e IS'
    : 'Outros';

  return {
    dataHora: dataHora,
    sender: sender,
    recipient: recipient,
    info: info,
    subject: subject,
    texto: corpoTexto,
    purpose: purpose
  };
}

/**
 * Extrai a lista de candidatos do corpo da mensagem: itens em lista não
 * enumerada, precedidos por matrícula no formato 000000-0.
 *
 * A lista costuma vir diagramada em colunas (duas ou mais por página), o
 * que faz com que, no texto extraído, mais de um candidato às vezes caia
 * na mesma linha física. Por isso o corte de cada item NÃO usa fim de
 * linha como delimitador: ele sempre para no próximo código de matrícula
 * (\d{6}-\d) encontrado, esteja ele na mesma linha ou não.
 */
function extrairCandidatos(texto) {
  var inicio = texto.search(/candidatos\s+abaixo\s+relacionados/i);
  var fim = texto.search(/\bDOIS\s*[-–—]/i);
  var trecho = texto.substring(
    inicio >= 0 ? inicio : 0,
    fim >= 0 ? fim : texto.length
  );

  var candidatos = [];
  var regexItem = /(\d{6}-\d)\s+([\s\S]+?)(?=\d{6}-\d|$)/g;
  var m;

  while ((m = regexItem.exec(trecho)) !== null) {
    var id = m[1];
    var nome = m[2]
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/-\s*$/, '')
      .trim()
      .replace(/;\s*e$/i, '')
      .replace(/[;.]$/, '')
      .trim();
    if (nome) candidatos.push({ id: id, nome: nome });
  }

  return candidatos;
}

/**
 * Reconstrói o bloco de candidatos dentro do texto da mensagem usando a
 * lista já corretamente separada (um candidato por linha), no lugar do
 * trecho original — que pode ter candidatos colados na mesma linha física
 * por causa do layout em colunas do PDF. Mantém o restante do texto
 * (introdução e itens DOIS/TRÊS/QUATRO) intocado. Reaproveita
 * aplicarPontuacao() para manter o mesmo padrão de pontuação (";", "; e",
 * ".") usado no restante do projeto.
 */
function normalizarTextoComCandidatos(texto, candidatos) {
  if (!texto || !candidatos || candidatos.length === 0) return texto;

  var inicio = texto.search(/candidatos\s+abaixo\s+relacionados/i);
  var fim = texto.search(/\bDOIS\s*[-–—]/i);
  if (inicio < 0 || fim < 0 || fim <= inicio) return texto;

  var fimIntroducao = texto.indexOf(':', inicio);
  if (fimIntroducao < 0 || fimIntroducao >= fim) return texto;

  var antes = texto.substring(0, fimIntroducao + 1);
  var depois = texto.substring(fim);

  var linhas = candidatos.map(function(c) { return '- ' + c.id + ' ' + c.nome; });
  var blocoCandidatos = aplicarPontuacao(linhas, false).join('\n');

  return antes + '\n' + blocoCandidatos + '\n' + depois;
}

/**
 * Grava novos candidatos na aba "candidatos" (colunas id e candidato) e
 * reordena TODAS as linhas (novas e já existentes) em ordem alfabética
 * crescente pelo nome. Retorna a quantidade de novos candidatos e a lista
 * completa já ordenada (usada em seguida para sincronizar
 * "candidatosDataBase" na mesma ordem).
 */
function gravarExaminee(ss, candidatos) {
  var aba = ss.getSheetByName('candidatos');
  if (!aba) throw new Error('Aba "candidatos" não encontrada.');

  var existentes = lerParesIdNome(aba);
  var idsExistentes = {};
  existentes.forEach(function(c) { idsExistentes[c.id] = true; });

  var novos = 0;
  candidatos.forEach(function(c) {
    if (!idsExistentes[c.id]) {
      existentes.push({ id: c.id, nome: c.nome });
      idsExistentes[c.id] = true;
      novos++;
    }
  });

  existentes.sort(function(a, b) {
    return a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' });
  });

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 2).clearContent();
  }
  if (existentes.length > 0) {
    var linhas = existentes.map(function(c) { return [c.id, c.nome]; });
    aba.getRange(2, 1, linhas.length, 2).setValues(linhas);
  }

  return { novos: novos, listaOrdenada: existentes };
}

/**
 * Utilitário: lê os pares {id, nome} das colunas A e B de uma aba
 * (a partir da linha 2), ignorando linhas sem ID.
 */
function lerParesIdNome(aba) {
  var ultimaLinha = aba.getLastRow();
  var pares = [];
  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 2).getValues().forEach(function(linha) {
      var id = String(linha[0]).trim();
      var nome = String(linha[1]).trim();
      if (id) pares.push({ id: id, nome: nome });
    });
  }
  return pares;
}

/**
 * Reescreve a coluna id da aba "candidatosDataBase" seguindo exatamente a
 * mesma ordem de "listaOrdenada" (a lista já ordenada de "candidatos"),
 * preservando os dados das demais colunas de cada candidato já existente
 * e deixando em branco as colunas de candidatos novos. Elimina linhas
 * órfãs/em branco que causavam o início dos dados fora da linha 2.
 */
function gravarExamineeDataBase(ss, listaOrdenada) {
  var aba = ss.getSheetByName('candidatosDataBase');
  if (!aba) throw new Error('Aba "candidatosDataBase" não encontrada.');

  // id + 9 colunas de dados (dataAgendamento, reagendamento, status,
  // finalizado, recurso, dataLaudo, Laudo, nº TIS, termoRecursoUrl)
  var NUM_COLUNAS = 10;
  var ultimaLinha = aba.getLastRow();
  var dadosPorId = {};

  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, NUM_COLUNAS).getValues().forEach(function(linha) {
      var id = String(linha[0]).trim();
      if (id) dadosPorId[id] = linha.slice(1);
    });
  }

  var novos = 0;
  var idsNaLista = {};
  var linhasFinais = listaOrdenada.map(function(c) {
    idsNaLista[c.id] = true;
    if (dadosPorId[c.id]) {
      return [c.id].concat(dadosPorId[c.id]);
    }
    novos++;
    return [c.id].concat(new Array(NUM_COLUNAS - 1).fill(''));
  });

  // Preserva (ao final) qualquer ID com dados que não esteja na lista de
  // "candidatos", em vez de descartar silenciosamente.
  Object.keys(dadosPorId).forEach(function(id) {
    if (!idsNaLista[id]) {
      linhasFinais.push([id].concat(dadosPorId[id]));
    }
  });

  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, NUM_COLUNAS).clearContent();
  }
  if (linhasFinais.length > 0) {
    aba.getRange(2, 1, linhasFinais.length, NUM_COLUNAS).setValues(linhasFinais);
  }

  return novos;
}

/**
 * Insere o registro da mensagem na primeira linha de dados da aba
 * "mensagens" (logo abaixo do cabeçalho), usando a Data-Hora como ID único.
 */
function gravarMensagem(ss, dadosMsg, urlArquivo) {
  var aba = ss.getSheetByName('mensagens');
  if (!aba) throw new Error('Aba "mensagens" não encontrada.');

  var idsExistentes = coletarIdsExistentes(aba);
  if (idsExistentes[dadosMsg.dataHora]) {
    throw new Error('Já existe uma mensagem registrada com o ID "' + dadosMsg.dataHora + '".');
  }

  aba.insertRowBefore(2);
  aba.getRange(2, 1, 1, 8).setValues([[
    dadosMsg.dataHora,
    urlArquivo,
    dadosMsg.purpose,
    dadosMsg.sender,
    dadosMsg.recipient,
    dadosMsg.info,
    dadosMsg.subject,
    dadosMsg.texto
  ]]);
}

/**
 * Utilitário: retorna um mapa {id: true} com os IDs já presentes na
 * coluna A de uma aba (a partir da linha 2).
 */
function coletarIdsExistentes(aba) {
  var mapa = {};
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 1).getValues().forEach(function(linha) {
      var id = String(linha[0]).trim();
      if (id) mapa[id] = true;
    });
  }
  return mapa;
}


// =========================================================================
// DATAS DE AGENDAMENTO (agendamentos): PERÍODO JRS, DIAS ÚTEIS E FERIADOS
// =========================================================================

/**
 * Extrai o período de agendamento da JRS no texto da mensagem, no padrão
 * "03AGO a 14SET2026 (JRS)". O ano do início é opcional no texto; quando
 * ausente, assume-se o mesmo ano do fim do período.
 */
function extrairPeriodoJRS(texto) {
  var meses = {
    'JAN': 0, 'FEV': 1, 'MAR': 2, 'ABR': 3, 'MAI': 4, 'JUN': 5,
    'JUL': 6, 'AGO': 7, 'SET': 8, 'OUT': 9, 'NOV': 10, 'DEZ': 11
  };

  var regex = /(\d{1,2})\s*([A-ZÇ]{3})\s*(\d{4})?\s*a\s*(\d{1,2})\s*([A-ZÇ]{3})\s*(\d{4})\s*\(\s*JRS\s*\)/i;
  var m = texto.match(regex);
  if (!m) return null;

  var mesIni = meses[m[2].toUpperCase()];
  var mesFim = meses[m[5].toUpperCase()];
  if (mesIni === undefined || mesFim === undefined) return null;

  var diaIni = parseInt(m[1], 10);
  var diaFim = parseInt(m[4], 10);
  var anoFim = parseInt(m[6], 10);
  var anoIni = m[3] ? parseInt(m[3], 10) : anoFim;

  return {
    inicio: new Date(anoIni, mesIni, diaIni),
    fim: new Date(anoFim, mesFim, diaFim)
  };
}

/**
 * Calcula a data da Páscoa (Domingo) para um determinado ano, pelo
 * algoritmo Anônimo Gregoriano (Meeus/Jones/Butcher).
 */
function calcularPascoa(ano) {
  var a = ano % 19;
  var b = Math.floor(ano / 100);
  var c = ano % 100;
  var d = Math.floor(b / 4);
  var e = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var mes = Math.floor((h + l - 7 * m + 114) / 31); // 3 = Março, 4 = Abril
  var dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}

/**
 * Retorna a lista de feriados nacionais de um ano, incluindo os feriados
 * fixos e os móveis calculados a partir da Páscoa (Carnaval, Quarta-feira
 * de Cinzas, Sexta-feira Santa e Corpus Christi).
 */
function obterFeriadosNacionais(ano) {
  var feriados = [];

  var somarDias = function(data, dias) {
    var d = new Date(data.getTime());
    d.setDate(d.getDate() + dias);
    return d;
  };

  var fixos = [
    [0, 1],   // 01/01 - Confraternização Universal
    [3, 21],  // 21/04 - Tiradentes
    [4, 1],   // 01/05 - Dia do Trabalho
    [8, 7],   // 07/09 - Independência do Brasil
    [9, 12],  // 12/10 - Nossa Senhora Aparecida
    [10, 2],  // 02/11 - Finados
    [10, 15], // 15/11 - Proclamação da República
    [10, 20], // 20/11 - Dia Nacional de Zumbi e da Consciência Negra
    [11, 25]  // 25/12 - Natal
  ];
  fixos.forEach(function(f) {
    feriados.push(new Date(ano, f[0], f[1]));
  });

  var pascoa = calcularPascoa(ano);
  feriados.push(somarDias(pascoa, -48)); // Carnaval (segunda-feira)
  feriados.push(somarDias(pascoa, -47)); // Carnaval (terça-feira)
  feriados.push(somarDias(pascoa, -46)); // Quarta-feira de Cinzas
  feriados.push(somarDias(pascoa, -2));  // Sexta-feira Santa
  feriados.push(pascoa);                 // Domingo de Páscoa
  feriados.push(somarDias(pascoa, 60));  // Corpus Christi

  return feriados;
}

/**
 * Formata uma data como chave "AAAA-MM-DD", independente de fuso horário.
 */
function formatarChaveData(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

/**
 * Calcula os dias úteis (segunda a sexta) entre duas datas (inclusive),
 * excluindo os feriados nacionais (fixos e móveis) do(s) ano(s) do período.
 */
function calcularDiasUteis(dataInicial, dataFinal) {
  var dias = [];
  var mapaFeriadosPorAno = {};

  var atual = new Date(dataInicial.getFullYear(), dataInicial.getMonth(), dataInicial.getDate());
  var fim = new Date(dataFinal.getFullYear(), dataFinal.getMonth(), dataFinal.getDate());

  while (atual.getTime() <= fim.getTime()) {
    var ano = atual.getFullYear();
    if (!mapaFeriadosPorAno[ano]) {
      mapaFeriadosPorAno[ano] = {};
      obterFeriadosNacionais(ano).forEach(function(d) {
        mapaFeriadosPorAno[ano][formatarChaveData(d)] = true;
      });
    }

    var diaSemana = atual.getDay(); // 0 = Domingo, 6 = Sábado
    var eFeriado = !!mapaFeriadosPorAno[ano][formatarChaveData(atual)];

    if (diaSemana !== 0 && diaSemana !== 6 && !eFeriado) {
      dias.push(new Date(atual.getTime()));
    }

    atual.setDate(atual.getDate() + 1);
  }

  return dias;
}

/**
 * Nomes dos dias da semana em pt-BR, no índice retornado por Date#getDay()
 * (0 = Domingo ... 6 = Sábado). calcularDiasUteis() só gera dias úteis
 * (segunda a sexta), então na prática só os índices 1 a 5 são usados aqui.
 */
var NOMES_DIAS_SEMANA = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

/**
 * Grava as datas úteis calculadas na aba "agendamentos": a data em si
 * na coluna "data" e o nome do dia da semana (Segunda...Sexta) na coluna
 * "diaDaSemana", preservando o status "ativa" de datas já existentes e
 * sem duplicar datas.
 */
function gravarDatasAgendamento(ss, diasUteis) {
  var aba = ss.getSheetByName('agendamentos');
  if (!aba) throw new Error('Aba "agendamentos" não encontrada.');

  var ultimaLinha = aba.getLastRow();
  var mapaAtivo = {};

  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 3).getValues().forEach(function(linha) {
      var data = linha[0];
      if (data instanceof Date && !isNaN(data.getTime())) {
        mapaAtivo[formatarChaveData(data)] = linha[2] === true;
      }
    });
  }

  var totalAntes = Object.keys(mapaAtivo).length;

  diasUteis.forEach(function(d) {
    var chave = formatarChaveData(d);
    if (!(chave in mapaAtivo)) mapaAtivo[chave] = false;
  });

  var chaves = Object.keys(mapaAtivo).sort();

  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 2).clearContent();
    aba.getRange(2, 3, ultimaLinha - 1, 1).clearContent();
  }

  if (chaves.length > 0) {
    var linhasDataDia = chaves.map(function(chave) {
      var p = chave.split('-');
      // Meio-dia (em vez de meia-noite) evita que o deslocamento de fuso
      // horário entre a interpretação UTC do runtime V8 e o fuso da
      // planilha (America/Recife) empurre a data para o dia anterior ao
      // ser exibida (ex.: segunda-feira aparecendo como domingo).
      var data = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
      return [data, NOMES_DIAS_SEMANA[data.getDay()]];
    });
    var linhasAtivo = chaves.map(function(chave) {
      return [mapaAtivo[chave]];
    });

    aba.getRange(2, 1, linhasDataDia.length, 2).setValues(linhasDataDia);
    aba.getRange(2, 3, linhasAtivo.length, 1).setValues(linhasAtivo);
  }

  return chaves.length - totalAntes;
}


// =========================================================================
// AGENDAMENTO DE INSPEÇÕES DE SAÚDE (IS)
// =========================================================================

/**
 * 1. Abre o modal de configuração do agendamento (AgendamentoIS.html),
 *    com o resumo (total de candidatos pendentes, período e dias úteis
 *    disponíveis) já calculado.
 */
function abrirModalAgendamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendentes = listarCandidatosPendentesAgendamento(ss);

  if (pendentes.length === 0) {
    mostrarAlertaGenerico('Aviso', 'Não há candidatos pendentes de agendamento (todos já têm uma data em "dataAgendamento", em "candidatosDataBase").');
    return;
  }

  var datas = listarDatasDisponiveis(ss);

  var contexto = {
    totalCandidatos: pendentes.length,
    periodoInicioFormatado: datas.length ? formatarDataSimples(datas[0].data) : '',
    periodoFimFormatado: datas.length ? formatarDataSimples(datas[datas.length - 1].data) : '',
    diasUteisDisponiveis: datas.length
  };

  var htmlTemplate = HtmlService.createTemplateFromFile('AgendamentoIS');
  htmlTemplate.contexto = contexto;

  var htmlOutput = htmlTemplate.evaluate()
    .setWidth(600)
    .setHeight(650)
    .setTitle('Inspeção de Saúde - Marinha do Brasil');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * Lista os candidatos de "candidatosDataBase" que ainda não têm uma data
 * em "dataAgendamento", já com o nome (de "candidatos") anexado. A ordem
 * segue a mesma ordem de "candidatosDataBase" (que espelha "candidatos",
 * já ordenada alfabeticamente por nome).
 */
function listarCandidatosPendentesAgendamento(ss) {
  var abaDataBase = ss.getSheetByName('candidatosDataBase');
  var abaExaminee = ss.getSheetByName('candidatos');
  if (!abaDataBase) throw new Error('Aba "candidatosDataBase" não encontrada.');
  if (!abaExaminee) throw new Error('Aba "candidatos" não encontrada.');

  var mapaNomes = {};
  lerParesIdNome(abaExaminee).forEach(function(c) { mapaNomes[c.id] = c.nome; });

  var ultimaLinha = abaDataBase.getLastRow();
  if (ultimaLinha < 2) return [];

  var pendentes = [];
  abaDataBase.getRange(2, 1, ultimaLinha - 1, 2).getValues().forEach(function(linha) {
    var id = String(linha[0]).trim();
    var jaAgendado = linha[1] instanceof Date && !isNaN(linha[1].getTime());
    if (id && !jaAgendado) {
      pendentes.push({ id: id, nome: mapaNomes[id] || '' });
    }
  });

  return pendentes;
}

/**
 * Lista as datas presentes em "agendamentos" (colunas "data" e
 * "diaDaSemana"), em ordem crescente.
 */
function listarDatasDisponiveis(ss) {
  var aba = ss.getSheetByName('agendamentos');
  if (!aba) throw new Error('Aba "agendamentos" não encontrada.');

  var ultimaLinha = aba.getLastRow();
  var datas = [];

  if (ultimaLinha >= 2) {
    aba.getRange(2, 1, ultimaLinha - 1, 2).getValues().forEach(function(linha) {
      if (linha[0] instanceof Date && !isNaN(linha[0].getTime())) {
        datas.push({ data: linha[0], diaSemana: String(linha[1]).trim() });
      }
    });
  }

  datas.sort(function(a, b) { return a.data.getTime() - b.data.getTime(); });
  return datas;
}

/**
 * 2. Marca "ativa" (coluna C de "agendamentos") como TRUE para as
 *    datas cujo dia da semana está entre os selecionados, e FALSE para
 *    as demais, refletindo a configuração escolhida pelo usuário.
 */
function ativarDatasPorDiaSemana(ss, diasSemanaSelecionados) {
  var aba = ss.getSheetByName('agendamentos');
  if (!aba) throw new Error('Aba "agendamentos" não encontrada.');

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  var diasSemanaColuna = aba.getRange(2, 2, ultimaLinha - 1, 1).getValues();
  var valoresAtivo = diasSemanaColuna.map(function(linha) {
    return [diasSemanaSelecionados.indexOf(String(linha[0]).trim()) !== -1];
  });

  aba.getRange(2, 3, valoresAtivo.length, 1).setValues(valoresAtivo);
}

/**
 * 3. Verifica se a combinação "quantidade de IS por dia" + "dias da
 *    semana escolhidos" é suficiente para agendar todos os candidatos
 *    pendentes. Se não for, sugere aumentar a quantidade por dia ou o
 *    número de dias da semana usados. Se for, retorna a distribuição
 *    completa (data + candidatos) para exibição no modal.
 *    Chamada via google.script.run a partir de AgendamentoIS.html.
 */
function verificarViabilidadeAgendamento(quantidadePorDia, diasSemanaSelecionados) {
  quantidadePorDia = parseInt(quantidadePorDia, 10);
  if (!quantidadePorDia || quantidadePorDia < 1) {
    throw new Error('Informe uma quantidade válida de IS por dia (mínimo 1).');
  }
  if (!diasSemanaSelecionados || diasSemanaSelecionados.length === 0) {
    throw new Error('Selecione ao menos um dia da semana.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ativarDatasPorDiaSemana(ss, diasSemanaSelecionados);

  var candidatos = listarCandidatosPendentesAgendamento(ss);
  if (candidatos.length === 0) {
    return { viavel: false, mensagem: 'Não há candidatos pendentes de agendamento.' };
  }

  var datas = listarDatasDisponiveis(ss).filter(function(d) {
    return diasSemanaSelecionados.indexOf(d.diaSemana) !== -1;
  });
  if (datas.length === 0) {
    return { viavel: false, mensagem: 'Nenhuma das datas disponíveis em "agendamentos" cai nos dias da semana selecionados.' };
  }

  var capacidadeTotal = datas.length * quantidadePorDia;
  if (capacidadeTotal < candidatos.length) {
    var qtdePorDiaSugerida = Math.ceil(candidatos.length / datas.length);
    var diasNecessariosSugeridos = Math.ceil(candidatos.length / quantidadePorDia);
    var mensagem = 'Com ' + quantidadePorDia + ' IS/dia em ' + datas.length + ' dia(s) disponível(is), ' +
      'cabem apenas ' + capacidadeTotal + ' candidatos, mas há ' + candidatos.length + ' pendentes.<br><br>' +
      'Sugestões: aumente para pelo menos <b>' + qtdePorDiaSugerida + ' IS por dia</b> (mantendo os mesmos dias da semana), ' +
      'ou selecione dias da semana suficientes para ter ao menos <b>' + diasNecessariosSugeridos + ' data(s) disponível(is)</b> ' +
      '(mantendo ' + quantidadePorDia + ' IS por dia).';
    return { viavel: false, mensagem: mensagem };
  }

  var agendamento = distribuirCandidatosNasDatas(candidatos, datas, quantidadePorDia);

  return {
    viavel: true,
    agendamento: agendamento.map(function(item) {
      return {
        dataFormatada: formatarDataSimples(item.data),
        diaSemana: item.diaSemana,
        candidatos: item.candidatos.map(function(c) { return c.nome; })
      };
    })
  };
}

/**
 * Distribui os candidatos (na ordem recebida) pelas datas disponíveis,
 * preenchendo cada data até "quantidadePorDia" antes de passar para a
 * próxima, na ordem cronológica das datas.
 */
function distribuirCandidatosNasDatas(candidatos, datas, quantidadePorDia) {
  var resultado = [];
  var indice = 0;

  for (var i = 0; i < datas.length && indice < candidatos.length; i++) {
    var grupo = candidatos.slice(indice, indice + quantidadePorDia);
    if (grupo.length === 0) break;
    resultado.push({ data: datas[i].data, diaSemana: datas[i].diaSemana, candidatos: grupo });
    indice += grupo.length;
  }

  return resultado;
}

/**
 * 4. Confirma o agendamento: recalcula a mesma distribuição (determinística
 *    a partir dos mesmos parâmetros já validados em
 *    verificarViabilidadeAgendamento), grava a data de cada candidato na
 *    coluna "dataAgendamento" de "candidatosDataBase" e exibe o modal com
 *    a minuta da mensagem de agendamento.
 *    Chamada via google.script.run a partir de AgendamentoIS.html.
 */
function confirmarAgendamentos(quantidadePorDia, diasSemanaSelecionados) {
  quantidadePorDia = parseInt(quantidadePorDia, 10);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidatos = listarCandidatosPendentesAgendamento(ss);
  var datas = listarDatasDisponiveis(ss).filter(function(d) {
    return diasSemanaSelecionados.indexOf(d.diaSemana) !== -1;
  });

  var agendamento = distribuirCandidatosNasDatas(candidatos, datas, quantidadePorDia);
  if (agendamento.length === 0) {
    throw new Error('Nenhum agendamento para confirmar. Verifique as datas novamente.');
  }

  gravarDatasAgendamentoCandidatos(ss, agendamento);
  preencherAbaPrincipal(ss, agendamento);

  var textoMinuta = gerarTextoMinutaAgendamento(ss, agendamento);

  var htmlTemplate = HtmlService.createTemplateFromFile('Modal');
  htmlTemplate.textoFinal = textoMinuta;
  htmlTemplate.titulo = 'MENSAGEM DE AGENDAMENTOS DA IS - Minuta gerada';
  htmlTemplate.fecharComAlerta = false;

  var htmlOutput = htmlTemplate.evaluate()
    .setWidth(750)
    .setHeight(800)
    .setTitle('Inspeção de Saúde - Marinha do Brasil');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * Grava, para cada candidato agendado, a data escolhida na coluna
 * "dataAgendamento" (coluna B) de "candidatosDataBase", localizando a
 * linha pelo "id".
 */
function gravarDatasAgendamentoCandidatos(ss, agendamento) {
  var aba = ss.getSheetByName('candidatosDataBase');
  if (!aba) throw new Error('Aba "candidatosDataBase" não encontrada.');

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  var mapaLinhaPorId = {};
  aba.getRange(2, 1, ultimaLinha - 1, 1).getValues().forEach(function(linha, indice) {
    var id = String(linha[0]).trim();
    if (id) mapaLinhaPorId[id] = indice + 2;
  });

  agendamento.forEach(function(item) {
    item.candidatos.forEach(function(c) {
      var linha = mapaLinhaPorId[c.id];
      if (linha) aba.getRange(linha, 2).setValue(item.data);
    });
  });
}

/**
 * Preenche a tabela "principal" da aba "Principal" com os candidatos
 * agendados nesta confirmação: uma linha por candidato, preenchendo
 * "dataAgendamento", "Matricula" e "Candidato" — anexando a partir da
 * primeira linha vazia (sem sobrescrever agendamentos de confirmações
 * anteriores). Antes de escrever, garante que existam linhas
 * suficientes (na grade e, quando possível, no limite da própria
 * tabela estruturada). Também aplica a formatação da coluna "Data":
 * para cada dia de agendamento, mescla a coluna em duas metades (data
 * em cima alinhada embaixo, dia da semana embaixo alinhado em cima —
 * ou uma célula só com as duas linhas quando há um único candidato
 * naquele dia) e pinta todas as colunas da tabela daquelas linhas com
 * fundo branco/cinza claro alternado por dia, continuando a
 * alternância a partir do último grupo já existente na aba (em vez de
 * sempre recomeçar do branco).
 *
 * A posição das colunas é localizada dinamicamente pelo cabeçalho (não
 * fixa por índice), então funciona independente da ordem/posição em que
 * as colunas "Data", "dataAgendamento", "Matricula" e "Candidato"
 * estejam na tabela.
 */
function preencherAbaPrincipal(ss, agendamento) {
  var aba = ss.getSheetByName('Principal');
  if (!aba) throw new Error('Aba "Principal" não encontrada.');

  var estrutura = localizarTabelaPrincipal(aba);
  var linhaInicio = proximaLinhaVaziaTabelaPrincipal(aba, estrutura);

  var linhas = [];
  agendamento.forEach(function(item) {
    item.candidatos.forEach(function(c) {
      linhas.push({ data: item.data, id: c.id, nome: c.nome });
    });
  });
  if (linhas.length === 0) return;

  var linhasNecessarias = linhaInicio + linhas.length - 1;
  garantirCapacidadeTabelaPrincipal(ss, aba, linhasNecessarias);

  aba.getRange(linhaInicio, estrutura.colDataAgendamento, linhas.length, 1)
    .setValues(linhas.map(function(l) { return [l.data]; }));
  aba.getRange(linhaInicio, estrutura.colMatricula, linhas.length, 1)
    .setValues(linhas.map(function(l) { return [l.id]; }));
  aba.getRange(linhaInicio, estrutura.colCandidato, linhas.length, 1)
    .setValues(linhas.map(function(l) { return [l.nome]; }));

  aplicarFormatacaoColunaData(aba, estrutura, agendamento, linhaInicio);
}

/**
 * Retorna a primeira linha vazia (sem "Matricula") da tabela
 * "principal", a partir da linha seguinte ao cabeçalho — o ponto
 * correto para ANEXAR novos candidatos sem sobrescrever agendamentos
 * já gravados em confirmações anteriores.
 */
function proximaLinhaVaziaTabelaPrincipal(aba, estrutura) {
  var primeiraLinhaDados = estrutura.linhaCabecalho + 1;
  var ultimaLinhaComConteudo = aba.getLastRow();

  if (ultimaLinhaComConteudo < primeiraLinhaDados) return primeiraLinhaDados;

  var valores = aba.getRange(primeiraLinhaDados, estrutura.colMatricula, ultimaLinhaComConteudo - primeiraLinhaDados + 1, 1).getValues();
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === '') return primeiraLinhaDados + i;
  }

  return ultimaLinhaComConteudo + 1;
}

/**
 * Garante que a aba tenha linhas físicas suficientes até
 * "linhasNecessarias" (inserindo na grade quando faltar) e, quando
 * possível, estende também o limite da tabela estruturada "principal"
 * (via serviço avançado Sheets v4) para que as novas linhas fiquem
 * dentro dos limites reais da tabela, não apenas soltas na grade.
 * Se a tabela não puder ser localizada/estendida por algum motivo
 * (ex.: API sem suporte), a escrita dos dados ainda funciona
 * normalmente — só o contorno visual da tabela pode precisar de um
 * ajuste manual (arrastar a alça no canto) depois.
 */
function garantirCapacidadeTabelaPrincipal(ss, aba, linhasNecessarias) {
  if (aba.getMaxRows() < linhasNecessarias) {
    aba.insertRowsAfter(aba.getMaxRows(), linhasNecessarias - aba.getMaxRows());
  }

  try {
    var spreadsheetId = ss.getId();
    var sheetId = aba.getSheetId();

    var resposta = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: 'sheets(properties(sheetId),tables(tableId,name,range))'
    });

    var folhaAlvo = null;
    (resposta.sheets || []).forEach(function(folha) {
      if (folha.properties && folha.properties.sheetId === sheetId) folhaAlvo = folha;
    });
    if (!folhaAlvo) return;

    var tabela = null;
    (folhaAlvo.tables || []).forEach(function(t) {
      if (t.name === 'principal') tabela = t;
    });
    if (!tabela || !tabela.range) return;
    if (tabela.range.endRowIndex >= linhasNecessarias) return;

    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        updateTable: {
          table: {
            tableId: tabela.tableId,
            range: {
              sheetId: sheetId,
              startRowIndex: tabela.range.startRowIndex,
              endRowIndex: linhasNecessarias,
              startColumnIndex: tabela.range.startColumnIndex,
              endColumnIndex: tabela.range.endColumnIndex
            }
          },
          fields: 'range'
        }
      }]
    }, spreadsheetId);
  } catch (erroTabela) {
    Logger.log('Não foi possível estender a tabela estruturada "principal": ' + erroTabela.message);
  }
}

/**
 * Localiza dinamicamente, dentro das primeiras 20 linhas/colunas da aba
 * "Principal", o cabeçalho da tabela "principal" (procurando as colunas
 * "Matricula" e "Candidato") e retorna a linha do cabeçalho e as colunas
 * de "Data", "dataAgendamento", "Matricula", "Candidato", "Status",
 * "Observações", "Nº TIS" e "Data laudo". As quatro últimas retornam -1
 * quando não encontradas (colunas opcionais, usadas pela sincronização
 * com "candidatosDataBase").
 */
function localizarTabelaPrincipal(aba) {
  var linhasBusca = Math.min(20, aba.getMaxRows());
  var colunasBusca = Math.min(20, aba.getMaxColumns());
  var valores = aba.getRange(1, 1, linhasBusca, colunasBusca).getValues();

  for (var l = 0; l < valores.length; l++) {
    var linha = valores[l];
    var colMatricula = -1, colData = -1, colDataAgendamento = -1, colCandidato = -1;
    var colStatus = -1, colObservacoes = -1, colNumTIS = -1, colDataLaudo = -1;

    for (var c = 0; c < linha.length; c++) {
      var textoOriginal = linha[c];
      var texto = normalizarTexto(textoOriginal);
      if (texto === 'matricula') colMatricula = c + 1;
      else if (texto === 'dataagendamento') colDataAgendamento = c + 1;
      else if (texto === 'data') colData = c + 1;
      else if (texto === 'candidato') colCandidato = c + 1;
      else if (texto === 'status') colStatus = c + 1;
      else if (texto === 'observacoes') colObservacoes = c + 1;
      else if (texto === 'datalaudo') colDataLaudo = c + 1;
      else if (/tis/i.test(String(textoOriginal))) colNumTIS = c + 1;
    }

    if (colMatricula !== -1 && colCandidato !== -1) {
      return {
        linhaCabecalho: l + 1,
        colData: colData !== -1 ? colData : 1,
        colDataAgendamento: colDataAgendamento !== -1 ? colDataAgendamento : Math.max(colMatricula - 1, 1),
        colMatricula: colMatricula,
        colCandidato: colCandidato,
        colStatus: colStatus,
        colObservacoes: colObservacoes,
        colNumTIS: colNumTIS,
        colDataLaudo: colDataLaudo
      };
    }
  }

  throw new Error('Não foi possível localizar o cabeçalho da tabela "principal" (colunas "Matricula"/"Candidato") na aba "Principal".');
}

/**
 * Normaliza texto para comparação de cabeçalhos: minúsculas, sem
 * acentos e sem espaços.
 */
function normalizarTexto(texto) {
  return String(texto)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .trim();
}

/**
 * Retorna a última coluna com cabeçalho preenchido a partir da linha do
 * cabeçalho da tabela (varre até 30 colunas), para saber até onde pintar
 * o fundo alternado de cada linha.
 */
function obterUltimaColunaTabela(aba, linhaCabecalho) {
  var largura = Math.min(30, aba.getMaxColumns());
  var valoresCabecalho = aba.getRange(linhaCabecalho, 1, 1, largura).getValues()[0];
  var ultima = 1;
  for (var i = 0; i < valoresCabecalho.length; i++) {
    if (String(valoresCabecalho[i]).trim() !== '') ultima = i + 1;
  }
  return ultima;
}

/**
 * Aplica, para cada dia de agendamento, o fundo alternado (branco /
 * #f6f6f6) em todas as colunas da tabela, e a mescla da coluna "Data"
 * (data em cima, dia da semana embaixo). A alternância de cor continua
 * a partir da cor da última linha já existente acima de "linhaInicio",
 * em vez de sempre recomeçar do branco. Também destaca a borda superior
 * da primeira linha e a borda inferior da última linha de cada bloco de
 * data, para marcar visualmente a separação entre os dias.
 */
function aplicarFormatacaoColunaData(aba, estrutura, agendamento, linhaInicio) {
  var ultimaColuna = obterUltimaColunaTabela(aba, estrutura.linhaCabecalho);

  var corAnterior = linhaInicio > estrutura.linhaCabecalho + 1
    ? aba.getRange(linhaInicio - 1, estrutura.colData).getBackground()
    : '#f6f6f6';

  var linhaAtual = linhaInicio;

  agendamento.forEach(function(item) {
    var qtde = item.candidatos.length;
    if (qtde === 0) return;

    var corGrupo = corAnterior === '#f6f6f6' ? '#ffffff' : '#f6f6f6';
    corAnterior = corGrupo;

    aba.getRange(linhaAtual, 1, qtde, ultimaColuna).setBackground(corGrupo);

    aba.getRange(linhaAtual, 1, 1, ultimaColuna)
      .setBorder(true, null, null, null, null, null, '#434343', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    aba.getRange(linhaAtual + qtde - 1, 1, 1, ultimaColuna)
      .setBorder(null, null, true, null, null, null, '#434343', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    var dataFormatada = formatarDataSimples(item.data);
    var diaSemanaFormatado = item.diaSemana.toLowerCase() + '-feira';

    if (qtde === 1) {
      var celulaUnica = aba.getRange(linhaAtual, estrutura.colData);
      celulaUnica.setValue(dataFormatada + '\n' + diaSemanaFormatado);
      celulaUnica.setVerticalAlignment('middle');
      celulaUnica.setHorizontalAlignment('center');
      celulaUnica.setFontWeight('bold');
    } else {
      var metadeSuperior = Math.ceil(qtde / 2);
      var metadeInferior = qtde - metadeSuperior;

      var rangeData = aba.getRange(linhaAtual, estrutura.colData, metadeSuperior, 1);
      rangeData.merge();
      rangeData.setValue(dataFormatada);
      rangeData.setVerticalAlignment('bottom');
      rangeData.setHorizontalAlignment('center');
      rangeData.setFontWeight('bold');

      var rangeDiaSemana = aba.getRange(linhaAtual + metadeSuperior, estrutura.colData, metadeInferior, 1);
      rangeDiaSemana.merge();
      rangeDiaSemana.setValue(diaSemanaFormatado);
      rangeDiaSemana.setVerticalAlignment('top');
      rangeDiaSemana.setHorizontalAlignment('center');
      rangeDiaSemana.setFontWeight('normal');
      rangeDiaSemana.setFontColor('#666666');
    }

    linhaAtual += qtde;
  });
}

/**
 * Gera o texto da minuta da MENSAGEM DE AGENDAMENTO, no padrão:
 *
 * {{data-hora MSG inicial}}, PTC:
 *
 * ALFA - As IS de Ingresso dos Candidatos a {{nomeConcurso}} estão
 * agendadas conforme:
 *
 * UNO - {{1º dia}} às 7h30:
 * - id candidato;
 * ...
 * - id candidato; e
 * - id candidato.
 *
 * DOIS - {{2º dia}} às 7h30:
 * ...
 *
 * BRAVO - CFM o item 3.1.2 da DGPM-406 (9ª Revisão)... BT
 */
function gerarTextoMinutaAgendamento(ss, agendamento) {
  var dadosMsgInicial = obterDadosMensagemInicial(ss);
  var dataHoraInicial = dadosMsgInicial ? dadosMsgInicial.dataHora : 'R-000000Z/MMM/AAAA';
  var nomeConcurso = dadosMsgInicial ? extrairNomeConcurso(dadosMsgInicial.subject) : 'NÃO INFORMADO';

  var linhas = [];
  linhas.push(dataHoraInicial + ', PTC:');
  linhas.push('');
  linhas.push('ALFA - As IS de Ingresso dos Candidatos a ' + nomeConcurso + ' estão agendadas conforme:');
  linhas.push('');

  agendamento.forEach(function(item, indice) {
    var marcador = numeroItemLista(indice + 1);
    linhas.push(marcador + ' - ' + formatarDataDDMMMAAAA(item.data) + ' às 7h30:');

    var itensCandidatos = item.candidatos.map(function(c) {
      return '- ' + c.id + ' ' + c.nome;
    });
    linhas.push(aplicarPontuacao(itensCandidatos, false).join('\n'));
    linhas.push('');
  });

  linhas.push('BRAVO - CFM o item 3.1.2 da DGPM-406 (9ª Revisão), Os candidatos que não comparecerem das respectivas datas de agendamentos de suas IS ou não apresentarem a totalidade dos exames previstos no edital do certame da data agendada, terão suas IS concluídas e assinadas tempestivamente com laudos, respectivamente, de "faltou" ou "Insuficiência Documental Médica" BT');

  return linhas.join('\n');
}

/**
 * Localiza, na aba "mensagens", a mensagem mais recente com propósito
 * "Apresentação e IS" (a mensagem inicial de apresentação de candidatos),
 * retornando seu ID (Data-Hora) e o Assunto.
 */
function obterDadosMensagemInicial(ss) {
  var aba = ss.getSheetByName('mensagens');
  if (!aba) throw new Error('Aba "mensagens" não encontrada.');

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return null;

  var dados = aba.getRange(2, 1, ultimaLinha - 1, 7).getValues(); // dataHora, fileUrl, proposito, remetente, destinatario, informacao, assunto
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][2]).trim() === 'Apresentação e IS') {
      return { dataHora: String(dados[i][0]).trim(), subject: String(dados[i][6]).trim() };
    }
  }

  return null;
}

/**
 * Extrai o identificador do concurso (ex.: "CPAEAM/2026") do Assunto da
 * mensagem, no padrão "SIGLA/AAAA".
 */
function extrairNomeConcurso(subject) {
  var m = subject.match(/([A-ZÇ]{2,10}\/\d{4})/);
  return m ? m[1] : subject;
}

/**
 * Formata uma data no padrão militar sem separadores: ddMMMaaaa
 * (ex.: 14AGO2026).
 */
function formatarDataDDMMMAAAA(data) {
  var meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];
  var dia = String(data.getDate()).padStart(2, '0');
  var mes = meses[data.getMonth()];
  var ano = data.getFullYear();
  return dia + mes + ano;
}

/**
 * Número cardinal por extenso em pt-BR, maiúsculo (suporta 1 a 999).
 */
function numeroCardinalExtenso(n) {
  var unidades = ['', 'UM', 'DOIS', 'TRÊS', 'QUATRO', 'CINCO', 'SEIS', 'SETE', 'OITO', 'NOVE'];
  var dezA19 = ['DEZ', 'ONZE', 'DOZE', 'TREZE', 'QUATORZE', 'QUINZE', 'DEZESSEIS', 'DEZESSETE', 'DEZOITO', 'DEZENOVE'];
  var dezenas = ['', '', 'VINTE', 'TRINTA', 'QUARENTA', 'CINQUENTA', 'SESSENTA', 'SETENTA', 'OITENTA', 'NOVENTA'];
  var centenas = ['', 'CENTO', 'DUZENTOS', 'TREZENTOS', 'QUATROCENTOS', 'QUINHENTOS', 'SEISCENTOS', 'SETECENTOS', 'OITOCENTOS', 'NOVECENTOS'];

  if (n < 10) return unidades[n];
  if (n < 20) return dezA19[n - 10];
  if (n < 100) {
    var d = Math.floor(n / 10);
    var u = n % 10;
    return dezenas[d] + (u > 0 ? ' E ' + unidades[u] : '');
  }
  if (n === 100) return 'CEM';
  if (n < 1000) {
    var c = Math.floor(n / 100);
    var resto = n % 100;
    return centenas[c] + (resto > 0 ? ' E ' + numeroCardinalExtenso(resto) : '');
  }

  return String(n);
}

/**
 * Marcador numérico usado nas listas das mensagens navais: "UNO" para 1
 * (em vez de "UM", convenção da Marinha para evitar ambiguidade com o
 * artigo "um"), e o cardinal por extenso normal para os demais.
 */
function numeroItemLista(n) {
  return n === 1 ? 'UNO' : numeroCardinalExtenso(n);
}


// =========================================================================
// SINCRONIZAÇÃO "Principal" -> "candidatosDataBase" (CRUD front-end)
// =========================================================================

/**
 * Texto do "Laudo" (candidatosDataBase, coluna H) correspondente a cada
 * opção da coluna "Status" da tabela "principal".
 */
var MAPA_LAUDO_POR_STATUS = {
  'APTO': 'Apto para Ingresso',
  'INAPTO': 'Inapto para Ingresso',
  'FALTOU': 'IS não concluída por não comparecimento',
  'INSUF DOCUMENTAL': 'IS não concluída por Insuficiência Documental Médica'
};

/**
 * Cria (se ainda não existir) o gatilho onEdit INSTALÁVEL responsável pela
 * sincronização Status/Nº TIS da tabela "principal" com "candidatosDataBase".
 * Precisa ser instalável (não simples) porque exibe alertas (SpreadsheetApp.getUi()),
 * o que um gatilho simples onEdit não tem permissão para fazer. Deve ser
 * executado uma única vez, a partir do menu (função acionada pelo usuário
 * tem autorização plena; um gatilho simples como onOpen não tem).
 */
function instalarGatilhoOnEditPrincipal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jaExiste = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'aoEditarPrincipalInstalavel' && t.getEventType() === ScriptApp.EventType.ON_EDIT;
  });

  if (jaExiste) {
    mostrarAlertaGenerico('Automação já ativa', 'A sincronização automática da tabela Principal com "candidatosDataBase" já está ativada nesta planilha.');
    return;
  }

  ScriptApp.newTrigger('aoEditarPrincipalInstalavel')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  mostrarAlertaGenerico('Automação ativada', 'A partir de agora, edições de <b>Status</b> e <b>Nº TIS</b> na tabela Principal serão sincronizadas automaticamente com "candidatosDataBase".');
}

/**
 * Gatilho onEdit instalável (ver instalarGatilhoOnEditPrincipal). Só age
 * sobre edições de uma única célula, dentro das linhas de dados da tabela
 * "principal" (aba "Principal"), nas colunas "Status" e "Nº TIS".
 */
function aoEditarPrincipalInstalavel(e) {
  if (!e || !e.range) return;

  var aba = e.range.getSheet();
  if (aba.getName() !== 'Principal') return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  var estrutura;
  try {
    estrutura = localizarTabelaPrincipal(aba);
  } catch (erroEstrutura) {
    return;
  }

  var linha = e.range.getRow();
  var coluna = e.range.getColumn();
  if (linha <= estrutura.linhaCabecalho) return;

  if (coluna === estrutura.colNumTIS) {
    sincronizarNumTISPrincipal(aba, estrutura, linha, e.value);
  } else if (coluna === estrutura.colStatus) {
    processarEdicaoStatusPrincipal(aba, estrutura, linha, e);
  }
}

/**
 * Localiza a linha (índice 1-based) do candidato de "id" na aba
 * "candidatosDataBase", ou -1 se não encontrado.
 */
function localizarLinhaCandidatosDataBasePorId(aba, id) {
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return -1;

  var ids = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

/**
 * Sincroniza a edição da coluna "Nº TIS" da tabela "principal" para a
 * coluna "nº TIS" (coluna I) de "candidatosDataBase". Uma célula limpa
 * também limpa o valor em "candidatosDataBase".
 */
function sincronizarNumTISPrincipal(aba, estrutura, linha, novoValor) {
  var id = String(aba.getRange(linha, estrutura.colMatricula).getValue()).trim();
  if (!id) return;

  var abaDataBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('candidatosDataBase');
  if (!abaDataBase) return;

  var linhaDataBase = localizarLinhaCandidatosDataBasePorId(abaDataBase, id);
  if (linhaDataBase === -1) return;

  abaDataBase.getRange(linhaDataBase, 9).setValue(novoValor || '');
}

/**
 * Processa a edição da coluna "Status" da tabela "principal":
 * - Se a célula for limpa, limpa também "status", "finalizado", "dataLaudo"
 *   e "Laudo" em "candidatosDataBase", e "Data laudo" em "Principal".
 * - Se um status válido (APTO/INAPTO/FALTOU/INSUF DOCUMENTAL) for
 *   selecionado, grava "status"/"finalizado"/"dataLaudo"/"Laudo" em
 *   "candidatosDataBase" e a "Data laudo" de hoje em "Principal".
 * - Para INAPTO, confirma com o usuário antes de gravar (revertendo a
 *   célula ao valor anterior se recusado) e, em seguida, oferece gerar o
 *   Termo de Cientificação de Recurso.
 */
function processarEdicaoStatusPrincipal(aba, estrutura, linha, e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = String(aba.getRange(linha, estrutura.colMatricula).getValue()).trim();
  if (!id) return;

  var abaDataBase = ss.getSheetByName('candidatosDataBase');
  if (!abaDataBase) return;

  var linhaDataBase = localizarLinhaCandidatosDataBasePorId(abaDataBase, id);
  if (linhaDataBase === -1) {
    SpreadsheetApp.getUi().alert('Candidato com matrícula "' + id + '" não foi encontrado em "candidatosDataBase". A edição de Status não foi sincronizada.');
    return;
  }

  var novoValor = String(e.value || '').trim().toUpperCase();

  if (!novoValor) {
    abaDataBase.getRange(linhaDataBase, 4).setValue('');
    abaDataBase.getRange(linhaDataBase, 5).setValue(false);
    abaDataBase.getRange(linhaDataBase, 7).setValue('');
    abaDataBase.getRange(linhaDataBase, 8).setValue('');
    if (estrutura.colDataLaudo !== -1) aba.getRange(linha, estrutura.colDataLaudo).setValue('');
    return;
  }

  var laudoTexto = MAPA_LAUDO_POR_STATUS[novoValor];
  if (!laudoTexto) return;

  var nomeCandidato = String(aba.getRange(linha, estrutura.colCandidato).getValue()).trim();
  var ui = SpreadsheetApp.getUi();
  var hoje = new Date();
  hoje.setHours(12, 0, 0, 0);

  if (novoValor === 'INAPTO') {
    var resposta = ui.alert(
      'Confirmar Status',
      'Registrar ' + nomeCandidato + ' como inapto hoje (' + formatarDataSimples(hoje) + ')?',
      ui.ButtonSet.YES_NO
    );
    if (resposta !== ui.Button.YES) {
      aba.getRange(linha, estrutura.colStatus).setValue(e.oldValue || '');
      return;
    }
  }

  abaDataBase.getRange(linhaDataBase, 4).setValue(novoValor);
  abaDataBase.getRange(linhaDataBase, 5).setValue(true);
  abaDataBase.getRange(linhaDataBase, 7).setValue(hoje);
  abaDataBase.getRange(linhaDataBase, 8).setValue(laudoTexto);
  if (estrutura.colDataLaudo !== -1) aba.getRange(linha, estrutura.colDataLaudo).setValue(hoje);

  if (novoValor === 'INAPTO') {
    var respostaTermo = ui.alert(
      'Termo de Recurso',
      'Gerar o termo de Cientificação de Recurso para ' + nomeCandidato + '?',
      ui.ButtonSet.YES_NO
    );
    if (respostaTermo === ui.Button.YES) {
      gerarTermoRecursoIndividual(id);
    }
  }
}

/**
 * Gera o Termo de Cientificação de Recurso para um único candidato
 * (identificado por "id"), buscando os dados em "candidatos" e
 * "candidatosDataBase". Segue o mesmo template/pasta usados em
 * processarGeracaoRecursos(), mas reaproveita um arquivo já existente em
 * vez de duplicá-lo. Ao final, marca "recurso" = "Sim" e grava a URL do
 * PDF em "termoRecursoUrl" (candidatosDataBase).
 */
function gerarTermoRecursoIndividual(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCandidatos = ss.getSheetByName('candidatos');
  var abaDataBase = ss.getSheetByName('candidatosDataBase');
  var abaPrincipal = ss.getSheetByName('Principal');
  if (!abaCandidatos || !abaDataBase) return;

  var mapaNomes = {};
  lerParesIdNome(abaCandidatos).forEach(function(c) { mapaNomes[c.id] = c.nome; });
  var candidato = mapaNomes[id];
  if (!candidato) {
    SpreadsheetApp.getUi().alert('Candidato com matrícula "' + id + '" não encontrado em "candidatos". Termo não gerado.');
    return;
  }

  var linhaDataBase = localizarLinhaCandidatosDataBasePorId(abaDataBase, id);
  if (linhaDataBase === -1) return;

  var dadosDataBase = abaDataBase.getRange(linhaDataBase, 1, 1, 10).getValues()[0];
  var dataLaudo = formatarDataSimples(dadosDataBase[6]); // coluna G: dataLaudo

  var nomeConcursoBruto = abaPrincipal ? (abaPrincipal.getRange('F3').getValue() || 'NÃO INFORMADO') : 'NÃO INFORMADO';
  var nomeConcurso = String(nomeConcursoBruto).replace(/CONCURSO\s+/i, '');

  var idPastaPai = '1_dJV8HP1WFXa5lSV-p0V0N22_YXWIRDa';
  var subPasta = DriveApp.getFolderById(idPastaPai);

  var nomeArquivoPdf = "Termo Recurso " + candidato + ".pdf";
  var arquivosExistentes = subPasta.getFilesByName(nomeArquivoPdf);
  var arquivoPdf;

  if (arquivosExistentes.hasNext()) {
    arquivoPdf = arquivosExistentes.next();
  } else {
    var idTemplate = '1CpgsInQSHnx_ji6NBfAiKmRmbfczKYW-M0QO4LZllgc';
    var dataHojeFormatada = formatarDataSimples(new Date());

    var docCopia = DriveApp.getFileById(idTemplate).makeCopy("Temp_Recurso_" + candidato);
    var docAberto = DocumentApp.openById(docCopia.getId());
    var body = docAberto.getBody();

    body.replaceText("\\{\\{Candidato\\}\\}", candidato);
    body.replaceText("\\{\\{Data Laudo\\}\\}", dataLaudo);
    body.replaceText("\\{\\{DATA_HOJE\\}\\}", dataHojeFormatada);

    docAberto.saveAndClose();

    var pdfBlob = docCopia.getAs("application/pdf");
    pdfBlob.setName(nomeArquivoPdf);
    arquivoPdf = subPasta.createFile(pdfBlob);

    docCopia.setTrashed(true);
  }

  abaDataBase.getRange(linhaDataBase, 6).setValue('Sim');
  abaDataBase.getRange(linhaDataBase, 10).setValue(arquivoPdf.getUrl());

  mostrarAlertaGenerico(
    'Termo Gerado',
    'Termo de Cientificação de Recurso gerado para <b>' + candidato + '</b> (' + nomeConcurso + ').<br><br>' +
    '<a href="' + arquivoPdf.getUrl() + '" target="_blank">Abrir termo</a>'
  );
}
