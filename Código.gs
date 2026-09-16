/**
 * Cria o menu personalizado assim que a planilha for aberta.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✏️ Termos')
    .addItem('🛑 Cientificação de Recurso', 'iniciarGeracaoRecursos')
    .addSeparator()
    .addItem('📄 Registrar Mensagem (PDF)', 'iniciarUploadMensagem')
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
 * Grava novos candidatos na aba "examinee" (colunas ID e examinee) e
 * reordena TODAS as linhas (novas e já existentes) em ordem alfabética
 * crescente pelo nome. Retorna a quantidade de novos candidatos e a lista
 * completa já ordenada (usada em seguida para sincronizar
 * "examineedataBase" na mesma ordem).
 */
function gravarExaminee(ss, candidatos) {
  var aba = ss.getSheetByName('examinee');
  if (!aba) throw new Error('Aba "examinee" não encontrada.');

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
 * Reescreve a coluna ID da aba "examineedataBase" seguindo exatamente a
 * mesma ordem de "listaOrdenada" (a lista já ordenada de "examinee"),
 * preservando os dados das demais colunas de cada candidato já existente
 * e deixando em branco as colunas de candidatos novos. Elimina linhas
 * órfãs/em branco que causavam o início dos dados fora da linha 2.
 */
function gravarExamineeDataBase(ss, listaOrdenada) {
  var aba = ss.getSheetByName('examineedataBase');
  if (!aba) throw new Error('Aba "examineedataBase" não encontrada.');

  var NUM_COLUNAS = 9; // id + 8 colunas de dados (schedulingDate ... appealRequestUrl)
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
  // "examinee", em vez de descartar silenciosamente.
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
 * "messages" (logo abaixo do cabeçalho), usando a Data-Hora como ID único.
 */
function gravarMensagem(ss, dadosMsg, urlArquivo) {
  var aba = ss.getSheetByName('messages');
  if (!aba) throw new Error('Aba "messages" não encontrada.');

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
// DATAS DE AGENDAMENTO (schedulingDates): PERÍODO JRS, DIAS ÚTEIS E FERIADOS
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
 * Grava as datas úteis calculadas na aba "schedulingDates" (coluna A),
 * preservando o status "active" (coluna C) de datas já existentes e sem
 * duplicar datas. A coluna B (weekDay) nunca é escrita pelo script: ela
 * já tem sua própria fórmula MAP+LAMBDA (configurada diretamente na
 * planilha) que preenche o dia da semana a partir da coluna A.
 */
function gravarDatasAgendamento(ss, diasUteis) {
  var aba = ss.getSheetByName('schedulingDates');
  if (!aba) throw new Error('Aba "schedulingDates" não encontrada.');

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
    aba.getRange(2, 1, ultimaLinha - 1, 1).clearContent();
    aba.getRange(2, 3, ultimaLinha - 1, 1).clearContent();
  }

  if (chaves.length > 0) {
    var linhasData = chaves.map(function(chave) {
      var p = chave.split('-');
      // Meio-dia (em vez de meia-noite) evita que o deslocamento de fuso
      // horário entre a interpretação UTC do runtime V8 e o fuso da
      // planilha (America/Recife) empurre a data para o dia anterior ao
      // ser exibida (ex.: segunda-feira aparecendo como domingo).
      return [new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0)];
    });
    var linhasAtivo = chaves.map(function(chave) {
      return [mapaAtivo[chave]];
    });

    aba.getRange(2, 1, linhasData.length, 1).setValues(linhasData);
    aba.getRange(2, 3, linhasAtivo.length, 1).setValues(linhasAtivo);
  }

  return chaves.length - totalAntes;
}
