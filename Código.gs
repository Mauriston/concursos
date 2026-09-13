/**
 * Cria o menu personalizado assim que a planilha for aberta.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✏️ Termos')
    .addItem('🛑 Cientificação de Recurso', 'iniciarGeracaoRecursos') // NOVO ITEM ADICIONADO
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


// =========================================================================
// GERAÇÃO DE TERMOS DE CIENTIFICAÇÃO EM LOTE (PDF ÚNICO)
// =========================================================================

function iniciarGeracaoTermos() {
  var htmlTemplate = HtmlService.createTemplateFromFile('Alerta');
  htmlTemplate.titulo = 'Gerar Termos de Cientificação';
  htmlTemplate.mensagem = 'Deseja gerar os Termos de Cientificação para todos os candidatos da aba "Dados Pessoais"?';
  htmlTemplate.tipo = 'termos_confirmacao';
  
  var htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(400).setTitle('Inspeção de Saúde');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function processarGeracaoTermos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var abaPrincipal = ss.getSheetByName('Principal');
    var abaDadosPessoais = ss.getSheetByName('Dados Pessoais');
    
    var nomeConcursoBruto = abaPrincipal.getRange('F3').getValue() || "NÃO INFORMADO";
    var nomeConcurso = String(nomeConcursoBruto).replace(/CONCURSO\s+/i, '');

    var dadosPrincipal = abaPrincipal.getRange('A12:C' + abaPrincipal.getLastRow()).getValues();
    var mapaDatas = {};
    for (var i = 0; i < dadosPrincipal.length; i++) {
      var data = dadosPrincipal[i][0]; 
      var insc = String(dadosPrincipal[i][2]).trim(); 
      if (insc) mapaDatas[insc] = formatarDataSimples(data);
    }

    var dadosPessoais = abaDadosPessoais.getDataRange().getValues();
    var cabecalhoDP = dadosPessoais[0];
    
    var idxInsc = cabecalhoDP.indexOf("Inscrição");
    var idxNome = cabecalhoDP.indexOf("Nome");
    var idxId = cabecalhoDP.indexOf("Identidade");
    var idxEmissor = cabecalhoDP.indexOf("Emissor");

    if (idxInsc === -1 || idxNome === -1) {
      mostrarAlertaGenerico("Erro", "Colunas 'Inscrição' ou 'Nome' não encontradas na aba 'Dados Pessoais'.");
      return;
    }

    var idTemplate = '1uSdvvwxjPDkvGm7kjCrDGaGnPVF3_pDu4rN6C6tJFiI';
    var tempFile = DriveApp.getFileById(idTemplate).makeCopy("Temp_Termos_" + new Date().getTime());
    
    var docTemp = DocumentApp.openById(tempFile.getId());
    var bodyTemp = docTemp.getBody();
    
    var docTemplate = DocumentApp.openById(idTemplate);
    var bodyTemplate = docTemplate.getBody();
    
    var candidatosProcessados = 0;
    var primeiroCandidato = true;

    for (var r = 1; r < dadosPessoais.length; r++) {
      var linha = dadosPessoais[r];
      var candInsc = String(linha[idxInsc]).trim();
      
      if (!candInsc) continue; 
      
      var candNome = linha[idxNome];
      var candId = linha[idxId];
      var candEmissor = linha[idxEmissor];
      var candData = mapaDatas[candInsc] || "___/___/_____"; 
      
      if (!primeiroCandidato) {
        bodyTemp.appendPageBreak();
        for (var j = 0; j < bodyTemplate.getNumChildren(); j++) {
          var elemento = bodyTemplate.getChild(j).copy();
          var tipo = elemento.getType();
          if (tipo === DocumentApp.ElementType.PARAGRAPH) bodyTemp.appendParagraph(elemento);
          else if (tipo === DocumentApp.ElementType.TABLE) bodyTemp.appendTable(elemento);
          else if (tipo === DocumentApp.ElementType.LIST_ITEM) bodyTemp.appendListItem(elemento);
        }
      }

      bodyTemp.replaceText("\\{\\{Nome\\}\\}", candNome);
      bodyTemp.replaceText("\\{\\{Concurso\\}\\}", nomeConcurso);
      bodyTemp.replaceText("\\{\\{Inscrição\\}\\}", candInsc);
      bodyTemp.replaceText("\\{\\{Identidade\\}\\}", candId);
      bodyTemp.replaceText("\\{\\{Emissor\\}\\}", candEmissor);
      bodyTemp.replaceText("\\{\\{Data\\}\\}", candData);

      primeiroCandidato = false;
      candidatosProcessados++;
    }

    docTemp.saveAndClose(); 

    if (candidatosProcessados === 0) {
      tempFile.setTrashed(true);
      mostrarAlertaGenerico("Aviso", "Nenhum candidato válido encontrado.");
      return;
    }

    var idPastaDestino = '1BC4YZRU-vS7C-0826QRJprlRb9ITiN39';
    var pastaDestino = DriveApp.getFolderById(idPastaDestino);
    
    var nomeArquivoFinal = "Termos Cientificacao " + nomeConcurso + ".pdf";
    var pdfBlob = tempFile.getAs("application/pdf");
    pdfBlob.setName(nomeArquivoFinal);
    
    var novoPdf = pastaDestino.createFile(pdfBlob);
    var urlArquivo = novoPdf.getUrl();

    tempFile.setTrashed(true);

    var mensagemSucesso = 'O ficheiro PDF com ' + candidatosProcessados + ' termos foi gerado e salvo com sucesso!<br><br><b>Documento:</b> <a href="' + urlArquivo + '" target="_blank">' + nomeArquivoFinal + '</a><br><br><i>Clique no link acima para abrir o ficheiro.</i>';
    mostrarAlertaGenerico("Processo Concluído", mensagemSucesso);

  } catch (erro) {
    mostrarAlertaGenerico("Erro", "Ocorreu um erro ao gerar os termos: <br><br>" + erro.message);
  }
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
