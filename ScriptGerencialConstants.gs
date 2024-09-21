// Colunas gerais
const colNomeGeral = 3;
const colEmailGeral = 4;
const colTelGeral = 5;
// Colunas planilha Interesse
const colNomeInteresse = colNomeGeral;
const colEmailInteresse = colEmailGeral;
const colTelInteresse = colTelGeral;
const colCidadeInteresse = 8;
const colEstadoInteresse = 9;
const colWhatsInteresse = 13;
const colRespondeuMarcoZeroInteresse = 14;
const colSituacaoInteresse = 15;
// Colunas planilha Marco Zero
const colNomeMarcoZero = colNomeGeral;
const colEmailMarcoZero = colEmailGeral;
const colTelMarcoZero = colTelGeral;
const colRespondeuInteresseMarcoZero = 13;
const colWhatsMarcoZero = 14;
// Colunas planilha Gerencial
const colEmailsTelefonesAlternativosGerencial = 1;
const colTerminouCursoGerencial = 2;
const colNomeGerencial = colNomeGeral;
const colEmailGerencial = colEmailGeral;
const colTelGerencial = colTelGeral;
const colCidadeGerencial = 6;
const colEstadoGerencial = 7;
const colWhatsGerencial = 8;
const colRespondeuInteresseGerencial = 9;
const colRespondeuMarcoZeroGerencial = 10;
const colSituacaoGerencial = 11;
const colLinkMapaGerencial = 12;
const colTextoMapaGerencial = 13;
const colPrazoEnvioMapaGerencial = 14;
const colComentarioEnviadoMapaGerencial = 15;
const colMensagemVerificacaoMapaGerencial = 16;
const colRespondeuMarcoFinalGerencial = 17;
const colEnviouReflexaoMarcoFinalGerencial = 18;
const colPrazoEnvioMarcoFinalGerencial = 19;
const colComentarioEnviadoMarcoFinalGerencial = 20;
const colDataCertificadoGerencial = 21;
const colLinkCertificadoGerencial = 22;
const colLinkTestadoCertificadoGerencial = 23;
const colEntrouGrupoCertificadoGerencial = 24;
// Colunas planilha Envio Mapa
const colDataEnvioMapa = 1;
const colNomeEnvioMapa = colNomeGeral;
const colEmailEnvioMapa = colEmailGeral;
const colTelEnvioMapa = colTelGeral;
const colLinkMapa = 9;
const colTextoMapa = 10;
const colComentarioEnviadoMapa = 11;
const colPrazoEnvioMapa = 12;
const colMensagemVerificacaoMapa = 13;
// Colunas planilha Marco Final
const colNomeMarcoFinal = colNomeGeral;
const colEmailMarcoFinal = colEmailGeral;
const colTelMarcoFinal = colTelGeral;
const colEnviouReflexaoMarcoFinal = 13;
const colPrazoEnvioMarcoFinal = 14;
const colComentarioEnviadoMarcoFinal = 15;
// Colunas planilha Envio Certificado
const colNomeCertificado = colNomeGeral;
const colEmailCertificado = colEmailGeral;
const colTelCertificado = colTelGeral;
const colDataCertificado = 7;
const colLinkCertificado = 8;
const colLinkTestadoCertificado = 9;
const colEntrouGrupoCertificado = 10;


// Variáveis de otimização
const ultimaLinhaAnalisadaInteresse = 2;
const ultimaLinhaAnalisadaMarcoZero = 2;
const ultimaLinhaAnalisadaEnvioMapa = 2;
const ultimaLinhaAnalisadaMarcoFinal = 2;
const ultimaLinhaAnalisadaCertificado = 2;
const ultimaLinhaAnalisadaWhatsGerencial = 2;


// Seleciona a planilha de Confirmação de Interesse e a aba
const urlInteresse = "https://docs.google.com/spreadsheets/d/1TztdPoYhZ6t_ExftBE3gtugVyPSKRDx0IPLZTAJNmlI/edit?gid=320866237#gid=320866237";
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheets()[0]

// Seleciona a planilha do Marco Zero e a aba
const urlMarcoZero = "https://docs.google.com/spreadsheets/d/1--p65M1CNQlUz1vCLWovFWWqflwZVvTnMCeWm5mj3Gs/edit?gid=861556083#gid=861556083"
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheets()[0]

// Seleciona a planilha do Envio do Mapa e a aba
const urlEnvioMapa = "https://docs.google.com/spreadsheets/d/1FzBClWA5X2YvIkCDhMcKI8t8RnDDrzTAFJ7QJ-mADTA/edit?gid=0#gid=0"
const planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
const abaEnvioMapa = planilhaEnvioMapa.getSheets()[0]

// Seleciona a planilha do Marco Final e a aba
const urlMarcoFinal = "https://docs.google.com/spreadsheets/d/1vjuBOGuX3T0mIZX3Ac4dLFGcEBJTT2_YKjeD_fD24VI/edit?gid=0#gid=0"
const planilhaMarcoFinal = SpreadsheetApp.openByUrl(urlMarcoFinal);
const abaMarcoFinal = planilhaMarcoFinal.getSheets()[0]

// Seleciona a planilha do Envio do Certificado e a aba
const urlCertificado = "https://docs.google.com/spreadsheets/d/1vjuBOGuX3T0mIZX3Ac4dLFGcEBJTT2_YKjeD_fD24VI/edit?gid=0#gid=0"
const planilhaCertificado = SpreadsheetApp.openByUrl(urlCertificado);
const abaCertificado = planilhaCertificado.getSheets()[0]

// Seleciona a planilha Gerencial e a aba
const planilhaGerencial = SpreadsheetApp.getActiveSpreadsheet();
const abaGerencial = planilhaGerencial.getSheets()[0]

// Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimaLinhaEnvioMapa = abaEnvioMapa.getLastRow();
const ultimaLinhaMarcoFinal = abaMarcoFinal.getLastRow();
const ultimaLinhaCertificado = abaCertificado.getLastRow();
const ultimalinhaGerencial = abaGerencial.getLastRow();
const ultimaColunaGerencial = abaGerencial.getLastColumn();

// Variável genérica da planilha ativa
const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
const abaAtiva = planilhaAtiva.getSheets()[0]
const ultimaLinhaAtiva = abaAtiva.getLastRow();
const ultimaColunaAtiva = abaAtiva.getLastColumn();
const colEmailAtiva = colEmailGeral;
const colTelAtiva = colTelGeral;

// Objeto para mapear as variáveis para cada aba, para que seja possível utilizar o argumento da aba desejada
const objetoMap = new Map([
    [abaInteresse, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaInteresse,
        ultimaLinha: ultimaLinhaInteresse,
        colEmail: colEmailInteresse,
        ImportarDadosPlanilha: ImportarDadosInteresse
    }],
    [abaMarcoZero, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoZero,
        ultimaLinha: ultimaLinhaMarcoZero,
        colEmail: colEmailMarcoZero,
        ImportarDadosPlanilha: ImportarDadosMarcoZero
    }],
    [abaEnvioMapa, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaEnvioMapa,
        ultimaLinha: ultimaLinhaEnvioMapa,
        colEmail: colEmailEnvioMapa,
        ImportarDadosPlanilha: ImportarDadosEnvioMapa
    }],
    [abaMarcoFinal, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoFinal,
        ultimaLinha: ultimaLinhaMarcoFinal,
        colEmail: colEmailMarcoFinal,
        ImportarDadosPlanilha: ImportarDadosMarcoFinal
    }],
    [abaCertificado, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaCertificado,
        ultimaLinha: ultimaLinhaCertificado,
        colEmail: colEmailCertificado,
        ImportarDadosPlanilha: ImportarDadosCertificado
    }]
]);