// Função que recebe o nome da coluna e transforma em número (Ex.: A = 1; Z = 26; AA = 27; AB = 28)
function Coluna(letras) {
    let numero = 0;
    for (let i = 0; i < letras.length; i++) {
        numero = numero * 26 + (letras[i].toUpperCase().charCodeAt(0) - 64);
    }
    return numero;
}


// Colunas Gerais
const colNomeGeral = Coluna('C');
const colEmailGeral = Coluna('D');
const colTelGeral = Coluna('E');

// Colunas planilha Interesse
const colNomeInteresse = colNomeGeral;
const colEmailInteresse = colEmailGeral;
const colTelInteresse = colTelGeral;

const colCidadeInteresse = Coluna('H');
const colEstadoInteresse = Coluna('I');
const colWhatsInteresse = Coluna('M');
const colRespondeuMarcoZeroInteresse = Coluna('N');
const colSituacaoInteresse = Coluna('O');

// Colunas planilha Marco Zero
const colNomeMarcoZero = colNomeGeral;
const colEmailMarcoZero = colEmailGeral;
const colTelMarcoZero = colTelGeral;

const colRespondeuInteresseMarcoZero = Coluna('M');
const colWhatsMarcoZero = Coluna('N');

// Colunas planilha Gerencial
const colNomeGerencial = colNomeGeral;
const colEmailGerencial = colEmailGeral;
const colTelGerencial = colTelGeral;

const colEmailsTelefonesAlternativosGerencial = Coluna('A');
const colTerminouCursoGerencial = Coluna('B');
const colCidadeGerencial = Coluna('F');
const colEstadoGerencial = Coluna('G');
const colWhatsGerencial = Coluna('H');
const colRespondeuInteresseGerencial = Coluna('I');
const colRespondeuMarcoZeroGerencial = Coluna('J');
const colSituacaoGerencial = Coluna('K');
const colLinkMapaGerencial = Coluna('L');
const colTextoMapaGerencial = Coluna('M');
const colPrazoEnvioMapaGerencial = Coluna('N');
const colComentarioEnviadoMapaGerencial = Coluna('O');
const colMensagemVerificacaoMapaGerencial = Coluna('P');
const colRespondeuMarcoFinalGerencial = Coluna('Q');
const colEnviouReflexaoMarcoFinalGerencial = Coluna('R');
const colPrazoEnvioMarcoFinalGerencial = Coluna('S');
const colComentarioEnviadoMarcoFinalGerencial = Coluna('T');
const colDataCertificadoGerencial = Coluna('U');
const colLinkCertificadoGerencial = Coluna('V');
const colLinkTestadoCertificadoGerencial = Coluna('W');
const colEntrouGrupoCertificadoGerencial = Coluna('X');

// Colunas planilha Envio Mapa
const colNomeEnvioMapa = colNomeGeral;
const colEmailEnvioMapa = colEmailGeral;
const colTelEnvioMapa = colTelGeral;

const colDataEnvioMapa = Coluna('A');
const colLinkMapa = Coluna('I');
const colTextoMapa = Coluna('J');
const colComentarioEnviadoMapa = Coluna('K');
const colPrazoEnvioMapa = Coluna('L');
const colMensagemVerificacaoMapa = Coluna('M');

// Colunas planilha Marco Final
const colNomeMarcoFinal = colNomeGeral;
const colEmailMarcoFinal = colEmailGeral;
const colTelMarcoFinal = colTelGeral;

const colEnviouReflexaoMarcoFinal = Coluna('M');
const colPrazoEnvioMarcoFinal = Coluna('N');
const colComentarioEnviadoMarcoFinal = Coluna('O');

// Colunas planilha Envio Certificado
const colNomeCertificado = colNomeGeral;
const colEmailCertificado = colEmailGeral;
const colTelCertificado = colTelGeral;

const colDataCertificado = Coluna('G');
const colLinkCertificado = Coluna('H');
const colLinkTestadoCertificado = Coluna('I');
const colEntrouGrupoCertificado = Coluna('J');


// Variáveis de otimização (Possível futura implementação)
// Ideia: Armazenar a ultima linha analisada para reduzir o tamanho do loop, assim evitando analisar campos já analisados toda vez
const ultimaLinhaAnalisadaInteresse = 2;
const ultimaLinhaAnalisadaMarcoZero = 2;
const ultimaLinhaAnalisadaGerencial = 2;
const ultimaLinhaAnalisadaEnvioMapa = 2;
const ultimaLinhaAnalisadaMarcoFinal = 2;
const ultimaLinhaAnalisadaCertificado = 2;
const ultimaLinhaAnalisadaWhatsGerencial = 2;

// -- Links das planilhas no arquivo Links

// Seleciona a planilha de Confirmação de Interesse e a aba
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheets()[0]

// Seleciona a planilha do Marco Zero e a aba
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheets()[0]

// Seleciona a planilha do Envio do Mapa e a aba
const planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
const abaEnvioMapa = planilhaEnvioMapa.getSheets()[0]

// Seleciona a planilha do Marco Final e a aba
const planilhaMarcoFinal = SpreadsheetApp.openByUrl(urlMarcoFinal);
const abaMarcoFinal = planilhaMarcoFinal.getSheets()[0]

// Seleciona a planilha do Envio do Certificado e a aba
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
const ultimaLinhaGerencial = abaGerencial.getLastRow();
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
    [abaGerencial, {
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaGerencial,
        ultimaLinha: ultimaLinhaGerencial,
        colEmail: colEmailGerencial
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