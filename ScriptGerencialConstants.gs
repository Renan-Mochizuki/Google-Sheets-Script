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
const colAnotacaoInteresse = Coluna('P');

// Colunas planilha Marco Zero
const colNomeMarcoZero = colNomeGeral;
const colEmailMarcoZero = colEmailGeral;
const colTelMarcoZero = colTelGeral;

const colRespondeuInteresseMarcoZero = Coluna('M');
const colWhatsMarcoZero = Coluna('N');

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

// Colunas planilha Gerencial
const colNomeGerencial = colNomeGeral;
const colEmailGerencial = colEmailGeral;
const colTelGerencial = colTelGeral;

const colAnotacaoGerencial = Coluna('A');
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
const colRedirectInteresseGerencial = Coluna('Y');
const colRedirectMarcoZeroGerencial = Coluna('Z');
const colRedirectEnvioMapaGerencial = Coluna('AA');
const colRedirectMarcoFinalGerencial = Coluna('AB');
const colRedirectCertificadoGerencial = Coluna('AC');

// Outras variáveis
const tempoNotificacao = 30;
const corCampoSemDados = '#ababab';

// Variáveis de otimização (Possível futura implementação)
// Ideia: Armazenar a ultima linha analisada para reduzir o tamanho do loop, assim evitando analisar toda vez campos já analisados
// Essa ideia requer muito cuidado
const ultimaLinhaAnalisadaInteresse = 2;
const ultimaLinhaAnalisadaMarcoZero = 2;
const ultimaLinhaAnalisadaEnvioMapa = 2;
const ultimaLinhaAnalisadaMarcoFinal = 2;
const ultimaLinhaAnalisadaCertificado = 2;
const ultimaLinhaAnalisadaWhatsGerencial = 2;

// -- Links das planilhas estão no arquivo Links

// Seleciona as planilhas e a aba
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheets()[0]
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheets()[0]
const planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
const abaEnvioMapa = planilhaEnvioMapa.getSheets()[0]
const planilhaMarcoFinal = SpreadsheetApp.openByUrl(urlMarcoFinal);
const abaMarcoFinal = planilhaMarcoFinal.getSheets()[0]
const planilhaCertificado = SpreadsheetApp.openByUrl(urlCertificado);
const abaCertificado = planilhaCertificado.getSheets()[0]
const planilhaGerencial = SpreadsheetApp.getActiveSpreadsheet();
const abaGerencial = planilhaGerencial.getSheets()[0]

// Captura as últimas linhas e colunas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimaLinhaEnvioMapa = abaEnvioMapa.getLastRow();
const ultimaLinhaMarcoFinal = abaMarcoFinal.getLastRow();
const ultimaLinhaCertificado = abaCertificado.getLastRow();
const ultimaLinhaGerencial = abaGerencial.getLastRow();
const ultimaColunaInteresse = abaInteresse.getLastColumn();
const ultimaColunaMarcoZero = abaMarcoZero.getLastColumn();
const ultimaColunaEnvioMapa = abaEnvioMapa.getLastColumn();
const ultimaColunaMarcoFinal = abaMarcoFinal.getLastColumn();
const ultimaColunaCertificado = abaCertificado.getLastColumn();
const ultimaColunaGerencial = abaGerencial.getLastColumn();

// Variável genérica da planilha ativa
const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
const abaAtiva = planilhaAtiva.getSheets()[0]
const ultimaLinhaAtiva = abaAtiva.getLastRow();
const ultimaColunaAtiva = abaAtiva.getLastColumn();
const colNomeAtiva = colNomeGeral;
const colEmailAtiva = colEmailGeral;
const colTelAtiva = colTelGeral;

// Objeto que permite generalizar o código, passando a aba para o objeto, assim extraindo as variáveis respectivas da aba
const objetoMap = new Map([
    [abaInteresse, {
        nome: 'Interesse',
        url: urlInteresse,
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaInteresse,
        ultimaLinha: ultimaLinhaInteresse,
        ultimaColuna: ultimaColunaInteresse,
        colNome: colNomeInteresse,
        colEmail: colEmailInteresse,
        colTel: colTelInteresse,
        colCidade: colCidadeInteresse,
        colEstado: colEstadoInteresse,
        ImportarDadosPlanilha: ImportarDadosInteresse
    }],
    [abaMarcoZero, {
        nome: 'Marco Zero',
        url: urlMarcoZero,
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoZero,
        ultimaLinha: ultimaLinhaMarcoZero,
        ultimaColuna: ultimaColunaMarcoZero,
        colNome: colNomeMarcoZero,
        colEmail: colEmailMarcoZero,
        colTel: colTelMarcoZero,
        ImportarDadosPlanilha: ImportarDadosMarcoZero
    }],
    [abaEnvioMapa, {
        nome: 'Envio do Mapa',
        url: urlEnvioMapa,
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaEnvioMapa,
        ultimaLinha: ultimaLinhaEnvioMapa,
        ultimaColuna: ultimaColunaEnvioMapa,
        colNome: colNomeEnvioMapa,
        colEmail: colEmailEnvioMapa,
        colTel: colTelEnvioMapa,
        ImportarDadosPlanilha: ImportarDadosEnvioMapa
    }],
    [abaMarcoFinal, {
        nome: 'Marco Final',
        url: urlMarcoFinal,
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoFinal,
        ultimaLinha: ultimaLinhaMarcoFinal,
        ultimaColuna: ultimaColunaMarcoFinal,
        colNome: colNomeMarcoFinal,
        colEmail: colEmailMarcoFinal,
        colTel: colTelMarcoFinal,
        ImportarDadosPlanilha: ImportarDadosMarcoFinal
    }],
    [abaCertificado, {
        nome: 'Envio do Certificado',
        url: urlCertificado,
        ultimaLinhaAnalisada: ultimaLinhaAnalisadaCertificado,
        ultimaLinha: ultimaLinhaCertificado,
        ultimaColuna: ultimaColunaCertificado,
        colNome: colNomeCertificado,
        colEmail: colEmailCertificado,
        colTel: colTelCertificado,
        ImportarDadosPlanilha: ImportarDadosCertificado
    }],
    [abaGerencial, {
        nome: 'Gerencial',
        url: urlGerencial,
        ultimaLinha: ultimaLinhaGerencial,
        ultimaColuna: ultimaColunaGerencial,
        colNome: colNomeGerencial,
        colEmail: colEmailGerencial,
        colTel: colTelGerencial,
        colCidade: colCidadeGerencial,
        colEstado: colEstadoGerencial
    }],
    [abaAtiva, {
        nome: 'Ativa',
        ultimaLinha: ultimaLinhaAtiva,
        ultimaColuna: ultimaColunaAtiva,
        colNome: colNomeAtiva,
        colEmail: colEmailAtiva,
        colTel: colTelAtiva
    }]
]);

const estados = [
    "Amapá - AP",
    "Amazonas - AM",
    "Bahia - BA",
    "Ceará - CE",
    "Distrito Federal - DF",
    "Espírito Santo - ES",
    "Goiás - GO",
    "Maranhão - MA",
    "Mato Grosso - MT",
    "Mato Grosso do Sul - MS",
    "Minas Gerais - MG",
    "Pará - PA",
    "Paraíba - PB",
    "Paraná - PR",
    "Pernambuco - PE",
    "Piauí - PI",
    "Rio de Janeiro - RJ",
    "Rio Grande do Norte - RN",
    "Rio Grande do Sul - RS",
    "Rondônia - RO",
    "Roraima - RR",
    "Santa Catarina - SC",
    "São Paulo - SP",
    "Sergipe - SE",
    "Tocantins - TO",
    "Internacional"
];