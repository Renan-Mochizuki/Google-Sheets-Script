// AVISOS
// O código de escopo global (que não está dentro de uma função) é executado toda vez que um script inicia
// Por isso, é preciso tomar cuidado ao utilizar variáveis como ultimaLinha, pois ela não é atualizada durante
// a execução do script
// Nesse caso é necessário fazer aba.getLastRow() novamente na função

// ORDEM OBRIGATÓRIA DOS CAMPOS
// Para melhorar a performance, é necessário evitar chamando a função .getRange() repetidamente, por isso
// foi utilizado intervalos (arrays/matrizes), então os campos de certas planilhas devem seguir algumas regras de ordem descritas:
// (Caso houver uma mudança na ordem descrita abaixo, mudar nas funções da lógica de importação de cada planilha)
// Planilha Gerencial:
// -Nome, Email, Telefone, Cidade, Estado, Whats, RespondeuInteresse, RespondeuMarcoZero, Situacao
// -LinkMapa, TextoMapa, DataPrazoMapa, ComentarioEnviadoMapa, MensagemVerificacaoMapa
// -RespondeuMarcoFinal, EnviouReflexaoMarcoFinal, PrazoEnvioMarcoFinal,ComentarioEnviadoMarcoFinal
// -DataCertificado, LinkCertificado, LinkTestadoCertificado, EntrouGrupoCertificado
// Todas Planilhas: (Caso alguma planilha não seguir mais essa ordem, alterar VerificarRepetições)
// -Email, Telefone

// -- Variáveis do Constants --
// 	  Colunas, planilhas, abas, links,
//    estados, tempoNotificacao, corCampoSemDados
//    e objetoMap (utilizado para generalizar o código)

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

const colunasDeSimNao = [
  colTerminouCursoGerencial,
  colWhatsGerencial,
  colRespondeuInteresseGerencial,
  colRespondeuMarcoZeroGerencial,
  colComentarioEnviadoMapaGerencial,
  colRespondeuMarcoFinalGerencial,
  colEnviouReflexaoMarcoFinalGerencial,
  colComentarioEnviadoMarcoFinalGerencial,
  colLinkTestadoCertificadoGerencial,
  colEntrouGrupoCertificadoGerencial,
];

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
const abaInteresse = planilhaInteresse.getSheets()[0];
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheets()[0];
const planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
const abaEnvioMapa = planilhaEnvioMapa.getSheets()[0];
const planilhaMarcoFinal = SpreadsheetApp.openByUrl(urlMarcoFinal);
const abaMarcoFinal = planilhaMarcoFinal.getSheets()[0];
const planilhaCertificado = SpreadsheetApp.openByUrl(urlCertificado);
const abaCertificado = planilhaCertificado.getSheets()[0];
const planilhaGerencial = SpreadsheetApp.openByUrl(urlGerencial);
const abaGerencial = planilhaGerencial.getSheets()[0];

// Captura as últimas linhas e colunas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimaLinhaEnvioMapa = abaEnvioMapa.getLastRow();
const ultimaLinhaMarcoFinal = abaMarcoFinal.getLastRow();
const ultimaLinhaCertificado = abaCertificado.getLastRow();
// Apenas use essa variável uma vez a cada execução, pois ela não se atualiza sozinha
const ultimaLinhaGerencial = abaGerencial.getLastRow();
const ultimaColunaInteresse = abaInteresse.getLastColumn();
const ultimaColunaMarcoZero = abaMarcoZero.getLastColumn();
const ultimaColunaEnvioMapa = abaEnvioMapa.getLastColumn();
const ultimaColunaMarcoFinal = abaMarcoFinal.getLastColumn();
const ultimaColunaCertificado = abaCertificado.getLastColumn();
const ultimaColunaGerencial = abaGerencial.getLastColumn();

// Variável genérica da planilha ativa
const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
const abaAtiva = planilhaAtiva.getSheets()[0];
const ultimaLinhaAtiva = abaAtiva.getLastRow();
const ultimaColunaAtiva = abaAtiva.getLastColumn();
const colNomeAtiva = colNomeGeral;
const colEmailAtiva = colEmailGeral;
const colTelAtiva = colTelGeral;

// Objeto que permite generalizar o código, passando a aba para o objeto, assim extraindo as variáveis respectivas da aba
const objetoMap = new Map([
  [
    abaInteresse,
    {
      nomePlanilha: 'Interesse',
      url: urlInteresse,
      ultimaLinhaAnalisada: ultimaLinhaAnalisadaInteresse,
      ultimaLinha: ultimaLinhaInteresse,
      ultimaColuna: ultimaColunaInteresse,
      colNome: colNomeInteresse,
      colEmail: colEmailInteresse,
      colTel: colTelInteresse,
      colCidade: colCidadeInteresse,
      colEstado: colEstadoInteresse,
      ImportarDadosPlanilha: ImportarDadosInteresse,
    },
  ],
  [
    abaMarcoZero,
    {
      nomePlanilha: 'Marco Zero',
      url: urlMarcoZero,
      ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoZero,
      ultimaLinha: ultimaLinhaMarcoZero,
      ultimaColuna: ultimaColunaMarcoZero,
      colNome: colNomeMarcoZero,
      colEmail: colEmailMarcoZero,
      colTel: colTelMarcoZero,
      ImportarDadosPlanilha: ImportarDadosMarcoZero,
    },
  ],
  [
    abaEnvioMapa,
    {
      nomePlanilha: 'Envio do Mapa',
      url: urlEnvioMapa,
      ultimaLinhaAnalisada: ultimaLinhaAnalisadaEnvioMapa,
      ultimaLinha: ultimaLinhaEnvioMapa,
      ultimaColuna: ultimaColunaEnvioMapa,
      colNome: colNomeEnvioMapa,
      colEmail: colEmailEnvioMapa,
      colTel: colTelEnvioMapa,
      ImportarDadosPlanilha: ImportarDadosEnvioMapa,
    },
  ],
  [
    abaMarcoFinal,
    {
      nomePlanilha: 'Marco Final',
      url: urlMarcoFinal,
      ultimaLinhaAnalisada: ultimaLinhaAnalisadaMarcoFinal,
      ultimaLinha: ultimaLinhaMarcoFinal,
      ultimaColuna: ultimaColunaMarcoFinal,
      colNome: colNomeMarcoFinal,
      colEmail: colEmailMarcoFinal,
      colTel: colTelMarcoFinal,
      ImportarDadosPlanilha: ImportarDadosMarcoFinal,
    },
  ],
  [
    abaCertificado,
    {
      nomePlanilha: 'Envio do Certificado',
      url: urlCertificado,
      ultimaLinhaAnalisada: ultimaLinhaAnalisadaCertificado,
      ultimaLinha: ultimaLinhaCertificado,
      ultimaColuna: ultimaColunaCertificado,
      colNome: colNomeCertificado,
      colEmail: colEmailCertificado,
      colTel: colTelCertificado,
      ImportarDadosPlanilha: ImportarDadosCertificado,
    },
  ],
  [
    abaGerencial,
    {
      nomePlanilha: 'Gerencial',
      url: urlGerencial,
      ultimaLinha: ultimaLinhaGerencial,
      ultimaColuna: ultimaColunaGerencial,
      colNome: colNomeGerencial,
      colEmail: colEmailGerencial,
      colTel: colTelGerencial,
      colCidade: colCidadeGerencial,
      colEstado: colEstadoGerencial,
    },
  ],
  [
    abaAtiva,
    {
      nomePlanilha: 'Ativa',
      ultimaLinha: ultimaLinhaAtiva,
      ultimaColuna: ultimaColunaAtiva,
      colNome: colNomeAtiva,
      colEmail: colEmailAtiva,
      colTel: colTelAtiva,
    },
  ],
]);

// Função que remove acentos e normaliza uma string
function NormalizarString(str) {
  if (!str) return str;
  return str
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

// Função que calcula a distância de Levenshtein entre duas strings
function CalcularLevenshtein(str1, str2) {
  str1 = str1.toLowerCase();
  str2 = str2.toLowerCase();

  let costs = new Array();
  for (let i = 0; i <= str1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= str2.length; j++) {
      if (i == 0) costs[j] = j;
      else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (str1.charAt(i - 1) != str2.charAt(j - 1)) newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) costs[str2.length] = lastValue;
  }
  return costs[str2.length];
}

// Função que compara a similaridade entre duas strings
function CompararSimilaridade(str1, str2) {
  str1 = NormalizarString(str1);
  str2 = NormalizarString(str2);

  let longer = str1;
  let shorter = str2;
  if (str1.length < str2.length) {
    longer = str2;
    shorter = str1;
  }
  let longerLength = longer.length;

  if (longerLength == 0) return 1.0;

  const similarity = (longerLength - CalcularLevenshtein(longer, shorter)) / parseFloat(longerLength);

  if (similarity >= 0.5) return similarity;
  else return 0;
}

const estados = [
  'Acre - AC',
  'Alagoas - AL',
  'Amapá - AP',
  'Amazonas - AM',
  'Bahia - BA',
  'Ceará - CE',
  'Distrito Federal - DF',
  'Espírito Santo - ES',
  'Goiás - GO',
  'Maranhão - MA',
  'Mato Grosso - MT',
  'Mato Grosso do Sul - MS',
  'Minas Gerais - MG',
  'Pará - PA',
  'Paraíba - PB',
  'Paraná - PR',
  'Pernambuco - PE',
  'Piauí - PI',
  'Rio de Janeiro - RJ',
  'Rio Grande do Norte - RN',
  'Rio Grande do Sul - RS',
  'Rondônia - RO',
  'Roraima - RR',
  'Santa Catarina - SC',
  'São Paulo - SP',
  'Sergipe - SE',
  'Tocantins - TO',
  'Internacional',
];
