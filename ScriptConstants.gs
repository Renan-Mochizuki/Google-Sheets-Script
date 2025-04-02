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
const colTerminouCursoMapa = Coluna('N');

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
const colRedirectInteresseGerencial = Coluna('AA');
const colRedirectMarcoZeroGerencial = Coluna('AB');
const colRedirectEnvioMapaGerencial = Coluna('AC');
const colRedirectMarcoFinalGerencial = Coluna('AD');
const colRedirectCertificadoGerencial = Coluna('AE');

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

let planilhaInteresse, planilhaMarcoZero, planilhaEnvioMapa, planilhaMarcoFinal, planilhaCertificado, planilhaGerencial;

// Seleciona as planilhas e a aba
// Usando try devido a erro a abrir a planilha (o trigger onOpen(e) automático não consegue executar openByUrl)
try {
  planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
  planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
  planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
  planilhaMarcoFinal = SpreadsheetApp.openByUrl(urlMarcoFinal);
  planilhaCertificado = SpreadsheetApp.openByUrl(urlCertificado);
  planilhaGerencial = SpreadsheetApp.openByUrl(urlGerencial);
} catch {}

const abaInteresse = planilhaInteresse?.getSheets()[0];
const abaMarcoZero = planilhaMarcoZero?.getSheets()[0];
const abaEnvioMapa = planilhaEnvioMapa?.getSheets()[0];
const abaMarcoFinal = planilhaMarcoFinal?.getSheets()[0];
const abaCertificado = planilhaCertificado?.getSheets()[0];
const abaGerencial = planilhaGerencial?.getSheets()[0];

// Captura as últimas linhas e colunas
const ultimaLinhaInteresse = abaInteresse?.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero?.getLastRow();
const ultimaLinhaEnvioMapa = abaEnvioMapa?.getLastRow();
const ultimaLinhaMarcoFinal = abaMarcoFinal?.getLastRow();
const ultimaLinhaCertificado = abaCertificado?.getLastRow();
// Apenas use essa variável uma vez a cada execução, pois ela não se atualiza sozinha
const ultimaLinhaGerencial = abaGerencial?.getLastRow();
const ultimaColunaInteresse = abaInteresse?.getLastColumn();
const ultimaColunaMarcoZero = abaMarcoZero?.getLastColumn();
const ultimaColunaEnvioMapa = abaEnvioMapa?.getLastColumn();
const ultimaColunaMarcoFinal = abaMarcoFinal?.getLastColumn();
const ultimaColunaCertificado = abaCertificado?.getLastColumn();
const ultimaColunaGerencial = abaGerencial?.getLastColumn();

// Variável genérica da planilha ativa
const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
const abaAtiva = planilhaAtiva.getSheets()[0];
const ultimaLinhaAtiva = abaAtiva.getLastRow();
const ultimaColunaAtiva = abaAtiva.getLastColumn();
const colNomeAtiva = colNomeGeral;
const colEmailAtiva = colEmailGeral;
const colTelAtiva = colTelGeral;
const colCidadeAtiva = colCidadeGerencial;
const colEstadoAtiva = colEstadoGerencial;

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
      colCidade: colCidadeAtiva,
      colEstado: colEstadoAtiva,
    },
  ],
]);

// Função que separa os dados pelo ;
function SepararDados(dadosMultiplos) {
  if (dadosMultiplos) {
    return NormalizarString(dadosMultiplos).split(';');
  } else {
    return [dadosMultiplos];
  }
}

// Função que verifica se os campos passados do loop atual são válidos args = (nome, email, telefone, outros...)
function ValidarLoop(...args) {
  // Se o valor existe, e caso for string, não contém a palavra 'teste'
  const nomeValido = args[0] && (typeof args[0] === 'string' ? !args[0].toLowerCase().includes('teste') : true);
  const emailValido = args[1] && (typeof args[1] === 'string' ? !args[1].toLowerCase().includes('teste') : true);
  const telefoneValido = args[2] && (typeof args[2] === 'string' ? !args[2].toLowerCase().includes('teste') : true);

  // Retorna falso se qualquer um dos campos for inválido
  if (!nomeValido || !emailValido || !telefoneValido) return false;

  // Se outro parametro foi passado, verifique se qualquer um for nulo, retorne falso
  for (let i = 3; i < args.length; i++) {
    if (!args[i]) return false;
  }

  return true;
}

// Função que verificará se o email existe na planilha desejada e retornará a linha
function RetornarLinhaDados(nomeProcurado, emailProcurado, telefoneProcurado, dados) {
  // Separando os dados procurados pois ele pode ser um valor com mais de um email
  const nomesProcurados = SepararDados(nomeProcurado);
  const emailsProcurados = SepararDados(emailProcurado);
  const telefonesProcurados = SepararDados(telefoneProcurado);

  // dados é uma matriz, na qual possui as colunas nome, email, telefone e cada linha é um cadastro

  // Conferir cada linha da matriz dos dados
  for (let i = 0; i < dados.length; i++) {
    const nomeDados = dados[i][0];
    const emailDados = dados[i][1];
    const telefoneDados = dados[i][2];

    const similaridadeNome = VerificarLinhaDados(nomeDados, nomesProcurados);
    const similaridadeEmail = VerificarLinhaDados(emailDados, emailsProcurados);
    const similaridadeTelefone = VerificarLinhaDados(telefoneDados, telefonesProcurados);

    // Se (email e telefone forem iguais) ou (email e nome forem iguais, tendo que o telefone não é caso especial) ou (telefone e nome forem iguais)
    if (
      (similaridadeEmail >= 0.8 && similaridadeTelefone >= 0.8) ||
      (similaridadeEmail >= 0.8 && similaridadeNome >= 0.5 && similaridadeTelefone !== -1) ||
      (similaridadeTelefone >= 0.9 && similaridadeNome >= 0.6)
    ) {
      return i + 2; // Retorne o índice da array + 2 (Porque a array começa em 0 e a planilha em 2)
    }
  }
  // Se não for encontrado nenhum
  return false;
}

// Função genérica para verificar se os dados são iguais (com uma certa tolerância)
function VerificarLinhaDados(dados, valoresProcurados) {
  if (!dados) return false;

  for (let dadoPlanilha of dados.toString().split(';')) {
    // Se o dado passado for um email, retire o domínio (Exemplo: @gmail.com)
    if (dadoPlanilha.includes('@')) {
      dadoPlanilha = dadoPlanilha.split('@')[0];
    }

    for (let valorProcurado of valoresProcurados) {
      // Se o valor procurado for um email, retire o domínio (Exemplo: @gmail.com)
      if (valorProcurado.includes('@')) {
        valorProcurado = valorProcurado.split('@')[0];
      }
      const similaridade = CompararSimilaridade(valorProcurado, dadoPlanilha);

      // Caso especifico com telefone (pois foi achado um dado que falhava na verificação comum)
      // O telefone procurado é diferente e o um telefone apenas possui o +
      if (similaridade < 0.8 && valorProcurado.includes('+') != dadoPlanilha.includes('+')) {
        return -1;
      }

      // Se o valor procurado e o dado bruto forem iguais, retorne true
      return similaridade;
    }
  }
  return false;
}

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

// Função que retorna a turma mais recente (Ex: T04-2024 > T01-2024 > T02-2023)
function RetornarTurmaMaisRecente(string1, string2) {
  // Se uma string não existir, ou for 'ESPERA', retorne a outra
  if (!string1 || string1 === 'ESPERA') return string2;
  if (!string2 || string2 === 'ESPERA') return string1;

  // Separar os números antes do traço (T01-2024 => 01; 2024)
  const regex = /(\d+)-(\d+)/;
  const match1 = string1.match(regex);
  const match2 = string2.match(regex);

  if (!match1 || !match2) return string1;

  // Verificar qual ano é maior
  if (match1[2] > match2[2]) return string1;
  if (match1[2] < match2[2]) return string2;

  // Anos iguais, verificar qual turma é maior
  if (match1[1] > match2[1]) return string1;
  if (match1[1] < match2[1]) return string2;

  return string1;
}

// Função que compara dois valores de SIM ou NÃO e retorna o valor correto (priorizando o SIM)
function RetornarValorSimNao(valor1, valor2) {
  if (!valor1) return valor2;
  if (!valor2) return valor1;
  if (valor1 == 'SIM' || valor2 == 'SIM') return 'SIM';
  return valor1;
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
];
