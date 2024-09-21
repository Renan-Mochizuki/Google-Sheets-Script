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
const colNomeGerencial = colNomeGeral;
const colEmailGerencial = colEmailGeral;
const colTelGerencial = colTelGeral;
const colCidadeGerencial = 6;
const colEstadoGerencial = 7;
const colWhatsGerencial = 8;
const colRespondeuInteresseGerencial = 9;
const colRespondeuMarcoZeroGerencial = 10;
const colSituacaoGerencial = 11;
// Colunas planilha Envio Mapa


// Variáveis de otimização
const ultimaLinhaAnalisadaInteresse = 2;
const ultimaLinhaAnalisadaMarcoZero = 2
const ultimaLinhaAnalisadaWhatsGerencial = 2;


//Seleciona a planilha de Confirmação de Interesse e a aba
const urlInteresse = "https://docs.google.com/spreadsheets/d/1TztdPoYhZ6t_ExftBE3gtugVyPSKRDx0IPLZTAJNmlI/edit?gid=320866237#gid=320866237";
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha do Marco Zero e a aba
const urlMarcoZero = "https://docs.google.com/spreadsheets/d/1--p65M1CNQlUz1vCLWovFWWqflwZVvTnMCeWm5mj3Gs/edit?gid=861556083#gid=861556083"
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha Gerencial e a aba
const planilhaGerencial = SpreadsheetApp.getActiveSpreadsheet();
const abaGerencial = planilhaGerencial.getSheetByName("Gerencial");

//Seleciona a planilha do Marco Zero e a aba
const urlEnvioMapa = "https://docs.google.com/spreadsheets/d/1FzBClWA5X2YvIkCDhMcKI8t8RnDDrzTAFJ7QJ-mADTA/edit?gid=0#gid=0"
const planilhaEnvioMapa = SpreadsheetApp.openByUrl(urlEnvioMapa);
const abaEnvioMapa = planilhaEnvioMapa.getSheetByName("Respostas ao formulário 1");

//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimalinhaGerencial = abaGerencial.getLastRow();
const ultimaColunaGerencial = abaGerencial.getLastColumn();