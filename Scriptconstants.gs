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


//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimalinhaGerencial = abaGerencial.getLastRow();
const ultimaColunaGerencial = abaGerencial.getLastColumn();