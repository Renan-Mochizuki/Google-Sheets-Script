//Seleciona a planilha de Confirmação de Interesse e a aba
const urlInteresse = "*";
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha do Marco Zero e a aba
const urlMarcoZero = "*"
const planilhaMarcoZero = SpreadsheetApp.openByUrl(urlMarcoZero);
const abaMarcoZero = planilhaMarcoZero.getSheetByName("Respostas ao formulário 1");

//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();

// Colunas A,B,C,D...
const colNome = 2;
const colEmail = 3;
const colTel = 4;
const colEstaNaInteresse = 13;

// Função que verificará se o email existe
function VerificarExistencia(emailMarcoZero) {
	//Conferir todos os emails da planilha Interesse
	for (let j = 2; j <= ultimaLinhaInteresse; j++) {
		const emailInteresse = abaInteresse.getRange(j, colEmail).getValue();

		if (emailMarcoZero == emailInteresse) return true;
	}
	// Se não for encontrado nenhum 
	return false;
}

function VerificarInteresse() {
	//Pegar o email na planilha Marco Zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
		const celEstaNaInteresse = abaMarcoZero.getRange(i, colEstaNaInteresse);

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailMarcoZero) {
			celEstaNaInteresse.setValue("");
			continue;
		}

		if (VerificarExistencia(emailMarcoZero))
			celEstaNaInteresse.setValue("SIM");
		else
			celEstaNaInteresse.setValue("NÃO");
	}
}