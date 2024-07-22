function onOpen(e) {
	// Add a custom menu to the spreadsheet.
	SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
		.createMenu('Custom Menu')
		.addItem('First item', 'menuItem1')
		.addToUi();
}

//Seleciona a planilha de Confirmação de Interesse e a aba
const planilhaInteresse = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1TztdPoYhZ6t_ExftBE3gtugVyPSKRDx0IPLZTAJNmlI/edit?gid=320866237#gid=320866237");
const abaInteresse = planilhaInteresse.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha do Marco Zero e a aba
const planilhaMarcoZero = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1--p65M1CNQlUz1vCLWovFWWqflwZVvTnMCeWm5mj3Gs/edit?gid=861556083#gid=861556083");
const abaMarcoZero = planilhaMarcoZero.getSheetByName("Respostas ao formulário 1");

//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();

// Colunas A,B,C,D...
const colNome = 2;
const colEmail = 3;
const colTel = 4;
const colRespondeuMarcoZero = 14;
const colFormularioEnviado = 16;

//Função para enviar o formulário do Marco Zero
function EnviarMarcoZero() { //Estamos na planilha de Confirmação de Interesse

	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const email = abaInteresse.getRange(i, colEmail).getValue();
		const celEnviadoMarcoZero = abaInteresse.getRange(i, colFormularioEnviado);
		const foiEnviado = celEnviadoMarcoZero.getValue();

		// Se o campo email estiver vazio, limpe a célula e passe para o próximo
		if (!email) {
			celEnviadoMarcoZero.setValue("");
			continue;
		}

		if (!foiEnviado || foiEnviado == "NÃO") {
			MailApp.sendEmail({
				to: `${email}`,
				subject: "Formulário Marco Zero",
				body: "Responda o formulário do Marco Zero para dar continuidade a sua formação em Mapas Conceituais. Link: https://forms.gle/YQdMCoemkDiumzyG6"
			})
			celEnviadoMarcoZero.setValue("SIM");
		}
	}
}

// Função que verificará se o email existe
function VerificarExistencia(emailInteresse) {
	//Conferir todos os emails da planilha Marco Zero
	for (let j = 2; j <= ultimaLinhaMarcoZero; j++) {
		const emailMarcoZero = abaMarcoZero.getRange(j, colEmail).getValue();

		if (emailInteresse == emailMarcoZero) return true;
	}
	// Se não for encontrado nenhum 
	return false;
}

//Função para verificar quem respondeu o Marco Zero
function VerificarMarcoZero() {
	//Pegar o email na planilha Interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const celRespondeuMarcoZero = abaInteresse.getRange(i, colRespondeuMarcoZero);

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailInteresse) {
			celRespondeuMarcoZero.setValue("");
			continue;
		}

		if (VerificarExistencia(emailInteresse))
			celRespondeuMarcoZero.setValue("SIM");
		else
			celRespondeuMarcoZero.setValue("NÃO");
	}
}

/*
function onOpen(e){
  let ambiente = SpreadsheetApp.getUi();
  ambiente.createMenu("Ações")
	 .addItem("Enviar Marco Zero", "enviarMarcoZero")
	 .addToUi();
  
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
		.createMenu('Custom Menu')
		.addItem('First item', 'menuItem1')
		.addToUi();
}
*/