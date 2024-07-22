// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Verificar se respondeu o Marco Zero', 'VerificarMarcoZero')
		.addItem('Sincronizar campos WhatsApp preenchidos', 'ImportarEntrouWhats')
		.addItem('Enviar Marco Zero por Email', 'EnviarMarcoZero')
		.addToUi();
}

onOpen();

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
const colRespondeuMarcoZero = 14;
const colFormularioEnviado = 16;
const colWhatsInteresse = 13;
const colWhatsMarcoZero = 14;

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

// Função que verificará se o email existe na planilha Marco Zero e retornará a linha
const RetornarLinhaEmail = (emailInteresse) => {
	//Conferir todos os emails da planilha Marco Zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();

		if (emailInteresse == emailMarcoZero) return i;
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
		const valRespondeuMarcoZero = celRespondeuMarcoZero.getValue();

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailInteresse) {
			celRespondeuMarcoZero.setValue("");
			continue;
		}

		// Se o campo já estiver marcado com sim
		if (valRespondeuMarcoZero == "SIM") continue;

		if (RetornarLinhaEmail(emailInteresse))
			celRespondeuMarcoZero.setValue("SIM");
		else
			celRespondeuMarcoZero.setValue("NÃO");
	}
}

// Função que pegará quem entrou no whatsapp pela planilha de interesse e colocará nessa planilha
function ImportarEntrouWhats() {
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const celWhatsInteresse = abaInteresse.getRange(i, colWhatsInteresse);
		const valWhatsInteresse = celWhatsInteresse.getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		const linhaCampoMarcoZero = RetornarLinhaEmail(emailInteresse);

		// Se o email for encontrado na outra planilha
		if (linhaCampoMarcoZero) {
			const celWhatsMarcoZero = abaMarcoZero.getRange(linhaCampoMarcoZero, colWhatsMarcoZero);
			const valWhatsMarcoZero = celWhatsMarcoZero.getValue();

			// Se o campo dessa planilha estiver como sim e da outra como não, altere o campo da outra planilha
			if (valWhatsInteresse == "SIM" && valWhatsMarcoZero == "NÃO") {
				celWhatsMarcoZero.setValue("SIM");
				continue;
			}

			// Se o campo da outra planilha estiver como sim ou não especificamente, altere o campo dessa planilha
			if (valWhatsMarcoZero == "SIM")
				celWhatsInteresse.setValue("SIM");
			else if (valWhatsMarcoZero == "NÃO")
				celWhatsInteresse.setValue("NÃO");
		}
	}
}
