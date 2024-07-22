// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Verificar se está cadastrada na planilha de interesse', 'VerificarInteresse')
		.addItem('Importar campos do Whatsapp preenchidos', 'ImportarEntrouWhats')
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
const colEstaNaInteresse = 13;
const colCadastradoWhats = 14;
const colEntrouWhatsInteresse = 13;

// Função que verificará se o email existe na planilha Interesse e retornará a linha
const RetornarLinhaEmail = (emailMarcoZero) => {
	//Conferir todos os emails da planilha Interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();

		if (emailMarcoZero == emailInteresse) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

//Função para verificar se a pessoa está cadastrada na planilha de Interesse
function VerificarInteresse() {
	//Pegar o email na planilha Marco Zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
		const celEstaNaInteresse = abaMarcoZero.getRange(i, colEstaNaInteresse);
		const valEstaNaInteresse = celEstaNaInteresse.getValue();

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailMarcoZero) {
			celEstaNaInteresse.setValue("");
			continue;
		}

		// Se o campo já estiver marcado com sim ou sessão pública
		if (valEstaNaInteresse == "SIM" || valEstaNaInteresse == "S. PÚBLICA") continue;

		if (RetornarLinhaEmail(emailMarcoZero))
			celEstaNaInteresse.setValue("SIM");
		else
			celEstaNaInteresse.setValue("S. PÚBLICA");
	}
}

// Função que pegará quem entrou no whatsapp pela planilha de interesse e colocará nessa planilha
function ImportarEntrouWhats() {
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
		const celCadastradoWhats = abaMarcoZero.getRange(i, colCadastradoWhats);

		// Se o campo estiver vazio, passe para o próximo
		if (!emailMarcoZero)
			continue;

		const linhaEmailEncontradoMarcoZero = RetornarLinhaEmail(emailMarcoZero);

		if (linhaEmailEncontradoMarcoZero) {
			const celEntrouWhatsInteresse = abaInteresse.getRange(linhaEmailEncontradoMarcoZero, colEntrouWhatsInteresse);
			const valEntrouWhatsInteresse = celEntrouWhatsInteresse.getValue();

			if (valEntrouWhatsInteresse == "SIM")
				celCadastradoWhats.setValue("SIM");
			else if (valEntrouWhatsInteresse == "NÃO")
				celCadastradoWhats.setValue("NÃO");
		}
	}
}
