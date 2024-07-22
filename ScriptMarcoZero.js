// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Verificar se está cadastrada na planilha de interesse', 'VerificarInteresse')
		.addItem('Sincronizar campos WhatsApp preenchidos', 'ImportarEntrouWhats')
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
const colWhatsMarcoZero = 14;
const colWhatsInteresse = 13;

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
		const celWhatsMarcoZero = abaMarcoZero.getRange(i, colWhatsMarcoZero);
		const valWhatsMarcoZero = celWhatsMarcoZero.getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		const linhaCampoMarcoZero = RetornarLinhaEmail(emailMarcoZero);

		// Se o email for encontrado na outra planilha
		if (linhaCampoMarcoZero) {
			const celWhatsInteresse = abaInteresse.getRange(linhaCampoMarcoZero, colWhatsInteresse);
			const valWhatsInteresse = celWhatsInteresse.getValue();

			// Se o campo dessa planilha estiver como sim e da outra como não, altere o campo da outra planilha
			if (valWhatsMarcoZero == "SIM" && valWhatsInteresse == "NÃO") {
				celWhatsInteresse.setValue("SIM");
				continue;
			}

			// Se o campo da outra planilha estiver como sim ou não especificamente, altere o campo dessa planilha
			if (valWhatsInteresse == "SIM")
				celWhatsMarcoZero.setValue("SIM");
			else if (valWhatsInteresse == "NÃO")
				celWhatsMarcoZero.setValue("NÃO");
		}
	}
}
