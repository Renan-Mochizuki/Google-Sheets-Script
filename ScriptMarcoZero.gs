// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Verificar quem está cadastrado na planilha de interesse', 'VerificarInteresse')
		.addItem('Sincronizar campos do WhatsApp', 'SincronizarWhats')
		.addToUi();
}

onOpen();

//Seleciona a planilha de Confirmação de Interesse e a aba
const urlInteresse = "*";
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha do Marco Zero e a aba
const planilhaMarcoZero = SpreadsheetApp.getActiveSpreadsheet();
const abaMarcoZero = planilhaMarcoZero.getSheetByName("Respostas ao formulário 1");

//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();

// Colunas A,B,C,D...
const colNome = 3;
const colEmail = 4;
const colTel = 5;
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

		// Se o campo já estiver marcado passe para o próximo
		if (valEstaNaInteresse) continue;

		if (RetornarLinhaEmail(emailMarcoZero))
			celEstaNaInteresse.setValue("SIM");
		else
			celEstaNaInteresse.setValue("S. PÚBLICA");
	}
}

// Função que sincronizará quem entrou no whatsapp entre as duas planilhas
function SincronizarWhats() {
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const celWhatsInteresse = abaInteresse.getRange(i, colWhatsInteresse);
		const valWhatsInteresse = celWhatsInteresse.getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		const linhaCampoMarcoZero = RetornarLinhaEmail(emailInteresse);

		// Se o email for encontrado no Marco Zero
		if (linhaCampoMarcoZero) {
			const celWhatsMarcoZero = abaMarcoZero.getRange(linhaCampoMarcoZero, colWhatsMarcoZero);
			const valWhatsMarcoZero = celWhatsMarcoZero.getValue();

			// Se o campo do Marco Zero estiver vazio, altere o campo do Marco Zero com o valor do Interesse
			if (!valWhatsMarcoZero) {
				celWhatsMarcoZero.setValue(valWhatsInteresse);
				continue;
			}

			// Se o campo do Interesse estiver como sim e da outra como não, altere o campo do Marco Zero
			if (valWhatsInteresse == "SIM" && valWhatsMarcoZero == "NÃO") {
				celWhatsMarcoZero.setValue("SIM");
				continue;
			}

			// Se o campo do Marco Zero estiver como sim ou não (especificamente), altere o campo do Interesse
			if (valWhatsMarcoZero == "SIM")
				celWhatsInteresse.setValue("SIM");
			else if (valWhatsMarcoZero == "NÃO")
				celWhatsInteresse.setValue("NÃO");
		}
	}
}
