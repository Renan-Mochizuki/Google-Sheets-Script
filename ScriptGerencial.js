// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Importar Dados', 'ImportarDados')
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

//Seleciona a planilha Gerencial e a aba
const urlGerencial = "*";
const planilhaGerencial = SpreadsheetApp.openByUrl(urlGerencial);
const abaGerencial = planilhaGerencial.getSheetByName("Gerencial");

//Captura as últimas linhas
const ultimaLinhaInteresse = abaInteresse.getLastRow();
const ultimaLinhaMarcoZero = abaMarcoZero.getLastRow();
const ultimalinhaGerencial = abaGerencial.getLastRow();

// Colunas gerais
const colNome = 3;
const colEmail = 4;
const colTel = 5;
const colConfirmacaoTel = 6;
// Colunas planilha Interesse
const colCidadeInteresse = 8;
const colEstadoInteresse = 9;
const colWhatsInteresse = 13;
const colRespondeuMarcoZeroInteresse = 14;
const colSituacaoInteresse = 15;
const colFormEnviadoInteresse = 16;
const colDataNascInteresse = 7;
// Colunas planilha Marco Zero
const colRespondeuInteresseMarcoZero = 13;
const colWhatsMarcoZero = 14;
// Colunas planilha Gerencial
const colWhatsGerencial = 8;
const colRespondeuInteresseGerencial = 9;
const colRespondeuMarcoZeroGerencial = 10;
const colSituacaoGerencial = 11;
const colFormEnviadoGerencial = 12;
const colDataNascGerencial = 2;
const colCidadeGerencial = 6;
const colEstadoGerencial = 6;

//Função para verificar se o email já existe no Gerencial.
function VerRepeticao(emailVerificar) {
	let i = 2;
	do {
		const emailGerencial = abaGerencial.getRange(i, colEmail).getValue();
		if (emailVerificar == emailGerencial) return false; //Já existe esse email no Gerencial.
		i++;
	} while (i <= ultimalinhaGerencial)
	return true; //Não existe esse email no Gerenciameto.
}

// Função que verificará se o email existe na planilha Gerencial e retornará a linha
const RetornarLinhaEmailGerencial = (emailInformado) => {
	//Conferir todos os emails da planilha Gerencial
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const emailGerencial = abaGerencial.getRange(i, colEmail).getValue();

		if (emailInformado == emailGerencial) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

//Retorna a linha em que o campo do email está vazio
function RetornarEspacoVazio() {
	let i = 2;
	do {
		const celEmailGerencial = abaGerencial.getRange(i, colEmail).getValue();
		if (!celEmailGerencial) return i;
		i++;
	} while (i != 0)
}

// Função que importa dados da planilha interesse e do marco zero que não estão na de interesse
function ImportarDados() {
	let linhaVazia = RetornarEspacoVazio();

	SincronizarWhatsInteresse();
	VerificarMarcoZeroInteresse()

	// Loop da planilha interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailInteresse);

		// Se aquele email ainda não estiver registrado na planilha gerencial
		if (!linhaCampoGerencial) {
			// Pegando os campos data e hora, nome, email, telefone, cidade e estado
			const dataHoraInteresse = abaInteresse.getRange(i, 1).getValue();
			const intervaloInteresse = abaInteresse.getRange(i, colNome, 1, 3).getValues();
			const intervaloCidadeInteresse = abaInteresse.getRange(i, colCidadeInteresse, 1, 2).getValues();

			// Inserindo os campos na planilha gerencial
			abaGerencial.getRange(linhaVazia, colNome, 1, 3).setValues(intervaloInteresse);
			abaGerencial.getRange(linhaVazia, 1).setValue(dataHoraInteresse);
			abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setValues(intervaloCidadeInteresse);

			AtualizarCamposAdicionaisInteresse(i, linhaVazia);

			// Marcando a coluna respondeu interesse da gerencial como sim
			abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue("SIM");

			// Atualizando a nova linha vazia
			linhaVazia = RetornarEspacoVazio();
			continue;
		}
		// Se o email já estiver registrado na planilha gerencial
		AtualizarCamposAdicionaisInteresse(i, linhaCampoGerencial);
	}

	VerificarInteresseMarcoZero();

	// Loop da planilha marco zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const respondeuInteresseMarcoZero = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero).getValue();

		// Apenas faça algo se aquela pessoa não estiver na planilha de interesse
		if (respondeuInteresseMarcoZero != "SIM") {
			const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
			const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailMarcoZero);

			// Se aquele email ainda não estiver registrado na planilha gerencial
			if (!linhaCampoGerencial) {
				// Pegando os campos data e hora, nome, email, telefone e whats
				const dataHoraMarcoZero = abaMarcoZero.getRange(i, 1).getValue();
				const intervaloMarcoZero = abaMarcoZero.getRange(i, colNome, 1, 3).getValues();
				const whatsMarcoZero = abaMarcoZero.getRange(i, colWhatsMarcoZero).getValue();

				// Inserindo os campos na planilha gerencial
				abaGerencial.getRange(linhaVazia, 1).setValue(dataHoraMarcoZero);
				abaGerencial.getRange(linhaVazia, colNome, 1, 3).setValues(intervaloMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuMarcoZeroGerencial).setValue("SIM");
				abaGerencial.getRange(linhaVazia, colFormEnviadoGerencial).setValue("SIM");
				abaGerencial.getRange(linhaVazia, colWhatsGerencial).setValue(whatsMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
				continue;
			}
			// Se o email já estiver registrado na planilha gerencial
			const whatsMarcoZero = abaMarcoZero.getRange(i, colWhatsMarcoZero).getValue();
			abaGerencial.getRange(linhaCampoGerencial, colWhatsGerencial).setValue(whatsMarcoZero);
			abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
		}
	}
}

// Função que atualizará os campos adicionais da planilha gerencial a partir da planilha de interesse
const AtualizarCamposAdicionaisInteresse = (linhaInteresse, linhaInserir) => {
	const whatsInteresse = abaInteresse.getRange(linhaInteresse, colWhatsInteresse).getValue();
	const respMarcoZero = abaInteresse.getRange(linhaInteresse, colRespondeuMarcoZeroInteresse).getValue();
	const situacaoInteresse = abaInteresse.getRange(linhaInteresse, colSituacaoInteresse).getValue();
	const formEnviadoInteresse = abaInteresse.getRange(linhaInteresse, colFormEnviadoInteresse).getValue();

	abaGerencial.getRange(linhaInserir, colWhatsGerencial).setValue(whatsInteresse);
	abaGerencial.getRange(linhaInserir, colRespondeuMarcoZeroGerencial).setValue(respMarcoZero);
	abaGerencial.getRange(linhaInserir, colSituacaoGerencial).setValue(situacaoInteresse);
	abaGerencial.getRange(linhaInserir, colFormEnviadoGerencial).setValue(formEnviadoInteresse);
}


// Função que verificará se o email existe na planilha Marco Zero e retornará a linha
const RetornarLinhaEmailMarcoZero = (emailInteresse) => {
	//Conferir todos os emails da planilha Marco Zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();

		if (emailInteresse == emailMarcoZero) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

//Função para verificar quem respondeu o Marco Zero
const VerificarMarcoZeroInteresse = () => {
	//Pegar o email na planilha Interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const celRespondeuMarcoZero = abaInteresse.getRange(i, colRespondeuMarcoZeroInteresse);
		const valRespondeuMarcoZero = celRespondeuMarcoZero.getValue();

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailInteresse) {
			celRespondeuMarcoZero.setValue("");
			continue;
		}

		// Se o campo já estiver marcado com sim
		if (valRespondeuMarcoZero == "SIM") continue;

		if (RetornarLinhaEmailMarcoZero(emailInteresse))
			celRespondeuMarcoZero.setValue("SIM");
		else
			celRespondeuMarcoZero.setValue("NÃO");
	}
}

// Função que sincroniza os campos do whats do marco zero para a planilha interesse
const SincronizarWhatsInteresse = () => {
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
		const celWhatsInteresse = abaInteresse.getRange(i, colWhatsInteresse);
		const valWhatsInteresse = celWhatsInteresse.getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		const linhaCampoMarcoZero = RetornarLinhaEmailMarcoZero(emailInteresse);

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

// Função que verificará se o email existe na planilha Interesse e retornará a linha
const RetornarLinhaEmailInteresse = (emailMarcoZero) => {
	//Conferir todos os emails da planilha Interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();

		if (emailMarcoZero == emailInteresse) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

//Função para verificar se a pessoa está cadastrada na planilha de Interesse
function VerificarInteresseMarcoZero() {
	//Pegar o email na planilha Marco Zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
		const celEstaNaInteresse = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero);
		const valEstaNaInteresse = celEstaNaInteresse.getValue();

		// Se o campo estiver vazio, limpe a célula e passe para o próximo
		if (!emailMarcoZero) {
			celEstaNaInteresse.setValue("");
			continue;
		}

		// Se o campo já estiver marcado passe para o próximo
		if (valEstaNaInteresse) continue;

		if (RetornarLinhaEmailInteresse(emailMarcoZero))
			celEstaNaInteresse.setValue("SIM");
		else
			celEstaNaInteresse.setValue("S. PÚBLICA");
	}
}