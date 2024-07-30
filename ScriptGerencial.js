// Função para adicionar o menu
function onOpen(e) {
	SpreadsheetApp.getUi()
		.createMenu('Menu de Funções')
		.addItem('Importar Dados', 'ImportarDados')
		.addItem('Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
		.addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
		.addItem('Excluir todos os campos', 'LimparCampos')
		.addItem('Criar contatos', 'CriaContatos')
		.addToUi();
}

onOpen();

//Seleciona a planilha de Confirmação de Interesse e a aba
const urlInteresse = "*";
const planilhaInteresse = SpreadsheetApp.openByUrl(urlInteresse);
const abaInteresse = planilhaInteresse.getSheetByName("Respostas ao formulário 1");

//Seleciona a planilha do Marco Zero e a aba
const urlMarcoZero = "*";
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
const ultimaColunaGerencial = abaGerencial.getLastColumn();

// Colunas gerais
const colData = 1;
const colNome = 3;
const colEmail = 4;
const colTel = 5;
const colConfirmacaoTel = 6;
// Colunas planilha Interesse
const colDataNascInteresse = 7;
const colCidadeInteresse = 8;
const colEstadoInteresse = 9;
const colWhatsInteresse = 13;
const colRespondeuMarcoZeroInteresse = 14;
const colSituacaoInteresse = 15;
const colFormEnviadoInteresse = 16;
// Colunas planilha Marco Zero
const colRespondeuInteresseMarcoZero = 13;
const colWhatsMarcoZero = 14;
// Colunas planilha Gerencial
const colDataInteresseGerencial = 1;
const colDataMarcoZeroGerencial = 2;
const colCidadeGerencial = 6;
const colEstadoGerencial = 7;
const colWhatsGerencial = 8;
const colRespondeuInteresseGerencial = 9;
const colRespondeuMarcoZeroGerencial = 10;
const colSituacaoGerencial = 11;
const colFormEnviadoGerencial = 12;

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

// Retorna a linha em que o campo do email está vazio
function RetornarEspacoVazio() {
	let i = 2;
	do {
		const celEmailGerencial = abaGerencial.getRange(i, colEmail).getValue();
		if (!celEmailGerencial) return i;
		i++;
	} while (i != 0)
}

// Função que importa todos os campos da planilha de interesse
function ImportarDadosInteresse() {
	// Pegando a próxima linha vazia da planilha
	let linhaVazia = RetornarEspacoVazio();

	// Loop da planilha interesse
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();

		// Se não existir email, passe para o próximo
		if (!emailInteresse) continue;

		const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailInteresse);

		// Se aquele email ainda não estiver registrado na planilha gerencial
		if (!linhaCampoGerencial) {
			// Pegando os campos data e hora, nome, email, telefone, cidade e estado
			const dataHoraInteresse = abaInteresse.getRange(i, colData).getValue();
			const intervaloInteresse = abaInteresse.getRange(i, colNome, 1, 3).getValues();
			const intervaloCidadeInteresse = abaInteresse.getRange(i, colCidadeInteresse, 1, 2).getValues();

			// Inserindo os campos na planilha gerencial
			abaGerencial.getRange(linhaVazia, colNome, 1, 3).setValues(intervaloInteresse);
			abaGerencial.getRange(linhaVazia, colDataInteresseGerencial).setValue(dataHoraInteresse);
			abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setValues(intervaloCidadeInteresse);
			abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue("SIM");

			AtualizarCamposAdicionaisInteresse(i, linhaVazia);

			// Atualizando a nova linha vazia
			linhaVazia = RetornarEspacoVazio();
			continue;
		}

		// Se o email já estiver registrado na planilha gerencial
		AtualizarCamposAdicionaisInteresse(i, linhaCampoGerencial);
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

	// Se a pessoa tiver respondido o marco zero, pegue a data da resposta e insira
	if (respMarcoZero == 'SIM') {
		const emailInteresse = abaInteresse.getRange(linhaInteresse, colEmail).getValue();
		const linhaCampoMarcoZero = RetornarLinhaEmailMarcoZero(emailInteresse);
		const dataHoraMarcoZero = abaMarcoZero.getRange(linhaCampoMarcoZero, colData).getValue();
		abaGerencial.getRange(linhaInserir, colDataMarcoZeroGerencial).setValue(dataHoraMarcoZero);
	}
}

// Função que importa os campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero() {
	// Pegando a próxima linha vazia da planilha
	let linhaVazia = RetornarEspacoVazio();

	// Loop da planilha marco zero
	for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();

		// Se não existir email, passe para o próximo
		if (!emailMarcoZero) continue;

		// Pegando o campo se está cadastrada na planilha de interesse e pegando a linha desse email na planilha gerencial
		const respondeuInteresseMarcoZero = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero).getValue();
		const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailMarcoZero);

		// Se aquela pessoa não estiver na planilha de interesse
		if (respondeuInteresseMarcoZero != "SIM") {

			// Se aquele email ainda não estiver registrado na planilha gerencial
			if (!linhaCampoGerencial) {

				// Pegando os campos data e hora, nome, email, telefone e whats
				const dataHoraMarcoZero = abaMarcoZero.getRange(i, colData).getValue();
				const intervaloMarcoZero = abaMarcoZero.getRange(i, colNome, 1, 3).getValues();
				const whatsMarcoZero = abaMarcoZero.getRange(i, colWhatsMarcoZero).getValue();

				// Inserindo os campos na planilha gerencial
				abaGerencial.getRange(linhaVazia, colDataMarcoZeroGerencial).setValue(dataHoraMarcoZero);
				abaGerencial.getRange(linhaVazia, colNome, 1, 3).setValues(intervaloMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuMarcoZeroGerencial).setValue("SIM");
				abaGerencial.getRange(linhaVazia, colFormEnviadoGerencial).setValue("SIM");
				abaGerencial.getRange(linhaVazia, colWhatsGerencial).setValue(whatsMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

				// Pintando campos
				abaGerencial.getRange(linhaVazia, colDataInteresseGerencial).setBackground("#eeeeee");
				abaGerencial.getRange(linhaVazia, colCidadeGerencial).setBackground("#eeeeee");

				// Atualizando a nova linha vazia
				linhaVazia = RetornarEspacoVazio();
				continue;
			}

			// Se o email não estiver na planilha de interesse e já estiver registrado na planilha gerencial
			abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
		}
	}
}

// Função que importa dados da planilha interesse e do marco zero que não estão na de interesse
function ImportarDados() {
	// Chamando funções das planilhas para atualizar seus campos
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
	VerificarMarcoZeroInteresse()
	VerificarInteresseMarcoZero();

	ImportarDadosInteresse();
	ImportarDadosMarcoZero();
}

// Função que sincronizará quem entrou no whatsapp entre as três planilhas
function SincronizarWhatsGerencial() {
	// Sincronize as planilhas Interesse e Marco Zero, depois as planilhas Interesse e Gerencial e por fim, Interesse e Marco Zero de novo
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsGerencial, abaGerencial);
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
}

// Função para limpar toda a planilha
function LimparCampos() {
	// Loop das linhas
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		// Loop das colunas
		for (let j = 1; j <= ultimaColunaGerencial; j++) {
			const celula = abaGerencial.getRange(i, j)
			celula.setValue('');
			celula.setBackground('#ffffff');
		}
	}
}

// Função que completa todos os campos vazios adicionais com NÃO
function CompletarVaziosComNao() {
	// Loop das colunas
	for (let j = colWhatsGerencial; j <= ultimaColunaGerencial; j++) {

		// Se a coluna for a de situação, pule
		if (j == colSituacaoGerencial) continue;

		// Loop das linhas
		for (let i = 2; i <= ultimalinhaGerencial; i++) {
			const celula = abaGerencial.getRange(i, j)
			const valor = celula.getValue();
			if (!valor) celula.setValue("NÃO");
		}
	}
}

// Função que sincronizará um dado campo entre as planilhas Interesse e uma outra desejada, caso não for informada,
// A outra planilha será o Marco Zero
function SincronizarCampoPlanilhas(colInteresseDesejada, colPlanilhaDesejada, abaDesejada) {
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		// Se a aba desejada for a gerencial, use a função da gerencial, se não, use a função do marco zero
		const linhaCampoPlanilhaDesejada = abaDesejada == abaGerencial ? RetornarLinhaEmailGerencial(emailInteresse) : RetornarLinhaEmailMarcoZero(emailInteresse);
		const abaPlanilhaDesejada = abaDesejada == abaGerencial ? abaGerencial : abaMarcoZero;

		// Se o email for encontrado na outra planilha
		if (linhaCampoPlanilhaDesejada) {
			const celInteresse = abaInteresse.getRange(i, colInteresseDesejada);
			const valInteresse = celInteresse.getValue();
			const celPlanilhaDesejada = abaPlanilhaDesejada.getRange(linhaCampoPlanilhaDesejada, colPlanilhaDesejada);
			const valPlanilhaDesejada = celPlanilhaDesejada.getValue();

			// Se o campo do Interesse estiver vazio, altere o campo do Interesse com o valor da outra planilha
			if (!valInteresse) {
				celInteresse.setValue(valPlanilhaDesejada);
				continue;
			}

			// Se o campo da outra planilha estiver vazio, altere o campo da outra planilha com o valor do Interesse
			if (!valPlanilhaDesejada) {
				celPlanilhaDesejada.setValue(valInteresse);
				continue;
			}

			// Se o campo do Interesse estiver como sim e da outra como não, altere o campo da outra planilha
			if (valInteresse == "SIM" && valPlanilhaDesejada == "NÃO") {
				celPlanilhaDesejada.setValue("SIM");
				continue;
			}

			// Se o campo da outra planilha estiver como sim e da outra como não, altere o campo do Interesse
			if (valPlanilhaDesejada == "SIM" && valInteresse == "NÃO") {
				celInteresse.setValue("SIM");
				continue;
			}
		}
	}
}

function CriaContatos() {
	// for para percorrer todas as linhas
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		// verifica se esta cadastrado no whats ou não 
		whats = abaGerencial.getRange(i, colWhatsGerencial).getValue();
		if (whats === "NÃO") {
			// pega o nome da pessoa e já divide o nome e sobrenome para ficar certo quando for criar o contato
			let nomes = abaGerencial.getRange(i, colNome).getValue().toString().split(" ");
			let lengthNomes = nomes.length;
			// pega o valor do telefone
			let telefone = abaGerencial.getRange(i, colTel).getValue();
			// cria o contato 
			let novoContato = People.People.createContact({
				// coloca o nome e sobrenome
				names: [{
					givenName: nomes[0],
					familyName: nomes[lengthNomes - 1]
				}],
				// coloca o número de telefone
				phoneNumbers: [{
					value: telefone.toString()
				}]
			});
		}
	}

}