const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
	ui.createMenu('Menu de Funções')
		.addItem('📂 Importar Dados', 'ImportarDados')
		.addItem('🗘 Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
		.addItem('👤 Criar contatos', 'CriaContatos')
		.addItem('🗑️ Excluir todos os campos', 'LimparPlanilha')
		.addSeparator()
		.addSubMenu(ui.createMenu('Formatação da planilha')
			.addItem('Formatar campos telefone', 'FormatarLinhasTelefone')
			.addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
			.addItem('Remover linhas vazias', 'RemoverLinhasVazias'))
		.addToUi();
}

// SOBRE VARIÁVEIS E FUNÇÕES
// -- Variáveis de Colunas das planilhas: --
// 	  Ver arquivo constants
//
// -- Funções da Gerencial: --
//    RetornarLinhaEmailGerencial: Retorna a linha daquele email na planilha Gerencial, se não existir, retorna false
//    ImportarDados: chama outras funções para sincronizar as planilhas Interesse e Marco Zero e chama as funções de importação
//    ImportarDadosInteresse: Pega todos os dados da planilha de Interesse e move na Gerencial, se o registro já exisir, atualize os campos adicionais
//    ImportarDadosMarcoZero: Pega todos os dados da planilha de Marco Zero e move na Gerencial, se o registro já exisir, atualize os campos adicionais
//    AtualizarCamposAdicionaisInteresse: Atualiza os campos adicionais da planilha gerencial a partir da planilha de interesse
//    SincronizarWhatsGerencial: Sincroniza o campo do whatsapp entre todas as planilhas
//    SincronizarCampoPlanilhas: Sincroniza um campo escolhido entre as planilha Interesse e outra desejada, caso não informada, será o Marco Zero
//    CriaContatos (Função não finalizada): Cria contatos no Google People a partir dos dados da planilha Gerencial 
//
// -- Funções de formatação: --
//    LimparPlanilha: Limpa toda a planilha
//    CompletarVaziosComNao: Preenche todos os campos adicionais vazios com o texto "NÃO"
//    FormatarTelefone: Recebe um telefone em formato de texto e o retorna formatado e padronizado
//    FormatarLinhasTelefone: Faz uso da função FormatarTelefone para formatar todos telefones da planilha
//    RemoverLinhasVazias: Remove linhas que estiverem sem email
//
// -- Funções das planilhas Interesse e Marco Zero: --
//    RetornarLinhaEmailInteresse: Retorna a linha daquele email na planilha Interesse, se não existir, retorna false
//    VerificarMarcoZeroInteresse: Verifica quem respondeu o Marco Zero na planilha Interesse
//    RetornarLinhaEmailMarcoZero: Retorna a linha daquele email na planilha Marco Zero, se não existir, retorna false
//    VerificarInteresseMarcoZero: Verifica se a pessoa do Marco Zero está cadastrada na planilha de Interesse

// Função que verificará se o email existe na planilha Gerencial e retornará a linha
function RetornarLinhaEmailGerencial(emailInformado) {
	//Conferir todos os emails da planilha Gerencial
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const emailGerencial = abaGerencial.getRange(i, colEmailGerencial).getValue();

		if (emailInformado == emailGerencial) return i;
	}
	// Se não for encontrado nenhum 
	return false;
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

// Função que importa todos os campos da planilha de interesse
function ImportarDadosInteresse() {
	// Pegando a próxima linha vazia da planilha
	let linhaVazia = abaGerencial.getLastRow() + 1;

	// Loop da planilha interesse
	for (let i = ultimaLinhaAnalisadaInteresse; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmailInteresse).getValue();

		// Se não existir email, passe para o próximo
		if (!emailInteresse) continue;

		const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailInteresse);

		// Se aquele email ainda não estiver registrado na planilha gerencial
		if (!linhaCampoGerencial) {
			// Pegando os campos nome, email, telefone, cidade e estado
			const intervaloInteresse = abaInteresse.getRange(i, colNomeInteresse, 1, 3).getValues();
			const intervaloCidadeInteresse = abaInteresse.getRange(i, colCidadeInteresse, 1, 2).getValues();

			// Inserindo os campos na planilha gerencial
			abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 3).setValues(intervaloInteresse);
			abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setValues(intervaloCidadeInteresse);
			abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue("SIM");

			AtualizarCamposAdicionaisInteresse(i, linhaVazia);

			// Atualizando a nova linha vazia
			linhaVazia++;
			continue;
		}

		// Se o email já estiver registrado na planilha gerencial
		AtualizarCamposAdicionaisInteresse(i, linhaCampoGerencial);
	}
}

// Função que importa os campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero() {
	// Pegando a próxima linha vazia da planilha
	let linhaVazia = abaGerencial.getLastRow() + 1;

	// Loop da planilha marco zero
	for (let i = ultimaLinhaAnalisadaMarcoZero; i <= ultimaLinhaMarcoZero; i++) {
		const emailMarcoZero = abaMarcoZero.getRange(i, colEmailMarcoZero).getValue();

		// Se não existir email, passe para o próximo
		if (!emailMarcoZero) continue;

		// Pegando o campo se está cadastrada na planilha de interesse e pegando a linha desse email na planilha gerencial
		const respondeuInteresseMarcoZero = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero).getValue();
		const linhaCampoGerencial = RetornarLinhaEmailGerencial(emailMarcoZero);

		// Se aquela pessoa não estiver na planilha de interesse
		if (respondeuInteresseMarcoZero != "SIM") {

			// Se aquele email ainda não estiver registrado na planilha gerencial
			if (!linhaCampoGerencial) {

				// Pegando os campos nome, email, telefone e whats
				const intervaloMarcoZero = abaMarcoZero.getRange(i, colNomeMarcoZero, 1, 3).getValues();
				const whatsMarcoZero = abaMarcoZero.getRange(i, colWhatsMarcoZero).getValue();

				// Inserindo os campos na planilha gerencial
				abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 3).setValues(intervaloMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuMarcoZeroGerencial).setValue("SIM");
				abaGerencial.getRange(linhaVazia, colWhatsGerencial).setValue(whatsMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

				// Pintando campos
				abaGerencial.getRange(linhaVazia, colCidadeGerencial).setBackground("#eeeeee");

				// Atualizando a nova linha vazia
				linhaVazia++;
				continue;
			}

			// Se o email já estiver registrado na planilha gerencial mas não estiver na planilha de interesse
			abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
		}
	}
}

// Função que atualizará os campos adicionais da planilha gerencial a partir da planilha de interesse
function AtualizarCamposAdicionaisInteresse(linhaInteresse, linhaInserir) {
	const whatsInteresse = abaInteresse.getRange(linhaInteresse, colWhatsInteresse).getValue();
	const respMarcoZero = abaInteresse.getRange(linhaInteresse, colRespondeuMarcoZeroInteresse).getValue();
	const situacaoInteresse = abaInteresse.getRange(linhaInteresse, colSituacaoInteresse).getValue();

	abaGerencial.getRange(linhaInserir, colWhatsGerencial).setValue(whatsInteresse);
	abaGerencial.getRange(linhaInserir, colRespondeuMarcoZeroGerencial).setValue(respMarcoZero);
	abaGerencial.getRange(linhaInserir, colSituacaoGerencial).setValue(situacaoInteresse);
}

// Função que sincronizará quem entrou no whatsapp entre as três planilhas
function SincronizarWhatsGerencial() {
	// Sincronize as planilhas Interesse e Marco Zero, depois as planilhas Interesse e Gerencial e por fim, Interesse e Marco Zero de novo
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsGerencial, abaGerencial);
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
}

// Função que sincronizará um dado campo entre as planilhas Interesse e uma outra desejada, caso não for informada,
// A outra planilha será o Marco Zero
function SincronizarCampoPlanilhas(colInteresseDesejada, colPlanilhaDesejada, abaDesejada) {
	for (let i = 2; i <= ultimaLinhaInteresse; i++) {
		const emailInteresse = abaInteresse.getRange(i, colEmailInteresse).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailInteresse)
			continue;

		// Se a aba desejada for a gerencial, use a função da gerencial, se não, use a função do marco zero
		const linhaCampoPlanilhaDesejada = abaDesejada == abaGerencial ? RetornarLinhaEmailGerencial(emailInteresse) : RetornarLinhaEmailMarcoZero(emailInteresse);
		const abaPlanilhaDesejada = abaDesejada ?? abaMarcoZero;

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
	for (let i = ultimaLinhaAnalisadaWhatsGerencial; i <= ultimalinhaGerencial; i++) {
		// verifica se esta cadastrado no whats ou não 
		const celGerencialWhats = abaGerencial.getRange(i, colWhatsGerencial)
		const whats = celGerencialWhats.getValue();
		if (whats === "NÃO") {
			// pega o nome da pessoa e já divide o nome e sobrenome para ficar certo quando for criar o contato
			const nomes = abaGerencial.getRange(i, colNomeGerencial).getValue().toString().trim().split(" ");
			const lengthNomes = nomes.length;
			// pega o valor do telefone
			const telefone = abaGerencial.getRange(i, colTelGerencial).getValue();
			// cria o contato 
			const novoContato = People.People.createContact({
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
			celGerencialWhats.setValue("SIM");
		}
	}
}