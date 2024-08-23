const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
	ui.createMenu('Menu de Funções')
		.addItem('📂 Importar Dados', 'ImportarDados')
		.addItem('🗘 Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
		.addItem('👤 Criar contatos', 'CriaContatos')
		.addItem('🗑️ Excluir todos os campos', 'LimparCampos')
		.addSeparator()
		.addSubMenu(ui.createMenu('Formatação da planilha')
			.addItem('Formatar campos telefone', 'FormatarLinhasTelefone')
			.addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
			.addItem('Remover linhas vazias', 'RemoverLinhasVazias'))
		.addToUi();
}

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

// Variáveis de otimização
const ultimaLinhaAnalisadaInteresse = 2;
const ultimaLinhaAnalisadaMarcoZero = 2
const ultimaLinhaAnalisadaWhatsGerencial = 2;

// Email de envio do formulário
const assuntoEmail = `Formulário Marco Zero`;
const textoEmail = `Responda o formulário do Marco Zero para dar continuidade a sua formação em Mapas Conceituais. Link: https://forms.gle/YQdMCoemkDiumzyG6`;

// Função que verificará se o email existe na planilha Gerencial e retornará a linha
function RetornarLinhaEmailGerencial(emailInformado) {
	//Conferir todos os emails da planilha Gerencial
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const emailGerencial = abaGerencial.getRange(i, colEmail).getValue();

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
			linhaVazia++;
			continue;
		}

		// Se o email já estiver registrado na planilha gerencial
		AtualizarCamposAdicionaisInteresse(i, linhaCampoGerencial);
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
	let linhaVazia = abaGerencial.getLastRow() + 1;

	// Loop da planilha marco zero
	for (let i = ultimaLinhaAnalisadaMarcoZero; i <= ultimaLinhaMarcoZero; i++) {
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
				abaGerencial.getRange(linhaVazia, colWhatsGerencial).setValue(whatsMarcoZero);
				abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

				// Pintando campos
				abaGerencial.getRange(linhaVazia, colDataInteresseGerencial).setBackground("#eeeeee");
				abaGerencial.getRange(linhaVazia, colCidadeGerencial).setBackground("#eeeeee");

				// Atualizando a nova linha vazia
				linhaVazia++;
				continue;
			}

			// Se o email não estiver na planilha de interesse e já estiver registrado na planilha gerencial
			abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
		}
	}
}

// Função que remove todas linhas vazias no meio da planilha
function RemoverLinhasVazias() {
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const emailGerencial = abaGerencial.getRange(i, colEmail).getValue();
		if (!emailGerencial) {
			abaGerencial.deleteRow(i);
		}
	}
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
	// Janela de diálogo de confirmação da ação
	const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os campos?', ui.ButtonSet.YES_NO);

	if (response == ui.Button.YES) {
		// Loop das linhas
		// Verifica se há mais de uma linha para limpar
		if (ultimalinhaGerencial > 1) {
			// Define o intervalo que vai da segunda linha até a última linha e a última coluna com conteúdo
			const intervalo = abaGerencial.getRange(2, 1, ultimalinhaGerencial - 1, ultimaColunaGerencial);

			// Limpa o conteúdo do intervalo selecionado
			intervalo.clearContent();
			intervalo.setBackground('#ffffff');
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

function FormatarTelefone(textoTelefone) {
	// Remove todos os caracteres não numéricos, exceto o '+'
	let telefoneNumeros = textoTelefone.toString().replace(/[^\d+]/g, '');

	// Regex para separar o código de país e o resto do telefone
	const regex = /\+(\d{2})\s*(.*)/;
	const resultado = telefoneNumeros.match(regex);

	// Se houver um código de pais, remova o código do telefone
	if (resultado) {
		// Se o código de país for diferente de 55 (Brasil), retorna o texto original
		if (resultado[1] !== '55') return textoTelefone;
		telefoneNumeros = resultado[2];
	}

	switch (telefoneNumeros.length) {
		case 8: // Telefone 8 dígitos sem DDD
			return telefoneNumeros.replace(/(\d{4})(\d)/, '$1-$2');
		case 9: // Telefone 9 dígitos sem DDD
			return telefoneNumeros.replace(/(\d{5})(\d)/, '$1-$2');
		case 10: // Telefone 8 dígitos com DDD
			return telefoneNumeros.replace(/(\d{2})(\d{1})/, '($1) $2').replace(/(\d{4})(\d)/, '$1-$2');
		case 11: // Telefone 9 dígitos com DDD
			return telefoneNumeros.replace(/(\d{2})(\d{1})/, '($1) $2').replace(/(\d{5})(\d)/, '$1-$2');
	}

	// Retorna o telefone com apenas números para os demais casos
	return telefoneNumeros;
}

function FormatarLinhasTelefone() {
	// Loop das linhas
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const valorTelefone = abaGerencial.getRange(i, colTel).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!valorTelefone) continue;

		const telefoneFormatado = FormatarTelefone(valorTelefone)
		abaGerencial.getRange(i, colTel).setValue(telefoneFormatado);
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
			const nomes = abaGerencial.getRange(i, colNome).getValue().toString().trim().split(" ");
			const lengthNomes = nomes.length;
			// pega o valor do telefone
			const telefone = abaGerencial.getRange(i, colTel).getValue();
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