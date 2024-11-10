// -- Funções de formatação da planilha --

// Função para limpar toda a planilha
function LimparPlanilha() {
	// Janela de diálogo de confirmação da ação
	// const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os campos?', ui.ButtonSet.YES_NO);

	// if (response == ui.Button.YES) {
	// Verifica se há mais de uma linha para limpar
	// if (ultimaLinhaAtiva > 1) {
	// Define o intervalo que vai da segunda linha até a última linha e a última coluna com conteúdo
	const planilha = abaAtiva.getRange(2, 1, ultimaLinhaAtiva - 1, ultimaColunaAtiva);

	// Limpa o conteúdo do intervalo selecionado
	planilha.clearContent();
	planilha.setBackground('#ffffff');
  planilha.clearNote();
	// }
	// }
}

// Função que completa campos vazios adicionais da planilha gerencial com NÃO
function CompletarVaziosComNao() {
	const colunas = [colWhatsGerencial, colRespondeuInteresseGerencial, colRespondeuMarcoZeroGerencial, colRespondeuMarcoFinalGerencial, colEnviouReflexaoMarcoFinalGerencial];

	// Loop das colunas
	for (let j = 0; j < colunas.length; j++) {
		const coluna = colunas[j];

		// Loop das linhas
		for (let i = 2; i <= ultimaLinhaGerencial; i++) {
			const celula = abaGerencial.getRange(i, coluna)
			const valor = celula.getValue();
			if (!valor) celula.setValue("NÃO");
		}
	}
}

// Função que recebe um telefone digitado e retorna o telefone formatado
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
	} else if (telefoneNumeros.startsWith("55") && telefoneNumeros.length == 13) {
		// Para caso for digitado o código do Brasil sem o +
		telefoneNumeros = telefoneNumeros.substring(2); // Remove os dois primeiros caracteres
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

// Função que usa a função FormatarTelefone para formatar todos os campos da planilha ativa
function FormatarLinhasTelefoneAba(abaDesejada) {
	// Atribui as variáveis de acordo com a abaDesejada
	const { colTel, ultimaLinha } = objetoMap.get(abaDesejada);

	// Pega todos os valores da coluna desejada
	const telefones = abaDesejada.getRange(2, colTel, ultimaLinha, 1).getValues();

	// Percorre toda a matriz
	for (let i = 0; i < telefones.length; i++) {
		const tel = telefones[i][0];

		// Se o campo estiver vazio, passe para o próximo
		if (!tel) continue;

		// Formata os telefones
		telefones[i][0] = FormatarTelefone(tel)
	}

	abaDesejada.getRange(2, colTel, ultimaLinha, 1).setValues(telefones);
}

// Função que usa a função FormatarTelefone para formatar todos os campos da planilha ativa
function FormatarLinhasTelefone() {
	FormatarLinhasTelefoneAba(abaAtiva);
}

// Função que usa a função FormatarTelefone para formatar todas planilhas
function FormatarLinhasTelefoneTodasAbas() {
	Promise.all([
		FormatarLinhasTelefoneAba(abaInteresse),
		FormatarLinhasTelefoneAba(abaMarcoZero),
		FormatarLinhasTelefoneAba(abaEnvioMapa),
		FormatarLinhasTelefoneAba(abaMarcoFinal),
		FormatarLinhasTelefoneAba(abaCertificado)
	]);
}

// Função que remove todas linhas vazias no meio da planilha
function RemoverLinhasVazias() {
	// Pega todos os valores da coluna desejada
	const valColunas = abaAtiva.getRange(2, colEmailAtiva, ultimaLinhaAtiva, 1).getValues().flat();

	// Loop que percorre todos valores da coluna
	for (let i = 0; i < valColunas.length; i++) {
		if (!emailAtiva) {
			abaAtiva.deleteRow(i);
		}
	}
}

// Função para preencher o campo do estado a partir do campo cidade
function PreencherEstado() {
	// Atribui os variáveis de acordo com a abaAtiva
	const { colCidade, colEstado } = objetoMap.get(abaAtiva);
	Logger.log(colCidade, colEstado);
	// Loop das linhas
	for (let i = 2; i <= ultimaLinhaAtiva; i++) {
		const cidade = abaAtiva.getRange(i, colCidade).getValue();

		abaAtiva.getRange(i, colEstado).setValue(estado);
	}
}

// Função que exibe o HTML da interface com checkboxes para escolher quem quer esconder
function MostrarInterfaceEsconderLinhas() {
	const html = HtmlService.createHtmlOutputFromFile('InterfaceCheckboxes').setWidth(400).setHeight(300);
	SpreadsheetApp.getUi().showModalDialog(html, 'Escolha quem visualizar');
}

// Função que recebe as escolhas feitas na interface e chama a função EsconderLinhas como necessário
function ProcessarEscolhasEsconderLinhas(escolhas) {
	EsconderLinhas(colTerminouCursoGerencial, "SIM")
}

// Função que esconde todas as linhas que possuem um certo valor em uma coluna
function EsconderLinhas(colDesejada, valorAMostrar) {
	// Pega todos os valores da coluna desejada
	const valColunas = abaAtiva.getRange(2, colDesejada, ultimaLinhaAtiva, 1).getValues().flat();

	// Loop que percorre todos valores da coluna
	for (let i = 0; i < valColunas.length; i++) {
		// Se o valor da coluna for diferente do valorAMostrar, esconde a linha
		if (valColunas[i] != valorAMostrar) {
			abaAtiva.hideRows(i + 2);
		}
	}
}

// Função que revela todas as linhas escondidas
function MostrarTodasLinhas() {
	abaAtiva.showRows(2, ultimaLinhaAtiva);
}