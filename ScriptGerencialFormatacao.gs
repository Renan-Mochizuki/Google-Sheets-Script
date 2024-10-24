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
	for (let i = 2; i <= ultimaLinhaAtiva; i++) {
		const valorTelefone = abaAtiva.getRange(i, colTelAtiva).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!valorTelefone) continue;

		const telefoneFormatado = FormatarTelefone(valorTelefone)
		abaAtiva.getRange(i, colTelAtiva).setValue(telefoneFormatado);
	}
}

// Função que remove todas linhas vazias no meio da planilha
function RemoverLinhasVazias() {
	for (let i = 2; i <= ultimaLinhaAtiva; i++) {
		const emailAtiva = abaAtiva.getRange(i, colEmailAtiva).getValue();
		if (!emailAtiva) {
			abaAtiva.deleteRow(i);
		}
	}
}

// Função para preencher o campo do estado a partir do campo cidade
function PreencherEstado() {
	// Atribui os variáveis de acordo com a abaDesejada
	const { colCidade, colEstado } = objetoMap.get(abaAtiva) || {};
	Logger.log(colCidade, colEstado);
	// Loop das linhas
	for (let i = 2; i <= ultimaLinhaAtiva; i++) {
		const cidade = abaAtiva.getRange(i, colCidade).getValue();

		abaAtiva.getRange(i, colEstado).setValue(estado);
	}
}

function mostrarInterfaceComCheckboxes() {
  var html = HtmlService.createHtmlOutputFromFile('InterfaceCheckboxes')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Escolha quem visualizar');
}

function processarEscolhas(escolhas) {
	planilhaAtiva.toast('ativo', 'ativo', tempoNotificacao);
	EsconderLinhas(abaGerencial, colTerminouCursoGerencial, "SIM")
  }

function EsconderLinhas(abaDesejada, colDesejada, valorAMostrar) {

	const { ultimaLinha } = objetoMap.get(abaDesejada) || {};

	const valColunas = abaDesejada.getRange(2, colDesejada, ultimaLinha, 1).getValues().flat();

	
	for (let i = 0; i < valColunas.length; i++) {
		if (valColunas[i] != valorAMostrar) {
			abaDesejada.hideRows(i + 2);
		}
	}
}