// -- Funções de formatação da planilha gerencial --

// Função para limpar toda a planilha
function LimparPlanilha() {
	// Janela de diálogo de confirmação da ação
	const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os campos?', ui.ButtonSet.YES_NO);

	if (response == ui.Button.YES) {
		// Verifica se há mais de uma linha para limpar
		if (ultimalinhaGerencial > 1) {
			// Define o intervalo que vai da segunda linha até a última linha e a última coluna com conteúdo
			const planilha = abaGerencial.getRange(2, 1, ultimalinhaGerencial - 1, ultimaColunaGerencial);

			// Limpa o conteúdo do intervalo selecionado
			planilha.clearContent();
			planilha.setBackground('#ffffff');
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
		const valorTelefone = abaGerencial.getRange(i, colTelGerencial).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!valorTelefone) continue;

		const telefoneFormatado = FormatarTelefone(valorTelefone)
		abaGerencial.getRange(i, colTelGerencial).setValue(telefoneFormatado);
	}
}

// Função que remove todas linhas vazias no meio da planilha
function RemoverLinhasVazias() {
	for (let i = 2; i <= ultimalinhaGerencial; i++) {
		const emailGerencial = abaGerencial.getRange(i, colEmailGerencial).getValue();
		if (!emailGerencial) {
			abaGerencial.deleteRow(i);
		}
	}
}