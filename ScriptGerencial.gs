const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
	ui.createMenu('Menu de Funções')
		.addItem('📂 Importar Dados', 'Importar')
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
// 	  Ver arquivo Constants
//
// -- Funções da Gerencial: --
//    RetornarLinhaEmailPlanilha(emailProcurado, abaDesejada):
//    - retorna a linha daquele email na planilha desejada, se não existir, retorna false
//    Importar():
//    - chama outras funções para sincronizar as planilhas e chama as funções de importação de todos dados
//    ImportarDados(abaDesejada):
//    - função genérica para chamar a função de importação de dados de cada planilha
//    ImportarDadosInteresse(linhaAtual, linhaCampoGerencial, linhaVazia):
//    - pega todos os dados da Interesse e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosMarcoZero(linhaAtual, linhaCampoGerencial, linhaVazia):
//    - pega todos os dados do Marco Zero e move na Gerencial ou apenas o campo respondeu interesse
//    ImportarDadosEnvioMapa(linhaAtual, linhaCampoGerencial, linhaVazia):
//    - pega todos os dados do Envio do Mapa e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosMarcoFinal(linhaAtual, linhaCampoGerencial, linhaVazia):
//    - pega todos os dados do Marco Final e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosCertificado(linhaAtual, linhaCampoGerencial, linhaVazia):
//    - pega todos os dados do Certificado e move na Gerencial ou apenas atualiza os campos adicionais
//    SincronizarWhatsGerencial():
//    - sincroniza o campo do whatsapp entre todas as planilhas
//    SincronizarCampoPlanilhas(abaDesejada1, colDesejada1, abaDesejada2, colDesejada2):
//    - sincroniza um campo escolhido entre duas planilhas desejadas
//    VerificarEMarcarCadastroOutraPlanilha(abaParaRegistro, colParaRegistro, abaParaVerificar, valCustomizadoSim, valCustomizadoNao):
//    - verifica se a pessoa está cadastrada em uma planilha e marca em outra
//    CriaContatos(): (Função não finalizada)
//    - cria contatos no Google People a partir dos dados da planilha Gerencial
//
// -- Funções de formatação: --
//    LimparPlanilha():
//    - limpa toda a planilha
//    CompletarVaziosComNao():
//    - preenche todos os campos adicionais vazios da planilha gerencial com o texto "NÃO"
//    FormatarTelefone(textoTelefone):
//    - recebe um telefone em formato de texto e o retorna formatado e padronizado
//    FormatarLinhasTelefone():
//    - faz uso da função FormatarTelefone para formatar todos telefones da planilha
//    RemoverLinhasVazias():
//    - remove linhas que estiverem sem email

// Função que verificará se o email existe na planilha desejada e retornará a linha
function RetornarLinhaEmailPlanilha(emailProcurado, abaDesejada) {
	// Pegar variáveis da planilha desejada
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada) || {};

	//Conferir todos os emails da planilha desejada
	for (let i = 2; i <= ultimaLinha; i++) {
		const email = abaDesejada.getRange(i, colEmail).getValue();

		if (emailProcurado == email) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

// Função que executa as funções necessárias para importar todos os dados
function Importar() {
	const tempoNotificacao = 5;
	let linhaVazia, linhasAfetadas, totalLinhasAfetadas = 0;
	// Chamando funções das planilhas para atualizar seus campos
	planilhaAtiva.toast('Sincronizando campos Whats', 'Executando funções', tempoNotificacao);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
	// Verifica na planilha Interesse, quem respondeu o Marco Zero, e verifica na planilha Marco Zero, quem respondeu o Interesse
	planilhaAtiva.toast('Verificando respostas Marco Zero', 'Executando funções', tempoNotificacao);
	VerificarEMarcarCadastroOutraPlanilha(abaInteresse, colRespondeuMarcoZeroInteresse, abaMarcoZero);
	planilhaAtiva.toast('Verificando respostas Interesse', 'Executando funções', tempoNotificacao);
	VerificarEMarcarCadastroOutraPlanilha(abaMarcoZero, colRespondeuInteresseMarcoZero, abaInteresse, null, "S. PÚBLICA");

	planilhaAtiva.toast('Importando dados da Interesse', 'Executando funções', tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaInteresse));
	totalLinhasAfetadas += linhasAfetadas;
	planilhaAtiva.toast('Importando dados do Marco Zero', 'Executando funções', tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaMarcoZero));
	totalLinhasAfetadas += linhasAfetadas;
	planilhaAtiva.toast('Importando dados do Envio de Mapa', 'Executando funções', tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaEnvioMapa));
	totalLinhasAfetadas += linhasAfetadas;
	planilhaAtiva.toast('Importando dados do Marco Final', 'Executando funções', tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaMarcoFinal));
	totalLinhasAfetadas += linhasAfetadas;
	planilhaAtiva.toast('Importando dados do Envio do Certificado', 'Executando funções', tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaCertificado));
	totalLinhasAfetadas += linhasAfetadas;
	let linhasCriadas = linhaVazia - ultimaLinhaGerencial - 1;
	let mensagem = 'Fim da execução.\n' + linhasCriadas + ' linhas criadas\n' + totalLinhasAfetadas + ' linhas afetadas';
	planilhaAtiva.toast(mensagem, 'Execução finalizada', tempoNotificacao + 5);

}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
	// Pegando a próxima linha vazia da planilha gerencial
	let linhaVazia = abaGerencial.getLastRow() + 1;
	let linhasAfetadas = 0;

	// Atribui os variáveis de acordo com a abaDesejada
	const { ultimaLinhaAnalisada, ultimaLinha, colEmail, ImportarDadosPlanilha } = objetoMap.get(abaDesejada) || {};

	// Loop da planilha Desejada
	for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
		const email = abaDesejada.getRange(i, colEmail).getValue();

		// Se não existir email, passe para o próximo
		if (!email) continue;

		const linhaCampoGerencial = RetornarLinhaEmailPlanilha(email, abaGerencial);

		const novaLinhaCriada = ImportarDadosPlanilha(i, linhaCampoGerencial, linhaVazia);

		if (novaLinhaCriada) linhaVazia++
		else linhasAfetadas++;
	}

	return { linhaVazia, linhasAfetadas };
}

// Função com a lógica da importação dos campos da planilha de interesse
function ImportarDadosInteresse(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Pegando os valores da linha
	const valLinha = abaInteresse.getRange(linhaAtual, 1, 1, ultimaColunaInteresse).getValues()[0];

	// Pega os campos adicionais da planilha Interesse adicionando "SIM" para o campo "Respondeu Interesse" na Gerencial
	const intervaloAdicionais = [
		valLinha[colWhatsInteresse - 1],
		"SIM",
		valLinha[colRespondeuMarcoZeroInteresse - 1],
		valLinha[colSituacaoInteresse - 1]
	];

	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Considerando a ordem dos campos da planilha Gerencial (Ver arquivo Constants)
		const intervaloInserir = [
			valLinha[colNomeInteresse - 1],
			valLinha[colEmailInteresse - 1],
			valLinha[colTelInteresse - 1],
			valLinha[colCidadeInteresse - 1],
			valLinha[colEstadoInteresse - 1],
			...intervaloAdicionais
		]

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 9).setValues([intervaloInserir]);

		// Nova linha criada
		return true;
	}

	// Se o email já estiver registrado na planilha gerencial, atualize os campos adicionais
	abaGerencial.getRange(linhaCampoGerencial, colWhatsGerencial, 1, 4).setValues([intervaloAdicionais]);

	// Nenhuma linha criada
	return false;
}

// Função com a lógica da importação dos campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Pegando os valores da linha
	const valLinha = abaMarcoZero.getRange(linhaAtual, 1, 1, ultimaColunaMarcoZero).getValues()[0];

	// Pegando o campo se está cadastrada na planilha de interesse
	const respondeuInteresseMarcoZero = valLinha[colRespondeuInteresseMarcoZero - 1];

	// Se aquela pessoa já estava na planilha de interesse
	if (respondeuInteresseMarcoZero == "SIM") {
		// Nenhuma linha criada
		return false;
	}

	// Se aquela pessoa ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Considerando a ordem dos campos da planilha Gerencial (Ver arquivo Constants)
		const intervaloInserir = [
			valLinha[colNomeMarcoZero - 1],
			valLinha[colEmailMarcoZero - 1],
			valLinha[colTelMarcoZero - 1],
			null,
			null,
			valLinha[colWhatsMarcoZero - 1],
			respondeuInteresseMarcoZero,
			"SIM"
		]

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 8).setValues([intervaloInserir]);

		// Pintando campos cidade e estado
		abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 1).setBackground("#eeeeee");

		// Nova linha criada
		return true;
	}

	// Se a pessoa já estiver registrado na planilha gerencial mas não estiver na planilha de interesse
	abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

	// Nenhuma linha criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosEnvioMapa(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(linhaAtual);
		return false;
	}

	// Pegando os valores da linha
	const valLinha = abaEnvioMapa.getRange(linhaAtual, 1, 1, ultimaColunaEnvioMapa).getValues()[0];

	const dataMapa = valLinha[colDataEnvioMapa - 1];
	const prazoEnvioMapa = valLinha[colPrazoEnvioMapa - 1];
	// Caso ainda não existir prazo, calcular um novo adicionando 7 dias
	const dataPrazo = !prazoEnvioMapa && dataMapa ? new Date(dataMapa.setDate(dataMapa.getDate() + 7)) : prazoEnvioMapa;

	// Considerando a ordem dos campos da planilha Gerencial (Ver arquivo Constants)
	const intervaloInserir = [
		valLinha[colLinkMapa - 1],
		valLinha[colTextoMapa - 1],
		dataPrazo,
		(valLinha[colComentarioEnviadoMapa - 1] || '').toUpperCase(),
		valLinha[colMensagemVerificacaoMapa - 1]
	]

	abaGerencial.getRange(linhaCampoGerencial, colLinkMapaGerencial, 1, 5).setValues([intervaloInserir]);

	// Nenhuma linha nova criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosMarcoFinal(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(linhaAtual);
		return false;
	}

	// Pegando os valores da linha
	const valLinha = abaMarcoFinal.getRange(linhaAtual, 1, 1, ultimaColunaMarcoFinal).getValues()[0];

	// Considerando a ordem dos campos da planilha Gerencial (Ver arquivo Constants)
	const intervaloInserir = [
		"SIM",
		(valLinha[colEnviouReflexaoMarcoFinal - 1] || '').toUpperCase(),
		valLinha[colPrazoEnvioMarcoFinal - 1],
		(valLinha[colComentarioEnviadoMarcoFinal - 1] || '').toUpperCase()
	]

	abaGerencial.getRange(linhaCampoGerencial, colRespondeuMarcoFinalGerencial, 1, 4).setValues([intervaloInserir]);

	// Nenhuma linha criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosCertificado(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(linhaAtual);
		return false;
	}

	// Pegando os valores da linha
	const valLinha = abaCertificado.getRange(linhaAtual, 1, 1, ultimaColunaCertificado).getValues()[0];

	const valEntrouGrupo = valLinha[colEntrouGrupoCertificado - 1];
	const entrouGrupoCertificado = valEntrouGrupo && valEntrouGrupo != "Enviei email" ? valEntrouGrupo.toUpperCase() : valEntrouGrupo;

	// Considerando a ordem dos campos da planilha Gerencial (Ver arquivo Constants)
	const intervaloInserir = [
		valLinha[colDataCertificado - 1],
		valLinha[colLinkCertificado - 1],
		(valLinha[colLinkTestadoCertificado - 1] || '').toUpperCase(),
		entrouGrupoCertificado
	]

	abaGerencial.getRange(linhaCampoGerencial, colTerminouCursoGerencial).setValue("SIM");
	abaGerencial.getRange(linhaCampoGerencial, colDataCertificadoGerencial, 1, 4).setValues([intervaloInserir]);

	// Nenhuma linha criada
	return false;
}

// Função que irá lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero
function LidarComPessoaNaoCadastrada() {

}

// Função que sincronizará quem entrou no whatsapp entre as três planilhas
function SincronizarWhatsGerencial() {
	// Sincronize as planilhas Interesse e Marco Zero, depois as planilhas Interesse e Gerencial e por fim, Interesse e Marco Zero de novo
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaGerencial, colWhatsGerencial);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
}

// Função que sincronizará um dado campo entre duas planilhas passadas
function SincronizarCampoPlanilhas(abaDesejada1, colDesejada1, abaDesejada2, colDesejada2) {
	// Atribui os variáveis de acordo com a abaDesejada1
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada1) || {};

	for (let i = 2; i <= ultimaLinha; i++) {
		const emailDesejada1 = abaDesejada1.getRange(i, colEmail).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailDesejada1)
			continue;

		// Pegue a linha do campo na planilha desejada 2
		const linhaCampoDesejada2 = RetornarLinhaEmailPlanilha(emailDesejada1, abaDesejada2);

		// Se o email for encontrado na outra planilha
		if (linhaCampoDesejada2) {
			const celDesejada1 = abaDesejada1.getRange(i, colDesejada1);
			const valDesejada1 = celDesejada1.getValue();
			const celDesejada2 = abaDesejada2.getRange(linhaCampoDesejada2, colDesejada2);
			const valDesejada2 = celDesejada2.getValue();

			// Se o campo do Desejada1 estiver vazio, altere o campo do Desejada1 com o valor da outra planilha
			if (!valDesejada1) {
				celDesejada1.setValue(valDesejada2);
				continue;
			}

			// Se o campo da outra planilha estiver vazio, altere o campo da outra planilha com o valor do Desejada1
			if (!valDesejada2) {
				celDesejada2.setValue(valDesejada1);
				continue;
			}

			// Se o campo do Desejada1 estiver como sim e da outra como não, altere o campo da outra planilha
			if (valDesejada1 == "SIM" && valDesejada2 == "NÃO") {
				celDesejada2.setValue("SIM");
				continue;
			}

			// Se o campo da outra planilha estiver como sim e da outra como não, altere o campo do Desejada1
			if (valDesejada2 == "SIM" && valDesejada1 == "NÃO") {
				celDesejada1.setValue("SIM");
				continue;
			}
		}
	}
}

//Função que verifica se a pessoa está cadastrada na planilha para verificar, e registra em outra planilha
function VerificarEMarcarCadastroOutraPlanilha(abaParaRegistro, colParaRegistro, abaParaVerificar, valCustomizadoSim, valCustomizadoNao) {
	// Atribui os variáveis de acordo com a abaDesejada
	const { ultimaLinha, colEmail } = objetoMap.get(abaParaRegistro) || {};

	//Pegar o email na planilha Desejada
	for (let i = 2; i <= ultimaLinha; i++) {
		const celParaRegistro = abaParaRegistro.getRange(i, colParaRegistro);
		const valParaRegistro = celParaRegistro.getValue();

		// Se o campo já estiver marcado com sim, passe para o próximo
		if (valParaRegistro == "SIM") continue;

		const email = abaParaRegistro.getRange(i, colEmail).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!email) {
			continue;
		}

		if (RetornarLinhaEmailPlanilha(email, abaParaVerificar)) {
			celParaRegistro.setValue(valCustomizadoSim ?? "SIM");
		} else {
			celParaRegistro.setValue(valCustomizadoNao ?? "NÃO");
		}
	}
}

function CriaContatos() {
	// for para percorrer todas as linhas
	for (let i = ultimaLinhaAnalisadaWhatsGerencial; i <= ultimaLinhaGerencial; i++) {
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