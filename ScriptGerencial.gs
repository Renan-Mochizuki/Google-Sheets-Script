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
//    LidarComPessoaNaoCadastrada(linhaAtual, linhaVazia, abaDesejada):
//    - função genérica para lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero   
//    InserirRedirecionamentoPlanilha(linhaAtual, colInserir, urlInteresse, linhaDestino):
//    - insere um link em um campo para um campo específico em outra planilha
//    SincronizarWhatsGerencial():
//    - sincroniza o campo do whatsapp entre todas as planilhas
//    SincronizarCampoPlanilhas(abaDesejada1, colDesejada1, abaDesejada2, colDesejada2):
//    - sincroniza um campo escolhido entre duas planilhas desejadas
//    CompararValoresEMarcar(celDesejada1, celDesejada2):
//    - função genérica usada pela função SincronizarCampoPlanilhas para sincronizar dois campos de "SIM" ou "NÃO"
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

	// Pegando todos os emails
	const emails = abaDesejada.getRange(2, colEmail, ultimaLinha, 1).getValues();

	//Conferir todos os emails da planilha desejada
	for (let i = 0; i < ultimaLinha - 1; i++) {
		if (emailProcurado == emails[i][0]) return i;
	}
	// Se não for encontrado nenhum 
	return false;
}

// Função que verificará se o email existe na planilha desejada e retornará a linha
function RetornarLinhaEmailDados(emailProcurado, dados) {
	//Conferir todos os emails da planilha desejada
	for (let i = 0; i < dados.length; i++) {
		// Se o email for encontrado, retorne o indice da array + 2 (Porque a array começa em 0 e a planilha em 2)
		if (emailProcurado == dados[i]) return i + 2;
	}
	// Se não for encontrado nenhum 
	return false;
}

// Função que executa as funções necessárias para importar todos os dados
function Importar() {
	const tempoNotificacao = 5;
	const tituloToast = 'Executando funções';
	let linhaVazia, linhasAfetadas, totalLinhasAfetadas = 0;
	// Chamando funções das planilhas para atualizar seus campos
	// planilhaAtiva.toast('Sincronizando campos Whats', tituloToast, tempoNotificacao);
	// SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
	// // Verifica na planilha Interesse, quem respondeu o Marco Zero, e verifica na planilha Marco Zero, quem respondeu o Interesse
	// planilhaAtiva.toast('Verificando respostas Marco Zero', tituloToast, tempoNotificacao);
	// VerificarEMarcarCadastroOutraPlanilha(abaInteresse, colRespondeuMarcoZeroInteresse, abaMarcoZero);
	// planilhaAtiva.toast('Verificando respostas Interesse', tituloToast, tempoNotificacao);
	// VerificarEMarcarCadastroOutraPlanilha(abaMarcoZero, colRespondeuInteresseMarcoZero, abaInteresse, null, "S. PÚBLICA");

	planilhaAtiva.toast('Importando dados da Interesse', tituloToast, tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaInteresse));
	totalLinhasAfetadas += linhasAfetadas;

	planilhaAtiva.toast('Importando dados do Marco Zero', tituloToast, tempoNotificacao);
	({ linhaVazia, linhasAfetadas } = ImportarDados(abaMarcoZero));
	totalLinhasAfetadas += linhasAfetadas;

	// planilhaAtiva.toast('Importando dados do Envio de Mapa', tituloToast, tempoNotificacao);
	// ({ linhaVazia, linhasAfetadas } = ImportarDados(abaEnvioMapa));
	// totalLinhasAfetadas += linhasAfetadas;

	// planilhaAtiva.toast('Importando dados do Marco Final', tituloToast, tempoNotificacao);
	// ({ linhaVazia, linhasAfetadas } = ImportarDados(abaMarcoFinal));
	// totalLinhasAfetadas += linhasAfetadas;

	// planilhaAtiva.toast('Importando dados do Envio do Certificado', tituloToast, tempoNotificacao);
	// ({ linhaVazia, linhasAfetadas } = ImportarDados(abaCertificado));
	// totalLinhasAfetadas += linhasAfetadas;

	const linhasCriadas = linhaVazia - ultimaLinhaGerencial - 1;
	const mensagem = 'Fim da execução.\n' + linhasCriadas + ' linhas criadas\n' + totalLinhasAfetadas + ' linhas afetadas';
	planilhaAtiva.toast(mensagem, 'Execução finalizada', tempoNotificacao + 5);

}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
	// Pegando a próxima linha vazia da planilha gerencial
	let linhaVazia = abaGerencial.getLastRow() + 1;
	let linhasAfetadas = 0;

	// Atribui as variáveis de acordo com a abaDesejada
	const { ultimaLinhaAnalisada, ultimaLinha, colEmail, ImportarDadosPlanilha } = objetoMap.get(abaDesejada) || {};

	// Pegando todos os emails da abaGerencial
	const emails = abaGerencial.getRange(2, colEmailGerencial, ultimaLinhaGerencial, 1).getValues().flat();
	Logger.log('ultimaLinhaGerencial: '+ ultimaLinhaGerencial);
	Logger.log('emails: '+emails);
	// Loop da planilha Desejada
	for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
		const email = abaDesejada.getRange(i, colEmail).getValue();
		if(abaDesejada == abaMarcoZero) Logger.log(emails)
		// Se não existir email, passe para o próximo
		if (!email) continue;
		const linhaCampoGerencial = RetornarLinhaEmailDados(email, emails);

		const novoEmailCriado = ImportarDadosPlanilha(i, linhaCampoGerencial, linhaVazia);

		if (novoEmailCriado) {
			linhaVazia++
			// Insira o novo email na array de emails (Se o primeiro item estiver vazio, substitua ele)
			emails[0] ? emails.push(novoEmailCriado) : (emails[0] = novoEmailCriado);
		}
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
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectInteresseGerencial, urlInteresse, linhaAtual);

		// Nova linha criada
		const emailCriado = valLinha[colEmailInteresse - 1]
		return emailCriado;
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

	// Se aquela pessoa já estiver na gerencial
	if (linhaCampoGerencial) {
		InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);
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
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

		// Pintando campos cidade e estado e redirecionamento para interesse
		abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setBackground("#eeeeee");
		abaGerencial.getRange(linhaVazia, colRedirectInteresseGerencial).setBackground("#eeeeee");

		// Nova linha criada
		const emailCriado = valLinha[colEmailMarcoZero - 1]
		return emailCriado;
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
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectEnvioMapaGerencial, urlEnvioMapa, linhaAtual);

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
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoFinalGerencial, urlMarcoFinal, linhaAtual);

	// Nenhuma linha criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosCertificado(linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(linhaAtual, linhaVazia, abaCertificado);
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
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectCertificadoGerencial, urlCertificado, linhaAtual);

	// Nenhuma linha criada
	return false;
}

// Função que irá lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero
function LidarComPessoaNaoCadastrada(linhaAtual, linhaVazia, abaDesejada) {
	// Pegando os valores da linha
	const email = abaCertificado.getRange(linhaAtual, colEmailCertificado).getValue();
	Logger.log('Email não cadastrado: ' + email);
}

// Função que adiciona um link para redirecionamento na planilha gerencial
function InserirRedirecionamentoPlanilha(linhaInserir, colInserir, urlDestino, linhaDestino) {
	// Expressão regular para extrair o ID da planilha e o ID da aba pelo link
	const regex = /\/d\/([a-zA-Z0-9-_]+).*gid=([0-9]+)/;
	const matches = urlDestino.match(regex);

	if (!matches) return;

	const planilhaID = matches[1];
	const abaID = matches[2];
	const urlRedirecionamento = `https://docs.google.com/spreadsheets/d/${planilhaID}/edit#gid=${abaID}&range=A${linhaDestino}`;

	abaGerencial.getRange(linhaInserir, colInserir).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
	abaGerencial.getRange(linhaInserir, colInserir).setValue(urlRedirecionamento);
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
	// Atribui as variáveis de acordo com a abaDesejada1
	const { ultimaLinha: ultimaLinha1, colEmail: colEmail1 } = objetoMap.get(abaDesejada1) || {};

	// Atribui as variáveis de acordo com a abaDesejada2
	const { ultimaLinha: ultimaLinha2, colEmail: colEmail2 } = objetoMap.get(abaDesejada2) || {};

	// Pegando todos os emails da abaDesejada2
	const emails = abaDesejada2.getRange(2, colEmail2, ultimaLinha2, 1).getValues().flat();

	for (let i = 2; i <= ultimaLinha1; i++) {
		const emailDesejada1 = abaDesejada1.getRange(i, colEmail1).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!emailDesejada1)
			continue;

		// Pegue a linha do campo na planilha desejada 2
		const linhaCampoDesejada2 = RetornarLinhaEmailDados(emailDesejada1, emails);

		// Se o email for encontrado na outra planilha
		if (linhaCampoDesejada2) {
			const celDesejada1 = abaDesejada1.getRange(i, colDesejada1);
			const celDesejada2 = abaDesejada2.getRange(linhaCampoDesejada2, colDesejada2);

			CompararValoresEMarcar(celDesejada1, celDesejada2);
		}
	}
}

// Função genérica que compara dois campos que possam conter "SIM" ou "NÃO" e sincroniza eles
function CompararValoresEMarcar(celDesejada1, celDesejada2) {
	const valDesejada1 = celDesejada1.getValue();
	const valDesejada2 = celDesejada2.getValue();

	// Se o campo do Desejada1 estiver vazio, altere o campo do Desejada1 com o valor da outra planilha
	if (!valDesejada1) {
		celDesejada1.setValue(valDesejada2);
		return;
	}

	// Se o campo da outra planilha estiver vazio, altere o campo da outra planilha com o valor do Desejada1
	if (!valDesejada2) {
		celDesejada2.setValue(valDesejada1);
		return;
	}

	// Se o campo do Desejada1 estiver como sim e da outra como não, altere o campo da outra planilha
	if (valDesejada1 == "SIM" && valDesejada2 == "NÃO") {
		celDesejada2.setValue("SIM");
		return;
	}

	// Se o campo da outra planilha estiver como sim e da outra como não, altere o campo do Desejada1
	if (valDesejada2 == "SIM" && valDesejada1 == "NÃO") {
		celDesejada1.setValue("SIM");
		return;
	}
}

//Função que verifica se a pessoa está cadastrada na planilha para verificar, e registra em outra planilha
function VerificarEMarcarCadastroOutraPlanilha(abaParaRegistro, colParaRegistro, abaParaVerificar, valCustomizadoSim, valCustomizadoNao) {
	// Atribui as variáveis de acordo com a abaDesejada
	const { ultimaLinha: ultimaLinhaRegistro, colEmail: colEmailRegistro } = objetoMap.get(abaParaRegistro) || {};
	const { ultimaLinha: ultimaLinhaVerificar, colEmail: colEmailVerificar } = objetoMap.get(abaParaVerificar) || {};

	const emailsAbaParaVerificar = abaParaVerificar.getRange(2, colEmailVerificar, ultimaLinhaVerificar, 1).getValues().flat();

	//Pegar o email na planilha Desejada
	for (let i = 2; i <= ultimaLinhaRegistro; i++) {
		const celParaRegistro = abaParaRegistro.getRange(i, colParaRegistro);
		const valParaRegistro = celParaRegistro.getValue();

		// Se o campo já estiver marcado com sim, passe para o próximo
		if (valParaRegistro == "SIM") continue;

		const email = abaParaRegistro.getRange(i, colEmailRegistro).getValue();

		// Se o campo estiver vazio, passe para o próximo
		if (!email) {
			continue;
		}

		if (RetornarLinhaEmailDados(email, emailsAbaParaVerificar)) {
			celParaRegistro.setValue(valCustomizadoSim ?? "SIM");
		} else {
			celParaRegistro.setValue(valCustomizadoNao ?? "NÃO");
		}
	}
}


function VerificarRepeticoesGerencial() {
	VerificarRepeticoes(abaGerencial)
}

function VerificarRepeticoes(abaDesejada) {
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada) || {};
	const emails = abaDesejada.getRange(2, colEmail, ultimaLinha, 1).getValues().flat();

	for (let i = 0; i < emails.length; i++) {
		const email = emails[i];
		if (emails.indexOf(email) !== i) {
			Logger.log(email);
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