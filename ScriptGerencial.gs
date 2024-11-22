const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
	ui.createMenu('Menu de Funções')
		.addItem('📂 Importar Dados', 'Importar')
		.addItem('📞 Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
		.addItem('👤 Criar contatos', 'CriaContatos')
		.addItem('🗑️ Excluir todos os campos', 'LimparPlanilha')
		.addSeparator()
		.addSubMenu(ui.createMenu('Formatação da planilha')
			.addItem('Formatar campos telefone', 'FormatarLinhasTelefone')
			.addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
			.addItem('Remover linhas vazias', 'RemoverLinhasVazias')
			.addItem('Mostrar todas linhas', 'MostrarTodasLinhas')
			.addItem('Esconder linhas', 'MostrarInterfaceEsconderLinhas'))
		.addToUi();
}


// -- IMPORTANTE --
// VEJA OS COMENTÁRIOS DO ARQUIVO CONSTANTS


// Função que verificará se o email existe na planilha desejada e retornará a linha
function RetornarLinhaEmailDados(emailProcurado, dados) {
	//Conferir todos os emails da planilha desejada
	for (let i = 0; i < dados.length; i++) {
		if (!dados[i] || typeof dados[i] !== 'string') continue;

		// Se o email for encontrado, retorne o indice da array + 2 (Porque a array começa em 0 e a planilha em 2)
		if (emailProcurado.toLowerCase() == dados[i].toLowerCase()) return i + 2;
	}
	// Se não for encontrado nenhum 
	return false;
}

// Função que executa outras funções para importar os dados de cada planilha
function Importar() {
	const tituloToast = 'Executando funções';
	let totalLinhasAfetadas = 0;

	// Formatando os telefones de todas as planilhas
	planilhaAtiva.toast('Formatando telefones de todas planilhas', tituloToast, tempoNotificacao);
	FormatarLinhasTelefoneTodasAbas();
	// Chamando funções das planilhas para atualizar seus campos
	planilhaAtiva.toast('Sincronizando campos Whats', tituloToast, tempoNotificacao);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
	// Verifica na planilha Interesse, quem respondeu o Marco Zero, e verifica na planilha Marco Zero, quem respondeu o Interesse
	planilhaAtiva.toast('Verificando respostas Marco Zero', tituloToast, tempoNotificacao);
	VerificarEMarcarCadastroOutraPlanilha(abaInteresse, colRespondeuMarcoZeroInteresse, abaMarcoZero);
	planilhaAtiva.toast('Verificando respostas Interesse', tituloToast, tempoNotificacao);
	VerificarEMarcarCadastroOutraPlanilha(abaMarcoZero, colRespondeuInteresseMarcoZero, abaInteresse, null, "S. PÚBLICA");

	planilhaAtiva.toast(tituloToast, 'Importando dados da Interesse', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaInteresse);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Zero', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaMarcoZero);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Envio de Mapa', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaEnvioMapa);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Final', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaMarcoFinal);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Envio do Certificado', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaCertificado);

	const quantidadeLinhasCriadas = abaGerencial.getLastRow() + 1 - ultimaLinhaGerencial;
	const mensagem = 'Fim da execução.\n' + quantidadeLinhasCriadas + ' linhas criadas\n' + totalLinhasAfetadas + ' linhas afetadas';
	planilhaAtiva.toast(mensagem, 'Execução finalizada', tempoNotificacao);
}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
	// Pegando a próxima linha vazia da planilha gerencial
	// Obs.: Não se pode usar a variável ultimaLinhaGerencial, pois ela não se atualiza sozinha
	let linhaVazia = abaGerencial.getLastRow() + 1;
	let linhasAfetadas = 0;

	// Atribui as variáveis de acordo com a abaDesejada
	const { nome, ultimaLinhaAnalisada, ultimaLinha, ultimaColuna, colEmail, ImportarDadosPlanilha } = objetoMap.get(abaDesejada);

	// Chamando a função importar anotações apenas quando estivermos no Marco Zero pois os dados da interesse já estarão na gerencial
	if (abaDesejada == abaMarcoZero) ImportarNotas(abaInteresse);

	// Pegando todos os emails da abaGerencial e da abaDesejada
	// Utilizando flat() pois o getValues() de um intervalo retorna uma matriz e flat() a transforma em uma array
	const emails = abaGerencial.getRange(2, colEmailGerencial, abaGerencial.getLastRow(), 1).getValues().flat();

	// Loop para percorrer todas linhas da planilha Desejada
	for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
		// Pegando os valores da linha e definindo o primeiro item como null para podermos acessar os índices sem precisar subtrair 1
		const valLinha = [null, ...abaDesejada.getRange(i, 1, 1, ultimaColuna).getValues()[0]];

		let email = valLinha[colEmail];

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		email = email.trim(); // Remove espaços em branco

		// Toast da mensagem do progresso de execução da função
		if (i % 100 === 0) {
			const tituloToast = Math.round(i / ultimaLinha * 100) + '% concluído da função atual';
			planilhaAtiva.toast('Processo na linha ' + i + ' da planilha ' + nome, tituloToast, tempoNotificacao);
		}

		const linhaCampoGerencial = RetornarLinhaEmailDados(email, emails);
		const foiCastradoNovoEmail = ImportarDadosPlanilha(valLinha, i, linhaCampoGerencial, linhaVazia);

		if (foiCastradoNovoEmail) {
			linhaVazia++;
			// Insira o novo email na array de emails (Se o primeiro item estiver vazio, substitua o item vazio)
			emails[0] ? emails.push(foiCastradoNovoEmail) : (emails[0] = foiCastradoNovoEmail);
		}
		else linhasAfetadas++;
	}

	return linhasAfetadas;
}

// Função com a lógica da importação dos campos da planilha de interesse
function ImportarDadosInteresse(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Declarando uma array com os campos adicionais da planilha Interesse com "SIM" para o campo "Respondeu Interesse" na Gerencial
	const intervaloAdicionais = [
		valLinha[colWhatsInteresse],
		"SIM",
		valLinha[colRespondeuMarcoZeroInteresse],
		valLinha[colSituacaoInteresse]
	];

	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
		const intervaloInserir = [
			valLinha[colNomeInteresse],
			valLinha[colEmailInteresse],
			valLinha[colTelInteresse],
			valLinha[colCidadeInteresse],
			valLinha[colEstadoInteresse],
			...intervaloAdicionais
		]

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 9).setValues([intervaloInserir]);
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectInteresseGerencial, urlInteresse, linhaAtual);
		AdicionarAnotacaoGerencial(linhaVazia, valLinha[colAnotacaoInteresse]);

		// Nova linha criada
		const emailCriado = valLinha[colEmailInteresse]
		return emailCriado;
	} else {
		// Se o email já estiver registrado na planilha gerencial, atualize os campos adicionais
		abaGerencial.getRange(linhaCampoGerencial, colWhatsGerencial, 1, 4).setValues([intervaloAdicionais]);

		// Nenhuma linha criada
		return false;
	}
}

// Função com a lógica da importação dos campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Pegando o campo se está cadastrada na planilha de interesse
	const respondeuInteresseMarcoZero = valLinha[colRespondeuInteresseMarcoZero];

	// Se aquela pessoa ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
		const intervaloInserir = [
			valLinha[colNomeMarcoZero],
			valLinha[colEmailMarcoZero],
			valLinha[colTelMarcoZero],
			null,
			null,
			valLinha[colWhatsMarcoZero],
			respondeuInteresseMarcoZero,
			"SIM"
		]

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 8).setValues([intervaloInserir]);
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

		// Pintando campos cidade e estado, situação e redirecionamento para interesse
		abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setBackground(corCampoSemDados);
		abaGerencial.getRange(linhaVazia, colSituacaoGerencial).setBackground(corCampoSemDados);
		abaGerencial.getRange(linhaVazia, colRedirectInteresseGerencial).setBackground(corCampoSemDados);

		// Nova linha criada
		const emailCriado = valLinha[colEmailMarcoZero]
		return emailCriado;
	} else {
		// Se a pessoa já estiver registrado na planilha gerencial
		InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);
		abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

		// Nenhuma linha criada
		return false;
	}
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosEnvioMapa(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(valLinha, linhaAtual, linhaVazia, abaEnvioMapa);
		// Nova linha criada
		const emailCriado = valLinha[colEmailMarcoZero]
		return emailCriado;
	}

	const dataMapa = valLinha[colDataEnvioMapa];
	const prazoEnvioMapa = valLinha[colPrazoEnvioMapa];

	// Caso ainda não existir prazo, calcular um novo adicionando 7 dias
	const dataPrazo = (!prazoEnvioMapa && dataMapa) ? new Date(dataMapa.setDate(dataMapa.getDate() + 7)) : prazoEnvioMapa;
	const comentarioEnviadoMapa = (valLinha[colComentarioEnviadoMapa] || '').toUpperCase()

	// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
	const intervaloInserir = [
		valLinha[colLinkMapa],
		valLinha[colTextoMapa],
		dataPrazo,
		comentarioEnviadoMapa,
		valLinha[colMensagemVerificacaoMapa]
	]

	abaGerencial.getRange(linhaCampoGerencial, colLinkMapaGerencial, 1, 5).setValues([intervaloInserir]);
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectEnvioMapaGerencial, urlEnvioMapa, linhaAtual);

	// Nenhuma linha nova criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosMarcoFinal(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(valLinha, linhaAtual, linhaVazia, abaMarcoFinal);
		// Nova linha criada
		const emailCriado = valLinha[colEmailMarcoZero]
		return emailCriado;
	}

	const enviouReflexaoMarcoFinal = (valLinha[colEnviouReflexaoMarcoFinal] || '').toUpperCase();
	const comentarioEnviadoMarcoFinal = (valLinha[colComentarioEnviadoMarcoFinal] || '').toUpperCase()

	// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
	const intervaloInserir = [
		"SIM",
		enviouReflexaoMarcoFinal,
		valLinha[colPrazoEnvioMarcoFinal],
		comentarioEnviadoMarcoFinal
	]

	abaGerencial.getRange(linhaCampoGerencial, colRespondeuMarcoFinalGerencial, 1, 4).setValues([intervaloInserir]);
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoFinalGerencial, urlMarcoFinal, linhaAtual);

	// Nenhuma linha criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosCertificado(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(valLinha, linhaAtual, linhaVazia, abaCertificado);
		AdicionarAnotacaoGerencial(linhaAtual,);
		// Nova linha criada
		const emailCriado = valLinha[colEmailMarcoZero]
		return emailCriado;
	}

	const linkTestadoCertificado = (valLinha[colLinkTestadoCertificado] || '').toUpperCase();
	const valEntrouGrupo = valLinha[colEntrouGrupoCertificado];

	// Transforme o texto em maísculas se ele não for 'Enviei email'
	const entrouGrupoCertificado = (valEntrouGrupo && valEntrouGrupo != "Enviei email") ? valEntrouGrupo.toUpperCase() : valEntrouGrupo;

	// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
	const intervaloInserir = [
		valLinha[colDataCertificado],
		valLinha[colLinkCertificado],
		linkTestadoCertificado,
		entrouGrupoCertificado
	]

	abaGerencial.getRange(linhaCampoGerencial, colTerminouCursoGerencial).setValue("SIM");
	abaGerencial.getRange(linhaCampoGerencial, colDataCertificadoGerencial, 1, 4).setValues([intervaloInserir]);
	InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectCertificadoGerencial, urlCertificado, linhaAtual);

	// Nenhuma linha criada
	return false;
}

// Função que irá lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero
function LidarComPessoaNaoCadastrada(valLinha, linhaAtual, linhaVazia, abaDesejada) {

	// Atribui as variáveis de acordo com a abaDesejada
	const { colNome, colEmail, colTel, ImportarDadosPlanilha } = objetoMap.get(abaDesejada);

	Logger.log('Email não cadastrado: ' + valLinha[colEmail]);

	// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
	const intervaloInserir = [
		valLinha[colNome],
		valLinha[colEmail],
		valLinha[colTel]
	]

	abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 3).setValues([intervaloInserir]);

	// Preencher os outros dados da planilha
	ImportarDadosPlanilha(valLinha, linhaAtual, linhaVazia, linhaVazia + 1);
}

// Função que adiciona um link para redirecionamento na planilha gerencial
function InserirRedirecionamentoPlanilha(linhaInserir, colInserir, urlDestino, linhaDestino) {
	// Expressão regular para extrair o ID da planilha e o ID da aba pelo link daquela planilha
	const regex = /\/d\/([a-zA-Z0-9-_]+).*gid=([0-9]+)/;
	const matches = urlDestino.match(regex);

	// Se o link não estiver correto, finalize a função
	if (!matches) return;

	const planilhaID = matches[1];
	const abaID = matches[2];
	const urlRedirecionamento = `https://docs.google.com/spreadsheets/d/${planilhaID}/edit#gid=${abaID}&range=A${linhaDestino}`;

	// Adiciona um link para redirecionamento na planilha gerencial
	abaGerencial.getRange(linhaInserir, colInserir).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
	abaGerencial.getRange(linhaInserir, colInserir).setValue(urlRedirecionamento);
}

// Função que sincronizará quem entrou no whatsapp entre as três planilhas
function SincronizarWhatsGerencial() {
	// Sincronize as planilhas Interesse e Marco Zero, depois as planilhas Interesse e Gerencial e por fim, Interesse e Marco Zero de novo
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
	planilhaAtiva.toast('Primeiro processo de sincronização de Whats concluída', '33% concluído da função atual', tempoNotificacao);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaGerencial, colWhatsGerencial);
	planilhaAtiva.toast('Segundo processo de sincronização de Whats concluída', '67% concluído da função atual', tempoNotificacao);
	SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
}

// Função que sincronizará um dado campo entre duas planilhas passadas
function SincronizarCampoPlanilhas(abaDesejada1, colDesejada1, abaDesejada2, colDesejada2) {
	// Atribui as variáveis de acordo com a abaDesejada1
	const { ultimaLinha: ultimaLinha1, colEmail: colEmail1 } = objetoMap.get(abaDesejada1);

	// Atribui as variáveis de acordo com a abaDesejada2
	const { ultimaLinha: ultimaLinha2, colEmail: colEmail2 } = objetoMap.get(abaDesejada2);

	// Pegando todos os emails da abaDesejada2
	const emails1 = abaDesejada1.getRange(2, colEmail1, ultimaLinha1, 1).getValues().flat();
	const emails2 = abaDesejada2.getRange(2, colEmail2, ultimaLinha2, 1).getValues().flat();

	// Loop para percorrer as linhas da abaDesejada1
	for (let i = 0; i < emails1.length; i++) {
		let email = emails1[0];

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		email = email.trim(); // Remove espaços em branco

		// Pegue a linha do campo na planilha desejada 2
		const linhaCampoDesejada2 = RetornarLinhaEmailDados(email, emails2);

		// Se o email for encontrado na outra planilha
		if (linhaCampoDesejada2) {
			// i + 2, pois a array começa em 0 e a planilha começa em 2 
			const celDesejada1 = abaDesejada1.getRange(i + 2, colDesejada1);
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
	// Atribui as variáveis de acordo com as abaDesejadas
	const { ultimaLinha: ultimaLinhaRegistro, colEmail: colEmailRegistro, nome: nomeRegistro } = objetoMap.get(abaParaRegistro);
	const { ultimaLinha: ultimaLinhaVerificar, colEmail: colEmailVerificar, nome: nomeDestino } = objetoMap.get(abaParaVerificar);

	const emailsAbaRegistro = abaParaRegistro.getRange(2, colEmailRegistro, ultimaLinhaRegistro, 1).getValues().flat();
	const emailsAbaVerificar = abaParaVerificar.getRange(2, colEmailVerificar, ultimaLinhaVerificar, 1).getValues().flat();

	//Pegar o email na planilha Desejada
	for (let i = 2; i <= ultimaLinhaRegistro; i++) {
		const celParaRegistro = abaParaRegistro.getRange(i, colParaRegistro);
		const valParaRegistro = celParaRegistro.getValue();

		// Se o campo já estiver marcado com sim, passe para o próximo
		if (valParaRegistro == "SIM") continue;

		let email = abaParaRegistro.getRange(i, colEmailRegistro).getValue();

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		email = email.trim(); // Remove espaços em branco

		// Toast da mensagem do progresso de execução da função
		if (i % 100 === 0) {
			const tituloToast = Math.round(i / ultimaLinhaRegistro * 100) + '% concluído da função atual';
			const textoToast = 'Processo na linha ' + i + ' da verificação da planilha ' + nomeRegistro + ' para ' + nomeDestino;
			planilhaAtiva.toast(textoToast, tituloToast, tempoNotificacao);
		}

		if (RetornarLinhaEmailDados(email, emailsAbaVerificar)) {
			celParaRegistro.setValue(valCustomizadoSim ?? "SIM");
		} else {
			celParaRegistro.setValue(valCustomizadoNao ?? "NÃO");
		}
	}
}

// Função que adiciona uma anotação no campo de anotações da planilha gerencial
function AdicionarAnotacaoGerencial(linhaInserir, anotacaoInserir) {
	// Se a anotacaoInserir existir e não for uma data
	if (anotacaoInserir && !(anotacaoInserir instanceof Date)) {
		const anotacaoGerencial = abaGerencial.getRange(linhaInserir, colAnotacaoGerencial).getValue();
		// Adicione um ponto e vírgula, para adicionar outra anotação se aquela anotação ainda não existir
		if (anotacaoGerencial && !(anotacaoGerencial.split(';').includes(anotacaoInserir))) {
			anotacaoInserir = anotacaoGerencial + '; ' + anotacaoInserir;
		}
		abaGerencial.getRange(linhaInserir, colAnotacaoGerencial).setValue(anotacaoInserir);
	}
}

function SepararEmails(texto) {

}

// Função que importa as anotações
function ImportarNotas(abaDesejada) {
	// Atribui as variáveis de acordo com a abaDesejada
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada);

	// Pega todos os valores da coluna desejada
	const notasColunas = abaDesejada.getRange(2, colEmail, ultimaLinha, 1).getNotes().flat();

	for (let i = 0; i < notasColunas.length; i++) {
		const anotacao = notasColunas[i];

		if (!anotacao) continue;

		// i + 2, pois a array começa em 0 e a planilha começa em 2 
		AdicionarAnotacaoGerencial(i + 2, anotacao);

		// Regex para verificar se há um email escrito na anotação
		const regexEmail = /([A-Za-z0-9._%+-]+)@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/;
		const emailEncontrado = anotacao.match(regexEmail);

		if (emailEncontrado) {
			// Pegue a nota da abaGerencial, se já existir, adicione um ponto e vírgula e o email, se não, apenas atribua o email encontrado
			const notaGerencial = abaGerencial.getRange(i + 2, colEmail).getNote();
			const notaInserir = notaGerencial ? notaGerencial + '; ' + emailEncontrado[0] : emailEncontrado[0];

			abaGerencial.getRange(i + 2, colEmail).setNote(notaInserir);
		}
	}
}

// Função que chama a função de VerificarRepeticoes para ser utilizada no menu 
function VerificarRepeticoesGerencial() {
	VerificarRepeticoes(abaGerencial)
}

// Função que verifica se existe um email repetido
function VerificarRepeticoes(abaDesejada) {
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada);
	const emailsTelefones = abaDesejada.getRange(2, colEmail, ultimaLinha, 2).getValues();

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
					familyName: nomes[lengthNomes]
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