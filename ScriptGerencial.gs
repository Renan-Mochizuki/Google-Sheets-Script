const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
	ui.createMenu('Menu de Funções')
		.addItem('📂 Importar Dados', 'Importar')
		.addItem('📞 Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
		// .addItem('👤 Criar contatos', 'CriaContatos')
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
function RetornarLinhaDados(emailProcurado, telefoneProcurado, dados) {
	// Separando o email procurado pois ele pode ser um valor com mais de um email
	const emailsProcuradosSeparados = emailProcurado ? NormalizarString(emailProcurado).split(';') : [NormalizarString(emailProcurado)];
	const telefonesPorcuradosSeparados = telefoneProcurado ? telefoneProcurado.toString().split(';') : [telefoneProcurado];

	// Conferir todos os emails da planilha desejada
	for (let i = 0; i < dados.length; i++) {
		const emailDados = dados[i][0];
		const telefoneDados = dados[i][1];

		if (emailDados && typeof emailDados === 'string'){
			const emailsSeparados = emailDados.split(';');

			for (let emailSeparado of emailsSeparados){
				for(let emailProcuradoSeparado of emailsProcuradosSeparados){
					// Se o email for encontrado, retorne o indice da array + 2 (Porque a array começa em 0 e a planilha em 2)
					if (CompararSimilaridade(emailProcuradoSeparado, emailSeparado)) return i + 2;
				}
			}
		}
		
		if(telefoneDados){
			const telefonesSeparados = telefoneDados.toString().split(';');

			for (let telefoneSeparado of telefonesSeparados){
				for(let telefoneProcuradoSeparado of telefonesPorcuradosSeparados){
					// Se o telefone for encontrado, retorne o indice da array + 2 (Porque a array começa em 0 e a planilha em 2)
					if (CompararSimilaridade(telefoneProcuradoSeparado, telefoneSeparado, 0.9)) return i + 2;
				}
			}
		}
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

	planilhaAtiva.toast(tituloToast, 'Importando notas da Interesse', tempoNotificacao);
	ImportarNotas(abaInteresse);
	
	planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Zero', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaMarcoZero);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Envio de Mapa', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaEnvioMapa);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Final', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaMarcoFinal);

	planilhaAtiva.toast(tituloToast, 'Importando dados do Envio do Certificado', tempoNotificacao);
	totalLinhasAfetadas += ImportarDados(abaCertificado);

	const quantidadeLinhasCriadas = abaGerencial.getLastRow() - ultimaLinhaGerencial;
	const mensagem = 'Fim da execução.\n' + quantidadeLinhasCriadas + ' linhas criadas e ' + totalLinhasAfetadas + ' linhas já registradas analisadas';
	planilhaAtiva.toast(mensagem, 'Execução finalizada', tempoNotificacao);
}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
	// Pegando a próxima linha vazia da planilha gerencial
	// Obs.: Não se pode usar a variável ultimaLinhaGerencial, pois ela não se atualiza sozinha
	const ultimaLinhaGerencial = abaGerencial.getLastRow();
	let linhaVazia = ultimaLinhaGerencial + 1;
	let linhasAfetadas = 0;

	// Atribui as variáveis de acordo com a abaDesejada
	const { nome, ultimaLinhaAnalisada, ultimaLinha, ultimaColuna, colEmail, colTel, ImportarDadosPlanilha } = objetoMap.get(abaDesejada);

	// Pegando todos os emails e telefones da abaGerencial
	const emailsTelefones = abaGerencial.getRange(2, colEmailGerencial, ultimaLinhaGerencial, 2).getValues();

	// Loop para percorrer todas linhas da planilha Desejada
	for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
		// Pegando os valores da linha e definindo o primeiro item como null para podermos acessar os índices sem precisar subtrair 1
		const valLinha = [null, ...abaDesejada.getRange(i, 1, 1, ultimaColuna).getValues()[0]];

		const email = valLinha[colEmail];
		const telefone = valLinha[colTel]

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		// Toast da mensagem do progresso de execução da função
		if (i % 100 === 0) {
			const tituloToast = Math.round(i / ultimaLinha * 100) + '% concluído da função atual';
			planilhaAtiva.toast('Processo na linha ' + i + ' da planilha ' + nome, tituloToast, tempoNotificacao);
		}

		const linhaCampoGerencial = RetornarLinhaDados(email, telefone, emailsTelefones);
		const foiCastradoNovoEmail = ImportarDadosPlanilha(valLinha, i, linhaCampoGerencial, linhaVazia);

		if (foiCastradoNovoEmail) {
			linhaVazia++;
			// Insira o novo email e tel na matriz de dados (Se o primeiro item estiver vazio, substitua o item vazio)
			if(!emailsTelefones[0][0] && !emailsTelefones[0][1]){
				emailsTelefones[0] = [email, telefone];
				continue;
			}

			emailsTelefones.push([email, telefone]);
		}
		else linhasAfetadas++;
	}

	return linhasAfetadas;
}

// Função com a lógica da importação dos campos da planilha de interesse
function ImportarDadosInteresse(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Declarando uma array com os campos adicionais da planilha Interesse com "SIM" para o campo "Respondeu Interesse" na Gerencial
	// Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
	const intervaloInserir = [
		valLinha[colAnotacaoInteresse],
		null,
		valLinha[colNomeInteresse],
		valLinha[colEmailInteresse],
		valLinha[colTelInteresse],
		valLinha[colCidadeInteresse],
		valLinha[colEstadoInteresse],
		valLinha[colWhatsInteresse],
		"SIM",
		valLinha[colRespondeuMarcoZeroInteresse],
		valLinha[colSituacaoInteresse]
	]

	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colAnotacaoGerencial, 1, 11).setValues([intervaloInserir]);
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectInteresseGerencial, urlInteresse, linhaAtual);

		// Nova linha criada
		return true;
	} else {
		// Pegando os valores daquela linha da planilha gerencial, pois alguem pode responder mais de uma vez
		// Nesse caso, não definiremos o primeiro item como null pois queremos manter os índices originais
		const valLinhaGerencial = abaGerencial.getRange(linhaCampoGerencial, colAnotacaoGerencial, 1, 11).getValues()[0];

		// Juntando os dados já existentes da planilha gerencial com os novos dados
		const intervaloUnido = JuntarDados(valLinhaGerencial, intervaloInserir, colAnotacaoGerencial);
		
		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaCampoGerencial, colAnotacaoGerencial, 1, 11).setValues([intervaloUnido]);
		
		// Nenhuma linha criada
		return false;
	}
}

// Função com a lógica da importação dos campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
	// Pegando o campo se está cadastrada na planilha de interesse
	const respondeuInteresseMarcoZero = valLinha[colRespondeuInteresseMarcoZero];

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

	// Se aquela pessoa ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 8).setValues([intervaloInserir]);
		InserirRedirecionamentoPlanilha(linhaVazia, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

		// Pintando campos cidade e estado, situação e redirecionamento para interesse
		abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setBackground(corCampoSemDados);
		abaGerencial.getRange(linhaVazia, colSituacaoGerencial).setBackground(corCampoSemDados);
		abaGerencial.getRange(linhaVazia, colRedirectInteresseGerencial).setBackground(corCampoSemDados);

		// Nova linha criada
		return true;
	} else {
		// Se a pessoa já estiver registrado na planilha gerencial
		InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

		// Pegando os valores daquela linha da planilha gerencial, pois alguem pode responder mais de uma vez
		// Nesse caso, não definiremos o primeiro item como null pois queremos manter os índices originais
		const valLinhaGerencial = abaGerencial.getRange(linhaCampoGerencial, colNomeGerencial, 1, 8).getValues()[0];

		// Juntando os dados já existentes da planilha gerencial com os novos dados
		const intervaloUnido = JuntarDados(valLinhaGerencial, intervaloInserir, colNomeGerencial);

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaCampoGerencial, colNomeGerencial, 1, 8).setValues([intervaloUnido]);
		
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
		return true;
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
		return true;
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
		// Nova linha criada
		return true;
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
		"Pessoa não cadastrada nas outras planilhas",
		null,
		valLinha[colNome],
		valLinha[colEmail],
		valLinha[colTel]
	]

	abaGerencial.getRange(linhaVazia, colAnotacaoGerencial, 1, 5).setValues([intervaloInserir]);

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
	abaGerencial.getRange(linhaInserir, colInserir).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setValue(urlRedirecionamento);
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
	const { ultimaLinha: ultimaLinha1, colEmail: colEmail1, nome: nome1 } = objetoMap.get(abaDesejada1);
	// Atribui as variáveis de acordo com a abaDesejada2
	const { ultimaLinha: ultimaLinha2, colEmail: colEmail2, nome: nome2 } = objetoMap.get(abaDesejada2);

	// Pegando todos os emails da abaDesejada1 e abaDesejada2
	const emailsTelefones1 = abaDesejada1.getRange(2, colEmail1, ultimaLinha1, 2).getValues();
	const emailsTelefones2 = abaDesejada2.getRange(2, colEmail2, ultimaLinha2, 2).getValues();

	const colsDesejadas1 = abaDesejada1.getRange(2, colDesejada1, ultimaLinha1, 1).getValues();
	const colsDesejadas2 = abaDesejada2.getRange(2, colDesejada2, ultimaLinha2, 1).getValues();

	// Loop para percorrer as linhas da abaDesejada1
	for (let i = 0; i < emailsTelefones1.length; i++) {
		const email = emailsTelefones1[i][0];
		const telefone = emailsTelefones1[i][1];

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		// Pegue a linha do campo na planilha desejada 2
		const linhaCampoDesejada2 = RetornarLinhaDados(email, telefone, emailsTelefones2);

		// Se o email for encontrado na outra planilha
		if (linhaCampoDesejada2) {
			const valDesejada1 = colsDesejadas1[i][0];
			const valDesejada2 = colsDesejadas2[linhaCampoDesejada2 - 2][0];
			// Se o campo do Desejada1 estiver vazio, altere o campo do Desejada1 com o valor da outra planilha
			if (!valDesejada1) colsDesejadas1[i][0] = valDesejada2;
			// Se o campo da outra planilha estiver vazio, altere o campo da outra planilha com o valor do Desejada1
			else if (!valDesejada2) colsDesejadas2[linhaCampoDesejada2 - 2][0] = valDesejada1;
			// Se o campo do Desejada1 estiver como sim e da outra como não, altere o campo da outra planilha
			else if (valDesejada1 == "SIM") colsDesejadas2[linhaCampoDesejada2 - 2][0] = "SIM";
			// Se o campo da outra planilha estiver como sim e da outra como não, altere o campo do Desejada1
			else if (valDesejada2 == "SIM") colsDesejadas1[i][0] = "SIM";
		}
	}
	// Toast da mensagem do progresso de execução da função
	const tituloToast ='50% concluído da função atual';
	const textoToast = 'Sincronizando campo entre planilhas ' + nome1 + ' e ' + nome2;
	planilhaAtiva.toast(textoToast, tituloToast, tempoNotificacao);

	// Loop para percorrer as linhas da abaDesejada2 (Caso houver uma pessoa repetida na abaDesejada2)
	for (let i = 0; i < emailsTelefones2.length; i++) {
		const email = emailsTelefones2[i][0];
		const telefone = emailsTelefones2[i][1];

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		// Pegue a linha do campo na planilha desejada 1
		const linhaCampoDesejada1 = RetornarLinhaDados(email, telefone, emailsTelefones1);

		// Se o email for encontrado na outra planilha
		if (linhaCampoDesejada1) {
			const valDesejada1 = colsDesejadas1[linhaCampoDesejada1 - 2][0];
			const valDesejada2 = colsDesejadas2[i][0];
			// Se o campo do Desejada1 estiver vazio, altere o campo do Desejada1 com o valor da outra planilha
			if (!valDesejada1) colsDesejadas1[linhaCampoDesejada1 - 2][0] = valDesejada2;
			// Se o campo da outra planilha estiver vazio, altere o campo da outra planilha com o valor do Desejada1
			else if (!valDesejada2) colsDesejadas2[i][0] = valDesejada1;
			// Se o campo do Desejada1 estiver como sim e da outra como não, altere o campo da outra planilha
			else if (valDesejada1 == "SIM") colsDesejadas2[i][0] = "SIM";
			// Se o campo da outra planilha estiver como sim e da outra como não, altere o campo do Desejada1
			else if (valDesejada2 == "SIM") colsDesejadas1[linhaCampoDesejada1 - 2][0] = "SIM";
		}
	}
	
	// Inserindo os valores nas planilhas
	abaDesejada1.getRange(2, colDesejada1, ultimaLinha1, 1).setValues(colsDesejadas1);
	abaDesejada2.getRange(2, colDesejada2, ultimaLinha2, 1).setValues(colsDesejadas2);
}

//Função que verifica se a pessoa está cadastrada na planilha para verificar, e registra em outra planilha
function VerificarEMarcarCadastroOutraPlanilha(abaParaRegistro, colParaRegistro, abaParaVerificar, valCustomizadoSim, valCustomizadoNao) {
	// Atribui as variáveis de acordo com a abaDesejada1
	const { ultimaLinha: ultimaLinhaVerificar, colEmail: colEmailVerificar, nome: nomeVerificar } = objetoMap.get(abaParaVerificar);
	// Atribui as variáveis de acordo com a abaParaRegistro
	const { ultimaLinha: ultimaLinhaRegistro, colEmail: colEmailRegistro, nome: nomeRegistro } = objetoMap.get(abaParaRegistro);

	// Pegando todos os emails da abaParaVerificar e abaParaRegistro
	const emailsTelefonesVerificar = abaParaVerificar.getRange(2, colEmailVerificar, ultimaLinhaVerificar, 2).getValues();
	const emailsTelefonesRegistro = abaParaRegistro.getRange(2, colEmailRegistro, ultimaLinhaRegistro, 2).getValues();

	const colsRegistro = abaParaRegistro.getRange(2, colParaRegistro, ultimaLinhaRegistro, 1).getValues();

	// Loop para percorrer as linhas da abaParaRegistro
	for (let i = 0; i < emailsTelefonesRegistro.length; i++) {
		const email = emailsTelefonesRegistro[i][0];
		const telefone = emailsTelefonesRegistro[i][1];

		// Se não existir email, ou for o "teste" passe para o próximo
		if (!email || email.toLowerCase().includes("teste")) continue;

		// Toast da mensagem do progresso de execução da função
		if (i % 300 === 0) {
			const tituloToast = Math.round(i / ultimaLinhaRegistro * 100) + '% concluído da função atual';
			const textoToast = 'Processo na linha ' + i + ' da verificação da planilha ' + nomeRegistro + ' para ' + nomeVerificar;
			planilhaAtiva.toast(textoToast, tituloToast, tempoNotificacao);
		}

		const existeNaAbaVerificar = RetornarLinhaDados(email, telefone, emailsTelefonesVerificar);

		// Se o email for encontrado na outra planilha
		if (existeNaAbaVerificar) {
			colsRegistro[i][0] = valCustomizadoSim ?? "SIM";
		} else {
			colsRegistro[i][0] = valCustomizadoNao ?? "NÃO";
		}
	}
	
	// Inserindo os valores na abaParaRegistro
	abaParaRegistro.getRange(2, colParaRegistro, ultimaLinhaRegistro, 1).setValues(colsRegistro);
}

// Função que importa as anotações
function ImportarNotas(abaDesejada) {
	// Atribui as variáveis de acordo com a abaDesejada
	const { ultimaLinha, colEmail } = objetoMap.get(abaDesejada);
	const ultimaLinhaGerencial = abaGerencial.getLastRow();

	let anotacoesGerencial = abaGerencial.getRange(2, colAnotacaoGerencial, ultimaLinhaGerencial, 1).getValues().flat();
	let emailsTelefonesGerencial = abaGerencial.getRange(2, colEmailGerencial, ultimaLinhaGerencial, 2).getValues();

	const notasColunasAbaDesejada = abaDesejada.getRange(2, colEmail, ultimaLinha, 1).getNotes().flat();
	const emailsTelefonesAbaDesejada = abaDesejada.getRange(2, colEmail, ultimaLinha, 2).getValues();

	for (let i = 0; i < notasColunasAbaDesejada.length; i++) {
		const notaDesejada = notasColunasAbaDesejada[i];

		if (!notaDesejada) continue;

		const email = emailsTelefonesAbaDesejada[i][0];
		const telefone = emailsTelefonesAbaDesejada[i][1];

		if(!email || email.toLowerCase().includes("teste")) continue;

		const linhaCampoGerencial = RetornarLinhaDados(email, telefone, emailsTelefonesGerencial);

		if (!linhaCampoGerencial){
			planilhaAtiva.toast('Email não encontrado na planilha gerencial: ' + email, 'Erro', tempoNotificacao);
			continue;
		}

		const anotacaoGerencial = anotacoesGerencial[linhaCampoGerencial - 2];
		let notaInserir;

		// Se já existir uma anotação na gerencial, e ainda não conter a notaDesejada

		if (anotacaoGerencial){
			if(!(anotacaoGerencial.split(';').includes(notaDesejada))) {
				notaInserir = anotacaoGerencial + '; ' + notaDesejada;
			} else{
				notaInserir = anotacaoGerencial;
			}
		} else {
			notaInserir = notaDesejada;
		}

		anotacoesGerencial[linhaCampoGerencial - 2] = notaInserir;

		const emailGerencial = emailsTelefonesGerencial[linhaCampoGerencial - 2][0];
		const regex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
		const emailsDaNota = notaDesejada.match(regex) || [];
	
		for(let emailNota of emailsDaNota){
			if(!emailGerencial.includes(emailNota)){
				emailsTelefonesGerencial[linhaCampoGerencial - 2][0] = emailGerencial + '; ' + emailNota;
			}
		}
	}

	abaGerencial.getRange(2, colAnotacaoGerencial, ultimaLinhaGerencial, 1).setValues(anotacoesGerencial.map(nota => [nota])); // Revertendo o .flat()
	abaGerencial.getRange(2, colEmailGerencial, ultimaLinhaGerencial, 2).setValues(emailsTelefonesGerencial);
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


function JuntarDados(dadosLinha1, dadosLinha2, primeiraColunaDoIntervalo){

	const primeiraColuna = primeiraColunaDoIntervalo ?? colNomeGerencial;

	let dadosConcatenados = [];
	const colunasDeSimNao = [colTerminouCursoGerencial, colWhatsGerencial, colRespondeuInteresseGerencial, colRespondeuMarcoZeroGerencial, colComentarioEnviadoMapaGerencial, colRespondeuMarcoFinalGerencial, colEnviouReflexaoMarcoFinalGerencial, colComentarioEnviadoMarcoFinalGerencial, colLinkTestadoCertificadoGerencial, colEntrouGrupoCertificadoGerencial];

	for (let i = 0; i < dadosLinha1.length; i++) {
		const colunaAtual = i + primeiraColuna;
		let possuiSimilaridade = false;
		let dado1 = dadosLinha1[i];
		let dado2 = dadosLinha2[i];

		// Exceções especiais
		if(colunasDeSimNao.includes(colunaAtual)){
			dadosConcatenados.push(RetornarValorSimNao(dado1, dado2));
			continue;
		}
		if(colunaAtual == colSituacaoGerencial){
			dadosConcatenados.push(RetornarTurmaMaisRecente(dado1, dado2));
			continue;
		}

		// Se o dado1 não existir, adicione o dado2
		if(!dado1) {
			dadosConcatenados.push(dado2);
			continue;
		}
		if(dado2) {
			// Separe o texto pelo ; para caso o campo já tiver sido concatenado
			const textosSeparados1 = dado1.toString().split(';');
			
			// Loop para comparar a similaridade para cada um dos textos
			for(let texto of textosSeparados1){
				if(CompararSimilaridade(texto, dado2, 0.9)) possuiSimilaridade = true;
			}

			if(!possuiSimilaridade) {
				dadosConcatenados.push(dado1 + '; ' + dado2);
				continue;
			}
		}
		dadosConcatenados.push(dado1);
	}

	return dadosConcatenados;
}

function RetornarTurmaMaisRecente(string1, string2) {
	// Se uma string não existir, ou for 'ESPERA', retorne a outra
	if (!string1 || string1 === 'ESPERA') return string2;
	if (!string2 || string2 === 'ESPERA') return string1;

	// Separar os números antes do traço (T01-2024 => 01; 2024)
	const regex = /(\d+)-(\d+)/;
	const match1 = string1.match(regex);
	const match2 = string2.match(regex);

	if (!match1 || !match2) return string1;

	// Verificar qual ano é maior
	if(match1[2] > match2[2]) return string1;
	if(match1[2] < match2[2]) return string2;

	// Anos iguais, verificar qual turma é maior
	if(match1[1] > match2[1]) return string1;
	if(match1[1] < match2[1]) return string2;

	return string1;
}

function RetornarValorSimNao(valor1, valor2){
	if(!valor1) return valor2;
	if(!valor2) return valor1;
	if(valor1 == "SIM" || valor2 == "SIM") return "SIM";
	return valor1;
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