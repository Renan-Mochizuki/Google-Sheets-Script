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
// 	  Ver arquivo constants
//
// -- Funções da Gerencial: --
//    RetornarLinhaEmailGerencial: Retorna a linha daquele email na planilha Gerencial, se não existir, retorna false
//    Importar: chama outras funções para sincronizar as planilhas e chama as funções de importação de todos dados
//    ImportarDados: função genérica para chamar a função de importação de dados para cada planilha
//    ImportarDadosInteresse: Pega todos os dados da Interesse e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosMarcoZero: Pega todos os dados do Marco Zero e move na Gerencial ou apenas o campo respondeu interesse
//    ImportarDadosEnvioMapa: Pega todos os dados do Envio do Mapa e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosMarcoFinal: Pega todos os dados do Marco Final e move na Gerencial ou apenas atualiza os campos adicionais
//    ImportarDadosCertificado: Pega todos os dados do Certificado e move na Gerencial ou apenas atualiza os campos adicionais
//    AtualizarCamposAdicionaisInteresse: Atualiza os campos adicionais da planilha gerencial a partir da planilha de interesse
//    SincronizarWhatsGerencial: Sincroniza o campo do whatsapp entre todas as planilhas
//    SincronizarCampoPlanilhas: Sincroniza um campo escolhido entre as planilha Interesse e Marco Zero ou então, uma desejada
//    CriaContatos (Função não finalizada): Cria contatos no Google People a partir dos dados da planilha Gerencial 
//
// -- Funções de formatação: --
//    LimparPlanilha: Limpa toda a planilha
//    CompletarVaziosComNao: Preenche todos os campos adicionais vazios da planilha gerencial com o texto "NÃO"
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

// Função que executa as funções necessárias para importar todos os dados
function Importar() {
	// Chamando funções das planilhas para atualizar seus campos
	SincronizarCampoPlanilhas(colWhatsInteresse, colWhatsMarcoZero);
	VerificarMarcoZeroInteresse()
	VerificarInteresseMarcoZero();

	ImportarDados(abaInteresse);
	ImportarDados(abaMarcoZero);
	ImportarDados(abaEnvioMapa);
	ImportarDados(abaMarcoFinal);
	ImportarDados(abaCertificado);
}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
	// Pegando a próxima linha vazia da planilha
	let linhaVazia = abaGerencial.getLastRow() + 1;

	// Atribui os variáveis de acordo com a abaDesejada
	let { ultimaLinhaAnalisada, ultimaLinha, colEmail, ImportarDadosPlanilha } = objetoMap.get(abaDesejada) || {};

	// Loop da planilha Desejada
	for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
		const email = abaDesejada.getRange(i, colEmail).getValue();

		// Se não existir email, passe para o próximo
		if (!email) continue;

		const linhaCampoGerencial = RetornarLinhaEmailGerencial(email);

		const novaLinhaCriada = ImportarDadosPlanilha(i, linhaCampoGerencial, linhaVazia);

		if (novaLinhaCriada) linhaVazia++;
	}
}

// Função com a lógica da importação dos campos da planilha de interesse
function ImportarDadosInteresse(i, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		// Pegando os campos nome, email, telefone, cidade e estado
		const intervalo = abaInteresse.getRange(i, colNomeInteresse, 1, 3).getValues();
		const intervaloCidade = abaInteresse.getRange(i, colCidadeInteresse, 1, 2).getValues();

		// Inserindo os campos na planilha gerencial
		abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 3).setValues(intervalo);
		abaGerencial.getRange(linhaVazia, colCidadeGerencial, 1, 2).setValues(intervaloCidade);
		abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue("SIM");

		AtualizarCamposAdicionaisInteresse(i, linhaVazia);

		// Nova linha criada
		return true;
	}

	// Se o email já estiver registrado na planilha gerencial
	AtualizarCamposAdicionaisInteresse(i, linhaCampoGerencial);

	// Nenhuma linha criada
	return false
}

// Função com a lógica da importação dos campos do marco zero que não estão na planilha de interesse
function ImportarDadosMarcoZero(i, linhaCampoGerencial, linhaVazia) {
	// Pegando o campo se está cadastrada na planilha de interesse e pegando a linha desse email na planilha gerencial
	const respondeuInteresseMarcoZero = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero).getValue();

	// Se aquela pessoa não estiver na planilha de interesse
	if (respondeuInteresseMarcoZero != "SIM") {

		// Se aquele email ainda não estiver registrado na planilha gerencial
		if (!linhaCampoGerencial) {

			// Pegando os campos nome, email, telefone e whats
			const intervalo = abaMarcoZero.getRange(i, colNomeMarcoZero, 1, 3).getValues();
			const whats = abaMarcoZero.getRange(i, colWhatsMarcoZero).getValue();

			// Inserindo os campos na planilha gerencial
			abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 3).setValues(intervalo);
			abaGerencial.getRange(linhaVazia, colRespondeuMarcoZeroGerencial).setValue("SIM");
			abaGerencial.getRange(linhaVazia, colWhatsGerencial).setValue(whats);
			abaGerencial.getRange(linhaVazia, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);

			// Pintando campos
			abaGerencial.getRange(linhaVazia, colCidadeGerencial).setBackground("#eeeeee");
			abaGerencial.getRange(linhaVazia, colEstadoGerencial).setBackground("#eeeeee");

			// Nova linha criada
			return true;
		}

		// Se o email já estiver registrado na planilha gerencial mas não estiver na planilha de interesse
		abaGerencial.getRange(linhaCampoGerencial, colRespondeuInteresseGerencial).setValue(respondeuInteresseMarcoZero);
	}
	// Nenhuma linha criada
	return false
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosEnvioMapa(i, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(i);
		return false;
	}

	const dataMapa = abaEnvioMapa.getRange(i, colDataEnvioMapa).getValue();
	const linkMapa = abaEnvioMapa.getRange(i, colLinkMapa).getValue();
	const textoMapa = abaEnvioMapa.getRange(i, colTextoMapa).getValue();
	const comentarioEnviadoMapa = abaEnvioMapa.getRange(i, colComentarioEnviadoMapa).getValue().toUpperCase();
	const prazoEnvioMapa = abaEnvioMapa.getRange(i, colPrazoEnvioMapa).getValue();
	const mensagemVerificacaoMapa = abaEnvioMapa.getRange(i, colMensagemVerificacaoMapa).getValue();

	abaGerencial.getRange(linhaCampoGerencial, colLinkMapaGerencial).setValue(linkMapa);
	abaGerencial.getRange(linhaCampoGerencial, colTextoMapaGerencial).setValue(textoMapa);
	abaGerencial.getRange(linhaCampoGerencial, colComentarioEnviadoMapaGerencial).setValue(comentarioEnviadoMapa);
	abaGerencial.getRange(linhaCampoGerencial, colMensagemVerificacaoMapaGerencial).setValue(mensagemVerificacaoMapa);

	// Se ainda não existir o prazo para envio, coloque o prazo de 7 dias
	if (!prazoEnvioMapa && dataMapa) {
		const dataAtual = dataMapa;
		const dataPrazo = new Date(dataAtual.setDate(dataAtual.getDate() + 7));
		abaGerencial.getRange(linhaCampoGerencial, colPrazoEnvioMapaGerencial).setValue(dataPrazo);
	} else {
		abaGerencial.getRange(linhaCampoGerencial, colPrazoEnvioMapaGerencial).setValue(prazoEnvioMapa);
	}

	// Nenhuma linha nova criada
	return false;
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosMarcoFinal(i, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(i);
		return false;
	}

	const enviouReflexaoMarcoFinal = abaMarcoFinal.getRange(i, colEnviouReflexaoMarcoFinal).getValue().toUpperCase();
	const prazoEnvioMarcoFinal = abaMarcoFinal.getRange(i, colPrazoEnvioMarcoFinal).getValue();
	const comentarioEnviadoMarcoFinal = abaMarcoFinal.getRange(i, colComentarioEnviadoMarcoFinal).getValue().toUpperCase();

	abaGerencial.getRange(linhaCampoGerencial, colRespondeuMarcoFinalGerencial).setValue("SIM");
	abaGerencial.getRange(linhaCampoGerencial, colEnviouReflexaoMarcoFinalGerencial).setValue(enviouReflexaoMarcoFinal);
	abaGerencial.getRange(linhaCampoGerencial, colPrazoEnvioMarcoFinalGerencial).setValue(prazoEnvioMarcoFinal);
	abaGerencial.getRange(linhaCampoGerencial, colComentarioEnviadoMarcoFinalGerencial).setValue(comentarioEnviadoMarcoFinal);

	// Nenhuma linha criada
	return false
}

// Função com a lógica da importação dos campos do envio do mapa
function ImportarDadosCertificado(i, linhaCampoGerencial, linhaVazia) {
	// Se aquele email ainda não estiver registrado na planilha gerencial
	if (!linhaCampoGerencial) {
		LidarComPessoaNaoCadastrada(i);
		return false;
	}

	const dataCertificado = abaCertificado.getRange(i, colDataCertificado).getValue();
	const linkCertificado = abaCertificado.getRange(i, colLinkCertificado).getValue();
	const linkTestadoCertificado = abaCertificado.getRange(i, colLinkTestadoCertificado).getValue().toUpperCase();
	let entrouGrupoCertificado = abaCertificado.getRange(i, colEntrouGrupoCertificado).getValue();

	if (entrouGrupoCertificado != "Enviei email") {
		entrouGrupoCertificado = entrouGrupoCertificado.toUpperCase();
	}

	abaGerencial.getRange(linhaCampoGerencial, colTerminouCursoGerencial).setValue("SIM");
	abaGerencial.getRange(linhaCampoGerencial, colDataCertificadoGerencial).setValue(dataCertificado);
	abaGerencial.getRange(linhaCampoGerencial, colLinkCertificadoGerencial).setValue(linkCertificado);
	abaGerencial.getRange(linhaCampoGerencial, colLinkTestadoCertificadoGerencial).setValue(linkTestadoCertificado);
	abaGerencial.getRange(linhaCampoGerencial, colEntrouGrupoCertificadoGerencial).setValue(entrouGrupoCertificado);

	// Nenhuma linha criada
	return false
}

// Função que irá lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero
function LidarComPessoaNaoCadastrada() {

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