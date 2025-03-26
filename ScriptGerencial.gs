const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
  ui.createMenu('Menu de Funções')
    .addItem('📂 Importar Dados', 'Importar')
    .addItem('📞 Sincronizar campos do Whatsapp', 'SincronizarWhatsGerencial')
    .addItem('🗑️ Limpar Planilha', 'LimparPlanilha')
    .addItem('👁‍🗨 Mostrar todas linhas', 'MostrarTodasLinhas')
    .addItem('🔎 Filtrar visualização', 'MostrarInterfaceEsconderLinhas')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Formatação da planilha')
        .addItem('Formatar todos telefone', 'FormatarLinhasTelefone')
        .addItem('Preencher campos estado', 'PreencherEstado')
        .addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
        .addItem('Apagar todos os dados', 'ApagarTodosDados')
        .addItem('Remover linhas vazias', 'RemoverLinhasVazias')
    )
    .addToUi();
}

// -- IMPORTANTE --
// VEJA OS COMENTÁRIOS DO ARQUIVO CONSTANTS

// Função que separa os dados pelo ;
function SepararDados(dadosMultiplos) {
  if (dadosMultiplos) {
    return NormalizarString(dadosMultiplos).split(';');
  } else {
    return [dadosMultiplos];
  }
}

// Função que verifica se os campos passados do loop atual são válidos args = (nome, email, telefone, outros...)
function ValidarLoop(...args) {
  // Se o valor existe, e caso for string, não contém a palavra 'teste'
  const nomeValido = args[0] && (typeof args[0] === 'string' ? !args[0].toLowerCase().includes('teste') : true);
  const emailValido = args[1] && (typeof args[1] === 'string' ? !args[1].toLowerCase().includes('teste') : true);
  const telefoneValido = args[2] && (typeof args[2] === 'string' ? !args[2].toLowerCase().includes('teste') : true);

  // Retorna falso se qualquer um dos campos for inválido
  if (!nomeValido || !emailValido || !telefoneValido) return false;

  // Se outro parametro foi passado, verifique se qualquer um for nulo, retorne falso
  for (let i = 3; i < args.length; i++) {
    if (!args[i]) return false;
  }

  return true;
}

// Função que verificará se o email existe na planilha desejada e retornará a linha
function RetornarLinhaDados(nomeProcurado, emailProcurado, telefoneProcurado, dados) {
  // Separando os dados procurados pois ele pode ser um valor com mais de um email
  const nomesProcurados = SepararDados(nomeProcurado);
  const emailsProcurados = SepararDados(emailProcurado);
  const telefonesProcurados = SepararDados(telefoneProcurado);

  // dados é uma matriz, na qual possui as colunas nome, email, telefone e cada linha é um cadastro

  // Conferir cada linha da matriz dos dados
  for (let i = 0; i < dados.length; i++) {
    const nomeDados = dados[i][0];
    const emailDados = dados[i][1];
    const telefoneDados = dados[i][2];

    const similaridadeNome = VerificarLinhaDados(nomeDados, nomesProcurados);
    const similaridadeEmail = VerificarLinhaDados(emailDados, emailsProcurados);
    const similaridadeTelefone = VerificarLinhaDados(telefoneDados, telefonesProcurados);

    // Se (email e telefone forem iguais) ou (email e nome forem iguais, tendo que o telefone não é caso especial) ou (telefone e nome forem iguais)
    if (
      (similaridadeEmail >= 0.8 && similaridadeTelefone >= 0.8) ||
      (similaridadeEmail >= 0.8 && similaridadeNome >= 0.5 && similaridadeTelefone !== -1) ||
      (similaridadeTelefone >= 0.9 && similaridadeNome >= 0.6)
    ) {
      return i + 2; // Retorne o índice da array + 2 (Porque a array começa em 0 e a planilha em 2)
    }
  }
  // Se não for encontrado nenhum
  return false;
}

// Função genérica para verificar se os dados são iguais (com uma certa tolerância)
function VerificarLinhaDados(dados, valoresProcurados) {
  if (!dados) return false;

  for (let dadoPlanilha of dados.toString().split(';')) {
    // Se o dado passado for um email, retire o domínio (Exemplo: @gmail.com)
    if (dadoPlanilha.includes('@')) {
      dadoPlanilha = dadoPlanilha.split('@')[0];
    }

    for (let valorProcurado of valoresProcurados) {
      // Se o valor procurado for um email, retire o domínio (Exemplo: @gmail.com)
      if (valorProcurado.includes('@')) {
        valorProcurado = valorProcurado.split('@')[0];
      }
      const similaridade = CompararSimilaridade(valorProcurado, dadoPlanilha);

      // Caso especifico com telefone (pois foi achado um dado que falhava na verificação comum)
      // O telefone procurado é diferente e o um telefone apenas possui o +
      if (similaridade < 0.8 && valorProcurado.includes('+') != dadoPlanilha.includes('+')) {
        return -1;
      }

      // Se o valor procurado e o dado bruto forem iguais, retorne true
      return similaridade;
    }
  }
  return false;
}

// Função que executa outras funções para importar os dados de cada planilha
function Importar() {
  const tituloToast = 'Executando funções';
  let totalLinhasAnalisadas = 0;

  // Formatando os telefones de todas as planilhas
  planilhaAtiva.toast('Formatando telefones de todas planilhas', tituloToast, tempoNotificacao);
  FormatarLinhasTelefoneTodasAbas();

  planilhaAtiva.toast(tituloToast, 'Importando dados da Interesse', tempoNotificacao);
  totalLinhasAnalisadas += ImportarDados(abaInteresse);

  planilhaAtiva.toast(tituloToast, 'Importando notas da Interesse', tempoNotificacao);
  ImportarNotas(abaInteresse);

  const ultimaLinhaDepoisDaInteresse = abaGerencial.getLastRow();

  planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Zero', tempoNotificacao);
  totalLinhasAnalisadas += ImportarDados(abaMarcoZero);

  const ultimaLinhaDepoisDoMarcoZero = abaGerencial.getLastRow();
  const intervaloInicioPintar = ultimaLinhaDepoisDaInteresse + 1;
  const intervaloFimPintar = ultimaLinhaDepoisDoMarcoZero - intervaloInicioPintar + 1;

  // Pintando campos cidade e estado, situação e redirecionamento para interesse das pessoas de S. PÚBLICA (esses campos nunca terão valor)
  abaGerencial.getRange(intervaloInicioPintar, colCidadeGerencial, intervaloFimPintar, 2).setBackground(corCampoSemDados);
  abaGerencial.getRange(intervaloInicioPintar, colSituacaoGerencial, intervaloFimPintar, 1).setBackground(corCampoSemDados);
  abaGerencial.getRange(intervaloInicioPintar, colRedirectInteresseGerencial, intervaloFimPintar, 1).setBackground(corCampoSemDados);

  planilhaAtiva.toast(tituloToast, 'Importando dados do Envio de Mapa', tempoNotificacao);
  totalLinhasAnalisadas += ImportarDados(abaEnvioMapa);

  planilhaAtiva.toast(tituloToast, 'Importando dados do Marco Final', tempoNotificacao);
  totalLinhasAnalisadas += ImportarDados(abaMarcoFinal);

  planilhaAtiva.toast(tituloToast, 'Importando dados do Envio do Certificado', tempoNotificacao);
  totalLinhasAnalisadas += ImportarDados(abaCertificado);

  const quantidadeLinhasCriadas = abaGerencial.getLastRow() - ultimaLinhaGerencial;
  const mensagem = 'Fim da execução.\n' + quantidadeLinhasCriadas + ' linhas criadas e ' + totalLinhasAnalisadas + ' linhas já registradas analisadas';
  planilhaAtiva.toast(mensagem, 'Execução finalizada', tempoNotificacao);
}

// Função genérica de importação para todas planilhas
function ImportarDados(abaDesejada) {
  // Pegando a próxima linha vazia da planilha gerencial
  // Obs.: Não se pode usar a variável ultimaLinhaGerencial, pois ela não se atualiza sozinha
  const ultimaLinhaPlanilhaGerencial = abaGerencial.getLastRow();
  let linhaVazia = ultimaLinhaPlanilhaGerencial + 1;
  let linhasAfetadas = 0;

  // Atribui as variáveis de acordo com a abaDesejada
  const { nomePlanilha, ultimaLinhaAnalisada, ultimaLinha, ultimaColuna, colNome, colEmail, colTel, ImportarDadosPlanilha } = objetoMap.get(abaDesejada);

  // Armazenando todos os nomes, emails e telefones da abaGerencial em uma matriz
  const nomesEmailsTelefones = abaGerencial.getRange(2, colNomeGerencial, ultimaLinhaPlanilhaGerencial ?? 1, 3).getValues();

  // Loop para percorrer todas as linhas da planilha desejada
  for (let i = ultimaLinhaAnalisada; i <= ultimaLinha; i++) {
    // Armazendo a linha inteira da planilha desejada em uma array
    // Definimos o primeiro item como null para facilitar o acesso aos índices (sem precisar ficar subtraindo 1)
    const valLinha = abaDesejada.getRange(i, 1, 1, ultimaColuna).getValues()[0];
    valLinha.unshift(null);

    const nome = valLinha[colNome];
    const email = valLinha[colEmail];
    const telefone = valLinha[colTel];

    if (!ValidarLoop(nome, email, telefone)) continue;

    // Toast da mensagem do progresso de execução da função
    if (i % 100 === 0) planilhaAtiva.toast('Processo na linha ' + i + ' da planilha ' + nomePlanilha, Math.round((i / ultimaLinha) * 100) + '% concluído da função atual', tempoNotificacao);

    // Pegando a linha do campo na planilha gerencial (Se existir)
    const linhaCampoGerencial = RetornarLinhaDados(nome, email, telefone, nomesEmailsTelefones);
    const foiCastradoNaGerencial = ImportarDadosPlanilha(valLinha, i, linhaCampoGerencial, linhaVazia);

    // Se o registro atual da planilha desejada for cadastrada na gerencial, atualizaremos a matriz de dados da gerencial
    // (para evitar chamadas de abaGerencial.getRange)
    if (foiCastradoNaGerencial) {
      const novoRegistro = [nome, email, telefone];
      linhaVazia++;
      // Insira o novo email e tel na matriz de dados (Se o primeiro item estiver vazio, substitua o item vazio)
      if (!nomesEmailsTelefones[0][1]) {
        nomesEmailsTelefones[0] = novoRegistro;
        continue;
      }

      nomesEmailsTelefones.push(novoRegistro);
    } else linhasAfetadas++;
  }

  return linhasAfetadas;
}

// Função com a lógica da importação dos campos da planilha de interesse
function ImportarDadosInteresse(valLinha, linhaAtual, linhaCampoGerencial, linhaVazia) {
  // Declarando uma array com os campos adicionais da planilha Interesse
  // *Considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = [
    valLinha[colAnotacaoInteresse],
    null,
    valLinha[colNomeInteresse],
    valLinha[colEmailInteresse],
    valLinha[colTelInteresse],
    valLinha[colCidadeInteresse],
    valLinha[colEstadoInteresse],
    valLinha[colWhatsInteresse],
    'SIM',
    valLinha[colRespondeuMarcoZeroInteresse],
    valLinha[colSituacaoInteresse],
  ];

  // Se o registro ainda não estiver cadastrado na planilha gerencial
  if (!linhaCampoGerencial) {
    // Inserindo os campos na planilha gerencial
    abaGerencial.getRange(linhaVazia, colAnotacaoGerencial, 1, 11).setValues([intervaloInserir]);
    InserirRedirecionamentoPlanilha(linhaVazia, colRedirectInteresseGerencial, urlInteresse, linhaAtual);

    // Nova linha criada
    return true;
  }
  // Registro já cadastrado na planilha gerencial
  else {
    // Pegando os valores daquela linha da planilha gerencial, pois alguem pode responder mais de uma vez
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

  // Declarando o intervalo considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = [valLinha[colNomeMarcoZero], valLinha[colEmailMarcoZero], valLinha[colTelMarcoZero], null, null, valLinha[colWhatsMarcoZero], respondeuInteresseMarcoZero, 'SIM'];

  // Se o registro ainda não estiver cadastrado na planilha gerencial
  if (!linhaCampoGerencial) {
    // Inserindo os campos na planilha gerencial
    abaGerencial.getRange(linhaVazia, colNomeGerencial, 1, 8).setValues([intervaloInserir]);
    InserirRedirecionamentoPlanilha(linhaVazia, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

    // Nova linha criada
    return true;
  }
  // Registro já cadastrado na planilha gerencial
  else {
    // Pegando os valores daquela linha da planilha gerencial, pois alguem pode responder mais de uma vez
    const valLinhaGerencial = abaGerencial.getRange(linhaCampoGerencial, colNomeGerencial, 1, 8).getValues()[0];

    // Juntando os dados já existentes da planilha gerencial com os novos dados
    const intervaloUnido = JuntarDados(valLinhaGerencial, intervaloInserir, colNomeGerencial);

    // Inserindo os campos na planilha gerencial
    abaGerencial.getRange(linhaCampoGerencial, colNomeGerencial, 1, 8).setValues([intervaloUnido]);
    InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectMarcoZeroGerencial, urlMarcoZero, linhaAtual);

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
  const comentarioEnviadoMapa = (valLinha[colComentarioEnviadoMapa] || '').toUpperCase();

  // Caso ainda não existir prazo, calcular um novo adicionando 7 dias
  const dataPrazo = !prazoEnvioMapa && dataMapa ? new Date(dataMapa.setDate(dataMapa.getDate() + 7)) : prazoEnvioMapa;

  // Declarando o intervalo considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = [valLinha[colLinkMapa], valLinha[colTextoMapa], dataPrazo, comentarioEnviadoMapa, valLinha[colMensagemVerificacaoMapa]];

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
  const comentarioEnviadoMarcoFinal = (valLinha[colComentarioEnviadoMarcoFinal] || '').toUpperCase();

  // Declarando o intervalo considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = ['SIM', enviouReflexaoMarcoFinal, valLinha[colPrazoEnvioMarcoFinal], comentarioEnviadoMarcoFinal];

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
  const entrouGrupoCertificado = valEntrouGrupo && valEntrouGrupo != 'Enviei email' ? valEntrouGrupo.toUpperCase() : valEntrouGrupo;

  // Declarando o intervalo considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = [valLinha[colDataCertificado], valLinha[colLinkCertificado], linkTestadoCertificado, entrouGrupoCertificado];

  abaGerencial.getRange(linhaCampoGerencial, colTerminouCursoGerencial).setValue('SIM');
  abaGerencial.getRange(linhaCampoGerencial, colDataCertificadoGerencial, 1, 4).setValues([intervaloInserir]);
  InserirRedirecionamentoPlanilha(linhaCampoGerencial, colRedirectCertificadoGerencial, urlCertificado, linhaAtual);

  // Nenhuma linha criada
  return false;
}

// Função que irá lidar com pessoas que estão em formulários posteriores sem estar na de interesse ou marco zero
function LidarComPessoaNaoCadastrada(valLinha, linhaAtual, linhaVazia, abaDesejada) {
  // Atribui as variáveis de acordo com a abaDesejada
  const { colNome, colEmail, colTel, ImportarDadosPlanilha } = objetoMap.get(abaDesejada);

  // Declarando o intervalo considerando a ordem dos campos da planilha Gerencial (Ver ORDEM OBRIGATÓRIA DOS CAMPOS)
  const intervaloInserir = ['Não cadastrada nas outras planilhas', null, valLinha[colNome], valLinha[colEmail], valLinha[colTel]];

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

// Função que importa as anotações de uma planilha na coluna do email
function ImportarNotas(abaDesejada, colDesejada) {
  // Atribui as variáveis de acordo com a abaDesejada
  const { ultimaLinha, colEmail, colNome, ultimaColuna } = objetoMap.get(abaDesejada);
  let ultimaLinhaPlanilhaGerencial = abaGerencial.getLastRow();

  // Pegando dados
  const nomesEmailsTelefonesGerencial = abaGerencial.getRange(2, colNomeGerencial, ultimaLinhaPlanilhaGerencial, 3).getValues();
  const nomesEmailsTelefonesAbaDesejada = abaDesejada.getRange(2, colNome, ultimaLinha, 3).getValues();

  // Pegando todas as anotações da planilha gerencial
  const anotacoesGerencial = abaGerencial.getRange(2, colAnotacaoGerencial, ultimaLinhaPlanilhaGerencial, 1).getValues().flat();

  // Pegando as notas da planilha desejada
  const notasColunasAbaDesejada = abaDesejada
    .getRange(2, colDesejada ?? colEmail, ultimaLinha, 1)
    .getNotes()
    .flat();

  // Loop para percorrer todas as notas da planilha desejada
  for (let i = 0; i < notasColunasAbaDesejada.length; i++) {
    const notaDesejada = notasColunasAbaDesejada[i];
    const nome = nomesEmailsTelefonesAbaDesejada[i][0];
    const email = nomesEmailsTelefonesAbaDesejada[i][1];
    const telefone = nomesEmailsTelefonesAbaDesejada[i][2];

    if (!ValidarLoop(nome, email, telefone, notaDesejada)) continue;

    const linhaCampoGerencial = RetornarLinhaDados(nome, email, telefone, nomesEmailsTelefonesGerencial);

    // Se aquele email não for encontrado na planilha gerencial
    if (!linhaCampoGerencial) {
      const valLinha = abaDesejada.getRange(i + 2, 1, 1, ultimaColuna).getValues()[0];
      valLinha.unshift(null);
      LidarComPessoaNaoCadastrada(valLinha, i + 2, ultimaLinhaPlanilhaGerencial + 1, abaDesejada);
      ultimaLinhaPlanilhaGerencial++;
      continue;
    }

    // Pegando a anotação daquele registro na gerencial
    const anotacaoGerencial = anotacoesGerencial[linhaCampoGerencial - 2];
    let notaInserir;

    // Se já existir uma anotação na gerencial
    if (anotacaoGerencial) {
      // Se a anotação da gerencial não conter a nota desejada
      if (!anotacaoGerencial.split(';').includes(notaDesejada)) {
        notaInserir = anotacaoGerencial + '; ' + notaDesejada;
      }
      // Se já conter a nota desejada, não altere nada
      else {
        notaInserir = anotacaoGerencial;
      }
    } else {
      notaInserir = notaDesejada;
    }

    // Inserindo notaInserir nas anotações da planilha gerencial
    anotacoesGerencial[linhaCampoGerencial - 2] = notaInserir;

    // Procurando emails nas anotações da nota desejada
    const emailGerencial = nomesEmailsTelefonesGerencial[linhaCampoGerencial - 2][1];
    const regex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    const emailsDaNota = notaDesejada.match(regex) || [];

    // Loop que percorre todos os emails presentes na nota desejada
    for (let emailNota of emailsDaNota) {
      // Se o email da gerencial ainda não conter esse email da nota desejada, adicione ele
      if (!emailGerencial.includes(emailNota)) {
        nomesEmailsTelefonesGerencial[linhaCampoGerencial - 2][1] = emailGerencial + '; ' + emailNota;
      }
    }
  }

  // Atualizando todos os dados de anotações e email na planilha gerencial
  abaGerencial.getRange(2, colAnotacaoGerencial, ultimaLinhaPlanilhaGerencial, 1).setValues(anotacoesGerencial.map((nota) => [nota])); // Revertendo o .flat()
  abaGerencial.getRange(2, colNomeGerencial, ultimaLinhaPlanilhaGerencial, 3).setValues(nomesEmailsTelefonesGerencial);
}

// Função que junta os dados de duas linhas em um só concatenando dados
function JuntarDados(dadosLinha1, dadosLinha2, primeiraColunaDoIntervalo) {
  const primeiraColuna = primeiraColunaDoIntervalo ?? colNomeGerencial;

  let dadosConcatenados = [];

  for (let i = 0; i < dadosLinha1.length; i++) {
    const dado1 = dadosLinha1[i];
    const dado2 = dadosLinha2[i];
    const colunaAtual = primeiraColuna + i;

    let possuiSimilaridade = false;

    // Exceções especiais
    if (colunasDeSimNao.includes(colunaAtual)) {
      dadosConcatenados.push(RetornarValorSimNao(dado1, dado2));
      continue;
    }
    if (colunaAtual === colSituacaoGerencial) {
      dadosConcatenados.push(RetornarTurmaMaisRecente(dado1, dado2));
      continue;
    }

    // Se o dado1 não existir, adicione o dado2
    if (!dado1) {
      dadosConcatenados.push(dado2 || '');
      continue;
    }
    if (dado2) {
      // Separe o texto pelo ; para caso o campo já tiver sido concatenado
      const textosSeparados1 = dado1.toString().split(';');

      // Loop para comparar a similaridade para cada um dos textos
      for (let texto of textosSeparados1) {
        if (CompararSimilaridade(texto, dado2) >= 0.9) {
          possuiSimilaridade = true;
          break;
        }
      }

      // Caso não o texto do dado1 não possua similaridade com o dado2, adicione o dado2
      if (!possuiSimilaridade) {
        // Caso especial para o estado
        if (colunaAtual === colEstadoGerencial) {
          dadosConcatenados.push(dado2.toString().trim());
          continue;
        }

        dadosConcatenados.push(dado1.toString().trim() + '; ' + dado2.toString().trim());
        continue;
      }
    }
    dadosConcatenados.push(dado1);
  }

  return dadosConcatenados;
}

// Função que retorna a turma mais recente (Ex: T04-2024 > T01-2024 > T02-2023)
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
  if (match1[2] > match2[2]) return string1;
  if (match1[2] < match2[2]) return string2;

  // Anos iguais, verificar qual turma é maior
  if (match1[1] > match2[1]) return string1;
  if (match1[1] < match2[1]) return string2;

  return string1;
}

// Função que compara dois valores de SIM ou NÃO e retorna o valor correto (priorizando o SIM)
function RetornarValorSimNao(valor1, valor2) {
  if (!valor1) return valor2;
  if (!valor2) return valor1;
  if (valor1 == 'SIM' || valor2 == 'SIM') return 'SIM';
  return valor1;
}

// Função que extrai a linha de uma url de redirect
function ExtrairLinhaRedirect(url) {
  const match = url.match(/(\d+)$/);
  return match ? match[1] : null;
}

// Função que formata uma array de strings, deixando apenas a primeira letra em caixa alta
function FormatarCaixaBaixa(array) {
  if (!Array.isArray(array)) return [];
  return array.map((str) => {
    if (typeof str !== 'string' || !str.trim()) return '';
    return str.trim().charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  });
}

// Função que atualiza os dados das planilhas originais para salvar as alterações
function FazerBackupOriginais() {
  for (let i = 2; i < ultimaLinhaGerencial; i++) {
    // Armazendo a linha inteira da planilha gerencial em uma array
    // Definimos o primeiro item como null para facilitar o acesso aos índices (sem precisar ficar subtraindo 1)
    const valLinha = abaGerencial.getRange(i, 1, 1, ultimaColunaGerencial).getValues()[0];
    valLinha.unshift(null);

    if (!ValidarLoop(valLinha[colNome], valLinha[colEmail], valLinha[colTel])) continue;

    if (i % 100 === 0) planilhaAtiva.toast('Processo na linha ' + i + ' da planilha gerencial', Math.round((i / ultimaLinhaGerencial) * 100) + '% concluído da função atual', tempoNotificacao);

    const numLinhaInteresse = ExtrairLinhaRedirect(valLinha[colRedirectInteresseGerencial]);
    const numLinhaMarcoZero = ExtrairLinhaRedirect(valLinha[colRedirectMarcoZeroGerencial]);
    const numLinhaEnvioMapa = ExtrairLinhaRedirect(valLinha[colRedirectEnvioMapaGerencial]);
    const numLinhaMarcoFinal = ExtrairLinhaRedirect(valLinha[colRedirectMarcoFinalGerencial]);
    const numLinhaCertificado = ExtrairLinhaRedirect(valLinha[colRedirectCertificadoGerencial]);

    if (numLinhaInteresse) {
      const intervaloInserir = [valLinha[colWhatsGerencial], valLinha[colRespondeuMarcoZeroGerencial], valLinha[colSituacaoGerencial]];
      if(intervaloInserir.every(item => !item)) continue;
      abaInteresse.getRange(numLinhaInteresse, colWhatsInteresse, 1, 3).setValues([intervaloInserir]);
    }

    if (numLinhaMarcoZero) {
      const intervaloInserir = [valLinha[colRespondeuInteresseGerencial], valLinha[colWhatsGerencial]];
      if(intervaloInserir.every(item => !item)) continue;
      abaMarcoZero.getRange(numLinhaMarcoZero, colRespondeuInteresseMarcoZero, 1, 2).setValues([intervaloInserir]);
    }

    if (numLinhaEnvioMapa) {
      const intervaloInserir = [valLinha[colComentarioEnviadoMapaGerencial], valLinha[colPrazoEnvioMapaGerencial], valLinha[colMensagemVerificacaoMapaGerencial], valLinha[colTerminouCursoGerencial]];
      if(intervaloInserir.every(item => !item)) continue;
      abaEnvioMapa.getRange(numLinhaEnvioMapa, colComentarioEnviadoMapa, 1, 4).setValues([intervaloInserir]);
    }

    if (numLinhaMarcoFinal) {
      const intervaloInserir = FormatarCaixaBaixa([valLinha[colEnviouReflexaoMarcoFinalGerencial], valLinha[colPrazoEnvioMarcoFinalGerencial], valLinha[colComentarioEnviadoMarcoFinalGerencial]]);
      if(intervaloInserir.every(item => !item)) continue;
      abaEnvioMapa.getRange(numLinhaMarcoFinal, colEnviouReflexaoMarcoFinal, 1, 3).setValues([intervaloInserir]);
    }

    if (numLinhaCertificado) {
      const intervaloInserir = FormatarCaixaBaixa([valLinha[colLinkTestadoCertificadoGerencial], valLinha[colEntrouGrupoCertificadoGerencial]]);
      if(intervaloInserir.every(item => !item)) continue;
      abaCertificado.getRange(numLinhaCertificado, colLinkTestadoCertificado, 1, 2).setValues([intervaloInserir]);
    }
  }
  planilhaAtiva.toast('Fim da execução', 'Backup concluído', tempoNotificacao);
}
