// -- Funções de formatação da planilha --

// Função para limpar toda a planilha
function LimparPlanilha() {
  // Janela de diálogo de confirmação da ação
  const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os campos? \n Os dados modificaveis dessa planilha serão salvos nas planilhas originais', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    // Verifica se há mais de uma linha para limpar
    if (ultimaLinhaAtiva <= 1) return;

    FazerBackupOriginais();

    // Define o intervalo que vai da segunda linha até a última linha e a última coluna com conteúdo
    const planilha = abaAtiva.getRange(2, 1, ultimaLinhaAtiva - 1, ultimaColunaAtiva);

    // Limpa o conteúdo do intervalo selecionado
    planilha.clearContent();
    planilha.setBackground('#ffffff');
    planilha.clearNote();
  }
}

// Função que completa campos vazios adicionais da planilha gerencial com NÃO
function CompletarVaziosComNao() {
  // Loop das colunas
  for (let j = 0; j < colunasDeSimNao.length; j++) {
    const coluna = colunasDeSimNao[j];

    // Loop das linhas
    for (let i = 2; i <= ultimaLinhaGerencial; i++) {
      const celula = abaGerencial.getRange(i, coluna);
      const valor = celula.getValue();
      if (!valor) celula.setValue('NÃO');
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
  } else if (telefoneNumeros.startsWith('55') && telefoneNumeros.length == 13) {
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

  telefones.forEach((linha, i) => {
    if (!linha[0]) return;

    telefones[i][0] = String(linha[0])
      .split('; ') // Divide os telefone separados por ;
      .filter((tel) => tel) // Remove valores vazios
      .map(FormatarTelefone) // Formata cada telefone
      .join('; '); // Une os telefones em uma string separados por ; novamente
  });

  abaDesejada.getRange(2, colTel, ultimaLinha, 1).setValues(telefones);
}

// Função que usa a função FormatarTelefone para formatar todos os campos da planilha ativa
function FormatarLinhasTelefone() {
  FormatarLinhasTelefoneAba(abaAtiva);
}

// Função que usa a função FormatarTelefone para formatar todas planilhas
function FormatarLinhasTelefoneTodasAbas() {
  FormatarLinhasTelefoneAba(abaInteresse);
  FormatarLinhasTelefoneAba(abaMarcoZero);
  FormatarLinhasTelefoneAba(abaEnvioMapa);
  FormatarLinhasTelefoneAba(abaMarcoFinal);
  FormatarLinhasTelefoneAba(abaCertificado);
}

// Função que remove todas linhas vazias no meio da planilha
function RemoverLinhasVazias() {
  // Pega todos os valores da coluna desejada
  const valColunas = abaAtiva.getRange(2, colEmailAtiva, ultimaLinhaAtiva, 1).getValues().flat();

  // Loop que percorre as linhas de baixo para cima
  for (let i = valColunas.length - 1; i >= 0; i--) {
    if (!valColunas[i]) {
      abaAtiva.deleteRow(i + 2); // Deleta a linha correspondente (i + 2 porque o índice começa em 0 e a planilha começa na linha 2)
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
  const html = HtmlService.createHtmlOutputFromFile('InterfaceCheckboxes').setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Escolha quem visualizar');
}

// Função que recebe as escolhas feitas na interface e esconde as linhas de acordo com elas
function ProcessarEscolhasEsconderLinhas(escolhas) {
  MostrarTodasLinhas();

  // Pega todos os valores necessários de acordo com as escolhas feitas
  const valores = {
    situacao: escolhas.situacao ? abaGerencial.getRange(2, colSituacaoGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    whats: escolhas.whats ? abaGerencial.getRange(2, colWhatsGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    terminouCurso: escolhas.terminouCurso ? abaGerencial.getRange(2, colTerminouCursoGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    linkTestadoCertificado: escolhas.linkTestadoCertificado ? abaGerencial.getRange(2, colLinkTestadoCertificadoGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    linkCertificado: escolhas.linkTestadoCertificado ? abaGerencial.getRange(2, colLinkCertificadoGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    comentarioEnviadoMapa: escolhas.comentarioEnviadoMapa ? abaGerencial.getRange(2, colComentarioEnviadoMapaGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    linkMapa: escolhas.comentarioEnviadoMapa ? abaGerencial.getRange(2, colLinkMapaGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    comentarioEnviadoMarcoFinal: escolhas.comentarioEnviadoMarcoFinal ? abaGerencial.getRange(2, colComentarioEnviadoMarcoFinalGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
    respondeuMarcoFinal: escolhas.comentarioEnviadoMarcoFinal ? abaGerencial.getRange(2, colRespondeuMarcoFinalGerencial, ultimaLinhaGerencial, 1).getValues().flat() : null,
  };

  // Loop que percorre todas as linhas
  for (let i = 0; i < ultimaLinhaGerencial; i++) {
    let esconderLinha = false;

    // Verifica cada condição
    if (escolhas.situacao && VerificarEsconderSituacao(escolhas.situacao, valores.situacao[i])) {
      esconderLinha = true;
    }
    if (escolhas.whats && VerificarEsconder(escolhas.whats, valores.whats[i])) {
      esconderLinha = true;
    }
    if (escolhas.terminouCurso && VerificarEsconder(escolhas.terminouCurso, valores.terminouCurso[i])) {
      esconderLinha = true;
    }
    if (escolhas.linkTestadoCertificado && (valores.linkCertificado[i] ? VerificarEsconder(escolhas.linkTestadoCertificado, valores.linkTestadoCertificado[i]) : true)) {
      esconderLinha = true;
    }
    if (escolhas.comentarioEnviadoMapa && (valores.linkMapa[i] ? VerificarEsconder(escolhas.comentarioEnviadoMapa, valores.comentarioEnviadoMapa[i])  : true)) {
      esconderLinha = true;
    }
    if (escolhas.comentarioEnviadoMarcoFinal && (valores.respondeuMarcoFinal[i] ? VerificarEsconder(escolhas.comentarioEnviadoMarcoFinal, valores.comentarioEnviadoMarcoFinal[i]) : true)) {
      esconderLinha = true;
    }

    if (esconderLinha) {
      abaGerencial.hideRows(i + 2);
    }
  }
}

// Função que verifica se a linha deve ser escondida ou não
function VerificarEsconder(escolha, valor) {
  if (!escolha) return false;
  return !((escolha === 'SIM' && valor === 'SIM') || (escolha === 'NÃO' && valor !== 'SIM'));
}

// Função que verifica se a linha da situação deve ser escondida ou não
function VerificarEsconderSituacao(escolha, valor) {
  if (!escolha) return false;
  return !((escolha === 'VAZIO' && !valor) || (escolha === valor));
}

// Função que revela todas as linhas escondidas
function MostrarTodasLinhas() {
  abaAtiva.showRows(2, ultimaLinhaAtiva);
}
