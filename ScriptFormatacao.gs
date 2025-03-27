// -- Funções de formatação da planilha --

// Função para limpar toda a planilha fazendo backup
function LimparPlanilha() {
  // Janela de diálogo de confirmação da ação
  const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os campos? \n Os dados modificáveis dessa planilha serão salvos nas planilhas originais', ui.ButtonSet.YES_NO);

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

// Função que apaga toda a planilha sem realizar backup
function ApagarTodosDados() {
  // Janela de diálogo de confirmação da ação
  const response = ui.alert('Confirmação', 'Você tem certeza que deseja excluir todos os dados? \n Todos os dados dessa planilha serão apagados e não serão salvos', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    // Verifica se há mais de uma linha para limpar
    if (ultimaLinhaAtiva <= 1) return;

    // Define o intervalo que vai da segunda linha até a última linha e a última coluna com conteúdo
    const planilha = abaAtiva.getRange(2, 1, ultimaLinhaAtiva - 1, ultimaColunaAtiva);

    // Limpa o conteúdo do intervalo selecionado
    planilha.clearContent();
    planilha.setBackground('#ffffff');
    planilha.clearNote();
  }
}

// Função que completa campos vazios adicionais da planilha com NÃO
function CompletarVaziosComNao() {
  // Loop das colunas
  for (let j = 0; j < colunasDeSimNao.length; j++) {
    const coluna = colunasDeSimNao[j];
    const valColuna = abaAtiva.getRange(2, coluna, ultimaLinhaAtiva, 1).getValues();

    // Loop das linhas
    for (let i = 0; i < valColuna.length; i++) {
      const valor = valColuna[i][0];
      if (!valor) valColuna[i][0] = 'NÃO';
    }

    abaAtiva.getRange(2, coluna, ultimaLinhaAtiva, 1).setValues(valColuna);
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
  const { colCidade } = objetoMap.get(abaAtiva);
  const cidadesEstados = abaAtiva.getRange(2, colCidade, ultimaLinhaAtiva, 2).getValues();

  // Loop das linhas
  for (let i = 0; i < cidadesEstados.length; i++) {
    const cidade = cidadesEstados[i][0];
    const estado = cidadesEstados[i][1];

    // Se a cidade estiver vazia, ou o estado já tiver preenchido
    if (!cidade || estado || typeof cidade !== 'string') continue;

    estados.forEach((estadoCompleto) => {
      const partes = estadoCompleto.split(' - ');
      for (let cidadeSeparada of cidade.split(';')) {
        if (cidadeSeparada.trim() === partes[0] || ContemUF(cidadeSeparada, partes[1])) {
          cidadesEstados[i][1] = estadoCompleto;
        }
      }
    });
  }

  abaAtiva.getRange(2, colCidade, ultimaLinhaAtiva, 2).setValues(cidadesEstados);
}

// Função que verifica se uma string contém uma UF
function ContemUF(str, uf) {
  if (typeof str !== 'string' || typeof uf !== 'string') return false;

  // Regex para verificar a UF isolada, permitindo espaços, barras, traços, etc.
  const regex = new RegExp(`(^|[\\s\\/\\-])${uf}($|[\\s\\/\\-])`, 'i');

  return regex.test(str);
}

// Função que exibe o HTML da interface para escolher quem quer esconder
function MostrarInterfaceEsconderLinhas() {
  const html = HtmlService.createHtmlOutputFromFile('Interface').setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Escolha quem visualizar');
}

// Função ProcessarEscolhasEsconderLinhas no script de cada arquivo

// Função que verifica se a linha deve ser escondida ou não
function VerificarEsconder(escolha, valor) {
  if (!escolha) return false;
  return !((escolha === 'SIM' && valor === 'SIM') || (escolha === 'NÃO' && valor !== 'SIM'));
}

// Função que verifica se a linha da situação deve ser escondida ou não
function VerificarEsconderSituacao(escolha, valor) {
  if (!escolha) return false;
  return !((escolha === 'VAZIO' && !valor) || escolha === valor);
}

// Função que revela todas as linhas escondidas
function MostrarTodasLinhas() {
  abaAtiva.showRows(2, ultimaLinhaAtiva);
}
