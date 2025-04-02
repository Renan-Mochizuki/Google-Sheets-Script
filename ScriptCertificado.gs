const ui = SpreadsheetApp.getUi();
// Função para adicionar o menu
function onOpen(e) {
  ui.createMenu('Menu de Funções')
    .addItem('🔄 Sincronizar planilha', 'SincronizarPlanilha')
    .addItem('👁‍🗨 Mostrar todas linhas', 'MostrarTodasLinhas')
    .addItem('🔎 Filtrar visualização', 'MostrarInterfaceEsconderLinhas')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Formatação da planilha')
        .addItem('Formatar todos telefone', 'FormatarLinhasTelefone')
        .addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
        .addItem('Remover linhas vazias', 'RemoverLinhasVazias')
    )
    .addToUi();
}

// -- IMPORTANTE --
// VEJA OS COMENTÁRIOS DO ARQUIVO CONSTANTS

// Função que irá sincronizar todos os campos adicionais da planilha
function SincronizarPlanilha() {
  planilhaAtiva.toast('Sincronizando link testado entre Certificado e Gerencial', 'Executando função', tempoNotificacao);
  SincronizarCampoPlanilhas(abaCertificado, colLinkTestadoCertificado, abaGerencial, colLinkTestadoCertificadoGerencial);
  planilhaAtiva.toast('Sincronizando entrou no grupo entre Certificado e Gerencial', 'Executando função', tempoNotificacao);
  SincronizarCampoPlanilhas(abaCertificado, colEntrouGrupoCertificado, abaGerencial, colEntrouGrupoCertificadoGerencial);
  planilhaAtiva.toast('A planilha foi sincronizada', 'Finalização da execução', tempoNotificacao);
}

// Função que recebe as escolhas feitas na interface e esconde as linhas de acordo com elas
function ProcessarEscolhasEsconderLinhas(escolhas) {
  MostrarTodasLinhas();

  // Pega todos os valores necessários de acordo com as escolhas feitas
  const valores = {
    linkTestadoCertificado: escolhas.linkTestadoCertificado ? abaAtiva.getRange(2, colLinkTestadoCertificado, ultimaLinhaAtiva, 1).getValues().flat() : null,
    entrouGrupoCertificado: escolhas.entrouGrupoCertificado ? abaAtiva.getRange(2, colEntrouGrupoCertificado, ultimaLinhaAtiva, 1).getValues().flat() : null,
  };

  // Loop que percorre todas as linhas
  for (let i = 0; i < ultimaLinhaAtiva; i++) {
    let esconderLinha = false;

    if (escolhas.linkTestadoCertificado && VerificarEsconder(escolhas.linkTestadoCertificado, valores.linkTestadoCertificado[i])) {
      esconderLinha = true;
    }
    if (escolhas.entrouGrupoCertificado && VerificarEsconder(escolhas.entrouGrupoCertificado, valores.entrouGrupoCertificado[i])) {
      esconderLinha = true;
    }

    if (esconderLinha) {
      abaAtiva.hideRows(i + 2);
    }
  }
}

// Apenas declarando as funções para evitar erros de constants
function ImportarDadosInteresse() {}
function ImportarDadosMarcoZero() {}
function ImportarDadosEnvioMapa() {}
function ImportarDadosMarcoFinal() {}
function ImportarDadosCertificado() {}
