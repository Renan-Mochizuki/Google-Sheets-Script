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
  planilhaAtiva.toast('Sincronizando enviou reflexão entre Marco Final e Gerencial', 'Executando função', tempoNotificacao);
  SincronizarCampoPlanilhas(abaMarcoFinal, colEnviouReflexaoMarcoFinal, abaGerencial, colEnviouReflexaoMarcoFinalGerencial);
  planilhaAtiva.toast('A planilha foi sincronizada', 'Finalização da execução', tempoNotificacao);
}

// Função que recebe as escolhas feitas na interface e esconde as linhas de acordo com elas
function ProcessarEscolhasEsconderLinhas(escolhas) {
  MostrarTodasLinhas();

  // Pega todos os valores necessários de acordo com as escolhas feitas
  const valores = {
    enviouReflexaoMarcoFinal: escolhas.enviouReflexaoMarcoFinal ? abaAtiva.getRange(2, colEnviouReflexaoMarcoFinal, ultimaLinhaAtiva, 1).getValues().flat() : null,
  };
  planilhaAtiva.toast(valores.enviouReflexaoMarcoFinal[1], 'Finalização da execução', tempoNotificacao);
  // Loop que percorre todas as linhas
  for (let i = 0; i < ultimaLinhaAtiva; i++) {
    let esconderLinha = false;

    if (escolhas.enviouReflexaoMarcoFinal && VerificarEsconder(escolhas.enviouReflexaoMarcoFinal, valores.enviouReflexaoMarcoFinal[i])) {
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
