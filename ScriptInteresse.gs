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
        .addItem('Preencher campos estado', 'PreencherEstado')
        .addItem('Completar campos vazios com NÃO', 'CompletarVaziosComNao')
        .addItem('Remover linhas vazias', 'RemoverLinhasVazias')
    )
    .addToUi();
}

// -- IMPORTANTE --
// VEJA OS COMENTÁRIOS DO ARQUIVO CONSTANTS

// Função que irá sincronizar todos os campos adicionais da planilha
function SincronizarPlanilha() {
  // Sincronize as planilhas Interesse e Marco Zero, depois as planilhas Interesse e Gerencial e por fim, Marco Zero e Gerencial
  planilhaAtiva.toast('Sincronizando Whats entre Interesse, Marco Zero e Gerencial', 'Executando função', tempoNotificacao);
  SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaMarcoZero, colWhatsMarcoZero);
  planilhaAtiva.toast('Primeiro processo de sincronização de Whats concluída', '33% concluído da função atual', tempoNotificacao);
  SincronizarCampoPlanilhas(abaInteresse, colWhatsInteresse, abaGerencial, colWhatsGerencial);
  planilhaAtiva.toast('Primeiro processo de sincronização de Whats concluída', '66% concluído da função atual', tempoNotificacao);
  SincronizarCampoPlanilhas(abaMarcoZero, colWhatsMarcoZero, abaGerencial, colWhatsGerencial);
  // Verificar quem respondeu Marco Zero
  VerificarEMarcarCadastroOutraPlanilha(abaInteresse, colRespondeuMarcoZeroInteresse, abaMarcoZero);
  // Sincronizar a situação entre Interesse e Gerencial
  SincronizarCampoPlanilhas(abaInteresse, colSituacaoInteresse, abaGerencial, colSituacaoGerencial);
  planilhaAtiva.toast('A planilha foi sincronizada', 'Finalização da execução', tempoNotificacao);
}

// Função que recebe as escolhas feitas na interface e esconde as linhas de acordo com elas
function ProcessarEscolhasEsconderLinhas(escolhas) {
  MostrarTodasLinhas();

  // Pega todos os valores necessários de acordo com as escolhas feitas
  const valores = {
    situacao: escolhas.situacao ? abaAtiva.getRange(2, colSituacaoInteresse, ultimaLinhaAtiva, 1).getValues().flat() : null,
    whats: escolhas.whats ? abaAtiva.getRange(2, colWhatsInteresse, ultimaLinhaAtiva, 1).getValues().flat() : null,
    respondeuMarcoZero: escolhas.respondeuMarcoZero ? abaAtiva.getRange(2, colRespondeuMarcoZeroInteresse, ultimaLinhaAtiva, 1).getValues().flat() : null,
  };

  // Loop que percorre todas as linhas
  for (let i = 0; i < ultimaLinhaAtiva; i++) {
    let esconderLinha = false;

    // Verifica cada condição
    if (escolhas.situacao && VerificarEsconderSituacao(escolhas.situacao, valores.situacao[i])) {
      esconderLinha = true;
    }
    if (escolhas.whats && VerificarEsconder(escolhas.whats, valores.whats[i])) {
      esconderLinha = true;
    }
    if (escolhas.respondeuMarcoZero && VerificarEsconder(escolhas.respondeuMarcoZero, valores.respondeuMarcoZero[i])) {
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
