// Função que sincronizará um dado campo entre duas planilhas passadas
function SincronizarCampoPlanilhas(abaDesejada1, colDesejada1, abaDesejada2, colDesejada2) {
  // Atribui as variáveis de acordo com as abas
  const { ultimaLinha: ultimaLinha1, colNome: colNome1, nomePlanilha: nomePlanilha1 } = objetoMap.get(abaDesejada1);
  const { ultimaLinha: ultimaLinha2, colNome: colNome2, nomePlanilha: nomePlanilha2 } = objetoMap.get(abaDesejada2);

  // Pegando todos os emails da abaDesejada1 e abaDesejada2
  const nomesEmailsTelefones1 = abaDesejada1.getRange(2, colNome1, ultimaLinha1, 3).getValues();
  const nomesEmailsTelefones2 = abaDesejada2.getRange(2, colNome2, ultimaLinha2, 3).getValues();

  const colsDesejadas1 = abaDesejada1.getRange(2, colDesejada1, ultimaLinha1, 1).getValues();
  const colsDesejadas2 = abaDesejada2.getRange(2, colDesejada2, ultimaLinha2, 1).getValues();

  // Loop para percorrer as linhas da abaDesejada1
  for (let i = 0; i < nomesEmailsTelefones1.length; i++) {
    const nome = nomesEmailsTelefones1[i][0];
    const email = nomesEmailsTelefones1[i][1];
    const telefone = nomesEmailsTelefones1[i][2];

    if (!ValidarLoop(nome, email, telefone)) continue;

    // Pegue a linha do campo na planilha desejada 2
    const linhaCampoDesejada2 = RetornarLinhaDados(nome, email, telefone, nomesEmailsTelefones2);

    // Se o email for encontrado na outra planilha
    if (linhaCampoDesejada2) {
      // Faremos linhaCampoDesejada - 2 pois a array começa em 0 e a planilha em 2
      const valDesejada1 = colsDesejadas1[i][0];
      const valDesejada2 = colsDesejadas2[linhaCampoDesejada2 - 2][0];

      // Caso especial para quando for sincronizado a situação entre as planilhas
      if ((abaDesejada1 == abaInteresse && colDesejada1 == colSituacaoInteresse) || (abaDesejada1 == abaGerencial && colDesejada1 == colSituacaoGerencial)) {
        const turmaMaisRecente = RetornarTurmaMaisRecente(valDesejada1, valDesejada2);
        colsDesejadas1[i][0] = turmaMaisRecente;
        colsDesejadas2[linhaCampoDesejada2 - 2][0] = turmaMaisRecente;
        continue;
      }

      // Lógica principal
      const valorSimNao = RetornarValorSimNao(valDesejada1, valDesejada2);
      colsDesejadas1[i][0] = valorSimNao;
      colsDesejadas2[linhaCampoDesejada2 - 2][0] = valorSimNao;
    }
  }
  // Toast da mensagem do progresso de execução da função
  const tituloToast = '50% concluído da função atual';
  const textoToast = 'Sincronizando campo entre planilhas ' + nomePlanilha1 + ' e ' + nomePlanilha2;
  planilhaAtiva.toast(textoToast, tituloToast, tempoNotificacao);

  // Loop para percorrer as linhas da abaDesejada2 (Caso houver uma pessoa repetida na abaDesejada2)
  for (let i = 0; i < nomesEmailsTelefones2.length; i++) {
    const nome = nomesEmailsTelefones2[i][0];
    const email = nomesEmailsTelefones2[i][1];
    const telefone = nomesEmailsTelefones2[i][2];

    if (!ValidarLoop(nome, email, telefone)) continue;

    // Pegue a linha do campo na planilha desejada 1
    const linhaCampoDesejada1 = RetornarLinhaDados(nome, email, telefone, nomesEmailsTelefones1);

    // Se o email for encontrado na outra planilha
    if (linhaCampoDesejada1) {
      // Faremos linhaCampoDesejada - 2 pois a array começa em 0 e a planilha em 2
      const valDesejada1 = colsDesejadas1[linhaCampoDesejada1 - 2][0];
      const valDesejada2 = colsDesejadas2[i][0];

      // Caso especial para quando for sincronizado a situação entre as planilhas
      if ((abaDesejada2 == abaInteresse && colDesejada2 == colSituacaoInteresse) || (abaDesejada2 == abaGerencial && colDesejada2 == colSituacaoGerencial)) {
        const turmaMaisRecente = RetornarTurmaMaisRecente(valDesejada1, valDesejada2);
        colsDesejadas1[linhaCampoDesejada1 - 2][0] = turmaMaisRecente;
        colsDesejadas2[i][0] = turmaMaisRecente;
        continue;
      }

      // Lógica principal
      const valorSimNao = RetornarValorSimNao(valDesejada1, valDesejada2);
      colsDesejadas1[linhaCampoDesejada1 - 2][0] = valorSimNao;
      colsDesejadas2[i][0] = valorSimNao;
    }
  }

  // Inserindo os valores nas planilhas
  abaDesejada1.getRange(2, colDesejada1, ultimaLinha1, 1).setValues(colsDesejadas1);
  abaDesejada2.getRange(2, colDesejada2, ultimaLinha2, 1).setValues(colsDesejadas2);
}

//Função que verifica se a pessoa está cadastrada na planilha para verificar, e registra em outra planilha
function VerificarEMarcarCadastroOutraPlanilha(abaParaRegistro, colParaRegistro, abaParaVerificar, valCustomizadoSim, valCustomizadoNao) {
  // Atribui as variáveis de acordo com a abaDesejada1
  const { ultimaLinha: ultimaLinhaVerificar, colNome: colNomeVerificar, nomePlanilha: nomePlanilhaVerificar } = objetoMap.get(abaParaVerificar);
  // Atribui as variáveis de acordo com a abaParaRegistro
  const { ultimaLinha: ultimaLinhaRegistro, colNome: colNomeRegistro, nomePlanilha: nomePlanilhaRegistro } = objetoMap.get(abaParaRegistro);

  // Pegando todos os emails da abaParaVerificar e abaParaRegistro
  const nomesEmailsTelefonesVerificar = abaParaVerificar.getRange(2, colNomeVerificar, ultimaLinhaVerificar, 3).getValues();
  const nomesEmailsTelefonesRegistro = abaParaRegistro.getRange(2, colNomeRegistro, ultimaLinhaRegistro, 3).getValues();

  const colsRegistro = abaParaRegistro.getRange(2, colParaRegistro, ultimaLinhaRegistro, 1).getValues();

  // Loop para percorrer as linhas da abaParaRegistro
  for (let i = 0; i < nomesEmailsTelefonesRegistro.length; i++) {
    const nome = nomesEmailsTelefonesRegistro[i][0];
    const email = nomesEmailsTelefonesRegistro[i][1];
    const telefone = nomesEmailsTelefonesRegistro[i][2];

    if (!ValidarLoop(nome, email, telefone)) continue;

    // Toast da mensagem do progresso de execução da função
    if (i % 300 === 0) {
      const tituloToast = Math.round((i / ultimaLinhaRegistro) * 100) + '% concluído da função atual';
      const textoToast = 'Processo na linha ' + i + ' da verificação da planilha ' + nomePlanilhaRegistro + ' para ' + nomePlanilhaVerificar;
      planilhaAtiva.toast(textoToast, tituloToast, tempoNotificacao);
    }

    const existeNaAbaVerificar = RetornarLinhaDados(nome, email, telefone, nomesEmailsTelefonesVerificar);

    // Se o email for encontrado na outra planilha
    if (existeNaAbaVerificar) {
      colsRegistro[i][0] = valCustomizadoSim ?? 'SIM';
    } else {
      colsRegistro[i][0] = valCustomizadoNao ?? 'NÃO';
    }
  }

  // Inserindo os valores na abaParaRegistro
  abaParaRegistro.getRange(2, colParaRegistro, ultimaLinhaRegistro, 1).setValues(colsRegistro);
}
