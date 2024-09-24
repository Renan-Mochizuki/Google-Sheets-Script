// -- Funções das planilhas Interesse e Marco Zero --

//Função para verificar quem respondeu o Marco Zero
const VerificarMarcoZeroInteresse = () => {
    //Pegar o email na planilha Interesse
    for (let i = 2; i <= ultimaLinhaInteresse; i++) {
        const email = abaInteresse.getRange(i, colEmailInteresse).getValue();
        const celRespondeuMarcoZero = abaInteresse.getRange(i, colRespondeuMarcoZeroInteresse);
        const valRespondeuMarcoZero = celRespondeuMarcoZero.getValue();

        // Se o campo estiver vazio, limpe a célula e passe para o próximo
        if (!email) {
            celRespondeuMarcoZero.setValue("");
            continue;
        }

        // Se o campo já estiver marcado com sim
        if (valRespondeuMarcoZero == "SIM") continue;

        if (RetornarLinhaEmailPlanilha(email, abaMarcoZero)) {
            celRespondeuMarcoZero.setValue("SIM");
        } else {
            celRespondeuMarcoZero.setValue("NÃO");
        }
    }
}

//Função para verificar se a pessoa está cadastrada na planilha de Interesse
function VerificarInteresseMarcoZero() {
    //Pegar o email na planilha Marco Zero
    for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
        const email = abaMarcoZero.getRange(i, colEmailMarcoZero).getValue();
        const celEstaNaInteresse = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero);
        const valEstaNaInteresse = celEstaNaInteresse.getValue();

        // Se o campo estiver vazio, limpe a célula e passe para o próximo
        if (!email) {
            celEstaNaInteresse.setValue("");
            continue;
        }

        // Se o campo já estiver marcado como "SIM" passe para o próximo
        if (valEstaNaInteresse == "SIM") continue;

        if (RetornarLinhaEmailPlanilha(email, abaInteresse))
            celEstaNaInteresse.setValue("SIM");
        else
            celEstaNaInteresse.setValue("S. PÚBLICA");
    }
}