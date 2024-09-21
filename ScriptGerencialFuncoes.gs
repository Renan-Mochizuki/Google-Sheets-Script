// -- Funções das planilhas Interesse e Marco Zero --

// Função que verificará se o email existe na planilha Interesse e retornará a linha
const RetornarLinhaEmailInteresse = (emailProcurado) => {
    //Conferir todos os emails da planilha Interesse
    for (let i = 2; i <= ultimaLinhaInteresse; i++) {
        const emailInteresse = abaInteresse.getRange(i, colEmailInteresse).getValue();

        if (emailProcurado == emailInteresse) return i;
    }
    // Se não for encontrado nenhum 
    return false;
}

//Função para verificar quem respondeu o Marco Zero
const VerificarMarcoZeroInteresse = () => {
    //Pegar o email na planilha Interesse
    for (let i = 2; i <= ultimaLinhaInteresse; i++) {
        const emailInteresse = abaInteresse.getRange(i, colEmailInteresse).getValue();
        const celRespondeuMarcoZero = abaInteresse.getRange(i, colRespondeuMarcoZeroInteresse);
        const valRespondeuMarcoZero = celRespondeuMarcoZero.getValue();

        // Se o campo estiver vazio, limpe a célula e passe para o próximo
        if (!emailInteresse) {
            celRespondeuMarcoZero.setValue("");
            continue;
        }

        // Se o campo já estiver marcado com sim
        if (valRespondeuMarcoZero == "SIM") continue;

        if (RetornarLinhaEmailMarcoZero(emailInteresse)) {
            celRespondeuMarcoZero.setValue("SIM");
        } else {
            celRespondeuMarcoZero.setValue("NÃO");
        }
    }
}

// Função que verificará se o email existe na planilha Marco Zero e retornará a linha
const RetornarLinhaEmailMarcoZero = (emailProcurado) => {
    //Conferir todos os emails da planilha Marco Zero
    for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
        const emailMarcoZero = abaMarcoZero.getRange(i, colEmailMarcoZero).getValue();

        if (emailProcurado == emailMarcoZero) return i;
    }
    // Se não for encontrado nenhum 
    return false;
}

//Função para verificar se a pessoa está cadastrada na planilha de Interesse
function VerificarInteresseMarcoZero() {
    //Pegar o email na planilha Marco Zero
    for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
        const emailMarcoZero = abaMarcoZero.getRange(i, colEmailMarcoZero).getValue();
        const celEstaNaInteresse = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero);
        const valEstaNaInteresse = celEstaNaInteresse.getValue();

        // Se o campo estiver vazio, limpe a célula e passe para o próximo
        if (!emailMarcoZero) {
            celEstaNaInteresse.setValue("");
            continue;
        }

        // Se o campo já estiver marcado como "SIM" passe para o próximo
        if (valEstaNaInteresse == "SIM") continue;

        if (RetornarLinhaEmailInteresse(emailMarcoZero))
            celEstaNaInteresse.setValue("SIM");
        else
            celEstaNaInteresse.setValue("S. PÚBLICA");
    }
}