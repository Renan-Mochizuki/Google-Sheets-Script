// Função que verificará se o email existe na planilha Interesse e retornará a linha
const RetornarLinhaEmailInteresse = (emailMarcoZero) => {
    //Conferir todos os emails da planilha Interesse
    for (let i = 2; i <= ultimaLinhaInteresse; i++) {
        const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();

        if (emailMarcoZero == emailInteresse) return i;
    }
    // Se não for encontrado nenhum 
    return false;
}

//Função para verificar quem respondeu o Marco Zero
const VerificarMarcoZeroInteresse = () => {
    //Pegar o email na planilha Interesse
    for (let i = 2; i <= ultimaLinhaInteresse; i++) {
        const emailInteresse = abaInteresse.getRange(i, colEmail).getValue();
        const celEnviadoMarcoZero = abaInteresse.getRange(i, colFormEnviadoInteresse);
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
            celEnviadoMarcoZero.setValue("SIM");
        } else {
            celRespondeuMarcoZero.setValue("NÃO");
        }
    }
}

// Função que verificará se o email existe na planilha Marco Zero e retornará a linha
const RetornarLinhaEmailMarcoZero = (emailInteresse) => {
    //Conferir todos os emails da planilha Marco Zero
    for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
        const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();

        if (emailInteresse == emailMarcoZero) return i;
    }
    // Se não for encontrado nenhum 
    return false;
}

//Função para verificar se a pessoa está cadastrada na planilha de Interesse
function VerificarInteresseMarcoZero() {
    //Pegar o email na planilha Marco Zero
    for (let i = 2; i <= ultimaLinhaMarcoZero; i++) {
        const emailMarcoZero = abaMarcoZero.getRange(i, colEmail).getValue();
        const celEstaNaInteresse = abaMarcoZero.getRange(i, colRespondeuInteresseMarcoZero);
        const valEstaNaInteresse = celEstaNaInteresse.getValue();

        // Se o campo estiver vazio, limpe a célula e passe para o próximo
        if (!emailMarcoZero) {
            celEstaNaInteresse.setValue("");
            continue;
        }

        // Se o campo já estiver marcado passe para o próximo
        if (valEstaNaInteresse) continue;

        if (RetornarLinhaEmailInteresse(emailMarcoZero))
            celEstaNaInteresse.setValue("SIM");
        else
            celEstaNaInteresse.setValue("S. PÚBLICA");
    }
}