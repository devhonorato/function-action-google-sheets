function onEdit(e) {
    // Variavel para verificar qual é a guia ativa
    // Desnecessário // const GuiaAtiva = SpreadsheetApp.getActive().getSheetName();

    // Se for diferente de "Página1", ele interrompe o resto do código e retorna false.
    // Desnecessário // if(GuiaAtiva != "Página1"){
    // Desnecessário //	    return false;
    // Desnecessário // }

    // Pega a guia com os dados da "Página1"
    const GuiaDados =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");

    // Pega o numero da linha editada/ativa
    let Linha = GuiaDados.getActiveRange().getRowIndex();

    // Verifica se o numero da linha é maior que 1, pois o numero 1 é o cabeçalho da planilha
    if (Linha > 1) {

        // Pega o valor da linha ativa na coluna 3 e guarda na variavel CelulaData
        CelulaData = GuiaDados.getRange(Linha, 3).getValue();

        // Verifica se a CelulaData é igual à ""
        if (CelulaData == "") {
            // Pega o valor da linha ativa na coluna 2 e guarda na variavel Email
            const Email = GuiaDados.getRange(Linha, 2).getValue();

            // Verifica se o Email é diferente de ""
            if (Email != "") {
                // Gera uma nova data (data atual)
                const Data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

                // Inseri a data na linha coluna 3
                GuiaDados.getRange(Linha, 3).setValue(Data);

                // Variavel para verificar se as proximas linhas estão vazias ou preenchidas
                let verificar = false;

                // Inicialmente essa verificação é sempre falsa e caso seja verdadeira ela sai do loop e finaliza o código
                while (verificar != true) {
                    // Linha atual mais 1 para pegar os valores da proxima linha
                    Linha = Linha + 1;

                    // Pega o valor da 'nova' linha na coluna 3 e guarda na variavel CelulaDataNew
                    CelulaDataNew = GuiaDados.getRange(Linha, 3).getValue();

                    // Verifica se a CelulaDataNew é igual à ""
                    if (CelulaDataNew == "") {
                        // Pega o valor da 'nova' linha na coluna 2 e guarda na variavel EmailNew
                        const EmailNew = GuiaDados.getRange(Linha, 2).getValue();

                        // Pega o valor da 'nova' linha na coluna 1 e guarda na variavel NomeNew
                        const NomeNew = GuiaDados.getRange(Linha, 1).getValue();

                        // Verifica se o EmailNew é diferente de ""
                        if (EmailNew != "") {
                            // Gera uma nova DataNew (data atual)
                            const DataNew = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

                            // Inseri a data na linha coluna 3
                            GuiaDados.getRange(Linha, 3).setValue(DataNew);
                        }

                        // Se NomeNew, CelulaDataNew e EmailNew for igual à "", a vraiavel verificar recebe true e no proximo loop o código não será executado
                        if (NomeNew == "" && CelulaDataNew == "" && EmailNew == "") {
                            verificar = true;
                        }
                    }
                }
            }
        }
    }
}
