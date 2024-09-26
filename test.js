function escreverPorExtenso(vlr) {
    var Num = parseFloat(vlr);
    var result = ""; // Use a variable to store the result instead of direct DOM manipulation

    if (vlr == 0) {
        result = "zero";
    } else if (vlr > 999 * 1e12) {
        return "Numero maior que 999 trilhões";
    } else {
        var inteiro = parseInt(vlr); // parte inteira do valor
        if (inteiro < 1000000000000000) {
            var resto = Num.toFixed(2) - inteiro.toFixed(2); // parte fracionária do valor
            resto = resto.toFixed(2);
            var vlrS = inteiro.toString();

            var cont = vlrS.length;
            var extenso = "";
            var auxnumero;
            var auxnumero2;
            var auxnumero3;

            var unidade = ["", "um", "dois", "três", "quatro", "cinco",
                "seis", "sete", "oito", "nove", "dez", "onze",
                "doze", "treze", "quatorze", "quinze", "dezesseis",
                "dezessete", "dezoito", "dezenove"];

            var centena = ["", "cento", "duzentos", "trezentos",
                "quatrocentos", "quinhentos", "seiscentos",
                "setecentos", "oitocentos", "novecentos"];

            var dezena = ["", "", "vinte", "trinta", "quarenta", "cinquenta",
                "sessenta", "setenta", "oitenta", "noventa"];

            var qualificaS = ["reais", "mil", "milhão", "bilhão", "trilhão"];
            var qualificaP = ["reais", "mil", "milhões", "bilhões", "trilhões"];

            for (var i = cont; i > 0; i--) {
                var verifica1 = "";
                var verifica2 = 0;
                var verifica3 = 0;
                auxnumero2 = "";
                auxnumero3 = "";
                auxnumero = 0;
                // auxnumero2 = vlrS.substr(cont - i, 1); .substr is deprecated
                auxnumero2 = vlrS.substring(cont - i, cont - i + 1);
                auxnumero = parseInt(auxnumero2);

                if ((i == 14) || (i == 11) || (i == 8) || (i == 5) || (i == 2)) {
                    // auxnumero2 = vlrS.substr(cont - i, 2); .subst is deprecated
                    auxnumero2 = vlrS.substring(cont - i, cont - i + 2);
                    auxnumero = parseInt(auxnumero2);
                }

                if ((i == 15) || (i == 12) || (i == 9) || (i == 6) || (i == 3)) {
                    extenso = extenso + centena[auxnumero];
                    // auxnumero2 = vlrS.substr(cont - i + 1, 1); .substr is deprecated
                    auxnumero2 = vlrS.substring(cont - i + 1, cont - i + 2);
                    // auxnumero3 = vlrS.substr(cont - i + 2, 1); .substr is deprecated
                    auxnumero3 = vlrS.substring(cont - i + 2, cont - i + 3);

                    if ((auxnumero2 != "0") || (auxnumero3 != "0"))
                        extenso += " e ";
                } else if (auxnumero > 19) {
                    // auxnumero2 = vlrS.substr(cont - i, 1); .substr is deprecated
                    auxnumero2 = vlrS.substring(cont - i, cont - i + 1);

                    auxnumero = parseInt(auxnumero2);
                    extenso = extenso + dezena[auxnumero];
                    // auxnumero3 = vlrS.substr(cont - i + 1, 1); .substr is deprecated
                    auxnumero3 = vlrS.substring(cont - i + 1, cont - i + 2);

                    if ((auxnumero3 != "0") && (auxnumero2 != "1"))
                        extenso += " e ";
                } else if ((auxnumero <= 19) && (auxnumero > 0) && ((i == 14) || (i == 11) || (i == 8) || (i == 5) || (i == 2))) {
                    extenso = extenso + unidade[auxnumero];
                } else if ((auxnumero < 10) && ((i == 13) || (i == 10) || (i == 7) || (i == 4) || (i == 1))) {
                    auxnumero3 = vlrS.substring(cont - i - 1, cont - i);
                    // Ensure to account for cases where there's a leading zero
                    if ((auxnumero3 != "1" || auxnumero3 === "") && auxnumero > 0) {
                        extenso = extenso + unidade[auxnumero];
                    }
                }                

                if (i % 3 == 1) {
                    verifica3 = cont - i;
                    if (verifica3 == 0)
                        // verifica1 = vlrS.substr(cont - i, 1); .substr is deprecated
                        verifica1 = vlrS.substring(cont - i, cont - i + 1);

                    if (verifica3 == 1)
                        // verifica1 = vlrS.substr(cont - i - 1, 2); .substr is deprecated
                        verifica1 = vlrS.substring(cont - i - 1, cont - i + 1);

                    if (verifica3 > 1)
                        // verifica1 = vlrS.substr(cont - i - 2, 3); .substr is deprecated
                        verifica1 = vlrS.substring(cont - i - 2, cont - i + 1);

                    verifica2 = parseInt(verifica1);

                    if (i == 13) {
                        if (verifica2 == 1) {
                            extenso = extenso + " " + qualificaS[4] + " ";
                        } else if (verifica2 != 0) { extenso = extenso + " " + qualificaP[4] + " "; }
                    }
                    if (i == 10) {
                        if (verifica2 == 1) {
                            extenso = extenso + " " + qualificaS[3] + " ";
                        } else if (verifica2 != 0) { extenso = extenso + " " + qualificaP[3] + " "; }
                    }
                    if (i == 7) {
                        if (verifica2 == 1) {
                            extenso = extenso + " " + qualificaS[2] + " ";
                        } else if (verifica2 != 0) { extenso = extenso + " " + qualificaP[2] + " "; }
                    }
                    if (i == 4) {
                        if (verifica2 == 1) {
                            extenso = extenso + " " + qualificaS[1] + " ";
                        } else if (verifica2 != 0) { extenso = extenso + " " + qualificaP[1] + " "; }
                    }
                    if (i == 1) {
                        if (verifica2 == 1) {
                            extenso = extenso + " " + qualificaS[0] + " ";
                        } else { extenso = extenso + " " + qualificaP[0] + " "; }
                    }
                }
            }
            resto = resto * 100;
            var aexCent = 0;
            if (resto <= 19 && resto > 0)
                extenso += " e " + unidade[resto] + " centavos";
            if (resto > 19) {
                aexCent = parseInt(resto / 10);

                extenso += " e " + dezena[aexCent];
                resto = resto - (aexCent * 10);

                if (resto != 0)
                    extenso += " e " + unidade[resto] + " centavos";
                else extenso += " centavos";
            }
            
            console.log(inteiro);
            // Handle singular for "real" and "centavo"
            if (inteiro === 1) {
                extenso = extenso.replace("reais", "real");
            }
            if (resto === 1) {
                extenso = extenso.replace("centavos", "centavo");
            }
            // Handle special case for "cem reais" (100 reais)
            if (inteiro === 100) {
                extenso = extenso.replace("cento reais", "cem reais");
            }
            if (inteiro === 0) {
                extenso = extenso.replace("reais  e ", "");
                extenso = extenso.replace("reais e ", "");
            }
            // // Handle special case for "mil reais" (1000 reais)
            // if (inteiro === 1000) {
            //     extenso = extenso.replace("um mil reais", "mil reais");
            // }

            extenso.trim(); // Remove whitespace from both sides of the string

            result = extenso; // Assign the final result to the variable
        }
    }
    return result; // Return the result instead of using document.getElementById
}

vlr = 101.18;
console.log(escreverPorExtenso(vlr));