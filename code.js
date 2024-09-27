// This is a Google Workspace Add-on (GWAO)
// it's so the add-on can work accross the all documents on the user Docs app.
// however, the user needs to click on the extension icon to iniate it, which is a limitation.
// Editor Add-ons would work fine with just an onOpen(e) simple trigger for this extension.
// However, Editor Add-ons needs to be bound and activated by the user for each document.
// And activating it in document templates from Google Docs wouldn't transfer to the copies it creates.
// It's using the CardService instead of HtmlService for the UI (sidebar).

const doc = DocumentApp.getActiveDocument();


// Uses the CardService to build the add-on UI and add the buttons
function onHomepageOpen(e) {
  return createHomePage();
}

function createHomePage() {
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Extenso para Docs'))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("Escolha uma ação:"))
        .addWidget(
          CardService.newTextButton()
            .setText('Cifra por extenso')
            .setOnClickAction(CardService.newAction().setFunctionName('converterParaExtenso'))
        )
        .addWidget(
          CardService.newTextButton()
            .setText('Atualizar data')
            .setOnClickAction(CardService.newAction().setFunctionName('updateLastDateToToday'))
        )
    )
    .build();

  return card;
}

// Helper function to create a notification card
function createNotificationCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Ação finalizada'))
    .addSection(
      CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(message))
    )
    .build();
}


/* For HtmlService implementation option:
  function onOpen(e) {
    DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Cifra por extenso', 'converterParaExtenso')
      .addToUi();

    showSidebar();

    updateLastDateToToday();
  }

  function showSidebar() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Extenso para Docs')
      .setWidth(300);
    DocumentApp.getUi().showSidebar(htmlOutput);
  }
*/


function updateLastDateToToday() {
  const body = doc.getBody();
  
  // Get today's date
  const today = new Date();
  
  // Format the date in Portuguese
  const options = { day: '2-digit', month: 'long', year: 'numeric' };
  const formattedDate = new Intl.DateTimeFormat('pt-BR', options).format(today);
  
  // Convert the formatted date to lowercase
  const formattedDateLowerCase = formattedDate.toLowerCase();
  
  // Regular expression to match dates in the format "d de MMMM de yyyy" or "dº de MMMM de yyyy"
  const datePattern = /\d{1,2}º? de [a-zçáéíóúãõ]+ de \d{4}/gi;
  
  // Get the text of the document
  const text = body.getText();
  
  // Find all matches of the date pattern
  const matches = text.match(datePattern);
  
  if (matches && matches.length > 0) {
    // Get the last match
    const lastMatch = matches[matches.length - 1];
    
    // Replace the last matched date with the formatted current date
    body.replaceText(lastMatch, formattedDateLowerCase);
  }
}

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
        auxnumero2 = vlrS.substring(cont - i, cont - i + 1);
        auxnumero = parseInt(auxnumero2);

        if ((i == 14) || (i == 11) || (i == 8) || (i == 5) || (i == 2)) {
          auxnumero2 = vlrS.substring(cont - i, cont - i + 2);
          auxnumero = parseInt(auxnumero2);
        }

        if ((i == 15) || (i == 12) || (i == 9) || (i == 6) || (i == 3)) {
          extenso = extenso + centena[auxnumero];
          auxnumero2 = vlrS.substring(cont - i + 1, cont - i + 2);
          auxnumero3 = vlrS.substring(cont - i + 2, cont - i + 3);

          if ((auxnumero2 != "0") || (auxnumero3 != "0"))
              extenso += " e ";
        } else if (auxnumero > 19) {
          auxnumero2 = vlrS.substring(cont - i, cont - i + 1);

          auxnumero = parseInt(auxnumero2);
          extenso = extenso + dezena[auxnumero];
          auxnumero3 = vlrS.substring(cont - i + 1, cont - i + 2);

          if ((auxnumero3 != "0") && (auxnumero2 != "1"))
              extenso += " e ";
        } else if ((auxnumero <= 19) && (auxnumero > 0) && ((i == 14) || (i == 11) || (i == 8) || (i == 5) || (i == 2))) {
          extenso = extenso + unidade[auxnumero];
        } else if ((auxnumero < 10) && ((i == 13) || (i == 10) || (i == 7) || (i == 4) || (i == 1))) {
          auxnumero3 = vlrS.substring(cont - i - 1, cont - i);
          console.log(auxnumero3);
          console.log(extenso);
          if (auxnumero3 !== "" && auxnumero3 !== "0" && auxnumero3 !== "1")
            extenso=extenso+unidade[auxnumero];
              console.log(extenso);

          // Original:
          // else if((auxnumero<10)&&((i==13)||(i==10)||(i==7)||(i==4)||(i==1))) {
          //   auxnumero3 = vlrS.substr(cont-i-1,1);
          //   if((auxnumero3!="1")&&(auxnumero3!=""))
          //   extenso=extenso+unidade[auxnumero];
          // }
        }

        if (i % 3 == 1) {
            verifica3 = cont - i;
            if (verifica3 == 0)
              verifica1 = vlrS.substring(cont - i, cont - i + 1);

            if (verifica3 == 1)
              verifica1 = vlrS.substring(cont - i - 1, cont - i + 1);

            if (verifica3 > 1)
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

      // Handle singular for "real" and "centavo"
      if (inteiro === 1) {
          extenso = extenso.replace("reais", "real");
      }
      if (resto === 1) {
          extenso = extenso.replace("centavos", "centavo");
      }
      // Handle special case for "cem reais" (100 reais)
      if (inteiro === 100) {
        extenso = extenso.replace("cento  reais", "cem reais");  
        extenso = extenso.replace("cento reais", "cem reais");
      }
      // Handle special case for "zero reais" (100 reais)
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


function converterParaExtenso() {
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getRangeElements();

    for (let i = 0; i < elements.length; i++) {
      let element = elements[i];

      Logger.log('element[i] looping is' + element);

      if (element.getElement().editAsText) {
        const textElement = element.getElement().asText();
        const startOffset = element.getStartOffset();
        const endOffset = element.getEndOffsetInclusive();
        const selectedText = textElement.getText().substring(startOffset, endOffset + 1);

        Logger.log('textElement is ' + textElement);
        Logger.log('startOffset is ' + startOffset);
        Logger.log('endOffset is ' + endOffset);
        Logger.log('selectedText is ' + selectedText);
        
        // const pattern = /(?:\s|R\$|\$)(\d{1,3}(?:\.\d{3})*,\d{2,4})/g;
        // Updated regex pattern to support both formats: with or without the thousands separator
        const pattern = /(?:\s|R\$|\$)(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3},\d{2})/g;
       
        let newText = selectedText;
        let match;

        while ((match = pattern.exec(selectedText)) !== null) {
          let currencyString = match[1];
          
          // Remove periods and replace commas for decimal format        
          let textContent = currencyString.replace(/\./g, '').replace(',', '.');

          // Call escreverPorExtenso to convert the number to text
          let convertedText = escreverPorExtenso(textContent);

          
          // Clean up any unnecessary spaces
          convertedText = convertedText.replace(/\s{2,}/g, ' ').trim();
          while (convertedText.includes("( ")) {
            convertedText = convertedText.replace("( ", "(");
          }
          while (convertedText.includes(" )")) {
            convertedText = convertedText.replace(" )", ")");
          }
          
          newText = newText.replace(currencyString, `${currencyString} (${convertedText})`).trim();
          Logger.log('newText is ' + newText);
        }

        if (!selectedText.includes('R$')) {
          DocumentApp.getUi().alert('Se não for valor em reais ajuste o texto manualmente!');
        }

        if (newText !== selectedText) {
          // Inserir o texto convertido no mesmo lugar do texto selecionado
          Logger.log('newText after loop is ' + newText);
          newText += ' '; // Concatenar um espaço no final para posicionar o cursor
          textElement.deleteText(startOffset, endOffset);
          textElement.insertText(startOffset, newText);

          // Colocar o cursos no final de newText para seguir escrevendo, colocando - 1 para selecionar o espaço recém concatenado
          const newEndOffset = startOffset + newText.length - 1;
          const rangeBuilder = doc.newRange();
          rangeBuilder.addElement(textElement, newEndOffset, newEndOffset);
          doc.setSelection(rangeBuilder.build());

          return;
        }
        
        DocumentApp.getUi().alert('Selecione a cifra inteira, desde R$ e verifique se há separador de milhares');
        Logger.log('String selecionada é incompatível');
        return;
      }
    }
  }
  DocumentApp.getUi().alert('Nenhum texto compatível selecionado');
  Logger.log('Nenhum texto compatível selecionado');
  return;
}