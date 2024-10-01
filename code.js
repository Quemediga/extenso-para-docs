// This is a Google Workspace Add-on (GWAO)
  // it's so the add-on can work accross the all documents on the user Docs app.
  // The user needs to click on the extension icon to iniate it, which is a limitation of any GWAO.
// On the other hand, an Editor Add-on would work fine with just an onOpen(e) simple trigger for this script.
  // However, Editor Add-ons needs to be bound and activated by the user for each document, unlike GWAO.
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
  
  // Format the date in pt-br
  const options = { day: '2-digit', month: 'long', year: 'numeric' };
  const formattedDate = new Intl.DateTimeFormat('pt-BR', options).format(today);
  
  // Convert the formatted date to lowercase
  const formattedDateLowerCase = formattedDate.toLowerCase();
  
  // Regular expression to also match dates in the format "d de MMMM de yyyy" and "dº de MMMM de yyyy"
  const datePattern = /\d{1,2}º? de [a-zçáéíóúãõ]+ de \d{4}/gi;
  
  // Get the text of the whole document
  const text = body.getText();
  
  // Find all matches of the date pattern
  const matches = text.match(datePattern);
  
  if (matches && matches.length > 0) {
    // Get the last match of the date pattern
    const lastMatch = matches[matches.length - 1];
    
    // Replace the last matched date with the formatted current date
    body.replaceText(lastMatch, formattedDateLowerCase);
  }
}

// This is an algorithm I contributed with on Stack Overflow
// I contributed by propely handling single-digits, solving special cases and adding proper formatting
function escreverPorExtenso(vlr) {
  var Num = parseFloat(vlr);
  var result = ""; // Variable to store the result

  if (vlr == 0) {
    result = "zero reais";
  } else if (vlr > 70000000000000) {
    DocumentApp.getUi().alert('Cifras acima dos 70 trilhões não são compatíveis.');
    console.log("cifras acima dos 70 trilhões não são compatíveis");
    return "";
  } else {
    var inteiro = parseInt(vlr); // Integer part of the number
    if (inteiro <= 70000000000000) {
      var resto = Num.toFixed(2) - inteiro.toFixed(2); // Decimal part of the number
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
          if ((auxnumero3 !== "1" || auxnumero > 9) && auxnumero3 !== "0")
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
            console.log('Verifica2 for trillions: ' + verifica2);
            if (verifica2 == 1) {
              extenso = extenso + " " + qualificaS[4] + " ";
            } else if (verifica2 != 0) {
              extenso = extenso + " " + qualificaP[4] + " ";
            }
          }
          if (i == 10) {
            if (verifica2 == 1) {
              extenso = extenso + " " + qualificaS[3] + " ";
            } else if (verifica2 != 0) {
              extenso = extenso + " " + qualificaP[3] + " ";
            }
          }
          if (i == 7) {
            if (verifica2 == 1) {
              extenso = extenso + " " + qualificaS[2] + " ";
            } else if (verifica2 != 0) {
              extenso = extenso + " " + qualificaP[2] + " ";
            }
          }
          if (i == 4) {
            if (verifica2 == 1) {
              extenso = extenso + " " + qualificaS[1] + " ";
            } else if (verifica2 != 0) {
              extenso = extenso + " " + qualificaP[1] + " ";
            }
          }
          if (i == 1) {
            if (verifica2 == 1) {
              extenso = extenso + " " + qualificaS[0] + " ";
            } else {
              extenso = extenso + " " + qualificaP[0] + " ";
            }
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

      // Handle special cases
      if (inteiro === 1) {
          extenso = extenso.replace("reais", "real");
      }
      if (resto === 1) {
          extenso = extenso.replace("centavos", "centavo");
      }
      if (inteiro === 100) {
        extenso = extenso.replace("cento  reais", "cem reais");  
        extenso = extenso.replace("cento reais", "cem reais");
      }
      if (inteiro === 0) {
        extenso = extenso.replace("reais  e ", "");
        extenso = extenso.replace("reais e ", "");
      }
      // // Handle special case for "mil reais" (1000 reais) (optional)
      // if (inteiro === 1000) {
      //     extenso = extenso.replace("um mil reais", "mil reais");
      // }

      // Proper formatting
      while (extenso.includes("  ")) {
        extenso = extenso.replace("  ", " ");
      }
      extenso.trim();

      console.log('extenso result is: ' + extenso);

      result = extenso;
    }
  }
  return result;
}


function converterParaExtenso() {
  const selection = doc.getSelection();
  
  if (selection) {
    const elements = selection.getRangeElements();

    for (let i = 0; i < elements.length; i++) {
      let element = elements[i];

      console.log('element[i] looping is: ' + element);

      if (element.getElement().editAsText) {
        const textElement = element.getElement().asText();
        const startOffset = element.getStartOffset();
        const endOffset = element.getEndOffsetInclusive();
        const selectedText = textElement.getText().substring(startOffset, endOffset + 1);

        console.log('textElement is: ' + textElement);
        console.log('startOffset is: ' + startOffset);
        console.log('endOffset is: ' + endOffset);
        console.log('selectedText is: ' + selectedText);
        
        const pattern = /(?:\s|R\$|\$)(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3},\d{2})/g;
       
        let newText = selectedText;
        let match;

        while ((match = pattern.exec(selectedText)) !== null) {
          let currencyString = match[1];
          
          // Prepare the R$ cypher to parseFloat it in escreverPorExtenso()
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

          // Now I create the actual output
          newText = newText.replace(currencyString, `${currencyString} (${convertedText})`).trim();
          console.log('newText output from the loop is: ' + newText);

        }

        // Return feedback to the user
        while ((match = pattern.exec(selectedText)) === null) {
          return DocumentApp.getUi().alert('Selecione apenas uma cifra e verifique se há o espaço entre $ e os números');
        }
        if (!selectedText.includes('R$')) {
          DocumentApp.getUi().alert('Se não for valor em reais ajuste o texto manualmente!');
        }

        // Insert the convertedText at the same position as the the currencyString in Google Docs
        if (newText !== selectedText) {
          console.log('newText when ready is: ' + newText);
          // Concatenate a space at the end of the string to position the cursor at it (so the user doesn't have to leave the keyboard)
          newText += ' ';
          // Insert the text on it's original position on the document
          textElement.deleteText(startOffset, endOffset);
          textElement.insertText(startOffset, newText);

          // Put the cursor at the end of newText, then position -1 to select the space we concatenated to it (so the user can keep on writting)
          // The user will be able to press Shift + Tab (move to the Close button) and Enter to close the GAWO UI and the cursor will be positioned already
          const newEndOffset = startOffset + newText.length - 1;
          const rangeBuilder = doc.newRange();
          rangeBuilder.addElement(textElement, newEndOffset, newEndOffset);
          doc.setSelection(rangeBuilder.build());

          return;
        }
        
        DocumentApp.getUi().alert('Selecione a cifra inteira, desde R$, e verifique se há o espaço depois de $ e o separador de milhares');
        console.log('String selecionada é incompatível');
        return;
      }
    }
  }
  DocumentApp.getUi().alert('Nenhum texto compatível selecionado');
  console.log('Nenhum texto compatível selecionado');
  return;
}