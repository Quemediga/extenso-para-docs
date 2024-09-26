I still need to handle a new bug for single digits:
- 102.18 to "cento e doisdois reais  e dezoito centavos" (incorrect).
- 2.18 to "dois reais  e dezoito centavos" (correct).
- 101.18 to "cento e umum reais  e dezoito centavos" (incorrect).

I'll also have a runtime error that terminates the app (user needs to refresh the page) if:
- the user selects more than one "R$ x.xxx,xx" structures.
- the user selects a cipher without the "." thousands separator (for pt-br), like in "R$ 1000,00".
- if there is no space between "$" and the numbers string, like in "R$1.000,00".