option explicit

dim dataRow    as integer
dim formulaRow as integer
dim formulaCol as integer

sub main() ' {

    fillData

    range(cells(1, 1), cells(dataRow, 1)).name = "someNumbers"
    range(cells(1, 2), cells(dataRow, 2)).name = "spellings"

    formulaCol = 4
  '
  ' 39 is a number that exists in someNumbers
  '
    addFormula "match(39, someNumbers)"
    addFormula "match(39, someNumbers, 0)"
    addFormula "match(39, someNumbers, 1)"

  '
  ' 42 is a number that does not exist in someNumbers
  '
    addFormula "match(42, someNumbers)"
    addFormula "match(42, someNumbers, 0)"
    addFormula "match(42, someNumbers, 1)"


    formulaRow = 0
    formulaCol = 7
  '
  ' fifty-three is a string that exists in spellings
  '
    addFormula "match(""fifty-three"", spellings)"
    addFormula "match(""fifty-three"", spellings, 0)"
    addFormula "match(""fifty-three"", spellings, 1)"

  '
  ' thirty is a string that does not exist in spellings
  '
    addFormula "match(""thirty"", spellings)"
    addFormula "match(""thirty"", spellings, 0)"
    addFormula "match(""thirty"", spellings, 1)"

    range(cells(1,1), cells(1,8)).entireColumn.autoFit

    cells(dataRow+2, 10).select

end sub ' }

sub fillData() ' {

    addDataRow 12, "twelve"
    addDataRow 17, "seventeen"
    addDataRow 21, "twenty-one"
    addDataRow 25, "twenty-five"
    addDataRow 31, "thirty-one"
    addDataRow 39, "thirty-nine"
    addDataRow 53, "fifty-three"
    addDataRow 74, "seventy-four"
    addDataRow 99, "ninety-ine"

end sub ' }

sub addDataRow(value as integer, text as string) ' {

    dataRow = dataRow + 1
    cells(dataRow, 1) = value
    cells(dataRow, 2) = text

end sub ' }

sub addFormula(text as string) ' {

    formulaRow = formulaRow + 1
    cells(formulaRow, formulaCol  )             =       text
    cells(formulaRow, formulaCol+1).formulaR1C1 = "=" & text
    cells(formulaRow, formulaCol  ).font.name   = "Courier New"

end sub ' }
