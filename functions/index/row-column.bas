option explicit

sub main() ' {

    testData

  '
  ' Get the value of the 4th row and 2nd column
  ' in the specified range (r2c2 to r6c4), thus
  ' evaluating to "5/3"
  '
    cells(2, 6).formulaR1C1 = "=index(r2c2:r6c4, 4, 2)"

end sub ' }

sub testData() ' {


    activeSheet.usedRange.clearContents

    dim r as long
    dim c as long

    for r = 2 to 6
    for c = 2 to 4
      '
      ' Specify the cell as text (although called numberFormat...)
      '
        cells(r, c).numberFormat = "@"

      '
      ' Set cell's text to row slash column.
      '
        cells(r, c) = r & "/" & c
    next c
    next r

end sub ' }
