option explicit

sub main() ' {

    dim rangeData as range

    set rangeData = testData

  '
  ' Determine the position of the vertical bar within the »text«:
  '
    rangeData.offset(0, 1).formulaR1C1 = "=find(""|"", rc[-1])"

  '
  ' Use the position along with left to extract the porition of »text« that
  ' is to the left of the vertical bar:
  '
    rangeData.offset(0, 2).formulaR1C1 = "=left(rc[-2], rc[-1] - 1)"

  '
  ' Also extract the portion of »text« to the right of the vertical bar:
  '
    rangeData.offset(0, 3).formulaR1C1 = "=right(rc[-3], len(rc[-3]) - rc[-2]) "

end sub ' }

function testData() as range' {

    activeSheet.cells.clearContents

    cells(1, 1) = "abc|defg hij"
    cells(2, 1) = "kl|mnopq|rst"
    cells(3, 1) = "uvwx|yz"

    set testData = range(cells(1, 1), cells(3, 1))

end function ' }
