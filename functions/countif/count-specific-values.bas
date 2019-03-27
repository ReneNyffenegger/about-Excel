option explicit

sub main() ' {

    testData

  '
  ' Count values of foo, bar and baz:
  '
    cells(1, 2) = "foo"
    cells(1, 3) = "bar"
    cells(1, 4) = "baz"

    range(cells(2,2), cells(2,4)).formulaR1C1 = "=countif(r3c1:r8c1, r[-1]c)"

end sub ' }

sub testData() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Insert a few values
 '
    cells(3, 1) = "foo"
    cells(4, 1) = "bar"
    cells(5, 1) = "foo"
    cells(6, 1) = "baz"
    cells(7, 1) = "baz"
    cells(8, 1) = "foo"

end sub ' }
