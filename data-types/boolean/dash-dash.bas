option explicit

sub main() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Fill some boolean values into the first column:
 '
    cells(1,1) = false
    cells(2,1) = true
    cells(3,1) = true
    cells(4,1) = false
    cells(5,1) = true

  '
  ' Use dash-dash (or minus minus) to turn boolean values into 0 and 1 in
  ' second column:
  '
    range(cells(1,2), cells(5, 2)).formulaR1C1 = "= -- rc[-1]"

end sub ' }
