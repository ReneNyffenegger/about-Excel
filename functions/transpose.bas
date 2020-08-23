option explicit

sub main() ' {

    cells(2,2) = "ABC" : cells(2,3) = 111
    cells(3,2) = "DEF" : cells(3,3) = 222
    cells(4,2) = "GHI" : cells(4,3) = 333
    cells(5,2) = "JKL" : cells(5,3) = 444

    range(cells(2,5), cells(3,8)).formulaArray = "= transpose(R2C2:R5C3)"

  '
  ' Align text with numbers:
  '
    range(cells(2,5), cells(2,8)).horizontalAlignment = xlRight

end sub ' }
