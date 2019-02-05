option explicit

sub main() ' {

    testdata

    cells(2, 8).formula = "=offset(b2, 3, 2)
    cells(3, 8).formula = "=offset(c3, 3, 2)

end sub ' }

sub testdata() ' {

  '
  ' Clear testdata from previous run
  '
    activeSheet.cells.clearContents

  '
  ' Fill two dimensional array
  '
    dim x, y as long
    for x = 1 to 5: for y = 1 to 5
        cells(y + 1, x + 1).numberFormat = "@"
        cells(y + 1, x + 1) = y & " - " & x
    next y: next x

end sub ' }
