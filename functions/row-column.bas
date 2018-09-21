option explicit

sub main() ' {

    activeSheet.cells.clearContents

    cells(2, 2).formula = "=row()"       '  2
    cells(3, 2).formula = "=column()"    '  2

    cells(4, 2).formula = "=row(j6)"     '  6
    cells(5, 2).formula = "=column(j6)"  ' 10

end sub ' }
