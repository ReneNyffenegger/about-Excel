option explicit

sub main() ' {

    activeSheet.cells.clearContents

    cells(2, 2).formula = "=address(10, 5)"            ' Address of 10th row, 5th column:                                     $E$10
    cells(3, 2).formula = "=address(row(), column())"  ' Address of THIS cell:                                                $B$3
    cells(4, 2).formula = "=address(10, 5, 4)"         ' Relative address of 10th row, 5th column:                            E10
    cells(5, 2).formula = "=address(10, 5, 4, 0)"      ' Relative address of 10th row, 5th column, R1C1 reference style R1C1: R[10]C[5]

end sub ' }
