option explicit

sub main() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents
    cells(1,1) = 5 : cells(1,2) = 3
    cells(2,1) = 7 : cells(2,2) = 2
    cells(3,1) = 4 : cells(3,2) = 6
    cells(4,1) = 1 : cells(4,2) = 9

    cells(5,3).formulaR1C1 = "= sumproduct( r1c1:r4c1 , r1c2:r4c2 )"

end sub ' }
