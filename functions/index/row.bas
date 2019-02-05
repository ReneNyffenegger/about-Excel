option explicit

sub main() ' {

    testData

    range(cells(1,4), cells(5,4)).formulaR1C1 = "=index(r1c1:r5c1,rc[-1])"
end sub ' }

sub testData() ' {

 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

    cells(1, 1) = "one"
    cells(2, 1) = "two"
    cells(3, 1) = "three"
    cells(4, 1) = "four"
    cells(5, 1) = "five"

    cells(1, 3) = 4
    cells(2, 3) = 2
    cells(3, 3) = 4
    cells(4, 3) = 1
    cells(5, 3) = 2

end sub ' }
