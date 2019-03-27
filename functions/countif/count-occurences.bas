option explicit

sub main() ' {

    testData

    range(cells(1,2), cells(6,2)).formulaR1C1 = "=countif(r1c1:r6c1,rc[-1])"

end sub ' }

sub testData() ' {

 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Insert a few values
 '
    cells(1, 1) = "foo"
    cells(2, 1) = "bar"
    cells(3, 1) = "foo"
    cells(4, 1) = "baz"
    cells(5, 1) = "baz"
    cells(6, 1) = "foo"
end sub ' }
