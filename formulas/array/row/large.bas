option explicit

sub main() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

    rnd(-280870)
    randomize(1)
    dim r as long
    for r = 1 to 10
        cells(r, 1) = cLng(rnd(r) * 900 + 100)
    next r

    range(cells(1,2), cells(10,2)).formulaArray = "=large(r1c1:r10c1, row(r1:r10))"

end sub ' }
