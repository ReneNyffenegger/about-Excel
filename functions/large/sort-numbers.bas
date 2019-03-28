option explicit

sub main() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

    randomize(700828)
    dim r as long
    for r = 1 to 10
        cells(r, 1)             =   cLng(rnd(-r) * 900 + 100)
        cells(r, 2).formulaR1C1 = "=large(r1c1:r10c1, " &  r &")"
    next r

end sub ' }
