option explicit

sub main() ' {

    cells(2,1) =  18.3
    cells(3,1) ="'18.3"
    cells(4,1) ="text"
    cells(5,1) = true
    cells(6,1) = false

    cells(1,2) = "value"       : range(cells(2,2), cells(6,2)).formulaR1C1 = "=value(rc1)"
    cells(1,3) = "numberValue" : range(cells(2,3), cells(6,3)).formulaR1C1 = "=numberValue(rc1)"
    cells(1,4) = "n"           : range(cells(2,4), cells(6,4)).formulaR1C1 = "=n(rc1)"
    cells(1,5) = "t"           : range(cells(2,5), cells(6,5)).formulaR1C1 = "=t(rc1)"
    cells(1,6) = "text"        : range(cells(2,6), cells(6,6)).formulaR1C1 = "=text(rc1, ""000.00"")"

    range(cells(1,2), cells(1,6)).font.bold = true

    activeSheet.usedRange.columns.autofit

end sub ' }
