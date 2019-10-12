option explicit

sub main() ' {

    cells(1, 1) = 500
    cells(1, 2).formulaR1C1 = "=dec2bin(rc[-1])"

    cells(2, 1) = 512
    cells(2, 2).formulaR1C1 = "=dec2bin(rc[-1])"

    cells(3, 1) = 123456
    cells(3, 2).formulaR1C1 = "=dec2bin(int(rc[-1]/512), 9) & " & _
                               "dec2bin(mod(rc[-1],512), 9)"

    with columns(2)

       .horizontalAlignment = xlRight
       .select
        with selection.font
             .name = "Courier New"
             .bold = true
        end with

       .autoFit
    end with

    cells(5, 5).select

end sub ' }
