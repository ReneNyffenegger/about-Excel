option explicit

sub main() ' {

    testData
    removeOutliers

end sub ' }

sub testData() ' {

    cells(1, 1) =  28.6
    cells(2, 1) =  57.8
    cells(3, 1) =  53.6
    cells(4, 1) =  32.3
    cells(5, 1) = 123.9 ' Outlier!
    cells(6, 1) =  45.1
    cells(7, 1) =  30.4

    cells(9, 1).formulaR1C1 = "=average(r1c1:r7c1)"

end sub ' }

sub removeOutliers() ' {

    range(cells(1, 2), cells(7, 2)).formulaR1C1 = "=if(rc[-1]>100, """", rc[-1])"

    cells(9, 2).formulaR1C1 = "=average(r1c2:r7c2)"

end sub ' }
