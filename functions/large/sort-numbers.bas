option explicit

sub main() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Fill some data
 '
    cells(1,1) = "foo" : cells(1,2) = "bar": cells(1,3) =  1
    cells(2,1) = "bar" : cells(2,2) = "baz": cells(2,3) =  4
    cells(3,1) = "foo" : cells(3,2) = "foo": cells(3,3) =  2
    cells(4,1) = "baz" : cells(4,2) = "foo": cells(4,3) = 17
    cells(5,1) = "baz" : cells(5,2) = "bar": cells(5,3) =  8
    cells(6,1) = "bar" : cells(6,2) = "foo": cells(6,3) = 18
    cells(7,1) = "bar" : cells(7,2) = "bar": cells(7,3) = 22
    cells(8,1) = "bar" : cells(8,2) = "baz": cells(8,3) =  5

  '
  ' Calculate sum of numbers where value in first and second columns are equal:
  '
    cells(9,3).formulaR1C1 = "= sumproduct( -- ( r1c1:r8c1 = r1c2:r8c2 ), r1c3:r8c3 )"

end sub ' }
