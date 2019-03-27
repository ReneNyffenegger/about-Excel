option explicit

sub main() ' {

    testData

  '
  ' Count values of foo, bar and baz:
  '
    range(cells(1,2), cells(1, 4)).formulaArray = "=countifs(r1c1:r12c1, {""foo"", ""bar"", ""baz""} )"

end sub ' }

sub testData() ' {
 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Insert a few values
 '
    cells( 2, 1) = "bar"
    cells( 3, 1) = "bla"
    cells( 4, 1) = "bar"
    cells( 5, 1) = "foo"
    cells( 6, 1) = "xyz"
    cells( 7, 1) = "baz"
    cells( 8, 1) = "bar"
    cells( 9, 1) = "abc"
    cells(10, 1) = "foo"
    cells(11, 1) = "xxx"
    cells(12, 1) = "bar"

end sub ' }
