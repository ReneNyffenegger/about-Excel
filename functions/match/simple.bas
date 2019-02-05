option explicit

sub main() ' {

    testdata

    cells(7, 1).formula = "=match(""Baz"", a1:a5)"

end sub ' }

sub testdata() ' {

  '
  ' Clear testdata from previous run
  '
    activeSheet.cells.clearContents

    cells(1, 1) = "Foo"
    cells(2, 1) = "Bar"
    cells(3, 1) = "Baz"
    cells(4, 1) = "The other"
    cells(5, 1) = "The same"

end sub ' }
