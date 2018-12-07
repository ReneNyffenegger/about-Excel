option explicit


sub main() ' {

  '
  ' Create test data...
  '
    testData

  ' -----------------------------------------------------------------------------------------
  '
  ' and turn data into an excel table:
  '
    dim excelTable as listObject
    set excelTable = activeSheet.listObjects.add(xlSrcRange, range(cells(2, 2), cells(6, 4)))
    excelTable.name = "tq84Tab"

  ' -----------------------------------------------------------------------------------------
  '
  ' Use rows(…) and columns(…) to show the number of rows and columns
  ' in the created data table:
  '
    cells(8,2) = "Total rows"   : cells(8, 4) = "=rows(tq84Tab)"
    cells(9,2) = "Total columns": cells(9, 4) = "=columns(tq84Tab)"

end sub ' }

sub testData() ' {

    cells(2, 2) = "colOne": cells(2, 3) = "colTwo": cells(2, 4) = "colThree"
    cells(3, 2) = "abc"   : cells(3, 3) =      22 : cells(3, 4) =      4024
    cells(4, 2) = "def"   : cells(4, 3) =      18 : cells(4, 4) =      3218
    cells(5, 2) = "ghi"   : cells(5, 3) =      21 : cells(5, 4) =      2973
    cells(6, 2) = "jkl"   : cells(6, 3) =      24 : cells(6, 4) =      3831

end sub ' }
