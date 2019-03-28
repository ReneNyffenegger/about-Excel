option explicit

sub main() ' {

    activeSheet.usedRange.clearContents

  '
  ' One dimensional horizontal array constant
  ' Note: the elements are separated by commas, not semicolons.
  '
    range(cells(1,2), cells(1, 6)).formulaArray = "={ ""H1"", ""H2"", ""H3"", ""H4"", ""H5"" }"

  '
  ' One dimensional vertical array constant
  ' Note: the elements are separated by semicolons, not commas.
  '
    range(cells(2,1), cells(6, 1)).formulaArray = "={ ""V1""; ""V2""; ""V3""; ""V4""; ""V5"" }"

  '
  ' Two dimensional array constant
  ' Note: rows are separated by semicolons, other elements by commas.
  '
    range(cells(3,3), cells(5,5)).formulaArray = "={ ""3/3"",""4/3"",""5/3"" ; ""3/4"",""4/4"",""5/4"" ; ""3/5"",""4/5"",""5/5"" }"

end sub ' }
