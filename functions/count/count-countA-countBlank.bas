option explicit

sub main() ' {
  '
  ' Data
  '
    cells( 3, 2).resize(1,6).value          = array(    0 ,    1 ,    2 ,      3 ,     4 ,     5 )
    cells( 4, 2).resize(1,6).value          = array("zero", "one", "two", "three", "four", "five")
    cells( 5, 2).resize(1,6).value          = array(""    , " "  , ""   , ""     , chr(9), ""    )
    cells( 6, 2).resize(1,6).value          = array("=7/0", ""   , ""   , "=na()", ""    , ""    )
  '
  ' Make data visually stand out:
  '
    cells( 3, 2).resize(4,6).interior.color = rgb(240, 210, 110)

  '
  ' Formulas
  '
    cells( 2, 9).resize(1,3).value          = array( "count"               ,   "countA"               ,  "countBlank"               )
    cells( 3, 9).resize(1,3).formulaR1C1    = array("=count(r[0]c2:r[0]c7)",  "=countA(r[0]c2:r[0]c7)", "=countBlank(r[0]c2:r[0]c7)")
    cells( 4, 9).resize(1,3).formulaR1C1    = array("=count(r[0]c2:r[0]c7)",  "=countA(r[0]c2:r[0]c7)", "=countBlank(r[0]c2:r[0]c7)")
    cells( 5, 9).resize(1,3).formulaR1C1    = array("=count(r[0]c2:r[0]c7)",  "=countA(r[0]c2:r[0]c7)", "=countBlank(r[0]c2:r[0]c7)")
    cells( 6, 9).resize(1,3).formulaR1C1    = array("=count(r[0]c2:r[0]c7)",  "=countA(r[0]c2:r[0]c7)", "=countBlank(r[0]c2:r[0]c7)")

    activeSheet.usedRange.columns.autofit

end sub ' }
