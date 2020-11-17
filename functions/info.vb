option explicit

sub main() ' {


  '
  ' The following evaluates to something like
  '   C:\Users\REN~1\AppData\Local\Temp\[bla.xlsm]Sheet1
  '
    cells(3, 2) = "directory"
    cells(4, 2) = "numfile"
    cells(5, 2) = "origin"
    cells(6, 2) = "osVersion"
    cells(7, 2) = "recalc"
    cells(8, 2) = "release"
    cells(9, 2) = "system"

    cells(2,3)  = "info(...)"
    range(cells(3,3), cells(9,3)).formulaR1C1 = "=info(RC[-1])"

    cells(3,4) = "cell(""filename"")
    cells(3,5).formula="=cell(""filename"")"

    activeSheet.usedRange.columns.autoFit

  '
  ' In order for info("directory") to return a value,
  ' we need to save the current workbook:
  '
    activeWorkbook.saveAs _
       fileName   := environ("TEMP") & "\" & "info.xlsm" , _
       fileFormat := xlOpenXMLWorkbookMacroEnabled

end sub ' }
