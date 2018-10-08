option explicit

sub main() ' {

  '
  ' In order for cell("filename", …) to return a value,
  ' we need to save the current workbook:
  '

'   activeWorkbook.saveAs _
'      fileName   := environ("TEMP") & "\" & "bla.xlsm" , _
'      fileFormat := xlOpenXMLWorkbookMacroEnabled


  '
  ' The following evaluates to something like
  '   C:\Users\REN~1\AppData\Local\Temp\[bla.xlsm]Sheet1
  '
    cells(2, 2).formula = "=cell(""filename"", a1)"


end sub ' }
