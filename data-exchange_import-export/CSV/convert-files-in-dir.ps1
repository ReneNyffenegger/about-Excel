set-strictMode -version latest

# add-type -path 'C:\Program Files (x86)\Microsoft Office\Office16\DCF\Microsoft.Office.Interop.Excel.dll'
$assembly = [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Excel")

$xls = new-object Microsoft.Office.Interop.Excel.ApplicationClass
$xls.Visible = $true

foreach ($csvFile in get-childItem ~\work\2021-03-29\*.csv) {
   write-host "csv file = $csvFile"

   $wb = $xls.Workbooks.Open($csvFile)

   $xlsmFile = $csvFile -replace '\.[^.]+$', '.xlsm'

   if (test-path $xlsmFile) {
      remove-item $xlsmFile
   }
   $wb.SaveAs(
       $xlsmFile,
     #
     #  Following constant specifies format for .xlsm
     # (Compare with https://renenyffenegger.ch/notes/Microsoft/dot-net/namespaces-classes/Microsoft/Office/Interop/_application_/Constants)
      [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
   )
}

$xls.Quit()
$xls = $null
