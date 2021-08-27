add-type -path 'C:\Program Files (x86)\Microsoft Office\Office16\DCF\Microsoft.Office.Interop.Excel.dll'

function export-CSV {

   param (
       [string] $dir,       # Must not end in a (back-)slash
       [string] $sheetName
   )

   $dir = resolve-path $dir

   $xls = new-object Microsoft.Office.Interop.Excel.ApplicationClass
   $xls.visible       = $true
   $xls.displayAlerts = $false # don't display message box when exporting a worksheet as CSV.


   foreach ($wbFile in get-childItem $dir\*.xls*) {
      $wb = $xls.workbooks.open($wbFile.fullName)

      try {
         $sh = $wb.sheets($sheetName)
      }
      catch {
         if ($_.exception.message -match 'Invalid index.') {
            write-host "Expected sheet not found in $($wb.name)"
            $wb.close()
            continue
         }
         throw $_
      }
      $sh.select()
      $csvFile = "$dir\$($wbFile.basename).csv"
      $wb.saveAs($csvFile, 6, $false)
      write-host "$csvFile was saved"
      $wb.close()
   }

   $xls.quit()
}

export-CSV "~/ZZZ/Excel/Export-CSV" 'Approval_Logs'

#
# Make sure Excel process is stopped/terminated when
# sheets are exported
# The following two methods Collect() and WaitForPendingFinalizers()
#   must be called in the scope that called the scope
#   where interop is used
#   https://stackoverflow.com/a/25135685/180275
#
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
