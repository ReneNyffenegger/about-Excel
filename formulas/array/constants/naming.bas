option explicit

sub main() ' {

 '
 '  Clear active sheet's data:
 '
    activeSheet.usedRange.clearContents

 '
 '  Name an array constant:
 '
    activeWorkbook.names.add "days", refersTo := "={""Sun"", ""Mon"", ""Tue"", ""Wed"", ""Thu"", ""Fri"", ""Sat""}"

 '
 '  Insert an array formula referring to the named constant:
 '
    range(cells(1,1), cells(1,7)).formulaArray = "=days"

end sub ' }
