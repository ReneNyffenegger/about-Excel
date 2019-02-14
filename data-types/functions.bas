option explicit

global curRow as long
global curCol as long

sub main() ' {

    curRow = 2

    enterText "Some Text"
    enterText  42
    enterText "2019-02-17"
    enterText "TRUE"
    enterText "'TRUE"
    curRow = curRow + 1                     ' Leave empty
    enterText "=NA()"
    enterText "=inexistingFunc()"
    enterText "=7/0"
    enterText "=d2"
    enterText "="""""

    curCol = 3

    addFormula "type"       , false
    addFormula "isText"     , true
    addFormula "isNonText"  , true
    addFormula "isLogical"  , true
    addFormula "isErr"      , true
    addFormula "isNumber"   , true
'   addFormula "isEmpty"    , true
    addFormula "isNA"       , true
    addFormula "isFormula"  , true
    addFormula "isBlank"    , true
    addFormula "isRef"      , true

    dim c as long: for c = 2 to curCol
        columns(c).autofit
    next c

end sub ' }

sub enterText(text as string) ' {

    curRow = curRow + 1

    cells(curRow, 2) = "'" & text
    cells(curRow, 3) =       text

end sub ' }

sub addFormula(formula as string, showTrue as boolean) ' {

    curCol = curCol + 1

    cells(2, curCol) = formula

    formula = formula & "(RC3)"

    if showTrue then
       formula = "if(" & formula & ", unichar(10003) , """")"  ' unichar(…) to insert a character from the full unicode range.
    end if

    range(cells(3, curCol), cells(curRow, curCol)).FormulaR1C1 = "=" & formula

end sub ' }
