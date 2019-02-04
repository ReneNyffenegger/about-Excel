option explicit

sub deleteRowsWithCriteria(searchRange as range, formula as string) ' {

    dim cell       as range
    dim foundCells as range

    for each cell in searchRange ' {

        dim cellValue as string

        if typeName(cell.value) = "String" then
           cellValue = """" & cell.value & """"
        else
           cellValue =        cell.value
        end if

        dim formula_ as string
        formula_ = replace(formula, "@", cellValue)
    '   debug.print(formula_)

        if application.evaluate(formula_) then
           if foundCells is nothing then
              set foundCells = cell
           else
              set foundCells = application.union(foundCells, cell)
           end if
        end if

    next cell ' }

    if not foundCells is nothing then
       foundCells.entireRow.delete
    end if

end sub ' }
