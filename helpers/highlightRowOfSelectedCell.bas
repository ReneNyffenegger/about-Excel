global lastHighligtedRow as long

sub highlightRowOfSelectedCell()

    if lastHighligtedRow > 0 Then
       rows(lastHighligtedRow).Interior.ColorIndex = 0 ' no fill?
    end if

    rows(selection.Row).interior.color = rgb(230, 230, 230)
    lastHighligtedRow = selection.Row

end sub
