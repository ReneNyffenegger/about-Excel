global lastHighlightedRow as long

sub highlightRowOfSelectedCell()

    if lastHighlightedRow > 0 Then
       rows(lastHighlightedRow).Interior.ColorIndex = 0 ' no fill?
    end if

    rows(selection.Row).interior.color = rgb(230, 230, 230)
    lastHighlightedRow = selection.Row

end sub
