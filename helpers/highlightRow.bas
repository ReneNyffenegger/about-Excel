option explicit

declare function GetCursorPos Lib "user32" (lpPoint as POINTAPI) as long

type POINTAPI
     x as long
     Y as long
end  type

global lastHighligtedRow as long

sub highlightRow

    dim cursorPos as POINTAPI
    GetCursorPos cursorPos

    dim rng as range

    set rng = activeWindow.rangeFromPoint(cursorPos.x, cursorPos.Y)

    if lastHighligtedRow > 0 then
       rows(lastHighligtedRow).interior.colorIndex = 0 ' no fill?
    end if

    rows(rng.row).interior.color = rgb(230, 230, 230)
    lastHighligtedRow = rng.row

end sub
