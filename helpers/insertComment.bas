'
'  Assign Alt-C to the macro:
'     application.onKey "%c", "insertComment"
'
sub insertComment() ' {

    dim cmt as comment

    set cmt = selection.addComment

    cmt.visible = true

    cmt.text text := ""
end sub ' }
