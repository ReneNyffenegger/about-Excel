'
'  Assign Alt-C to the macro:
'     application.onKey "%c", "insertComment"
'
sub insertComment() ' {

    dim cmt as comment

    set cmt = selection.addComment

    cmt.visible = true
    cmt.text text := ""

  '
  ' Select the comment so that it's possible
  ' to type-write into it right away:
  '
    cmt.shape.select true

end sub ' }
