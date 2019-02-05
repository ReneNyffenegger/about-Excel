option explicit

sub main() ' {

    cells(2, 2).formula = "=index({""one""; ""two""; ""three""; ""four""; ""five""}, a1)"
    cells(1, 1).value   =  4

  '
  ' It seems that index with a hardcoded array does not
  ' get updated automatically.
  '
    activeSheet.calculate

end sub ' }
