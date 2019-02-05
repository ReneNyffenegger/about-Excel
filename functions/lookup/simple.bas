option explicit

sub main() ' {

    testdata

    cells(14, 1) = "September"
    cells(14, 2).formula = "=lookup(a14, a1:a12, b1:b12)"

    columns(1).autoFit


end sub ' }

sub testdata() ' {

    activeSheet.cells.clearContents

    cells( 1, 1) = "January"   : cells( 1, 2) =  10.1
    cells( 2, 1) = "February"  : cells( 2, 2) =   9.3
    cells( 3, 1) = "March"     : cells( 3, 2) =  11.5
    cells( 4, 1) = "April"     : cells( 4, 2) =   8.6
    cells( 5, 1) = "May"       : cells( 5, 2) =  10.2
    cells( 6, 1) = "June"      : cells( 6, 2) =   9.8
    cells( 7, 1) = "July"      : cells( 7, 2) =  11.6
    cells( 8, 1) = "August"    : cells( 8, 2) =  11.8
    cells( 9, 1) = "September" : cells( 9, 2) =  10.7
    cells(10, 1) = "October"   : cells(10, 2) =  10.8
    cells(11, 1) = "November"  : cells(11, 2) =  10.1
    cells(12, 1) = "December"  : cells(12, 2) =   9.9

end sub ' }
