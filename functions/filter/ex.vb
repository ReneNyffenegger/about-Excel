option explicit

sub main() ' {
    testdata

    cells(1,6)              = "bar"

    dim data     as string : data     = "R2C3:R9C4"
    dim criteria as string : criteria = "R2C2:R9C2=R1C6"
    dim formula  as string : formula  = "=filter(" & data & "," & criteria & "," & """?"")"

    cells(2,6).formula2R1C1 = formula

    activesheet.usedRange.columns.autofit

end sub ' }


sub testdata() ' {

    range(cells(2,2), cells(2,4)) = array("foo", 1, "one"  )
    range(cells(3,2), cells(3,4)) = array("bar", 2, "two"  )
    range(cells(4,2), cells(4,4)) = array("foo", 3, "three")
    range(cells(5,2), cells(5,4)) = array("baz", 4, "four" )
    range(cells(6,2), cells(6,4)) = array("baz", 5, "five" )
    range(cells(7,2), cells(7,4)) = array("bar", 6, "six"  )
    range(cells(8,2), cells(8,4)) = array("foo", 7, "seven")
    range(cells(9,2), cells(9,4)) = array("bar", 8, "eight")

end sub ' }
