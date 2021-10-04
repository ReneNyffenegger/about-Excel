option explicit

sub main() ' {

    dim rng as range
    set rng = range(cells(1,1), cells(1,3))

    rng.numberFormat = "@"
    rng              = array("01", "+2", "1E4")

    dim cel as range
    for each cel in rng
        cel.errors(xlNumberAsText).ignore = true
    next cel

end sub ' }
