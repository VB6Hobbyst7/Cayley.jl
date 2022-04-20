
function test_shockmodel()
    model = XVA.buildmodel(1)
    shockmodel(model, 1.1, 1)
end

@memoize LRU{Tuple{Any,Any,Any},Any}(maxsize = 10) function shockmodel(model::AbstractDict,
    fxshock::Real, fxvolshock::Real)::AbstractDict

    @info("In shockmodel with fxshock = $fxshock fxvolshock = $fxvolshock, this call not in memoize cache")
    shockedmodel = deepcopy(model)
    for c in setdiff(shockedmodel["Currencies"], [shockedmodel["Numeraire"]])
        shockedmodel["spot"][c] /= fxshock
        if fxvolshock != 1
            newfxvols = deepcopy(shockedmodel["fxatmvol"][c])
            newfxvols["data"] *= fxvolshock
            XVA.addfxvol_xccyhw!(shockedmodel, c, convert(Dict, newfxvols))
        end
    end
    shockedmodel
end