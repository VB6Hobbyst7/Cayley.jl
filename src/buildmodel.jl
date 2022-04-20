using XVA
using OrderedCollections


function cayleybuildmodels(filename::String, filenameh::String, fxshock::Real, fxvolshock::Real)
    cayleymodel = cayleybuildmodel(filename)
    cayleymodelshocked = shockmodel(cayleymodel, fxshock, fxvolshock)
    cayleymodelhistoric = cayleybuildmodel(filenameh)
    cayleymodelhistoricshocked = shockmodel(cayleymodelhistoric, fxshock, fxvolshock)
    cayleymodel, cayleymodelshocked, cayleymodelhistoric, cayleymodelhistoricshocked
end

function pack4models(cayleymodel::AbstractDict, cayleymodelshocked::AbstractDict,
    cayleymodelhistoric::AbstractDict, cayleymodelhistoricshocked::AbstractDict,
    fxshock::Real, fxvolshock::Real)

    Dict("cayleymodel" => barebones(cayleymodel, 1, 1),
        "cayleymodelshocked" => barebones(cayleymodelshocked, fxshock, fxvolshock),
        "cayleymodelhistoric" => barebones(cayleymodelhistoric, 1, 1),
        "cayleymodelhistoricshocked" => barebones(cayleymodelhistoricshocked, fxshock, fxvolshock))
end


#cayleymodel = Cayley.cayleybuildmodel("c:/temp/CayleyMarket.json")
#TODO move memoizing into XVA package?
"""
    cayleybuildmodel(filename::String)
Build models required for valuations. Models are global to the Cayley namespace.
This is a port of R code in file CayleyBuildModel.R. See https://github.com/PGS62/XVA-R
TODO add models cayleymodelhistoric and cayleymodelhistoricshocked, but simplest
to change VBA code to export two market data files.
"""
function cayleybuildmodel(filename::String)
    thehash = bytes2hex(SHA.sha256(read(filename)))
    buildmodelwrap(filename, thehash)
end

"""
    buildmodelwrap(filename::String , filehash::String)
Wrap XVA.buildmodel, but with memoizing. Could consider putting the memoizing into XVA?
"""
@memoize LRU{Tuple{Any,Any},Any}(maxsize = 10) function buildmodelwrap(filename::String, filehash::String)
    @info "In buildmodel, this call not in memoize cache"
    XVA.buildmodel(filename, true, true)
end

"""
    barebones(model::AbstractDict)
Returns the small amount of data within a model that we need to keep track of in the VBA code
"""
function barebones(model::AbstractDict, fxshock::Real, fxvolshock::Real)
    res = Dict{String,Any}()
    res["Numeraire"] = model["Numeraire"]
    res["spot"] = model["spot"]
    res["AnchorDate"] = model["AnchorDate"]
    res["EURUSD"] = model["spot"]["EUR"] / model["spot"]["USD"]
    @assert model["Numeraire"] == "EUR" "model Numeraire must be EUR, but it is $(model["Numeraire"])"
    res["EURUSD3YVol"] = XVA.interp1d(convert(Dict, model["fxatmvol"]["USD"]), 3)
    res["EURUSDForwards"] = keyforwardrates(model)
    res["fxshock"] = fxshock
    res["fxvolshock"] = fxvolshock
    res
end

"""
The VBA code needs to know certain EURUSD forward rates in order to be able to execute at-the-money
hedges. 
"""
function keyforwardrates(model::AbstractDict)
    res = OrderedDict{Date,Float64}()
    for i = 0:10
        dte::Date = model["AnchorDate"] + Day(i * 365)
        res[dte] = forwardrate(model, "EUR", "USD", dte)
    end
    res
end

"""

"""
function forwardrate(model::AbstractDict, ccy1::String, ccy2::String, enddate::Date)
    time = XVA.datetotime(enddate, model["AnchorDate"])
    fundingspread = 0.0
    spotrate = model["spot"][ccy1] / model["spot"][ccy2]
    df1 = XVA.BtPtT_xccyhw(nothing, model, ccy1, time, true, fundingspread, true)
    df2 = XVA.BtPtT_xccyhw(nothing, model, ccy2, time, true, fundingspread, true)
    (spotrate*df1/df2)[1]
end

"""
    EURUSDforwardrates(model::AbstractDict, xldates)
Wrapped by VBA function EURUSDForwardRates in the Cayley workbook. Return has two dimensions
with one column, giving forward rates to the passed in xldates (Ints representing dates,
excel style).
"""
function EURUSDforwardrates(model::AbstractDict, xldates::Vector{Int64})
    times = XVA.exceldatetotime.(xldates, model["AnchorDate"])
    fundingspread = 0.0
    ccy1 = "EUR"
    ccy2 = "USD"
    spotrate = model["spot"][ccy1] / model["spot"][ccy2]
    df1 = XVA.BtPtT_xccyhw(nothing, model, ccy1, times, true, fundingspread, true)
    df2 = XVA.BtPtT_xccyhw(nothing, model, ccy2, times, true, fundingspread, true)
    #Note transpose below, so that 
    Matrix((spotrate .* df1 ./ df2)')
end



