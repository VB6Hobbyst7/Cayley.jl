
function debug_pfeprofilefortrades()
    model = XVA.buildmodel("c:/Temp/Cayley/CayleyMarket.json", true, true)
    trades = XVA.loadtrades("c:/Temp/Cayley/CayleyTrades.csv")
    numsims = 255
    timegap = 1 / 12
    timeend = 5
    pfepercentile = 0.95
    isshortfall = false
    usethreads = !Sys.iswindows()
    reportcurrency = "USD"
    calcbyproduct = true
    pfeprofilefortrades(model, numsims, trades, timegap, timeend, pfepercentile,
        isshortfall, usethreads, reportcurrency, calcbyproduct)
end

function test_pfeprofilefortrades()
    model = XVA.loadmodel(4)
    trades = XVA.loadtrades(4)
    control = XVA.loadcontrol(4)
    control["PartitionByTrade"] = false
    numsims = control["NumSims"]
    timegap = control["TimeGap"]
    timeend = model["TStar"]
    pfepercentile = control["PFEPercentile"]
    isshortfall = false
    usethreads = !Sys.iswindows()
    reportcurrency = model["Numeraire"]
    calcbyproduct = false
    res1 = pfeprofilefortrades(model, numsims, trades, timegap, timeend, pfepercentile,
        isshortfall, usethreads, reportcurrency, calcbyproduct)

    #=test against xva_main, which calls `pfecore`, from which `pfeprofilefortrades` is a 
     copy-paste-edit.=#
    res2 = xva_main(4, !Sys.iswindows())["PartyExposures"]["BARC_GB_LON"]
    res3 = hcat(timetoexceldate.(res2["Time"], model["AnchorDate"]),
        res2["Time"], res2["PFE"])

    #also test that passing in the file name duplicated doubles the PFE
    res4 = pfeprofilefortrades(model, numsims, fill(pathfromsetno(4, "trades.csv"), 2),
        timegap, timeend, pfepercentile, isshortfall, usethreads, reportcurrency, 
        calcbyproduct)

    (res3 ≈ res1) && (res4 .* [1 1 0.5] ≈ res1)

end

"""
    pfeprofilefortrades(model::AbstractDict, numsims::Int64, trades::DataFrame,
    timegap::Real, timeend::Real, pfepercentile::Float64, isshortfall::Bool, 
    usethreads::Bool, reportcurrency::String, calcbyproduct::Bool)

This function designed to be called from the "Cayley" project. There is (rather shameful)
copy-paste of code between function `pfecore` and this function. Differences include that
here we cope with PFE being defined in terms of expected shortfall (if `isshortfall ==
true`) and the return is simply a three column array (columns being excel-style date, time,
PFE) rather than a dictionary including data not relevant for Cayley. Also support a 
`reportcurrency` that may be different from `model["Numeraire"]`.

If `whattoreturn` is `arrayvals` (rather than the default of `pfe`) then the return is an
array of values where each row of the array is a "path" for the trades' value.
"""
function pfeprofilefortrades(model::AbstractDict, numsims::Int64, trades::DataFrame,
    timegap::Real, timeend::Real, pfepercentile::Float64, isshortfall::Bool,
    usethreads::Bool, reportcurrency::String, calcbyproduct::Bool)

    anchordate::Date = model["AnchorDate"]
    if !calcbyproduct
        arrayvals = arrayvalsfortrades(model, numsims, trades, timegap, timeend, usethreads,
                                       reportcurrency)
        pfeprofilefromarrayvals(arrayvals, pfepercentile, isshortfall, timegap, timeend,
                                anchordate)
    else
        trades1, trades2 = splitbyassetclass(trades)
        arrayvals1 = arrayvalsfortrades(model, numsims, trades1, timegap, timeend,
                                        usethreads, reportcurrency)
        arrayvals2 = arrayvalsfortrades(model, numsims, trades2, timegap, timeend,
                                        usethreads, reportcurrency)
        profile1 = pfeprofilefromarrayvals(arrayvals1, pfepercentile, isshortfall, timegap,
                                           timeend, anchordate)
        profile2 = pfeprofilefromarrayvals(arrayvals2, pfepercentile, isshortfall, timegap,
                                           timeend, anchordate)
        addprofiles(profile1, profile2)
    end

end

function addprofiles(x, y)
    size(x, 2) == 3 || throw("Expected PFEProfile to have 3 columns but got $(size(x,2))")
    size(y, 2) == 3 || throw("Expected PFEProfile to have 3 columns but got $(size(y,2))")
    x[:, 1:2] == y[:, 1:2] || throw("Cannot add PFE Profiles as their left two columns don't match")
    hcat(x[:, 1:2], x[:, 3] .+ y[:, 3])
end


@memoize LRU{Tuple{Any,Any,Any},Any}(maxsize = 10) function producestatewrapper(model, adjustedtstrip, numsims)
    XVA.producestateforpfe_xccyhw(model, adjustedtstrip, numsims)
end


@memoize LRU{Tuple{Any,Any,Any,Any,Any,Any,Any},Any}(maxsize = 10) function arrayvalsfortrades(model::AbstractDict, numsims::Int64, trades::DataFrame,
    timegap::Real, timeend::Real, usethreads::Bool, reportcurrency::String)

    @info "In arrayvalsfortrade, this call not in memoize cache"
    numtrades = nrow(trades)

    anchordate::Date = model["AnchorDate"]
    maxtradelife = XVA.tradelife(trades, anchordate)
    tstrip = timegap:timegap:timeend
    numtimes = length(tstrip)
    numzeros = sum(tstrip .> maxtradelife) # The number of elements of tstrip for which we know that every trade
    # will have zero value since every trade will have matured.
    #adjustedtimeend::Float64 = maximum(ifelse.(tstrip .<= maxtradelife, tstrip, 0.0))

    if numzeros < numtimes
        # There exists at least one trade with a life greater than timegap
        adjustedtstrip::Vector{Float64} = filter(x -> x <= maxtradelife, tstrip)
        #state = XVA.producestateforpfe_xccyhw(model, adjustedtstrip, numsims)
        @time state = producestatewrapper(model, adjustedtstrip, numsims)

        function valuetradesforpfe(state, model::AbstractDict, trades::DataFrame,
            fundingspread::Real, usethreads::Bool, numtimes::Integer,
            reportcurrency::String)::Array{Float64,2}

            v = XVA.valuetrades(state, model, trades, fundingspread, usethreads)
            if reportcurrency != model["Numeraire"]
                v ./= XVA.BtStDt_xccyhw(state, model, reportcurrency)
            end
            v = reshape(v, numsims, length(adjustedtstrip))
            if size(v, 2) < numtimes
                v = hcat(v, fill(0.0, numsims, numtimes - size(v, 2)))
            end
            v
        end

        if numtrades > 0
            arrayvals = valuetradesforpfe(state, model, trades, 0.0, usethreads,
                numtimes, reportcurrency)
        end
    else
        # No trades have life greater than timegap
        if numtrades > 0
            arrayvals = fill(0.0, numsims, numtimes)
        end
    end

    if numtrades > 0
        initialvalue = XVA.valuetrades(nothing, model, trades, 0.0, usethreads)[1]
        if reportcurrency != model["Numeraire"]
            initialvalue = initialvalue / XVA.BtStDt_xccyhw(nothing, model, reportcurrency)[1]
        end
    else
        initialvalue = 0.0
    end
    # Tack on the time zero state
    if numtrades > 0
        arrayvals = hcat(repeat(initialvalue, numsims), arrayvals)
    else
        arrayvals = fill(0.0, numsims, numtimes + 1)
    end

    return (arrayvals)
end

"""
    profilevectortothreecols(profilevector,timegap,timeend,anchordate::Date)
When we pass the PFE profile back to Excel, we tack on two additional columns.
"""
function profilevectortothreecols(profilevector, timegap, timeend, anchordate::Date)
    time = 0:timegap:timeend
    dates = XVA.timetoexceldate.(time, anchordate)
    return (hcat(dates, time, profilevector))
end

function pfeprofilefromarrayvals(arrayvals, pfepercentile, isshortfall::Bool, timegap, timeend, anchordate::Date)
    pfe = quantileorshortfall(arrayvals, pfepercentile, isshortfall)
    profilevectortothreecols(pfe, timegap, timeend, anchordate)
end

"""
    function quantileorshortfall(v::AbstractVector, p, isshortfall::Bool)
Returns quantile or, if `isshortfall == false` the so-called shortfall, defined as
conditional mean of `v` conditional on `v` being "outside" quantile `p` i.e. above quantile
p if p > .5 or below quantile p if p <= 0.5.
"""
function quantileorshortfall(v::AbstractVector, p, isshortfall::Bool)
    q = quantile(v, p)
    if isshortfall
        if p > 0.5
            return (mean(filter(x -> x >= q, v)))
        else
            return (mean(filter(x -> x <= q, v)))
        end
    else
        return (q)
    end
end

"""
    quantileorshortfall(v::AbstractMatrix,p,isshortfall::Bool)
quantileorshortfall operates column-wise on a matrix.
"""
function quantileorshortfall(v::AbstractMatrix, p, isshortfall::Bool)
    [quantileorshortfall(view(v, :, i), p, isshortfall) for i = 1:size(v, 2)]
end

"""
    pfeprofilefortrades(model::AbstractDict, numsims::Int64,
        tradefiles::Union{String,Vector{String}}, timegap::Real, timeend::Real,
        pfepercentile::Float64, isshortfall::Bool, usethreads::Bool, reportcurrency::String,
        calcbyproduct::Bool)

Method which accepts trades in one or more CSV files.
"""
function pfeprofilefortrades(model::AbstractDict, numsims::Int64,
    tradefiles::Union{String,Vector{String}}, timegap::Real, timeend::Real,
    pfepercentile::Float64, isshortfall::Bool, usethreads::Bool, reportcurrency::String,
    calcbyproduct::Bool)

    if tradefiles isa String
        trades = XVA.loadtrades(tradefiles)
    else
        trades = XVA.loadtrades(tradefiles[1])
        for i = 2:length(tradefiles)
            append!(trades, XVA.loadtrades(tradefiles[i]), cols = :union)
        end
    end
    pfeprofilefortrades(model, numsims, trades, timegap, timeend, pfepercentile,
        isshortfall, usethreads, reportcurrency, calcbyproduct)
end

function test_valueportfolio(usethreads::Bool, throwonerror::Bool, returnvector::Bool)
    model = buildmodel("c:/Temp/Cayley/CayleyMarket.json", true, true)
    tradefile = "C:/Temp/Cayley/CayleyTradesForTradesViewer.csv"
    trades = XVA.loadtrades(tradefile)
    valueportfolio(model, trades, 0.0, "GBP", throwonerror, returnvector, usethreads)
end


function valueportfolio(model::AbstractDict, tradefile::String, fundingspread,
    reportcurrency::String, throwonerror::Bool, returnvector::Bool, usethreads::Bool)

    trades = XVA.loadtrades(tradefile)

    valueportfolio(model, trades, fundingspread, reportcurrency,
        throwonerror, returnvector, usethreads)
end

"""
    valueportfolio(model::AbstractDict, trades::DataFrame, fundingspread::Real,
        reportcurrency::String, usethreads::Bool)::Vector{Union{String,Float64}}
Values a portfolio (a DataFrame) of trades, with state = nothing, i.e. in the t = 0 initial
state of `model`. Return maybe either a vector (ordered as per `trades`) or a single number
according to `returnvector` argument. So the function is definitely not type stable!
"""
function valueportfolio(model::AbstractDict, trades::DataFrame, fundingspread::Real,
    reportcurrency::String, throwonerror::Bool, returnvector::Bool,
    usethreads::Bool)

    state = nothing
    fxrate = XVA.BtStDt_xccyhw(state, model, reportcurrency)[1]
    ntrades = nrow(trades)

    if throwonerror
        pvvector = Vector{Float64}(undef, ntrades)
    else
        pvvector = Vector{Union{String,Float64}}(undef, ntrades)
    end

    pvtotal = 0.0
    if usethreads
        tasks = Array{Task}(undef, ntrades)
        i = 0
        for trade in eachrow(trades)
            i += 1
            tasks[i] = @spawn XVA.valuetrade(state, model, trade, fundingspread)
        end
        for i = 1:ntrades
            try
                pvvector[i] = fetch(tasks[i])[1] / fxrate
                pvtotal += pvvector[i]
            catch e
                if throwonerror
                    rethrow(e)
                else
                    pvvector[i] = "$e"
                end
            end
        end
    else
        i = 0
        for trade in eachrow(trades)
            try
                i += 1
                pvvector[i] = XVA.valuetrade(state, model, trade, fundingspread)[1] / fxrate
                pvtotal += pvvector[i]
            catch e
                if throwonerror
                    rethrow(e)
                else
                    pvvector[i] = "$e"
                end
            end
        end
    end
    if returnvector
        return (pvvector)
    else
        return (pvtotal)
    end
end
