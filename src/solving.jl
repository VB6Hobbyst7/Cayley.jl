
function test_creditlimitminustwopfes()
    limittimes = [1, 2, 3, 4, 5, 7, 10]
    creditlimits = [35, 35, 35, 25, 20, 7, 1.5]
    timevector = collect(0:1/12:5)
    creditinterpmethod = "Linear"
    pfepercentile = 0.95
    isshortfall = false
    horizon = 8
    arrayvals1 = 0.5 .* (rand(255, 61) .+
                         transpose(interpcreditline(limittimes, creditlimits, timevector, "Linear")) .- 1)
    arrayvals2 = 0.5 .* (rand(255, 61) .+
                         transpose(interpcreditline(limittimes, creditlimits, timevector, "Linear")) .- 1)
    arrayvals1 = repeat(1/255:1/255:1, 255, 61) .+
                 transpose(interpcreditline(limittimes, creditlimits, timevector, "Linear")) .- 1
    arrayvals1 = repeat(0, 255, 61) .+
                 transpose(interpcreditline(limittimes, creditlimits, timevector, "Linear")) .- 1

    arrayvals2 = nothing
    #   arrayvals1 = 2 .* arrayvals1
    creditlimitminustwopfes(arrayvals1, arrayvals2, limittimes, creditlimits,
        creditinterpmethod, pfepercentile, isshortfall, timevector, horizon)

end


"""
    function creditlimitminustwopfes(arrayvals1::Matrix{Float64},
        arrayvals2::Union{Nothing,Matrix{Float64}},
        limittimes::Nums1, creditlimits::Nums1, creditinterpmethod::String,
        pfepercentile::Float64, isshortfall::Bool, timevector::Nums1, horizon)

Used in the objective function of headroomsolvercore. Calculates the PFE given two sets of 
trade values (arrayvals1 and arrayvals2) and then returns the minimum (over time <= horizon)
of the gap between credit limit and PFE, ignoring any period where the PFE is zero.
"""
function creditlimitminustwopfes(arrayvals1::Matrix{Float64},
    arrayvals2::Union{Nothing,Matrix{Float64}},
    limittimes::Nums1, creditlimits::Nums1, creditinterpmethod::String,
    pfepercentile::Float64, isshortfall::Bool, timevector::Nums1, horizon)

    if !isnothing(arrayvals2)
        size(arrayvals1) == size(arrayvals2) || throw("arrayvals1 and arrayvals2 must be \
                      the same size but got $(size(arrayvals1)) and $(size(arrayvals2))")
    end
    length(timevector) == size(arrayvals1)[2] || throw("length of timevector must be the \
                      same as the number of columns in arrayvals1 but got \
                      $(length(timevector)) and $(size(arrayvals1)[2])")

    creditlimitattimevector = interpcreditline(limittimes, creditlimits, timevector, creditinterpmethod)
    pfe1attimevector = quantileorshortfall(arrayvals1, pfepercentile, isshortfall)

    if isnothing(arrayvals2)
        pfeattimevector = pfe1attimevector
    else
        pfe2attimevector = quantileorshortfall(arrayvals2, pfepercentile, isshortfall)
        pfeattimevector = pfe1attimevector .+ pfe2attimevector
    end

    chooser = (timevector .<= horizon) .&& (pfeattimevector .!= 0) 

    differences = creditlimitattimevector .- pfeattimevector
    if any(chooser)
        res= minimum(differences[chooser])
    else
       res = differences[1]
    end

    return(res)

end

"""
    interpcreditline(limittimes::Nums1, creditlimits::Nums1,
        interptotimes::Nums01, creditinterpmethod::String)

Implements "interpolation" of credit lines. Either "Linear" between knots with flat
extrapolation or else "FlatToRight" i.e. extrapolate flat from each knot to the
next knot.
"""
function interpcreditline(limittimes::Nums1, creditlimits::Nums1,
    interptotimes::Nums01, creditinterpmethod::String)

    length(limittimes) == length(creditlimits) || throw("limittimes and creditlimits must \
            have the same length but got lengths \
            $(length(limittimes)) and $(length(creditlimits))")
    for i = 2:length(limittimes)
        if limittimes[i] <= limittimes[i-1]
            throw("limittimes must be strictly ascending, but element $(i-1) is not less \
                than element $i ($(limittimes[i-1]) and $(limittimes[i]) respectively)")
        end
    end
    if creditinterpmethod == "Linear"
        function intplin(x)
            pos = searchsortedfirst(limittimes, x)
            if pos == 1
                return (Float64(creditlimits[1]))
            elseif pos == length(limittimes) + 1
                return (Float64(creditlimits[end]))
            else
                x1 = limittimes[pos-1]
                x2 = limittimes[pos]
                y1 = creditlimits[pos-1]
                y2 = creditlimits[pos]
                return (((x2 - x) * y1 + (x - x1) * y2) / (x2 - x1))
            end
        end
        return (intplin.(interptotimes))
    elseif creditinterpmethod == "FlatToRight"
        function intpftr(x)
            pos = min(searchsortedfirst(limittimes, x), length(limittimes))
            creditlimits[pos]
        end
        return (intpftr.(interptotimes))
    else
        throw("creditinterpmethod must be 'Linear' or 'FlatToRight' but \
               got '$creditinterpmethod'")
    end
end

#=
TODO I don't think the argument notionalcapapplies is necessary
TODO basecurrency ~ reportcurrency
=#

function test_headroomsolver()
    model = XVA.buildmodel("C:/Users/phili/AppData/Local/Temp/@Cayley/CayleyMarket.json",
        true, true)
    existingtradesfile = "C:/Users/phili/AppData/Local/Temp/@Cayley/ExistingTrades"
    unithedgetradesfile = "C:/Users/phili/AppData/Local/Temp/@Cayley/UnitHedgeTrades"
    basecurrency = "GBP"
    limittimes = [1, 2, 3, 4, 5, 7, 10]
    creditlimits = [350000000, 350000000, 350000000, 250000000, 200000000, 70000000,
        15000000]
    creditinterpmethod = "FlatToRight"
    numsims = 255
    timegap = 1 / 12
    timeend = 5
    pfepercentile = 0.95
    isshortfall = false
    donotionalcap = false
    notionalcapfornewtrades = 0
    horizon = 8
    usethreads = true

    headroomsolver(model, existingtradesfile, unithedgetradesfile,
        basecurrency, limittimes, creditlimits, creditinterpmethod,
        numsims, timegap, timeend, pfepercentile, isshortfall, donotionalcap,
        notionalcapfornewtrades,horizon, usethreads)

end

"""
    function splitbyassetclass(trades::DataFrame)
Partitions a DataFrame of trades into a pair of DataFrames, the first containing the
interest rate trades, and the other the fx trades.
"""
function splitbyassetclass(trades::DataFrame)
    ratesvfs = ["InterestRateSwap", "CrossCurrencySwap", "CapFloor", "Swaption"]
    ratestrades = filter(:valuationfunction => in(ratesvfs), trades)
    fxtrades = filter(:valuationfunction => !in(ratesvfs), trades)
    ratestrades, fxtrades
end

"""
    function headroomsolvercalcbyproducteitherorbasis(model, existingtradesfile::String,
        unithedgetradesfile::String, basecurrency::String, limittimes::Nums1,
        creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
        pfepercentile, isshortfall::Bool, donotionalcap::Bool,
        notionalcapfornewtrades, calcbyproduct::Bool, horizons, usethreads::Bool)

A loop around `headroomsolvercalcbyproduct`. The trades listed in the input
`unithedgetradesfile` file are processed one by one, each being considered as a trade which
we wish to scale in size so as to exhaust credit lines (up to corresponding element of
`horizons`)
"""
function headroomsolvercalcbyproducteitherorbasis(model, existingtradesfile::String,
    unithedgetradesfile::String, basecurrency::String, limittimes::Nums1,
    creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
    pfepercentile, isshortfall::Bool, donotionalcap::Bool,
    notionalcapfornewtrades, calcbyproduct::Bool, horizons, usethreads::Bool)

    unithedgetrades = XVA.loadtrades(unithedgetradesfile)

    loopto = nrow(unithedgetrades)
    multiples = fill(0.0, loopto)

    outerres = Dict{String,Any}
    for i = 1:loopto
        res = headroomsolvercalcbyproduct(model, existingtradesfile,
            unithedgetrades[[i], :], basecurrency, limittimes, creditlimits,
            creditinterpmethod, numsims, timegap, timeend, pfepercentile, isshortfall,
            donotionalcap, notionalcapfornewtrades, calcbyproduct, horizons[i], usethreads)

        multiples[i] = res["multiple"]

        if i == loopto
            outerres = res
            outerres["multiple"] = multiples
        end

    end

    outerres

end

"""
    function headroomsolvercalcbyproduct(model, existingtradesfile::String,
        unithedgetrades::Union{String,DataFrame}, basecurrency::String, limittimes::Nums1,
        creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
        pfepercentile, isshortfall::Bool, donotionalcap::Bool, notionalcapfornewtrades,
        calcbyproduct::Bool, horizon, usethreads::Bool)

Wrap the relatively "clean" function `headroomsolver` to cope with those banks who's 
"Product Credit Limits" are on a "Global Limit & Calculation by Product" basis, i.e. they
calculate a PFE for rates trade, a PFE for fx trades, then add the PFEs and apply a limit to
that sum.
"""
function headroomsolvercalcbyproduct(model, existingtradesfile::String,
    unithedgetrades::Union{String,DataFrame}, basecurrency::String, limittimes::Nums1,
    creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
    pfepercentile, isshortfall::Bool, donotionalcap::Bool, notionalcapfornewtrades,
    calcbyproduct::Bool, horizon, usethreads::Bool)

    if !calcbyproduct
        return (headroomsolver(model, existingtradesfile, unithedgetrades,
            basecurrency, limittimes, creditlimits, creditinterpmethod,
            numsims, timegap, timeend, pfepercentile, isshortfall, donotionalcap,
            notionalcapfornewtrades, horizon, usethreads))
    else
        existingtrades = XVA.loadtrades(existingtradesfile)
        existingratestrades, existingfxtrades = splitbyassetclass(existingtrades)

        if nrow(existingratestrades) == 0
            return (headroomsolver(model, existingtrades, unithedgetrades,
                basecurrency, limittimes, creditlimits, creditinterpmethod,
                numsims, timegap, timeend, pfepercentile, isshortfall, donotionalcap,
                notionalcapfornewtrades, horizon, usethreads))
        end

        arrayvalsert = arrayvalsfortrades(model, numsims, existingratestrades, timegap,
            timeend, usethreads, basecurrency)

        pfeprofileert = quantileorshortfall(arrayvalsert, pfepercentile, isshortfall)
        timevector = 0:timegap:timeend

        #=In the call below, amend the limit structure by the pfe profile of the existing
        rates trades=#
        newlimittimes = timevector
        newcreditlimits = interpcreditline(limittimes, creditlimits, timevector,
            creditinterpmethod) .- pfeprofileert

        result = headroomsolver(model, existingfxtrades, unithedgetrades, basecurrency,
            newlimittimes, newcreditlimits, creditinterpmethod,
            numsims, timegap, timeend, pfepercentile, isshortfall,
            donotionalcap, notionalcapfornewtrades, horizon, usethreads)

        # add on the profile of the existing rates trades    
        pfeprofilewithoutet = result["pfeprofilewithoutet"]
        pfeprofilewithoutet[:, 3] .+= pfeprofileert
        pfeprofilewithet = result["pfeprofilewithet"]
        pfeprofilewithet[:, 3] .+= pfeprofileert

        return (result)

    end
end

"""
function headroomsolver(model, existingtrades::Union{String,DataFrame},
    unithedgetrades::Union{String,DataFrame}, basecurrency::String, limittimes::Nums1,
    creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
    pfepercentile, isshortfall::Bool, donotionalcap::Bool, notionalcapfornewtrades, 
    horizon, usethreads::Bool)

Mid-level function for trade-headroom solving. Calculates the (Monte Carlo) trade paths for 
both the existing trades and the "unit hedge trades" before handing off to
`headroomsolvercore`.
"""
function headroomsolver(model, existingtrades::Union{String,DataFrame},
    unithedgetrades::Union{String,DataFrame}, basecurrency::String, limittimes::Nums1,
    creditlimits::Nums1, creditinterpmethod::String, numsims, timegap, timeend,
    pfepercentile, isshortfall::Bool, donotionalcap::Bool, notionalcapfornewtrades, 
    horizon, usethreads::Bool)

    if existingtrades isa String
        existingtrades = XVA.loadtrades(existingtrades)
    end

    if unithedgetrades isa String
        unithedgetrades = XVA.loadtrades(unithedgetrades)
    end

    arrayvalsexistingtrades = arrayvalsfortrades(model, numsims, existingtrades,
        timegap, timeend, usethreads, basecurrency)

    arrayvalsuht = arrayvalsfortrades(model, numsims, unithedgetrades,
        timegap, timeend, usethreads, basecurrency)

    if donotionalcap
        maxallowedmultiple = notionalcapfornewtrades /
                             sum(unithedgetrades[:, :receivenotional])
    else
        maxallowedmultiple = Inf
    end

    anchordate::Date = model["AnchorDate"]

    headroomsolvercore(arrayvalsexistingtrades, arrayvalsuht, limittimes, creditlimits,
        creditinterpmethod, pfepercentile, isshortfall, timegap, timeend, anchordate,
        donotionalcap, maxallowedmultiple, horizon)

end

"""
    headroomsolvercore(arrayvalsexistingtrades, arrayvalsuht, limittimes, creditlimits,
        creditinterpmethod, pfepercentile, isshortfall, timegap, timeend, anchordate, 
        donotionalcap, maxallowedmultiple)

Low-level function for trade-headroom solving. By working with the arrayvals variables
the solver can be fast, i.e. sidestep the fact that the PFE of the union of two
portfolios is not generally the sum of the PFE of each portfolio.
"""
function headroomsolvercore(arrayvalsexistingtrades, arrayvalsuht, limittimes, creditlimits,
    creditinterpmethod, pfepercentile, isshortfall, timegap, timeend, anchordate,
    donotionalcap, maxallowedmultiple, horizon)

    pfeprofilewithoutet = pfeprofilefromarrayvals(arrayvalsexistingtrades, pfepercentile,
        isshortfall, timegap, timeend, anchordate)

    timevector = 0:timegap:timeend
    clmp = creditlimitminustwopfes(arrayvalsexistingtrades, nothing, limittimes,
        creditlimits, creditinterpmethod, pfepercentile,
        isshortfall, timevector, horizon)

    if clmp < 0
        return (Dict("multiple" => 0,
            "notionalcapapplies" => false,
            "pfeprofilewithoutet" => pfeprofilewithoutet,
            "pfeprofilewithet" => pfeprofilewithoutet))#because the extra trades are zero
    end

    function f(x)
        arrayvals1 = arrayvalsexistingtrades .+ x .* arrayvalsuht
        creditlimitminustwopfes(arrayvals1, nothing, limittimes, creditlimits,
            creditinterpmethod, pfepercentile, isshortfall, timevector, horizon)
    end

    multiple = find_zero(f, 1000, verbose = true)

    notionalcapapplies = false
    if donotionalcap
        if multiple > maxallowedmultiple
            multiple = maxallowedmultiple
            notionalcapapplies = true
        end
    end

    pfeprofilewithet =
        pfeprofilefromarrayvals(arrayvalsexistingtrades .+ multiple .* arrayvalsuht,
            pfepercentile, isshortfall, timegap, timeend, anchordate)

    res = Dict{String,Any}("multiple" => multiple,
        "notionalcapapplies" => notionalcapapplies,
        "pfeprofilewithoutet" => pfeprofilewithoutet,
        "pfeprofilewithet" => pfeprofilewithet)

    return (res)

end

"""
    function fxsolver(model::AbstractDict, tradesfile::String, basecurrency, limittimes,
        creditlimits, creditinterpmethod, numsims, timegap, timeend, pfepercentile,
        isshortfall::Bool, calcbyproduct::Bool, usethreads::Bool)

Calculates the "Fx Headroom", i.e. solves for an fxshock such that the gap between the PFE
of the trade portfolio and the credit lines is zero, where gap is the 
minimum PFE(t)-CreditLine(t) for 0 <= t <= timeend.

"""
function fxsolver(model::AbstractDict, tradesfile::String, basecurrency, limittimes,
    creditlimits, creditinterpmethod, numsims, timegap, timeend, pfepercentile,
    isshortfall::Bool, calcbyproduct::Bool, usethreads::Bool)

    trades = XVA.loadtrades(tradesfile)
    timevector = 0:timegap:timeend

    if calcbyproduct
        trades1, trades2 = splitbyassetclass(trades)
    else
        trades1, trades2 = trades, nothing
    end

    function f(x)
        shockedmodel = shockmodel(model, x, 1)
        arrayvals1 = arrayvalsfortrades(shockedmodel, numsims, trades1, timegap,
            timeend, usethreads, basecurrency)
        if calcbyproduct
            arrayvals2 = arrayvalsfortrades(shockedmodel, numsims, trades2, timegap,
                timeend, usethreads, basecurrency)
        else
            arrayvals2 = nothing
        end
        creditlimitminustwopfes(arrayvals1, arrayvals2, limittimes, creditlimits,
            creditinterpmethod, pfepercentile, isshortfall, timevector,timevector[end])
    end
    
    multiple = find_zero(f, 0.1, Order16(), atol = 100, verbose = true)

    shockedmodel = shockmodel(model, multiple, 1)
    pfeprofileunshockedfx = pfeprofilefortrades(model, numsims, trades, timegap, timeend,
        pfepercentile, isshortfall, usethreads,
        basecurrency, calcbyproduct)
    pfeprofileshockedfx = pfeprofilefortrades(shockedmodel, numsims, trades, timegap,
        timeend, pfepercentile, isshortfall,
        usethreads, basecurrency, calcbyproduct)

    return (Dict("multiple" => multiple,
        "fxroot" => shockedmodel["spot"]["EUR"] / shockedmodel["spot"]["USD"],
        "pfeprofileunshockedfx" => pfeprofileunshockedfx,
        "pfeprofileshockedfx" => pfeprofileshockedfx))

end

"""
    function fxvolsolver(model::AbstractDict, tradesfile::String, basecurrency, limittimes,
        creditlimits, creditinterpmethod, numsims, timegap, timeend, pfepercentile,
        isshortfall::Bool, calcbyproduct::Bool, usethreads::Bool)

Calculates the "FxVol Headroom", i.e. solves for an fxvolshock such that the gap between the
PFE of the trade portfolio and the credit lines is zero, where gap is the 
minimum PFE(t)-CreditLine(t) for 0 <= t <= timeend.        
"""
function fxvolsolver(model::AbstractDict, tradesfile::String, basecurrency, limittimes,
    creditlimits, creditinterpmethod, numsims, timegap, timeend, pfepercentile,
    isshortfall::Bool, calcbyproduct::Bool, usethreads::Bool)

    trades = XVA.loadtrades(tradesfile)
    timevector = 0:timegap:timeend

    if calcbyproduct
        trades1, trades2 = splitbyassetclass(trades)
    else
        trades1, trades2 = trades, nothing
    end

    function f(x)
        shockedmodel = shockmodel(model, 1, x)
        arrayvals1 = arrayvalsfortrades(shockedmodel, numsims, trades1, timegap,
            timeend, usethreads, basecurrency)
        if calcbyproduct
            arrayvals2 = arrayvalsfortrades(shockedmodel, numsims, trades2, timegap,
                timeend, usethreads, basecurrency)
        else
            arrayvals2 = nothing
        end
        creditlimitminustwopfes(arrayvals1, arrayvals2, limittimes, creditlimits,
            creditinterpmethod, pfepercentile, isshortfall, timevector, timevector[end])
    end

    multiple = 0.0
    succeeded = true
    try
    multiple = find_zero(f, 1, verbose = true, Order16(), atol = 0.5)
    catch
        succeeded = false    
    end

    if !succeeded
        multiple = find_zero(f, (.01,100), verbose = true,  atol = 0.5)
    end

    shockedmodel = shockmodel(model, 1, multiple)
    pfeprofileunshockedfx = pfeprofilefortrades(model, numsims, trades, timegap, timeend,
        pfepercentile, isshortfall, usethreads,
        basecurrency, calcbyproduct)
    pfeprofileshockedfx = pfeprofilefortrades(shockedmodel, numsims, trades, timegap,
        timeend, pfepercentile, isshortfall,
        usethreads, basecurrency, calcbyproduct)

    return (Dict("multiple" => multiple,
        "fxvolroot" => XVA.interp1d(convert(Dict, shockedmodel["fxatmvol"]["USD"]), 3),
        "pfeprofileunshockedfx" => pfeprofileunshockedfx,
        "pfeprofileshockedfx" => pfeprofileshockedfx))

end