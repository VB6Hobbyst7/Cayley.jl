
# as of 17 Dec 21 I think I will not use makedf but simply read .csv files
"""
    makedf(x::Union{Matrix{Any},Matrix{String}})::DataFrame
Converts input x to a DataFrame. Top row of X are type indicators, with syntax inherited
from R. CHAR = String, DOUBLE = Float64, BOOL = Bool, INT = Int64,
DateStr = number representing a date, as per Excel.
Second row of x becomes the column names of the returned DataFrame.
"""
function makedf(x::Union{Matrix{Any},Matrix{String}})::DataFrame
    nr, nc = size(x)
    nr = nr - 1
    df = DataFrame()
    for i = 1:nc
        df[!, Symbol(x[2, i])] = narrowcontents(x[1, i], x[3:end, i])
    end
    df
end

# sub of makedf, so may well not need this function.
function narrowcontents(typeindicator::String, contents)

    convertnonmissing(T::Type, x) = x isa Missing ? missing : convert(T, x)

    if typeindicator == "CHAR"
        return (convertnonmissing.(String, contents))
    elseif typeindicator == "DOUBLE"
        return (convertnonmissing.(Float64, contents))
    elseif typeindicator == "BOOL"
        return (convertnonmissing.(Bool, contents))
    elseif typeindicator == "INT"
        return (convertnonmissing.(Int64, contents))
    elseif typeindicator == "DATESTR"#misleading, not a string but a number, using Excel's date-number equivalence
        res = Vector{Union{Missing,Date}}(undef, length(contents))
        for i = 1:length(contents)
            if contents[i] isa Real
                if contents[i] == 0
                    res[i] = missing
                else
                    res[i] = XVA.exceldatetodate(contents[i])
                end
            else
                res[i] = missing
            end
        end
        return (res)
    else
        throw("typeindicator '$typeindicator' is not recognised")
    end
end