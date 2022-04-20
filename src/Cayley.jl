module Cayley

using XVA
using DataFrames
using Dates
using Memoize
using LRUCache
using SHA
using Statistics
using Roots

import Base.Threads.@spawn

const Nums01 = Union{T where T<:Real,Array{T,1} where T<:Real,AbstractRange{T} where T<:Real}
const Nums1 = Union{Array{T,1} where T<:Real,AbstractRange{T} where T<:Real}

include("shock.jl")
include("buildmodel.jl")
include("solving.jl")
include("pfe.jl")

function installme()
    Sys.iswindows() || throw("Cayley.installme can only be run from Julia on Windows")   
    installscript = normpath(joinpath(@__DIR__,"..","installer","install.vbs"))
    exefile = "C:/Windows/System32/wscript.exe"
    isfile(exefile) || throw("Cannot find Windows Script Host at '$exefile'")
    isfile(installscript) || throw("Cannot find install script at '$installscript'")
    run(`$exefile $installscript`,wait = false)
    println("Installer script has been launched, please respond to the dialogs there.")
    nothing
end

function create_system_image()
    filetoinclude = joinpath(@__DIR__,"create_system_image.jl")
    include(filetoinclude)
end

end # module
