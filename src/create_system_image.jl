if Sys.iswindows()
    sysimage_path = "c:/Users/Public/Solum/XVA_Windows.sox"
elseif Sys.islinux()
    sysimage_path = "/mnt/c/Users/Public/Solum/XVA_Linux.sox"
end

# Better to delete file, since if it's locked then the process fails, but only after quite
# some time..
isfile(sysimage_path) && rm(sysimage_path)

using Pkg
Pkg.activate()
Pkg.add("PackageCompiler")
using PackageCompiler
using XVA
packagefolder = pkgdir(XVA)
precompile_execution_file = joinpath(packagefolder,"src","precompile_execution_file.jl")
PackageCompiler.create_sysimage(["XVA"];sysimage_path,precompile_execution_file)