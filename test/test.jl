using COMCall
using Base.Test

@test CoInitialize()

@test ie = COM["InternetExplorer.Application"]
@test ie[:Navigate]("www.julialang.org")
