using COMCall
using Base.Test

CoInitialize()
iecls = COMCall.CLSIDFromProgID("InternetExplorer.Application")
ie = COM[iecls]
@test typeof(COMCall.QueryInterface(ie, COMCall.BaseIID.IDispatch)) ==  COMCall._IDispatch


# Goal:
#@test ie = COM["InternetExplorer.Application"]
#@test ie[:Navigate]("www.julialang.org")
