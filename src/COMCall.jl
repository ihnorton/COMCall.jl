module COMCall
include("types.jl")

######## Global state start
#const com_initialized = false

######## Global state end


######## Base COM API end

### High-level COM interface ###
export COM

# intended API:
#   ie = COM["InternetExplorer.Application"]
#   ie[:Navigate2]("www.julialang.org")

type COMGlobal
end
COM = nothing

function CoInitializeEx(concurrency_model)
    res = ccall( (:CoInitializeEx, l_ole32), Uint32, (DWORD,) concurrency_model)
    res == HRESULT.S_OK || error("Unable to initialize COM interface")
    global com_initialized::Bool = true
end
CoInitializeEx() = CoInitializeEx(COINIT.APARTMENTTHREADED)
CoInitialize() = CoInitializeEx()

function CLSIDFromString()
end

function CLSIDFromProgID(id::String)
    clsid = [CLSID()]
    res = ccall( (:CLSIDFromProgID, l_ole32), Uint32, (LPCOLESTR, LPCLSID), utf16(id), clsid)
    res == HRESULT.S_OK || error("Unable to locate program $id")
    return clsid[1]
end

function CoCreateInstance()
end
function getindex(c::COMGlobal, name::String)
end

function getindex(c::COMGlobal, clsid::CLSID)
end

#type COMObj
#    ptr::Ptr{Void}
#    
#    function COMObj(clsid)
#        
#    end
#end
    

function __init__()
    CoInitialize()
    global COM::COMGlobal = COMGlobal()
end

end # module COMCall
