#
# COM API wrappers
#

function CoInitializeEx(concurrency_model)
    res = ccall( (:CoInitializeEx, l_ole32), Uint32, (Csize_t, DWORD,), C_NULL, concurrency_model)
    if (res == HRESULT.S_FALSE)
        warn("COM interface already initialized")
    elseif (res != HRESULT.S_OK)
        error("Unable to initialize COM interface")
    end
end
CoInitializeEx() = CoInitializeEx(COINIT.APARTMENTTHREADED)
CoInitialize() = CoInitializeEx()

function CLSIDFromProgID(id::String)
    clsid = CLSID()
    res = ccall( (:CLSIDFromProgID, l_ole32), stdcall, Uint32, (LPCOLESTR, LPCLSID), utf16(id), clsid)
    res == HRESULT.S_OK || error("Unable to locate program $id")
    return clsid
end

function StringFromCLSID(clsid::CLSID)
    s = Ptr{Uint16}[C_NULL]
    res = ccall( (:StringFromCLSID, l_ole32), stdcall, Uint32,
                (REFCLSID, Ptr{Uint16}), clsid, s)
    return (res, s)
end

function CoCreateInstance(clsid::CLSID; clsctx=None)
    if (clsctx == None)
        clsctx = CLSCTX.SERVER
    end
    
    ppv = [C_NULL]
    res = ccall( (:CoCreateInstance, l_ole32), Uint32,
                (LPCLSID, LPUNKNOWN, DWORD, REFIID, Ptr{Ptr{Void}}),
                clsid, C_NULL, clsctx, BaseIID.IUnknown, ppv)
    if (res == REGDB.E_CLASSNOTREG)
        error("Class not registered ($clsid)")
    end

    COMInstance(ppv[1])
end

#
# High-level API
#

function getindex(c::COMGlobal, name::String)
    # TODO: error message based on the HRESULT or REGDB value
    clsid = CLSIDFromProgID(name)
    CoCreateInstance(clsid)
end

# Intended interface:
#   ie = COM["InternetExplorer.Application"]
#   ie[:Navigate2]("www.julialang.org")

function getindex(c::COMGlobal, clsid::CLSID)
    CoCreateInstance(clsid)
end

function getindex(c::IDispatch, x::Symbol)
    
end