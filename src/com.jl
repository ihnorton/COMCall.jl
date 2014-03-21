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
    clsid = [CLSID()]
    res = ccall( (:CLSIDFromProgID, l_ole32), stdcall, Uint32, (LPCOLESTR, LPCLSID), utf16(id), clsid)
    res == HRESULT.S_OK || error("Unable to locate program $id")
    return clsid[1]
end

# TODO
#function StringFromCLSID(clsid::CLSID)
#    s = Ptr{Uint16}[C_NULL]
#    res = ccall( (:StringFromCLSID, l_ole32), stdcall, Uint32,
#                (REFCLSID, Ptr{Uint16}), clsid, s)
#    return (res, s)
#end

function CoCreateInstance(clsid::CLSID; iid=None, clsctx=None)
    if (clsctx == None)
        clsctx = CLSCTX.SERVER
    end
    if (iid == None)
        iid = IID[IUnknown]
    end
    
    ppv = [C_NULL]
    res = ccall( (:CoCreateInstance, l_ole32), Uint32,
                (LPCLSID, LPUNKNOWN, DWORD, REFIID, Ptr{Ptr{Void}}),
                &clsid, C_NULL, clsctx, &iid, ppv)
    if (res == REGDB.E_CLASSNOTREG)
        error("Class not registered ($clsid)")
    end
    iid[ppv[1]]
end

################################################################################

#
# IUnknown interface
#
#   QueryInterface
#   AddRef
#   Release

function QueryInterface{T <: IUnknown}(this::Ptr{T}, clsid::CLSID; err=false)
    obj = [C_NULL]
    #fptr = getvtptr(this,1)
    #res = ccall( fptr, thiscall, Uint32, 
    #            (THISPTR,
    #            REFIID, Ptr{Ptr{Void}}),
    #            this, clsid, obj)
    res = @vcall(this, 1, HResult, clsid::REFIID, obj::Ptr{Ptr{Void}})
    if (res != HRESULT.S_OK)
        err && error("QueryInterface: $clsid not supported")
        return C_NULL
    end
    return getimpl(clsid)(obj[1])
end

function AddRef{T<:IUnknown}(this::Ptr{T})
    refcount = ccall( getvtptr(this,2), thiscall, Culong,
                      (Ptr{Void},), this)
end

function Release{T<:IUnknown}(this::Ptr{T})
    refcount = ccall( getvtptr(this,3), thiscall, Culong,
                      (Ptr{Void},), this)
end

#
# IDispatch methods
#
#   <: IUnknown
#       GetTypeInfoCount
#       GetTypeInfo
#       GetIDsOfNames
#       Invoke

function GetTypeInfoCount{T <: IDispatch}(this::Ptr{T})
    pctinfo = [zero(Cuint)]
    res = @vcall(this, 4, HResult, pctinfo::Ptr{Cuint})
    show(res)
    res != HRESULT.S_OK && error("failed GetTypeInfoCount")
    return pctinfo[1]
end

# TODO
type ITypeInfo
end

function GetTypeInfo{T <: IDispatch}(this::Ptr{T}, tinfokind::Cuint)
    lcid = C_NULL
    itypeinfo = Ptr{Ptr{ITypeInfo}}[C_NULL]
    
    res = ccall( getvtptr(this, 5), thiscall, Uint32,
                (THISPTR,
                Cuint, LCID, Ptr{Ptr{ITypeInfo}}),
                o.ptr, tinfokind, lcid, itypeinfo)
    res != HRESULT.S_OK && error("failed GetTypeInfo")
    return itypeinfo[1]
end

function GetIDsOfNames{T <: IDispatch}(this::Ptr{T}, numNames)
    lcid = DefaultLCID::LCID
    rgszNames = Ptr{LPOLESTR}[C_NULL]
    numNames = convert(Cuint, numNames)
    refiid::REFIID

    namesptr = Ptr{LPOLESTR}[C_NULL]
    res = ccall(getvtptr(this, 6), thiscall, Uint32,
                (THISPTR,
                REFIID, Ptr{LPOLESTR}, Cuint, LCID, Ptr{DISPID}),
                this,
                BaseIID.IID_NULL, rgszNames, cNames, lcid, rgDispId)
    res != HRESULT.S_OK && error("GetIDsOfNames error")
    return rgszNames[1]
end

################################################################################
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