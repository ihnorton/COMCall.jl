#
# GUID and CLSID
#

immutable t_data4
    d1::Cuchar; d2::Cuchar; d3::Cuchar; d4::Cuchar; d5::Cuchar; d6::Cuchar; d7::Cuchar; d8::Cuchar
end

immutable GUID
    data1::Culong
    data2::Cushort
    data3::Cushort
    data4::t_data4
end
GUID() = GUID(0,0,0,t_data4(0,0,0,0,0,0,0,0))


immutable CLSID
    guid::GUID
    CLSID(guid) = new(guid)
end
CLSID() = CLSID(GUID())


#
# Win API aliases
#
immutable IID
    guid::GUID
end
typealias REFIID Ptr{IID}
typealias LPIID Ptr{GUID}
typealias REFCLSID Ptr{CLSID}
typealias LPCLSID Ptr{CLSID}
typealias LPOLESTR Ptr{Cwchar_t}
typealias LPCOLESTR Ptr{Cwchar_t}
typealias DWORD Culong

typealias LPUNKNOWN Ptr{Void} # TODO: make this better

######## Base COM API start

baremodule COINIT
    APARTMENTTHREADED  = 0x2
    MULTITHREADED      = 0x0
    DISABLE_OLE1DDE    = 0x4
    SPEED_OVER_MEMORY  = 0x8
end

baremodule HRESULT
    S_OK          =   0x00000000
    S_FALSE       =   0x00000001
    E_UNEXPECTED  =   0x8000FFFF
    E_NOTIMPL     =   0x80004001
    E_OUTOFMEMORY =   0x8007000E
    E_INVALIDARG  =   0x80070057
    E_NOINTERFACE =   0x80004002
    E_POINTER     =   0x80004003
    E_HANDLE      =   0x80070006
    E_ABORT       =   0x80004004
    E_FAIL        =   0x80004005
    E_ACCESSDENIED=   0x80070005
    E_PENDING     =   0x8000000A
end # module HRESULT

baremodule REGDB
    E_FIRST           = 0x80040150
    E_LAST            = 0x8004015F
    S_FIRST           = 0x00040150
    S_LAST            = 0x0004015F
    E_READREGDB       = 0x80040150
    E_WRITEREGDB      = 0x80040151
    E_KEYMISSING      = 0x80040152
    E_INVALIDVALUE    = 0x80040153
    E_CLASSNOTREG     = 0x80040154
    E_IIDNOTREG       = 0x80040155
    E_BADTHREADINGMODEL = 0x80040156
end


baremodule CLSCTX
    import Base.|
  
    INPROC_SERVER           = 0x1
    INPROC_HANDLER          = 0x2
    LOCAL_SERVER            = 0x4
    INPROC_SERVER16         = 0x8
    REMOTE_SERVER           = 0x10
    INPROC_HANDLER16        = 0x20
    RESERVED1               = 0x40
    RESERVED2               = 0x80
    RESERVED3               = 0x100
    RESERVED4               = 0x200
    NO_CODE_DOWNLOAD        = 0x400
    RESERVED5               = 0x800
    NO_CUSTOM_MARSHAL       = 0x1000
    ENABLE_CODE_DOWNLOAD    = 0x2000
    NO_FAILURE_LOG          = 0x4000
    DISABLE_AAA             = 0x8000
    ENABLE_AAA              = 0x10000
    FROM_DEFAULT_CONTEXT    = 0x20000
    ACTIVATE_32_BIT_SERVER  = 0x40000
    ACTIVATE_64_BIT_SERVER  = 0x80000
    ENABLE_CLOAKING         = 0x100000
    APPCONTAINER            = 0x400000
    ACTIVATE_AAA_AS_IU      = 0x800000
    PS_DLL                  = 0x80000000
    
    SERVER = INPROC_SERVER | LOCAL_SERVER | REMOTE_SERVER
    ALL    = INPROC_HANDLER | SERVER
end # module CLSCTX

################################################################################

#
# CLSID
#

function CLSIDFromString(id::String)
    out = [GUID()]
    res = ccall( (:CLSIDFromString, l_ole32), Uint32,
                (LPCOLESTR, LPCLSID), utf16(id), out)
    res != HRESULT.S_OK && error("CLSIDFromString: Unable to convert $id to CLSID")
    return CLSID(out[1])
end
CLSID(id::String) = CLSIDFromString(id)

#
# Interfaces
#

abstract IUnknown
abstract IDispatch <: IUnknown
#TODO: more types

#
# Builtin IIDs
#

const BaseIIDs = {
    IUnknown => CLSID("{00000000-0000-0000-C000-000000000046}"),
    #RecordInfo => CLSID("{0000002F-0000-0000-C000-000000000046}"),
    #IRecordInfo => CLSID("{0000002F-0000-0000-C000-000000000046}"),
    IDispatch => CLSID("{00020400-0000-0000-C000-000000000046}"),
    #ITypeComp => CLSID("{00020403-0000-0000-C000-000000000046}"),
    #ITypeInfo => CLSID("{00020401-0000-0000-C000-000000000046}"),
    #ITypeInfo2 => CLSID("{00020412-0000-0000-C000-000000000046}"),
    #ITypeLib => CLSID("{00020402-0000-0000-C000-000000000046}"),
    #ITypeLib2 => CLSID("{00020411-0000-0000-C000-000000000046}"),
    None => CLSID("{00000000-0000-0000-0000-000000000000}")
    }
getindex(::Type{IID}, x::Type) = IID(BaseIIDs[x].guid)

function getindex(t::IID, x::Ptr{Void})
    k = None
    for k in keys(BaseIIDs)
        if BaseIIDs[k].guid == t break
        else continue
        end
    end
    convert(Ptr{k}, x)
end
#
# Misc
#

immutable LCID
    d::Uint32
end
convert(::Type{Uint32}, t::LCID) = t.d
