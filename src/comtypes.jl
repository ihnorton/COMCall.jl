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
    guid::Array{GUID,1}
    CLSID() = new([GUID()])
end

#
# Win API aliases
#
typealias IID GUID
typealias REFIID Ptr{IID}
typealias LPIID Ptr{GUID}
typealias REFCLSID Ptr{CLSID}
typealias LPCLSID Ptr{CLSID}

typealias LPCOLESTR Ptr{Cwchar_t}
typealias DWORD Culong

typealias LPUNKNOWN Ptr{Void} # TODO: make this better

######## Base COM API start

const l_ole32 = "ole32"

baremodule COINIT
    APARTMENTTHREADED  = 0x2
    MULTITHREADED      = 0x0
    DISABLE_OLE1DDE    = 0x4
    SPEED_OVER_MEMORY  = 0x8
end

baremodule HRESULT
    const S_OK          =   0x00000000
    const S_FALSE       =   0x00000001
    const E_UNEXPECTED  =   0x8000FFFF
    const E_NOTIMPL     =   0x80004001
    const E_OUTOFMEMORY =   0x8007000E
    const E_INVALIDARG  =   0x80070057
    const E_NOINTERFACE =   0x80004002
    const E_POINTER     =   0x80004003
    const E_HANDLE      =   0x80070006
    const E_ABORT       =   0x80004004
    const E_FAIL        =   0x80004005
    const E_ACCESSDENIED=   0x80070005
    const E_PENDING     =   0x8000000A
end # module HRESULT

baremodule REGDB
    const E_FIRST           = 0x80040150
    const E_LAST            = 0x8004015F
    const S_FIRST           = 0x00040150
    const S_LAST            = 0x0004015F
    const E_READREGDB       = 0x80040150
    const E_WRITEREGDB      = 0x80040151
    const E_KEYMISSING      = 0x80040152
    const E_INVALIDVALUE    = 0x80040153
    const E_CLASSNOTREG     = 0x80040154
    const E_IIDNOTREG       = 0x80040155
    const E_BADTHREADINGMODEL = 0x80040156
end


baremodule CLSCTX
    import Base.|
  
    const INPROC_SERVER           = 0x1
    const INPROC_HANDLER          = 0x2
    const LOCAL_SERVER            = 0x4
    const INPROC_SERVER16         = 0x8
    const REMOTE_SERVER           = 0x10
    const INPROC_HANDLER16        = 0x20
    const RESERVED1               = 0x40
    const RESERVED2               = 0x80
    const RESERVED3               = 0x100
    const RESERVED4               = 0x200
    const NO_CODE_DOWNLOAD        = 0x400
    const RESERVED5               = 0x800
    const NO_CUSTOM_MARSHAL       = 0x1000
    const ENABLE_CODE_DOWNLOAD    = 0x2000
    const NO_FAILURE_LOG          = 0x4000
    const DISABLE_AAA             = 0x8000
    const ENABLE_AAA              = 0x10000
    const FROM_DEFAULT_CONTEXT    = 0x20000
    const ACTIVATE_32_BIT_SERVER  = 0x40000
    const ACTIVATE_64_BIT_SERVER  = 0x80000
    const ENABLE_CLOAKING         = 0x100000
    const APPCONTAINER            = 0x400000
    const ACTIVATE_AAA_AS_IU      = 0x800000
    const PS_DLL                  = 0x80000000
    
    const SERVER = INPROC_SERVER | LOCAL_SERVER | REMOTE_SERVER
    ALL   = INPROC_HANDLER | SERVER
end # module CLSCTX

################################################################################

#
# CLSID
#

function CLSIDFromString(id::String)
    out = CLSID()
    res = ccall( (:CLSIDFromString, l_ole32), Uint32,
                (LPCOLESTR, LPCLSID), utf16(id), out.guid)
    res != HRESULT.S_OK && error("CLSIDFromString: Unable to convert $id to CLSID")
    return out
end
CLSID(id::String) = CLSIDFromString(id)

#
# Builtin IIDs
#

baremodule BaseIID
    import ..CLSID
    const IUnknown = CLSID("{00000000-0000-0000-C000-000000000046}")
    const RecordInfo = CLSID("{0000002F-0000-0000-C000-000000000046}")
    const IRecordInfo = CLSID("{0000002F-0000-0000-C000-000000000046}")
    const IDispatch = CLSID("{00020400-0000-0000-C000-000000000046}")
    const ITypeComp = CLSID("{00020403-0000-0000-C000-000000000046}")
    const ITypeInfo = CLSID("{00020401-0000-0000-C000-000000000046}")
    const ITypeInfo2 = CLSID("{00020412-0000-0000-C000-000000000046}")
    const ITypeLib = CLSID("{00020402-0000-0000-C000-000000000046}")
    const ITypeLib2 = CLSID("{00020411-0000-0000-C000-000000000046}")
    const IID_NULL = CLSID("{00000000-0000-0000-0000-000000000000}")
end

#
# COMObject / COMInstance
#

abstract COMObject

type COMInstance <: COMObject
    ptr::Ptr{Void}
end

#
# IUnknown interface
#
#   QueryInterface
#   AddRef
#   Release
abstract IUnknown

#
# IDispatch methods
#
#   <: IUnknown
#       GetTypeInfoCount
#       GetTypeInfo
#       GetIDsOfNames
#       Invoke

abstract IDispatch <: IUnknown


