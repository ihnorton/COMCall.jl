### Helpers start ###
macro L_str(s)
    quote utf16($s) end
end
### Helpers end ###

### GUID start ###

immutable t_data4
    d1::Cuchar; d2::Cuchar; d3::Cuchar; d4::Cuchar;
end

immutable GUID
    data1::Culong
    data2::Cushort
    data3::Cushort
    data4::t_data4
end
GUID() = GUID(0,0,0,t_data4(0,0,0,0))

typealias IID GUID
typealias LPIID Ptr{GUID}
typealias CLSID GUID
typealias LPCLSID Ptr{CLSID}

typealias LPCOLESTR Ptr{Cwchar_t}
typealias HRESULT Cuint
typealias DWORD Culong
### UUID end ###

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
