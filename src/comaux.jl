#
# Global state
#

type COMGlobal
    initialized::Bool
    COMGlobal() = new(false)
end

################################################################################
# Extras
################################################################################

get_default_lcid() = ccall( (:GetSystemDefaultLCID, l_kernel32), LCID, () )

getvtptr(o,idx) = unsafe_load(unsafe_load(pointer(Ptr{Ptr{Void}},o)),idx)

#
# L string helper
#
macro L_str(s)
    quote utf16($s) end
end

#
# Display and conversion
#

show(io::IO, x::GUID) = @printf(io, "%08X-%04hX-%04hX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX",
                     x.data1, x.data2, x.data3,
                     x.data4.d1,x.data4.d2,x.data4.d3,x.data4.d4,x.data4.d5,x.data4.d6,
                     x.data4.d7,x.data4.d8)

show(io::IO, x::CLSID) = print(io, "CLSID{", x.guid, "}")

convert(::Type{Ptr{CLSID}}, x::CLSID) = reinterpret(Ptr{CLSID}, x.guid)
convert(::Type{Ptr{GUID}}, x::CLSID) = pointer(x.guid)
convert(::Type{Ptr{Void}}, x::IUnknown) = x.ptr
convert(::Type{Ptr{IID}}, x::IID) = x.guid