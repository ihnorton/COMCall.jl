module COMCall

# Imports

import Base: show, convert, getindex

# Exports

export COM, getindex, BaseIID, QueryInterface, CoInitialize, CoCreateInstance


# Libraries

const l_ole32 = "ole32"
const l_kernel32 = "kernel32"


# Implementation

include("comtypes.jl")
include("comaux.jl")
include("com.jl")


# Module initialization  

COM = COMGlobal()
const DefaultLCID = get_default_lcid()

function __init__()
    CoInitialize()
    global COM
    COM.initialized = true
end

end # module COMCall
