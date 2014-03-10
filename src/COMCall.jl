module COMCall

#
# Imports
#
import Base: show, convert, getindex

#
# Exports
#
export COM, getindex, BaseIID, QueryInterface, CoInitialize, CoCreateInstance

#
# Implementation
#

include("comtypes.jl")
include("comaux.jl")
include("com.jl")

#
# Module initialization
#   

COM = COMGlobal()

function __init__()
    CoInitialize()
    global COM
    COM.initialized = true
end

end # module COMCall
