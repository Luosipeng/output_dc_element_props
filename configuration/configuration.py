"""
   Configurations for calling etap on local machines
"""

# the REST API address
base_address = "http://localhost:60000"



# the project name


# Default configurations
revision_name = "Base"
config_name = "Normal"
# study_case = "LF"
study_case = "TDLF"
presentation = "OLV1"
output_report = "Results"
get_online_data = False
online_config_only = False
what_if_commands = {"Commands":[
    "string"
    ]
}

elementTypes = [
  "AMMETER",
  "BATTERY",
  "BUS",
  "BusDuct",
  "BusWay",
  "CABLE",
  "CAPACITOR",
  "CHARGER",
  "CONTACTOR",
  "CONTINUATION",
  "CXFORM",
  "CXMOTOR",
  "CXNETWORK",
  "DCBUS",
  "DCCB",
  "DCCONVERTER",
  "DCDOUBLESWITCH",
  "DCFUSE",
  "DCIMPEDANCE",
  "DCLUMPLOAD",
  "DCMACHINE",
  "DCSINGLESWITCH",
  "DCSTLOAD",
  "DISTANCERELAY",
  "DOUBLESWITCH",
  "ELEMENTARYDIAGRAM",
  "FRELAY",
  "FUSE",
  "GRDD",
  "GRDRX",
  "GRDY",
  "GRID",
  "GROUND",
  "GroundEarthingAdapter",
  "GroundSwitch",
  "HARMFILTER",
  "HVCB",
  "HVDC_LINK",
  "IMPEDANCE",
  "INDMOTOR",
  "INVERTER",
  "LUMPEDLOAD",
  "LVCB",
  "MGSET",
  "MOV",
  "MRELAY",
  "MULTIFUNCTIONRELAY",
  "MULTIMETER",
  "MVSSTRELAY",
  "NOTABUS",
  "OCRELAY",
  "OLTEXTBOX",
  "OVERLOADHEATER",
  "OVERLOADRELAY",
  "PANEL",
  "PROXY",
  "PVArray",
  "PXFORM",
  "REACTOR",
  "RECLOSER",
  "RELAY32",
  "RELAY87",
  "SINGLESWITCH",
  "SPFEEDER",
  "STATICVAR",
  "STLOAD",
  "SYNGEN",
  "SYNMOTOR",
  "TAGLINK",
  "UNIVERSALRELAY",
  "UPS",
  "UTIL",
  "VFDRIVE",
  "VOLTMETER",
  "VRELAY",
  "WTGEN",
  "XFORM2W",
  "XFORM2W_DELTA",
  "XFORM3W",
  "XLINE",
  "ZIGZAGXFMR"
]