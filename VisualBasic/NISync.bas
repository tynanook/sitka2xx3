Attribute VB_Name = "NISync"
'****************************************************************************
'*                       NI-Sync
'*---------------------------------------------------------------------------
'*   Copyright (c) National Instruments 2003-2005.  All Rights Reserved.
'*---------------------------------------------------------------------------
'*
'* Title:    niSync.bas
'* Purpose:  NI-Sync
'*           Instrument Driver Declarations.
'*
'****************************************************************************

Option Explicit
Option Base 0

'****************************************************************************
'*---------------------------- Attribute Defines ---------------------------*
'****************************************************************************
Public Const NISYNC_ATTR_BASE = 1150000                     ' IVI_SPECIFIC_PUBLIC_ATTR_BASE

' Interface attributes
Public Const NISYNC_ATTR_INTF_NUM = (1150000 + 0)                    ' ViInt32
Public Const NISYNC_ATTR_SERIAL_NUM = (1150000 + 1)                  ' ViInt32

' Calibration Attributes
Public Const NISYNC_ATTR_PFI0_THRESHOLD = (1150000 + 100)            ' ViReal64
Public Const NISYNC_ATTR_PFI1_THRESHOLD = (1150000 + 101)            ' ViReal64
Public Const NISYNC_ATTR_PFI2_THRESHOLD = (1150000 + 102)            ' ViReal64
Public Const NISYNC_ATTR_PFI3_THRESHOLD = (1150000 + 103)            ' ViReal64
Public Const NISYNC_ATTR_PFI4_THRESHOLD = (1150000 + 104)            ' ViReal64
Public Const NISYNC_ATTR_PFI5_THRESHOLD = (1150000 + 105)            ' ViReal64
Public Const NISYNC_ATTR_OSCILLATOR_VOLTAGE = (1150000 + 106)        ' ViReal64
Public Const NISYNC_ATTR_CLK10_PHASE_ADJUST = (1150000 + 107)        ' ViReal64
Public Const NISYNC_ATTR_DDS_VCXO_VOLTAGE = (1150000 + 108)          ' ViReal64
Public Const NISYNC_ATTR_DDS_PHASE_ADJUST = (1150000 + 109)          ' ViReal64
Public Const NISYNC_ATTR_PFI0_1KOHM_ENABLE = (1150000 + 110)         ' ViBoolean
Public Const NISYNC_ATTR_PFI1_1KOHM_ENABLE = (1150000 + 111)         ' ViBoolean
Public Const NISYNC_ATTR_PFI2_1KOHM_ENABLE = (1150000 + 112)         ' ViBoolean
Public Const NISYNC_ATTR_PFI3_1KOHM_ENABLE = (1150000 + 113)         ' ViBoolean
Public Const NISYNC_ATTR_PFI4_1KOHM_ENABLE = (1150000 + 114)         ' ViBoolean
Public Const NISYNC_ATTR_PFI5_1KOHM_ENABLE = (1150000 + 115)         ' ViBoolean

' Synchronization Clock Attributes
Public Const NISYNC_ATTR_FRONT_SYNC_CLK_SRC = (1150000 + 200)        ' ViString
Public Const NISYNC_ATTR_REAR_SYNC_CLK_SRC = (1150000 + 201)         ' ViString
Public Const NISYNC_ATTR_SYNC_CLK_DIV1 = (1150000 + 202)             ' ViInt32
Public Const NISYNC_ATTR_SYNC_CLK_DIV2 = (1150000 + 203)             ' ViInt32
Public Const NISYNC_ATTR_SYNC_CLK_RST_PXITRIG_NUM = (1150000 + 204)  ' ViString
Public Const NISYNC_ATTR_SYNC_CLK_PFI0_FREQ = (1150000 + 205)        ' ViReal64
Public Const NISYNC_ATTR_SYNC_CLK_RST_DDS_CNTR_ON_PXITRIG = (1150000 + 206)      ' ViBoolean
Public Const NISYNC_ATTR_SYNC_CLK_RST_PFI0_CNTR_ON_PXITRIG = (1150000 + 207)     ' ViBoolean
Public Const NISYNC_ATTR_SYNC_CLK_RST_CLK10_CNTR_ON_PXITRIG = (1150000 + 208)    ' ViBoolean

' Trigger State Attributes
Public Const NISYNC_ATTR_TERMINAL_STATE_PXISTAR = (1150000 + 300)    ' ViInt32
Public Const NISYNC_ATTR_TERMINAL_STATE_PXITRIG = (1150000 + 301)    ' ViInt32
Public Const NISYNC_ATTR_TERMINAL_STATE_PFI = (1150000 + 302)        ' ViInt32

' Trigger Routing Attribute
Public Const NISYNC_ATTR_TRIG_ROUTE_ALLBUS = (1150000 + 303)         ' ViBoolean

' DDS Attributes
Public Const NISYNC_ATTR_DDS_FREQ = (1150000 + 400)                  ' ViReal64
Public Const NISYNC_ATTR_DDS_UPDATE_SOURCE = (1150000 + 401)         ' ViString
Public Const NISYNC_ATTR_DDS_INITIAL_DELAY = (1150000 + 402)         ' ViReal64

' Clk Attributes
Public Const NISYNC_ATTR_CLKIN_PLL_FREQ = (1150000 + 500)            ' ViReal64
Public Const NISYNC_ATTR_CLKIN_USE_PLL = (1150000 + 501)             ' ViBoolean
Public Const NISYNC_ATTR_CLKIN_PLL_LOCKED = (1150000 + 502)          ' ViBoolean
Public Const NISYNC_ATTR_CLKOUT_GAIN_ENABLE = (1150000 + 503)        ' ViBoolean
Public Const NISYNC_ATTR_PXICLK10_PRESENT = (1150000 + 504)          ' ViBoolean

' User LED Attributes
Public Const NISYNC_ATTR_USER_LED_STATE = (1150000 + 600)            ' ViBoolean

' 1588 Attributes
Public Const NISYNC_ATTR_1588_1588_CLK_ADJ_OFFSET = (1150000 + 715)  ' ViInt32
Public Const NISYNC_ATTR_1588_1588_CLK_ID = (1150000 + 723)          ' ViInt32
Public Const NISYNC_ATTR_1588_1588_CLK_STATE = (1150000 + 712)       ' ViInt32
Public Const NISYNC_ATTR_1588_AVAIL_TIMESTAMPS = (1150000 + 719)     ' ViInt32
Public Const NISYNC_ATTR_1588_CLK_RESOLUTION = (1150000 + 720)       ' ViInt32
Public Const NISYNC_ATTR_1588_1588_CLK_VARIANCE = (1150000 + 705)    ' ViInt32
Public Const NISYNC_ATTR_1588_EPOCH_NUMBER = (1150000 + 710)         ' ViInt32
Public Const NISYNC_ATTR_1588_GRANDMASTER_UUID = (1150000 + 724)     ' ViString
Public Const NISYNC_ATTR_1588_IP_ADDRESS = (1150000 + 700)           ' ViString
Public Const NISYNC_ATTR_1588_IS_PREFERRED_MASTER = (1150000 + 706)  ' ViBoolean
Public Const NISYNC_ATTR_1588_LINK_SPEED = (1150000 + 725)           ' ViInt32
Public Const NISYNC_ATTR_1588_MASTER_LAST_SEQ_NUM = (1150000 + 714)  ' ViInt32
Public Const NISYNC_ATTR_1588_OBSERVED_DRIFT = (1150000 + 722)       ' ViInt32
Public Const NISYNC_ATTR_1588_OFFSET_FROM_MASTER = (1150000 + 713)   ' ViReal64
Public Const NISYNC_ATTR_1588_ONE_WAY_DELAY = (1150000 + 721)        ' ViInt32
Public Const NISYNC_ATTR_1588_PARENT_UUID = (1150000 + 726)          ' ViString
Public Const NISYNC_ATTR_1588_PTP_SUBDOMAIN = (1150000 + 703)        ' ViString
Public Const NISYNC_ATTR_1588_STEPS_TO_GRANDMASTER = (1150000 + 716)  ' ViInt32
Public Const NISYNC_ATTR_1588_SYNC_INTERVAL = (1150000 + 709)        ' ViInt32
Public Const NISYNC_ATTR_1588_TIMESTAMP_BUF_SIZE = (1150000 + 718)   ' ViInt32
Public Const NISYNC_ATTR_1588_USE_BURST_MODE = (1150000 + 708)       ' ViBoolean
Public Const NISYNC_ATTR_1588_UTC_OFFSET = (1150000 + 711)           ' ViInt32
Public Const NISYNC_ATTR_1588_UUID = (1150000 + 717)                 ' ViString
' = (NISYNC_ATTR_BASE + 727)

'****************************************************************************
'*------------------------ Attribute Value Defines -------------------------*
'****************************************************************************

' Trigger Terminals Selectors
Public Const NISYNC_VAL_PXITRIG0 = "PXI_Trig0"
Public Const NISYNC_VAL_PXITRIG1 = "PXI_Trig1"
Public Const NISYNC_VAL_PXITRIG2 = "PXI_Trig2"
Public Const NISYNC_VAL_PXITRIG3 = "PXI_Trig3"
Public Const NISYNC_VAL_PXITRIG4 = "PXI_Trig4"
Public Const NISYNC_VAL_PXITRIG5 = "PXI_Trig5"
Public Const NISYNC_VAL_PXITRIG6 = "PXI_Trig6"
Public Const NISYNC_VAL_PXITRIG7 = "PXI_Trig7"
Public Const NISYNC_VAL_PXISTAR0 = "PXI_Star0"
Public Const NISYNC_VAL_PXISTAR1 = "PXI_Star1"
Public Const NISYNC_VAL_PXISTAR2 = "PXI_Star2"
Public Const NISYNC_VAL_PXISTAR3 = "PXI_Star3"
Public Const NISYNC_VAL_PXISTAR4 = "PXI_Star4"
Public Const NISYNC_VAL_PXISTAR5 = "PXI_Star5"
Public Const NISYNC_VAL_PXISTAR6 = "PXI_Star6"
Public Const NISYNC_VAL_PXISTAR7 = "PXI_Star7"
Public Const NISYNC_VAL_PXISTAR8 = "PXI_Star8"
Public Const NISYNC_VAL_PXISTAR9 = "PXI_Star9"
Public Const NISYNC_VAL_PXISTAR10 = "PXI_Star10"
Public Const NISYNC_VAL_PXISTAR11 = "PXI_Star11"
Public Const NISYNC_VAL_PXISTAR12 = "PXI_Star12"
Public Const NISYNC_VAL_PFI0 = "PFI0"
Public Const NISYNC_VAL_PFI1 = "PFI1"
Public Const NISYNC_VAL_PFI2 = "PFI2"
Public Const NISYNC_VAL_PFI3 = "PFI3"
Public Const NISYNC_VAL_PFI4 = "PFI4"
Public Const NISYNC_VAL_PFI5 = "PFI5"
Public Const NISYNC_VAL_GND = "Ground"
Public Const NISYNC_VAL_SYNC_CLK_FULLSPEED = "SyncClkFullSpeed"
Public Const NISYNC_VAL_SYNC_CLK_DIV1 = "SyncClkDivided1"
Public Const NISYNC_VAL_SYNC_CLK_DIV2 = "SyncClkDivided2"

' Trigger Terminal Synchronization Clock Selectors
Public Const NISYNC_VAL_SYNC_CLK_ASYNC = "SyncClkAsync"
' Public Const NISYNC_VAL_SYNC_CLK_FULLSPEED  DEFINED ABOVE
' Public Const NISYNC_VAL_SYNC_CLK_DIV1       DEFINED ABOVE
' Public Const NISYNC_VAL_SYNC_CLK_DIV2       DEFINED ABOVE

' Software Trigger Terminal Selectors
Public Const NISYNC_VAL_SWTRIG_GLOBAL = "GlobalSoftwareTrigger"

' Clock Terminal Selectors
Public Const NISYNC_VAL_CLK10 = "PXI_Clk10"
Public Const NISYNC_VAL_CLKIN = "ClkIn"
Public Const NISYNC_VAL_CLKOUT = "ClkOut"
Public Const NISYNC_VAL_OSCILLATOR = "Oscillator"
Public Const NISYNC_VAL_DDS = "DDS"

' "All Connected" Terminal Selector
Public Const NISYNC_VAL_ALL_CONNECTED = "AllConnected"

' Synchronization Clock Source Selectors
' Public Const NISYNC_VAL_PFI0                DEFINED ABOVE
' Public Const NISYNC_VAL_CLK10               DEFINED ABOVE
' Public Const NISYNC_VAL_DDS                 DEFINED ABOVE

' Trigger Terminal Connection Mode Definitions (invert, updateEdge, etc.)
Public Const NISYNC_VAL_DONT_INVERT = 0
Public Const NISYNC_VAL_INVERT = 1
Public Const NISYNC_VAL_UPDATE_EDGE_RISING = 0
Public Const NISYNC_VAL_UPDATE_EDGE_FALLING = 1

' DDS Update Signal Source Selectors
Public Const NISYNC_VAL_DDS_UPDATE_IMMEDIATE = "DDS_UpdateImmediate"
' Public Const NISYNC_VAL_PXITRIG0            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG1            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG2            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG3            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG4            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG5            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG6            DEFINED ABOVE
' Public Const NISYNC_VAL_PXITRIG7            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR0            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR1            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR2            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR3            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR4            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR5            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR6            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR7            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR8            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR9            DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR10           DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR11           DEFINED ABOVE
' Public Const NISYNC_VAL_PXISTAR12           DEFINED ABOVE
' Public Const NISYNC_VAL_PFI0                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI1                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI2                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI3                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI4                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI5                DEFINED ABOVE

' PCI-1588 Terminal Selectors
' Public Const NISYNC_VAL_PFI0                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI1                DEFINED ABOVE
' Public Const NISYNC_VAL_PFI2                DEFINED ABOVE
Public Const NISYNC_VAL_RTSI0 = "RTSI0"
Public Const NISYNC_VAL_RTSI1 = "RTSI1"
Public Const NISYNC_VAL_RTSI2 = "RTSI2"
Public Const NISYNC_VAL_RTSI3 = "RTSI3"
Public Const NISYNC_VAL_RTSI4 = "RTSI4"
Public Const NISYNC_VAL_RTSI5 = "RTSI5"
Public Const NISYNC_VAL_RTSI6 = "RTSI6"
Public Const NISYNC_VAL_RTSI7 = "RTSI7"

' Initial Time Source Definitions
Public Const NISYNC_VAL_INIT_TIME_SRC_SYSTEM_CLK = 0
Public Const NISYNC_VAL_INIT_TIME_SRC_MANUAL = 1

' Level Definitions = (output level)
Public Const NISYNC_VAL_LEVEL_LOW = 0
Public Const NISYNC_VAL_LEVEL_HIGH = 1

' Edge Definitions = (activeEdge, detectedEdge )
Public Const NISYNC_VAL_EDGE_RISING = 0
Public Const NISYNC_VAL_EDGE_FALLING = 1
Public Const NISYNC_VAL_EDGE_ANY = 2

' 1588 Clock State Definitions
Public Const NISYNC_VAL_1588_CLK_STATE_NOT_DEFINED = -1
Public Const NISYNC_VAL_1588_CLK_STATE_INIT = 0
Public Const NISYNC_VAL_1588_CLK_STATE_FAULT = 1
Public Const NISYNC_VAL_1588_CLK_STATE_DISABLE = 2
Public Const NISYNC_VAL_1588_CLK_STATE_LISTENING = 3
Public Const NISYNC_VAL_1588_CLK_STATE_PREMASTER = 4
Public Const NISYNC_VAL_1588_CLK_STATE_MASTER = 5
Public Const NISYNC_VAL_1588_CLK_STATE_PASSIVE = 6
Public Const NISYNC_VAL_1588_CLK_STATE_UNCALIBRATED = 7
Public Const NISYNC_VAL_1588_CLK_STATE_SLAVE = 8

' Sync Interval Definitions
Public Const NISYNC_VAL_SYNC_INTERVAL_1_SEC = 1
Public Const NISYNC_VAL_SYNC_INTERVAL_2_SEC = 2
Public Const NISYNC_VAL_SYNC_INTERVAL_8_SEC = 8
Public Const NISYNC_VAL_SYNC_INTERVAL_16_SEC = 16
Public Const NISYNC_VAL_SYNC_INTERVAL_64_SEC = 64

' Link Speed Definitions
Public Const NISYNC_VAL_LINK_SPEED_10MBITS = 10
Public Const NISYNC_VAL_LINK_SPEED_100MBITS = 100
Public Const NISYNC_VAL_LINK_SPEED_1GBIT = 1000

' 1588 Clock Identifer Definitions
Public Const NISYNC_VAL_1588_CLK_ID_ATOM = 1
Public Const NISYNC_VAL_1588_CLK_ID_GPS = 2
Public Const NISYNC_VAL_1588_CLK_ID_NTP = 3
Public Const NISYNC_VAL_1588_CLK_ID_HAND = 4
Public Const NISYNC_VAL_1588_CLK_ID_INIT = 5
Public Const NISYNC_VAL_1588_CLK_ID_DFLT = 6

' PTP Subdomain Definitions
Public Const NISYNC_VAL_1588_PTP_SUBDOMAIN_DFLT = 0
Public Const NISYNC_VAL_1588_PTP_SUBDOMAIN_ALT1 = 1
Public Const NISYNC_VAL_1588_PTP_SUBDOMAIN_ALT2 = 2
Public Const NISYNC_VAL_1588_PTP_SUBDOMAIN_ALT3 = 3

'- Defined values for action in niSync_CloseExtCal= ()  --------
Public Const NISYNC_VAL_EXT_CAL_ABORT = 0
Public Const NISYNC_VAL_EXT_CAL_COMMIT = 1

'****************************************************************************
'*---------------- Instrument Driver Function Declarations -----------------*
'****************************************************************************

'- Init and Close Functions -------------------------------------------
Public Declare Function niSync_init Lib "NISYNC.DLL" (ByVal resourceName As String, ByVal IDQuery As Integer, ByVal resetDevice As Integer, ByRef vi As Long) As Long
Public Declare Function niSync_close Lib "NISYNC.DLL" (ByVal vi As Long) As Long

'- Error Functions ----------------------------------------------------
Public Declare Function niSync_error_message Lib "NISYNC.DLL" (ByVal vi As Long, ByVal errorCode As Long, ByVal errorMessage As String) As Long

'- Utility Functions --------------------------------------------------
Public Declare Function niSync_reset Lib "NISYNC.DLL" (ByVal vi As Long) As Long
Public Declare Function niSync_self_test Lib "NISYNC.DLL" (ByVal vi As Long, ByRef selfTestResult As Integer, ByVal selfTestMessage As String) As Long
Public Declare Function niSync_revision_query Lib "NISYNC.DLL" (ByVal vi As Long, ByVal instrumentDriverRevision As String, ByVal firmwareRevision As String) As Long

'- FPGA Management Functions ------------------------------------------
Public Declare Function niSync_ConfigureFPGA Lib "NISYNC.DLL" (ByVal vi As Long, ByVal fpgaProgramPath As String) As Long

'- Trigger Terminal Connection Functions -------------------------------
Public Declare Function niSync_ConnectTrigTerminals Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String, ByVal syncClock As String, ByVal invert As Long, ByVal updateEdge As Long) As Long
Public Declare Function niSync_DisconnectTrigTerminals Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String) As Long
Public Declare Function niSync_GetTrigTerminalConnectionInfo Lib "NISYNC.DLL" (ByVal vi As Long, ByVal destTerminal As String, ByVal srcTerminal As String, ByVal syncClock As String, ByRef invert As Long, ByRef updateEdge As Long) As Long

'- Software Trigger Connection Functions --------------------------------
Public Declare Function niSync_ConnectSWTrigToTerminal Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String, ByVal syncClock As String, ByVal invert As Long, ByVal updateEdge As Long, ByVal delay As Double) As Long
Public Declare Function niSync_DisconnectSWTrigFromTerminal Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String) As Long
Public Declare Function niSync_GetSWTrigConnectionInfo Lib "NISYNC.DLL" (ByVal vi As Long, ByVal destTerminal As String, ByVal srcTerminal As String, ByVal syncClk As String, ByRef invert As Long, ByRef updateEdge As Long, ByRef delay As Double) As Long
Public Declare Function niSync_SendSoftwareTrigger Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String) As Long

'- Clk Terminal Functions -----------------------------------------------
Public Declare Function niSync_ConnectClkTerminals Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String) As Long
Public Declare Function niSync_DisconnectClkTerminals Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal destTerminal As String) As Long
Public Declare Function niSync_GetClkTerminalConnectionInfo Lib "NISYNC.DLL" (ByVal vi As Long, ByVal destTerminal As String, ByVal srcTerminal As String) As Long

'- Frequency Counting Functions ---------------------------------------
Public Declare Function niSync_MeasureFrequency Lib "NISYNC.DLL" (ByVal vi As Long, ByVal srcTerminal As String, ByVal duration As Double, ByRef actualDuration As Double, ByRef frequency As Double, ByRef error As Double) As Long

'- 1588 Functions -----------------------------------------------------

'- 1588 PTP and Time Functions ----------------------------------------
Public Declare Function niSync_StartPTP Lib "NISYNC.DLL" (ByVal vi As Long, ByVal initialTimeSource As Long, ByVal initialTimeSeconds As Long, ByVal initialTimeNanoseconds As Long, ByVal initialTimeFractionalNanoseconds As Integer) As Long
Public Declare Function niSync_StopPTP Lib "NISYNC.DLL" (ByVal vi As Long) As Long
Public Declare Function niSync_Get1588Time Lib "NISYNC.DLL" (ByVal vi As Long, ByRef timeSeconds As Long, ByRef timeNanoseconds As Long, ByRef timeFractionalNanoseconds As Integer) As Long

'- 1588 Future Time Events Functions ---------------------------------
Public Declare Function niSync_CreateFutureTimeEvent Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String, ByVal outputLevel As Long, ByVal timeSeconds As Long, ByVal timeNanoseconds As Long, ByVal timeFractionalNanoseconds As Integer) As Long
Public Declare Function niSync_ClearFutureTimeEvents Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String) As Long

'- 1588 Time Stamping Triggers Functions ------------------------------
Public Declare Function niSync_EnableTimeStampTrigger Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String, ByVal activeEdge As Long) As Long
Public Declare Function niSync_ReadTriggerTimeStamp Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String, ByVal timeout As Double, ByRef timeSeconds As Long, ByRef timeNanoseconds As Long, ByRef timeFractionalNanoseconds As Integer, ByRef detectedEdge As Long) As Long
Public Declare Function niSync_DisableTimeStampTrigger Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String) As Long

'- 1588 Clock Functions ----------------------------------
Public Declare Function niSync_CreateClock Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String, ByVal highTicks As Long, ByVal lowTicks As Long, ByVal startTimeSeconds As Long, ByVal startTimeNanoseconds As Long, ByVal startTimeFractionalNanoseconds As Integer, ByVal stopTimeSeconds As Long, ByVal stopTimeNanoseconds As Long, ByVal stopTimeFractionalNanoseconds As Integer) As Long
Public Declare Function niSync_ClearClock Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminal As String) As Long

'- Attribute Functions ------------------------------------------------
Public Declare Function niSync_GetAttributeViInt32 Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByRef value As Long) As Long
Public Declare Function niSync_GetAttributeViReal64 Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByRef value As Double) As Long
Public Declare Function niSync_GetAttributeViBoolean Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByRef value As Integer) As Long
Public Declare Function niSync_GetAttributeViString Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByVal bufferSize As Long, ByVal value As String) As Long

Public Declare Function niSync_SetAttributeViInt32 Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByVal value As Long) As Long
Public Declare Function niSync_SetAttributeViReal64 Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByVal value As Double) As Long
Public Declare Function niSync_SetAttributeViBoolean Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByVal value As Integer) As Long
Public Declare Function niSync_SetAttributeViString Lib "NISYNC.DLL" (ByVal vi As Long, ByVal terminalName As String, ByVal attr As Long, ByVal value As String) As Long

'- Calibration Functions ------------------------------------------

'- Calibration Utility Functions ----------------------------------
Public Declare Function niSync_GetExtCalLastDateAndTime Lib "NISYNC.DLL" (ByVal vi As Long, ByRef Year As Long, ByRef Month As Long, ByRef Day As Long, ByRef Hour As Long, ByRef minute As Long) As Long

Public Declare Function niSync_GetExtCalLastTemp Lib "NISYNC.DLL" (ByVal vi As Long, ByRef temp As Double) As Long

Public Declare Function niSync_GetExtCalRecommendedInterval Lib "NISYNC.DLL" (ByVal vi As Long, ByRef months As Long) As Long

Public Declare Function niSync_ChangeExtCalPassword Lib "NISYNC.DLL" (ByVal vi As Long, ByVal oldPassword As String, ByVal newPassword As String) As Long

Public Declare Function niSync_ReadCurrentTemperature Lib "NISYNC.DLL" (ByVal vi As Long, ByRef temperature As Double) As Long

'- Calibration Data Retrieval Functions-----------------------------
Public Declare Function niSync_CalGetOscillatorVoltage Lib "NISYNC.DLL" (ByVal vi As Long, ByRef voltage As Double) As Long

Public Declare Function niSync_CalGetClk10PhaseVoltage Lib "NISYNC.DLL" (ByVal vi As Long, ByRef voltage As Double) As Long

Public Declare Function niSync_CalGetDDSStartPulsePhaseVoltage Lib "NISYNC.DLL" (ByVal vi As Long, ByRef voltage As Double) As Long

Public Declare Function niSync_CalGetDDSInitialPhase Lib "NISYNC.DLL" (ByVal vi As Long, ByRef phase As Double) As Long

'- Calibration Session Management Functions (password required)-----
Public Declare Function niSync_InitExtCal Lib "NISYNC.DLL" (ByVal resourceName As String, ByVal Password As String, ByRef extCalVi As Long) As Long

Public Declare Function niSync_CloseExtCal Lib "NISYNC.DLL" (ByVal extCalVi As Long, ByVal action As Long) As Long

'- Calibration Adjustment Functions (password required)--------------
Public Declare Function niSync_CalAdjustOscillatorVoltage Lib "NISYNC.DLL" (ByVal extCalVi As Long, ByRef measuredVoltage As Double, ByVal oldVoltage As Double) As Long

Public Declare Function niSync_CalAdjustClk10PhaseVoltage Lib "NISYNC.DLL" (ByVal extCalVi As Long, ByRef measuredVoltage As Double, ByVal oldVoltage As Double) As Long

Public Declare Function niSync_CalAdjustDDSStartPulsePhaseVoltage Lib "NISYNC.DLL" (ByVal extCalVi As Long, ByRef measuredVoltage As Double, ByVal oldVoltage As Double) As Long

Public Declare Function niSync_CalAdjustDDSInitialPhase Lib "NISYNC.DLL" (ByVal extCalVi As Long, ByVal measuredPhase As Double, ByRef oldPhase As Double) As Long

'****************************************************************************
'*------------------------ Error And Completion Codes ----------------------*
'****************************************************************************
Public Const NISYNC_ERROR_BASE = &HBFFA4000                    ' IVI_SPECIFIC_PUBLIC_ATTR_BASE

Public Const NISYNC_ERROR_INV_PARAMETER = &HBFFF0078            ' VI_ERROR_INV_PARAMETER
Public Const NISYNC_ERROR_NSUP_ATTR = &HBFFF001D                ' VI_ERROR_NSUP_ATTR
Public Const NISYNC_ERROR_NSUP_ATTR_STATE = &HBFFF001E          ' VI_ERROR_NSUP_ATTR_STATE
Public Const NISYNC_ERROR_ATTR_READONLY = &HBFFF001F            ' VI_ERROR_ATTR_READONLY
Public Const NISYNC_ERROR_INVALID_DESCRIPTOR = (&HBFFA4000 + 1)
Public Const NISYNC_ERROR_INVALID_MODE = (&HBFFA4000 + 2)
Public Const NISYNC_ERROR_FEATURE_NOT_SUPPORTED = (&HBFFA4000 + 3)
Public Const NISYNC_ERROR_VERSION_MISMATCH = (&HBFFA4000 + 4)
Public Const NISYNC_ERROR_INTERNAL_SOFTWARE = (&HBFFA4000 + 5)
Public Const NISYNC_ERROR_FILE_IO = (&HBFFA4000 + 6)

Public Const NISYNC_ERROR_DRIVER_INITIALIZATION = (&HBFFA4000 + 10)
Public Const NISYNC_ERROR_DRIVER_TIMEOUT = (&HBFFA4000 + 11)

Public Const NISYNC_ERROR_READ_FAILURE = (&HBFFA4000 + 20)
Public Const NISYNC_ERROR_WRITE_FAILURE = (&HBFFA4000 + 21)
Public Const NISYNC_ERROR_DEVICE_NOT_FOUND = (&HBFFA4000 + 22)
Public Const NISYNC_ERROR_DEVICE_NOT_READY = (&HBFFA4000 + 23)
Public Const NISYNC_ERROR_INTERNAL_HARDWARE = (&HBFFA4000 + 24)
Public Const NISYNC_ERROR_OVERFLOW = (&HBFFA4000 + 25)
Public Const NISYNC_ERROR_REMOTE_DEVICE = (&HBFFA4000 + 26)

Public Const NISYNC_ERROR_FIRMWARE_LOAD = (&HBFFA4000 + 30)
Public Const NISYNC_ERROR_DEVICE_NOT_INITIALIZED = (&HBFFA4000 + 31)
Public Const NISYNC_ERROR_CLK10_NOT_PRESENT = (&HBFFA4000 + 32)

Public Const NISYNC_ERROR_PLL_NOT_PRESENT = (&HBFFA4000 + 40)
Public Const NISYNC_ERROR_DDS_NOT_PRESENT = (&HBFFA4000 + 41)
Public Const NISYNC_ERROR_DDS_ALREADY_STARTED = (&HBFFA4000 + 42)

Public Const NISYNC_ERROR_DEST_TERMINAL_IN_USE = &HBFFF0042     ' VI_ERROR_LINE_IN_USE
Public Const NISYNC_ERROR_SRC_TERMINAL_INVALID = (&HBFFA4000 + 50)
Public Const NISYNC_ERROR_DEST_TERMINAL_INVALID = (&HBFFA4000 + 51)
Public Const NISYNC_ERROR_TERMINAL_NOT_CONNECTED = (&HBFFA4000 + 52)
Public Const NISYNC_ERROR_SYNC_CLK_INVALID = (&HBFFA4000 + 53)

Public Const NISYNC_ERROR_CAL_INCORRECT_PASSWORD = (&HBFFA4000 + 60)
Public Const NISYNC_ERROR_CAL_PASSWORD_TOO_LARGE = (&HBFFA4000 + 61)
Public Const NISYNC_ERROR_CAL_NOT_PERMITTED = (&HBFFA4000 + 62)

Public Const NISYNC_ERROR_RSRC_UNAVAILABLE = (&HBFFA4000 + 70)
Public Const NISYNC_ERROR_RSRC_RESERVED = (&HBFFA4000 + 71)
Public Const NISYNC_ERROR_RSRC_NOT_RESERVED = (&HBFFA4000 + 72)
Public Const NISYNC_ERROR_HW_BUFFER_FULL = (&HBFFA4000 + 73)
Public Const NISYNC_ERROR_SW_BUFFER_FULL = (&HBFFA4000 + 74)
Public Const NISYNC_ERROR_SOCKET_FAILURE = (&HBFFA4000 + 75)
Public Const NISYNC_ERROR_SESSION_ABORTED = (&HBFFA4000 + 76)
Public Const NISYNC_ERROR_SESSION_ABORTING = (&HBFFA4000 + 77)
Public Const NISYNC_ERROR_TERMINAL_NOT_SPECIFIED = (&HBFFA4000 + 78)

Public Const NISYNC_ERROR_TIME_OVERFLOW = (&HBFFA4000 + 80)
Public Const NISYNC_ERROR_TIME_TOO_EARLY = (&HBFFA4000 + 81)
Public Const NISYNC_ERROR_TIME_TOO_LATE = (&HBFFA4000 + 82)
Public Const NISYNC_ERROR_PTP_ALREADY_STARTED = (&HBFFA4000 + 83)
Public Const NISYNC_ERROR_PTP_NOT_STARTED = (&HBFFA4000 + 84)
Public Const NISYNC_ERROR_INVALID_CLOCK_STATE = (&HBFFA4000 + 85)
Public Const NISYNC_ERROR_IP_ADDRESS = (&HBFFA4000 + 86)
Public Const NISYNC_ERROR_FUTURE_TIME_EVENT_TOO_SOON = (&HBFFA4000 + 87)
Public Const NISYNC_ERROR_CLOCK_PERIOD_TOO_SHORT = (&HBFFA4000 + 88)
Public Const NISYNC_ERROR_DUP_FUTURE_TIME_EVENT = (&HBFFA4000 + 89)
Public Const NISYNC_ERROR_SYNC_INTERVAL_MISMACH = (&HBFFA4000 + 90)
Public Const NISYNC_ERROR_INVALID_INITIAL_TIME = (&HBFFA4000 + 91)
Public Const NISYNC_ERROR_CLK_ADJ_TOO_LARGE = (&HBFFA4000 + 92)

Public Const NISYNC_ERRMSG_INV_PARAMETER = "A parameter for this operation is invalid."
Public Const NISYNC_ERRMSG_NSUP_ATTR = "The specified attribute is not supported."
Public Const NISYNC_ERRMSG_NSUP_ATTR_STATE = "The specified attribute state is not supported."
Public Const NISYNC_ERRMSG_ATTR_READONLY = "The specified attribute is read-only."
Public Const NISYNC_ERRMSG_INVALID_DESCRIPTOR = "The specified instrument descriptor is invalid."
Public Const NISYNC_ERRMSG_INVALID_MODE = "The mode for this operation is invalid."
Public Const NISYNC_ERRMSG_FEATURE_NOT_SUPPORTED = "This operation requires a feature that is not supported."
Public Const NISYNC_ERRMSG_VERSION_MISMATCH = "There is a version mismatch."
Public Const NISYNC_ERRMSG_INTERNAL_SOFTWARE = "An internal software error occurred."
Public Const NISYNC_ERRMSG_FILE_IO = "An error occurred while reading or writing a file."

Public Const NISYNC_ERRMSG_DRIVER_INITIALIZATION = "An error occurred while initializing the driver."
Public Const NISYNC_ERRMSG_DRIVER_TIMEOUT = "The driver timed out while performing an operation."

Public Const NISYNC_ERRMSG_READ_FAILURE = "A failure occurred while reading from the device."
Public Const NISYNC_ERRMSG_WRITE_FAILURE = "A failure occurred while writing to the device."
Public Const NISYNC_ERRMSG_DEVICE_NOT_FOUND = "The specified device was not found."
Public Const NISYNC_ERRMSG_DEVICE_NOT_READY = "The specified device is not ready."
Public Const NISYNC_ERRMSG_INTERNAL_HARDWARE = "An internal hardware error occurred."
Public Const NISYNC_ERRMSG_OVERFLOW = "An overflow condition occurred."
Public Const NISYNC_ERRMSG_REMOTE_DEVICE = "The specified device is a remote device.  Remote devices are not allowed."

Public Const NISYNC_ERRMSG_FIRMWARE_LOAD = "The firmware failed to load."
Public Const NISYNC_ERRMSG_DEVICE_NOT_INITIALIZED = "The device is not initialized."
Public Const NISYNC_ERRMSG_CLK10_NOT_PRESENT = "PXI_Clk10 is not present."

Public Const NISYNC_ERRMSG_PLL_NOT_PRESENT = "This device does not support a PLL."
Public Const NISYNC_ERRMSG_DDS_NOT_PRESENT = "The device does not support a DDS."
Public Const NISYNC_ERRMSG_DDS_ALREADY_STARTED = "The specified attribute cannot be set because the DDS is already running."

Public Const NISYNC_ERRMSG_DEST_TERMINAL_IN_USE = "The specified destination terminal is in use."
Public Const NISYNC_ERRMSG_SRC_TERMINAL_INVALID = "The specified source terminal is invalid for this operation."
Public Const NISYNC_ERRMSG_DEST_TERMINAL_INVALID = "The specified destination terminal is invalid for this operation."
Public Const NISYNC_ERRMSG_TERMINAL_NOT_CONNECTED = "The specified terminal is not connected."
Public Const NISYNC_ERRMSG_SYNC_CLK_INVALID = "The specified synchronization clock is invalid for this operation."

Public Const NISYNC_ERRMSG_CAL_INCORRECT_PASSWORD = "The supplied external calibration password is incorrect."
Public Const NISYNC_ERRMSG_CAL_PASSWORD_TOO_LARGE = "The external calibration password contains too many characters."
Public Const NISYNC_ERRMSG_CAL_NOT_PERMITTED = "The specified calibration operation is not permitted on this session type."

Public Const NISYNC_ERRMSG_RSRC_UNAVAILABLE = "A resource necessary to complete the specified operation is not available, and, therefore, the operation cannot be completed."
Public Const NISYNC_ERRMSG_RSRC_RESERVED = "A resource necessary to complete the specified operation is already reserved by a previous operation and cannot be shared, and, therefore, the operation cannot be completed."
Public Const NISYNC_ERRMSG_RSRC_NOT_RESERVED = "A resource necessary to complete the specified operation is not reserved and should have already been, and, therefore, the operation cannot be completed"
Public Const NISYNC_ERRMSG_HW_BUFFER_FULL = "A hardware buffer necessary to complete the specified operation is unexpectedly full and, therefore, the operation cannot be completed."
Public Const NISYNC_ERRMSG_SW_BUFFER_FULL = "A software buffer necessary to complete the specified operation is unexpectedly full and, therefore, the operation cannot be completed."
Public Const NISYNC_ERRMSG_SOCKET_FAILURE = "A network socket necessary to complete the specified operation has generated a failure, and, therefore, the operation cannot be completed."
Public Const NISYNC_ERRMSG_SESSION_ABORTED = "The specified operation cannot be performed because a session has been aborted or a device has been removed from the system. Handle this situation as required by the application and then, if appropriate, attempt to perform the operation again."
Public Const NISYNC_ERRMSG_SESSION_ABORTING = "The specified operation cannot be performed because a session is in the process of being aborted or a device is in the process of being removed from the system. Wait until the abort operation is complete and attempt to perform the operation again."
Public Const NISYNC_ERRMSG_TERMINAL_NOT_SPECIFIED = "The specified operation cannot be performed since the terminal was not specified."

Public Const NISYNC_ERRMSG_TIME_OVERFLOW = "A 1588 time value has overflowed.  The resulting value is not accurate."
Public Const NISYNC_ERRMSG_TIME_TOO_EARLY = "The specified time value is too early to be represented as a 1588 time value."
Public Const NISYNC_ERRMSG_TIME_TOO_LATE = "The specified time value is too late to be represetned as a 1588 time value."
Public Const NISYNC_ERRMSG_PTP_ALREADY_STARTED = "The Precision Time Protocol (PTP) has already been started on this device and, therefore, cannot be started again."
Public Const NISYNC_ERRMSG_PTP_NOT_STARTED = "The Precision Time Protocol (PTP) has not been started on this device and. therefore, cannot be stopped."
Public Const NISYNC_ERRMSG_INVALID_CLOCK_STATE = "The specified attribute cannot be set when the Precision Time Protocol (PTP) is in its current state."
Public Const NISYNC_ERRMSG_IP_ADDRESS = "The IP address for the specified device cannot be determined, and, therefore, the specified operation cannot be completed."
Public Const NISYNC_ERRMSG_FUTURE_TIME_EVENT_TOO_SOON = "The time for the specified future time event is too soon, or may be in the past, and cannot be programmed in the device before it would occur."
Public Const NISYNC_ERRMSG_CLOCK_PERIOD_TOO_SHORT = "A clock with the specified period is too short to be generated by the device."
Public Const NISYNC_ERRMSG_DUP_FUTURE_TIME_EVENT = "A future time event with the same time and same terminal as the specified future time event has already been created.  Multiple future time events on the same terminal at the same time cannot be created."
Public Const NISYNC_ERRMSG_SYNC_INTERVAL_MISMACH = "The specified sync interval for this 1588 clock is different than the sync interval specified for other 1588 clocks participating in the PTP.  Adjust the sync interval on this 1588 clock or the other 1588 clocks participating in the PTP to the same value."
Public Const NISYNC_ERRMSG_INVALID_INITIAL_TIME = "The specified initial time is invalid.  Initial times must be after 0 hours 1 January 2000 and before 0 hours 1 January 2100."
Public Const NISYNC_ERRMSG_CLK_ADJ_TOO_LARGE = "The specified 1588 clock adjustment offset is too large.  The clock adjustment cannot be more than +1 seconds or less than -1 seconds."

'****************************************************************************
'*---------------------------- End Include File ----------------------------*
'****************************************************************************

