Attribute VB_Name = "afDigitizerDll"
'= Aeroflex afDigitizer Component ===============================================
'
' File         afdigitizerDll.bas
'
' description  Interface exported by afDigitizerDll
'
'================================================================================
'
' Copyright (c) 2000-2014, Aeroflex Ltd.
'
'================================================================================
'
' This software is distributed under the terms of the Aeroflex Ltd.
' Software Licence And Warranty. See "software Licence and Warranty.pdf"
' in the distribution directory for further information
'
'================================================================================


'------------------------------------------------------------------------------------------------------
' Data types
'------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------
' Error Source
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_mtModuleType_t
    afDigitizerDll_mtAFDIGITIZER = -1
    afDigitizerDll_mtAF3010 = &H3010
    afDigitizerDll_mtAF3030 = &H3030
    afDigitizerDll_mtAF3070 = &H3070
    afDigitizerDll_mtPlugin = &H0
End Enum


'------------------------------------------------------------------------------------------------------
' Captured Data
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_sdtSampleDataType_t
    afDigitizerDll_sdtIFData = 0
    afDigitizerDll_sdtIQData = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Capture IQResolution Type
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_iqrIQResolution_t
    afDigitizerDll_iqr16Bit = 0
    afDigitizerDll_iqrAuto = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Capture IQResolution Type
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_ctmCaptureTimeoutMode_t
    afDigitizerDll_ctmAuto = 0
    afDigitizerDll_ctmUser = 1
End Enum



'------------------------------------------------------------------------------------------------------
' list address source
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_liAddressSource_t
    afDigitizerDll_liasManual = 0
    afDigitizerDll_liasExternal = 1
    afDigitizerDll_liasCounter = 2
    afDigitizerDll_liasExtSerial = 3
End Enum

'------------------------------------------------------------------------------------------------------
' Repeat Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_rmRepeatMode_t
    afDigitizerDll_rmSingle = 0
    afDigitizerDll_rmNTimes = 1
    afDigitizerDll_rmContinuous = 2
End Enum

'------------------------------------------------------------------------------------------------------
' list counter strobe
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_liCounterStrobe_t
    afDigitizerDll_licsExternal = 0
    afDigitizerDll_licsTimer = 1
End Enum


'------------------------------------------------------------------------------------------------------
' reference modes
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lormReferenceMode_t
    afDigitizerDll_lormOCXO = 0
    afDigitizerDll_lormInternal = 1
    afDigitizerDll_lormExternalDaisy = 2
    afDigitizerDll_lormExternalTerminated = 3
End Enum


'------------------------------------------------------------------------------------------------------
' LO Bandwidth
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lolbLoopBandwidth_t
    afDigitizerDll_lolbNormal = 0
    afDigitizerDll_lolbNarrow = 1
    afDigitizerDll_lolbUnspecified = 2
End Enum


'------------------------------------------------------------------------------------------------------
' LO Trigger Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lotmTriggerMode_t
    afDigitizerDll_lotmNone = 0
    afDigitizerDll_lotmAdvance = 1
    afDigitizerDll_lotmToggle = 2
    afDigitizerDll_lotmHop = 3
End Enum


'------------------------------------------------------------------------------------------------------
' single-trigger sources
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_tssTriggerSourceSingle_t
    afDigitizerDll_tssLBL0 = 0
    afDigitizerDll_tssLBR0 = 1
    afDigitizerDll_tssPTB0 = 2
    afDigitizerDll_tssPXI_STAR = 3
    afDigitizerDll_tssPTB1 = 4
    afDigitizerDll_tssPTB2 = 5
    afDigitizerDll_tssPTB3 = 6
    afDigitizerDll_tssPTB4 = 7
    afDigitizerDll_tssPTB5 = 8
    afDigitizerDll_tssPTB6 = 9
    afDigitizerDll_tssPTB7 = 10
    afDigitizerDll_tssLBR1 = 11
    afDigitizerDll_tssLBR2 = 12
    afDigitizerDll_tssLBR3 = 13
    afDigitizerDll_tssLBR4 = 14
    afDigitizerDll_tssLBR5 = 15
    afDigitizerDll_tssLBR6 = 16
    afDigitizerDll_tssLBR7 = 17
    afDigitizerDll_tssLBR8 = 18
    afDigitizerDll_tssLBR9 = 19
    afDigitizerDll_tssLBR10 = 20
    afDigitizerDll_tssLBR11 = 21
    afDigitizerDll_tssLBR12 = 22
End Enum


'------------------------------------------------------------------------------------------------------
' addressed-trigger sources
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_tsaTriggerSourceAddressed_t
    afDigitizerDll_tsaNONE = 0
    afDigitizerDll_tsaTRIG = 1
    afDigitizerDll_tsaLBR = 2
    afDigitizerDll_tsaSER = 3
End Enum


'------------------------------------------------------------------------------------------------------
' LVDS Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lmLVDSMode_t
    afDigitizerDll_lmInput = 0
    afDigitizerDll_lmTristate = 1
    afDigitizerDll_lmOutput = 3
End Enum


'------------------------------------------------------------------------------------------------------
' Modulation Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_mmModulationMode_t
    afDigitizerDll_mmUMTS = 0
    afDigitizerDll_mmGSM = 1
    afDigitizerDll_mmCDMA20001x = 2
    afDigitizerDll_mmEmu2319 = 4
    afDigitizerDll_mmGeneric = 5
End Enum


'------------------------------------------------------------------------------------------------------
' RF Input Source
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_isInputSource_t
    afDigitizerDll_isIFInput = 0
    afDigitizerDll_isRFInput = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Rf Front End Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_femFrontEndMode_t
    afDigitizerDll_femAuto = 0
    afDigitizerDll_femAutoIF = 1
    afDigitizerDll_femManual = 2
End Enum


'------------------------------------------------------------------------------------------------------
' RF External Reference
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_erExternalReference_t
    afDigitizerDll_erLockTo10MHz = 0
    afDigitizerDll_erFreeRun = 2
End Enum


'------------------------------------------------------------------------------------------------------
' Reference Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_rfrmReferenceMode_t
    afDigitizerDll_rfrmInternal = 2
    afDigitizerDll_rfrmExternalDaisy = 0
    afDigitizerDll_rfrmExternalPciBackplane = 1
End Enum


'------------------------------------------------------------------------------------------------------
' RF IF Filter Bypass
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_iffbIFFilterBypass_t
    afDigitizerDll_iffbDisable = 0
    afDigitizerDll_iffbEnable = 1
End Enum


'------------------------------------------------------------------------------------------------------
' RF Auto Temperature Optimization
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_atoAutoTemperatureOptimization_t
    afDigitizerDll_atoDisable = 0
    afDigitizerDll_atoEnable = 1
End Enum


'------------------------------------------------------------------------------------------------------
' RF Auto Flatness Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_afmAutoFlatnessMode_t
    afDigitizerDll_afmDisable = 0
    afDigitizerDll_afmEnable = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Trigger Source
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_tsTrigSource_t
    afDigitizerDll_tsPXI_TRIG_0 = 0
    afDigitizerDll_tsPXI_TRIG_1 = 1
    afDigitizerDll_tsPXI_TRIG_2 = 2
    afDigitizerDll_tsPXI_TRIG_3 = 3
    afDigitizerDll_tsPXI_TRIG_4 = 4
    afDigitizerDll_tsPXI_TRIG_5 = 5
    afDigitizerDll_tsPXI_TRIG_6 = 6
    afDigitizerDll_tsPXI_TRIG_7 = 7
    afDigitizerDll_tsPXI_STAR = 8
    afDigitizerDll_tsPXI_LBL_0 = 9
    afDigitizerDll_tsPXI_LBL_1 = 10
    afDigitizerDll_tsPXI_LBL_2 = 11
    afDigitizerDll_tsPXI_LBL_3 = 12
    afDigitizerDll_tsPXI_LBL_4 = 13
    afDigitizerDll_tsPXI_LBL_5 = 14
    afDigitizerDll_tsPXI_LBL_6 = 15
    afDigitizerDll_tsPXI_LBL_7 = 16
    afDigitizerDll_tsPXI_LBL_8 = 17
    afDigitizerDll_tsPXI_LBL_9 = 18
    afDigitizerDll_tsPXI_LBL_10 = 19
    afDigitizerDll_tsPXI_LBL_11 = 20
    afDigitizerDll_tsPXI_LBL_12 = 21
    afDigitizerDll_tsLVDS_MARKER_0 = 22
    afDigitizerDll_tsLVDS_MARKER_1 = 23
    afDigitizerDll_tsLVDS_MARKER_2 = 24
    afDigitizerDll_tsLVDS_MARKER_3 = 25
    afDigitizerDll_tsLVDS_AUX_0 = 26
    afDigitizerDll_tsLVDS_AUX_1 = 27
    afDigitizerDll_tsLVDS_AUX_2 = 28
    afDigitizerDll_tsLVDS_AUX_3 = 29
    afDigitizerDll_tsLVDS_AUX_4 = 30
    afDigitizerDll_tsLVDS_SPARE_0 = 31
    afDigitizerDll_tsSW_TRIG = 32
    afDigitizerDll_tsINT_TIMER = 34
    afDigitizerDll_tsINT_TRIG = 35
    afDigitizerDll_tsFRONT_SMB = 36
End Enum



'------------------------------------------------------------------------------------------------------
' Trigger Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_rsmReTrigSourceMode_t
    afDigitizerDll_rsmAuto = 0
    afDigitizerDll_rsmUser = 1
End Enum



'------------------------------------------------------------------------------------------------------
' Trigger Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_swtSwTrigMode_t
    afDigitizerDll_swtImmediate = 0
    afDigitizerDll_swtArmed = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Trigger Type
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_ttTrigType_t
    afDigitizerDll_ttEdge = 0
    afDigitizerDll_ttGate = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Trigger edge/Gate polarity
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_egpPolarity_t
    afDigitizerDll_egpPositive = 0
    afDigitizerDll_egpNegative = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Trigger IntTrigger Mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_itmIntTriggerMode_t
    afDigitizerDll_itmAbsolute = 0
    afDigitizerDll_itmRelative = 1
End Enum

'------------------------------------------------------------------------------------------------------
' Trigger IntTrigger Source
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_itsIntTriggerSource_t
    afDigitizerDll_itsIF = 0
    afDigitizerDll_itsIQ = 1
End Enum


'------------------------------------------------------------------------------------------------------
' Trigger IQ Trigger bandwith selection mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_iqbmIQTrigBWidthSelMode_t
    afDigitizerDll_iqbmNearestAbove = 0
    afDigitizerDll_iqbmNearestBelow = 1
End Enum




'------------------------------------------------------------------------------------------------------
' IF asc file format
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_afiASCFileIFFormat_t
    afDigitizerDll_afifNewlineSeparated = 0
End Enum


'------------------------------------------------------------------------------------------------------
' IF Bin file format
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_bfifBinFileIFFormat_t
    afDigitizerDll_bfifStandard = 0
End Enum


'------------------------------------------------------------------------------------------------------
' IQ ASC Format
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_afiqASCFileIQFormat_t
    afDigitizerDll_afiqInterleavedIQ = 0
    afDigitizerDll_afiqCommaSeparatedIQPair = 1
End Enum


'------------------------------------------------------------------------------------------------------
' IQ Bin file format
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_bfiqBinFileIQFormat_t
    afDigitizerDll_bfiqStandard = 0
End Enum


'------------------------------------------------------------------------------------------------------
' RF Routing Matrix
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_rmRoutingMatrix_t
    afDigitizerDll_rmPXI_TRIG_0 = 0
    afDigitizerDll_rmPXI_TRIG_1 = 1
    afDigitizerDll_rmPXI_TRIG_2 = 2
    afDigitizerDll_rmPXI_TRIG_3 = 3
    afDigitizerDll_rmPXI_TRIG_4 = 4
    afDigitizerDll_rmPXI_TRIG_5 = 5
    afDigitizerDll_rmPXI_TRIG_6 = 6
    afDigitizerDll_rmPXI_TRIG_7 = 7
    afDigitizerDll_rmPXI_STAR = 8
    afDigitizerDll_rmPXI_LBL_0 = 9
    afDigitizerDll_rmPXI_LBL_1 = 10
    afDigitizerDll_rmPXI_LBL_2 = 11
    afDigitizerDll_rmPXI_LBL_3 = 12
    afDigitizerDll_rmPXI_LBL_4 = 13
    afDigitizerDll_rmPXI_LBL_5 = 14
    afDigitizerDll_rmPXI_LBL_6 = 15
    afDigitizerDll_rmPXI_LBL_7 = 16
    afDigitizerDll_rmPXI_LBL_8 = 17
    afDigitizerDll_rmPXI_LBL_9 = 18
    afDigitizerDll_rmPXI_LBL_10 = 19
    afDigitizerDll_rmPXI_LBL_11 = 20
    afDigitizerDll_rmPXI_LBL_12 = 21
    afDigitizerDll_rmLVDS_MARKER_1 = 22
    afDigitizerDll_rmLVDS_MARKER_2 = 23
    afDigitizerDll_rmLVDS_MARKER_3 = 24
    afDigitizerDll_rmLVDS_MARKER_4 = 25
    afDigitizerDll_rmLVDS_AUX_0 = 26
    afDigitizerDll_rmLVDS_AUX_1 = 27
    afDigitizerDll_rmLVDS_AUX_2 = 28
    afDigitizerDll_rmLVDS_AUX_3 = 29
    afDigitizerDll_rmLVDS_AUX_4 = 30
    afDigitizerDll_rmLVDS_SPARE_0 = 31
    afDigitizerDll_rmLVDS_SPARE_1 = 32
    afDigitizerDll_rmLVDS_SPARE_2 = 33
    afDigitizerDll_rmGND = 34
    afDigitizerDll_rmINT_TRIG = 35
    afDigitizerDll_rmTIMER = 36
    afDigitizerDll_rmTIMER_SYNC = 37
    afDigitizerDll_rmFRONT_SMB = 38
    afDigitizerDll_rmCAPT_BUSY = 39
    afDigitizerDll_rmLA_IN_0 = 40
    afDigitizerDll_rmLA_IN_1 = 41
    afDigitizerDll_rmLA_IN_2 = 42
    afDigitizerDll_rmLA_IN_3 = 43
    afDigitizerDll_rmLA_IN_4 = 44
    afDigitizerDll_rmLA_IN_5 = 45
    afDigitizerDll_rmLA_IN_6 = 46
    afDigitizerDll_rmLA_IN_7 = 47
    afDigitizerDll_rmLSTB_IN = 48
    afDigitizerDll_rmLA_OUT_0 = 49
    afDigitizerDll_rmLA_OUT_1 = 50
    afDigitizerDll_rmLA_OUT_2 = 51
    afDigitizerDll_rmLA_OUT_3 = 52
    afDigitizerDll_rmLA_OUT_4 = 53
    afDigitizerDll_rmLA_OUT_5 = 54
    afDigitizerDll_rmLA_OUT_6 = 55
    afDigitizerDll_rmLA_OUT_7 = 56
    afDigitizerDll_rmSEQ_STB_IN = 57
    afDigitizerDll_rmSEQ_RESET = 58
    afDigitizerDll_rmSEQ_OUT_0 = 59
    afDigitizerDll_rmSEQ_OUT_1 = 60
    afDigitizerDll_rmSEQ_OUT_2 = 61
    afDigitizerDll_rmSEQ_OUT_3 = 62
    afDigitizerDll_rmSEQ_OUT_4 = 63
    afDigitizerDll_rmSEQ_OUT_5 = 64
    afDigitizerDll_rmSEQ_OUT_6 = 65
    afDigitizerDll_rmSEQ_OUT_7 = 66
    afDigitizerDll_rmSEQ_STB_OUT = 67
    afDigitizerDll_rmSEQ_START = 68
    afDigitizerDll_rmLST_BLANK = 69
    afDigitizerDll_rmSW_TRIG = 70
    afDigitizerDll_rmTIMER_TRIG = 71
    afDigitizerDll_rmLA_SERIAL_OUT = 72
    afDigitizerDll_rmLA_SERIAL_IN = 73
    afDigitizerDll_rm_TRIG_BUSY = 77
End Enum


'------------------------------------------------------------------------------------------------------
' RF Routing Scenario
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_rsRoutingScenario_t
    afDigitizerDll_rsNONE = 0
    afDigitizerDll_rsLVDS_AUX_TO_PXI_LBL = 1
End Enum


'------------------------------------------------------------------------------------------------------
' RF Routing Input
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_riRoutingInput_t
    afDigitizerDll_riGROUND = 0
End Enum


'------------------------------------------------------------------------------------------------------
' LO Position
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lopLOPosition_t
    afDigitizerDll_lopBelow = 0
    afDigitizerDll_lopAbove = 1
End Enum

Public Enum afDigitizerDll_lopmLOPositionMode_t
    afDigitizerDll_lopmManual = 0
    afDigitizerDll_lopmSemiAuto = 1
    afDigitizerDll_lopmAuto = 2
End Enum





'------------------------------------------------------------------------------------------------------
' Async IQ Capture Event
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_ciqeCaptureIQEvent_t
    afDigitizerDll_ciqeSetup = 0
    afDigitizerDll_ciqeTriggered = 1
    afDigitizerDll_ciqeCaptured = 2
    afDigitizerDll_ciqeCompleted = 3
    afDigitizerDll_ciqeError = 4
    afDigitizerDll_ciqeArmed = 5
    afDigitizerDll_ciqeNonHwInit = 6
    afDigitizerDll_ciqeTransfer = 7
End Enum


'------------------------------------------------------------------------------------------------------
' Bandwidth for a given flatness
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_frFlatnessRequired_t
    afDigitizerDll_frLessThanOnedB = 0
    afDigitizerDll_frLessThanQuarterdB
End Enum


'------------------------------------------------------------------------------------------------------
' LVDS Sampling Rate mode
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lsrmLvdsSamplingRateMode_t
    afDigitizerDll_lsrmLockedToCapture = 0
    afDigitizerDll_lsrmAutoAdjust = 1
End Enum


'------------------------------------------------------------------------------------------------------
' LVDS Clock Rate
'------------------------------------------------------------------------------------------------------
Public Enum afDigitizerDll_lcrLvdsClockRate_t
    afDigitizerDll_lcr62_5MHz = 1
    afDigitizerDll_lcr125MHz = 0
    afDigitizerDll_lcr180MHz = 2
End Enum






'------------------------------------------------------------------------------------------------------
' Function prototypes
'------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------
' Object eeprom Cache policy control
'------------------------------------------------------------------------------------------------------
Declare Function afDigitizerDll_EepromCacheEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal state As Long) As Long
Declare Function afDigitizerDll_EepromCacheEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pState As Long) As Long

Declare Function afDigitizerDll_EepromCachePathLength_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLen As Long) As Long
Declare Function afDigitizerDll_EepromCachePath_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pathBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_EepromCachePath_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal buffer As String) As Long



'------------------------------------------------------------------------------------------------------
' Object creation and destruction
'------------------------------------------------------------------------------------------------------
' methods
Declare Function afDigitizerDll_CreateObject Lib "afDigitizerDll_32.dll" (ByRef pDigitizerId As Long) As Long
Declare Function afDigitizerDll_DestroyObject Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long



'------------------------------------------------------------------------------------------------------
' Errors Information ed
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_ErrorCode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pErrorCode As Long) As Long
Declare Function afDigitizerDll_ErrorSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pErrorSource As Long) As Long
Declare Function afDigitizerDll_ErrorMessage_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pErrorMessageBuffer As String, ByVal bufferLen As Long) As Long



'------------------------------------------------------------------------------------------------------
' General
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_IsActive_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsActive As Long) As Long
Declare Function afDigitizerDll_LoRfSpeedSyncEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal enable As Long) As Long
Declare Function afDigitizerDll_LoRfSpeedSyncEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEnable As Long) As Long

' methods
Declare Function afDigitizerDll_BootInstrument Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal LoResource As String, ByVal RfResource As String, ByVal LoIsPlugin As Long) As Long
Declare Function afDigitizerDll_CloseInstrument Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_ClearErrors Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_GetVersion Lib "afDigitizerDll_32.dll" (ByRef version As Long) As Long




'------------------------------------------------------------------------------------------------------
' Calibrate LoNull
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Calibrate_LoNull_CurrentFreqCalRequired_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pcalRequired As Long) As Long
Declare Function afDigitizerDll_Calibrate_LoNull_FreqBandCalRequired_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pcalRequired As Long) As Long

' methods
Declare Function afDigitizerDll_Calibrate_LoNull_CurrentFreqCal Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_Calibrate_LoNull_FreqBandCal Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long


'------------------------------------------------------------------------------------------------------
' Capture IQ
'------------------------------------------------------------------------------------------------------
' Properties

Declare Function afDigitizerDll_Capture_IQ_ADCOverload_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pADCOverload As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_CaptComplete_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCaptComplete As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_CapturedSampleCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCapturedSampleCount As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_ListAddrCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pListAddrCount As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_ReclaimTimeout_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal timeoutMillisecs As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_ReclaimTimeout_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTimeoutMillisecs As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Resolution_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIQResolution As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Resolution_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal IQResolution As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_TriggerCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_TriggerDetected_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDetected As Long) As Long

' methods
Declare Function afDigitizerDll_Capture_IQ_Abort Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Cancel Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal capture As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_CaptMem Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numberOfIQSamples As Long, ByRef iBuffer As Single, ByRef qBuffer As Single) As Long
Declare Function afDigitizerDll_Capture_IQ_CaptMemWithKey Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numberOfIQSamples As Long, ByRef iBuffer As Single, ByRef qBuffer As Single, ByRef tag As Double, ByRef Key As Double) As Long
Declare Function afDigitizerDll_Capture_IQ_GetAbsSampleTime Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleNumber As Long, ByRef psampleTime As Double) As Long
Declare Function afDigitizerDll_Capture_IQ_GetCaptMemFromOffset Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal offset As Long, ByVal numberOfIQSamples As Long, ByRef iBuffer As Single, ByRef qBuffer As Single) As Long
Declare Function afDigitizerDll_Capture_IQ_GetCaptMemFromOffsetWithKey Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal offset As Long, ByVal numberOfIQSamples As Long, ByRef iBuffer As Single, ByRef qBuffer As Single, ByRef tag As Double, ByRef Key As Double) As Long
Declare Function afDigitizerDll_Capture_IQ_GetListAddrInfo Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal listEvent As Long, ByRef pListAddr As Integer, ByRef pStartSample As Long, ByRef pNumSamples As Long, ByRef pInvalidSamples As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_GetOutstandingBuffers Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Long, ByRef pCompleted As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_GetTriggerSampleNumber Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal triggerNumber As Long, ByRef pSampleNumber As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_GetBufferKey Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal capture As Long, ByRef tag As Double, ByRef Key As Double) As Long
Declare Function afDigitizerDll_Capture_IQ_TriggerArm Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal samples As Long) As Long

Declare Function afDigitizerDll_Capture_IQ_GetSampleCaptured Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleNumber As Long, ByRef pSampleCaptured As Long) As Long

'------------------------------------------------------------------------------------------------------
' Capture IQ Power
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Capture_IQ_Power_IsAvailable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsAvailable As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Power_NumOfSteps_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numOfSteps As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Power_NumOfSteps_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pNumOfSteps As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Power_NumMeasurementsAvail_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pNumMeasurementsAvail As Long) As Long

' methods
Declare Function afDigitizerDll_Capture_IQ_Power_SetParameters Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal stepLength As Long, ByVal measOffset As Long, ByVal measLength As Long) As Long
Declare Function afDigitizerDll_Capture_IQ_Power_GetParameters Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pStepLength As Long, ByRef pMeasOffset As Long, ByRef pMeasLength As Long) As Long

Declare Function afDigitizerDll_Capture_IQ_Power_GetAllMeasurements Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numMeasurements As Long, ByRef powers As Double) As Long
Declare Function afDigitizerDll_Capture_IQ_Power_GetSingleMeasurement Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal step As Long, ByRef pPower As Double) As Long



'------------------------------------------------------------------------------------------------------
' Capture IF
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Capture_IF_ADCOverload_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pADCOverload As Long) As Long
Declare Function afDigitizerDll_Capture_IF_CaptComplete_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCaptComplete As Long) As Long
Declare Function afDigitizerDll_Capture_IF_CapturedSampleCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCapturedSampleCount As Long) As Long
Declare Function afDigitizerDll_Capture_IF_ListAddrCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pListAddrCount As Long) As Long
Declare Function afDigitizerDll_Capture_IF_TriggerCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Long) As Long
Declare Function afDigitizerDll_Capture_IF_TriggerDetected_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDetected As Long) As Long

' IF methods
Declare Function afDigitizerDll_Capture_IF_Abort Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_Capture_IF_CaptMem Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numberOfIFSamples As Long, ByRef ifBuffer As Integer) As Long
Declare Function afDigitizerDll_Capture_IF_GetAbsSampleTime Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleNumber As Long, ByRef psampleTime As Double) As Long
Declare Function afDigitizerDll_Capture_IF_GetCaptMemFromOffset Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal offset As Long, ByVal numberOfIFSamples As Long, ByRef ifBuffer As Integer) As Long
Declare Function afDigitizerDll_Capture_IF_GetListAddrInfo Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal listEvent As Long, ByRef pListAddr As Integer, ByRef pStartSample As Long, ByRef pNumSamples As Long, ByRef pInvalidSamples As Long) As Long
Declare Function afDigitizerDll_Capture_IF_GetTriggerSampleNumber Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal triggerNumber As Long, ByRef pSampleNumber As Long) As Long
Declare Function afDigitizerDll_Capture_IF_TriggerArm Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal samples As Long) As Long

Declare Function afDigitizerDll_Capture_IF_GetSampleCaptured Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleNumber As Long, ByRef pSampleCaptured As Long) As Long


'------------------------------------------------------------------------------------------------------
' Capture General
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Capture_SampleDataType_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSampleDataType As Long) As Long
Declare Function afDigitizerDll_Capture_SampleDataType_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleDataType As Long) As Long

Declare Function afDigitizerDll_Capture_PipeliningEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEnable As Long) As Long
Declare Function afDigitizerDll_Capture_PipeliningEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal enable As Long) As Long

Declare Function afDigitizerDll_Capture_TimeoutMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTimeoutMode As Long) As Long
Declare Function afDigitizerDll_Capture_TimeoutMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal timeoutMode As Long) As Long

Declare Function afDigitizerDll_Capture_UserTimeout_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pUserTimeout As Long) As Long
Declare Function afDigitizerDll_Capture_UserTimeout_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal userTimeout As Long) As Long



'------------------------------------------------------------------------------------------------------
' List mode
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_ListMode_Available_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailable As Long) As Long

Declare Function afDigitizerDll_ListMode_ReTrigAvailable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailable As Long) As Long


Declare Function afDigitizerDll_ListMode_AddressSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pliAddSrc As Long) As Long
Declare Function afDigitizerDll_ListMode_AddressSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal liAddSrc As Long) As Long

Declare Function afDigitizerDll_ListMode_RepeatMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal mode As Long) As Long
Declare Function afDigitizerDll_ListMode_RepeatMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef mode As Long) As Long

Declare Function afDigitizerDll_ListMode_RepeatCount_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal count As Integer) As Long
Declare Function afDigitizerDll_ListMode_RepeatCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef count As Integer) As Long

Declare Function afDigitizerDll_ListMode_Counter_StartAddress_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pStartAddress As Integer) As Long
Declare Function afDigitizerDll_ListMode_Counter_StartAddress_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal startAddress As Integer) As Long

Declare Function afDigitizerDll_ListMode_Counter_StopAddress_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pStopAddress As Integer) As Long
Declare Function afDigitizerDll_ListMode_Counter_StopAddress_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal stopAddress As Integer) As Long

Declare Function afDigitizerDll_ListMode_Counter_StrobeSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCounterStrobeSrc As Long) As Long
Declare Function afDigitizerDll_ListMode_Counter_StrobeSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal counterStrobeSrc As Long) As Long

Declare Function afDigitizerDll_ListMode_Strobe_NegativeEdge_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pNegativeEdge As Long) As Long
Declare Function afDigitizerDll_ListMode_Strobe_NegativeEdge_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal negativeEdge As Long) As Long

Declare Function afDigitizerDll_ListMode_Channel_DwellInSamples_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pDwellInSamples As Long) As Long
Declare Function afDigitizerDll_ListMode_Channel_DwellInSamples_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal dwellInSamples As Long) As Long


Declare Function afDigitizerDll_ListMode_Channel_WaitForReTrig_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pWait As Long) As Long
Declare Function afDigitizerDll_ListMode_Channel_WaitForReTrig_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal wait As Long) As Long

Declare Function afDigitizerDll_ListMode_Channel_DiscardSamples_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pDiscard As Long) As Long
Declare Function afDigitizerDll_ListMode_Channel_DiscardSamples_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal discard As Long) As Long

Declare Function afDigitizerDll_ListMode_Channel_DiscardSamplesUntilReTrig_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pDiscard As Long) As Long
Declare Function afDigitizerDll_ListMode_Channel_DiscardSamplesUntilReTrig_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal discard As Long) As Long



'------------------------------------------------------------------------------------------------------
' LO
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_LO_Reference_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLoRefMode As Long) As Long
Declare Function afDigitizerDll_LO_Reference_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal loRefMode As Long) As Long
Declare Function afDigitizerDll_LO_ReferenceLocked_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef isLocked As Long) As Long
Declare Function afDigitizerDll_LO_LoopBandwidth_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLoopBandwidth As Long) As Long
Declare Function afDigitizerDll_LO_LoopBandwidth_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal loopBandwidth As Long) As Long

Declare Function afDigitizerDll_LO_Options_AvailableCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailableCount As Long) As Long

Declare Function afDigitizerDll_LO_Resource_FPGACount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Integer) As Long
Declare Function afDigitizerDll_LO_Resource_FPGAConfiguration_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pFPGAConfig As Integer) As Long
Declare Function afDigitizerDll_LO_Resource_FPGAConfiguration_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal FPGAConfig As Integer) As Long
Declare Function afDigitizerDll_LO_Resource_IsActive_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsActive As Long) As Long
Declare Function afDigitizerDll_LO_Resource_IsPlugin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsPlugin As Long) As Long
Declare Function afDigitizerDll_LO_Resource_PluginName_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pluginNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_LO_Resource_PluginName_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pluginNameBuffer As String) As Long
Declare Function afDigitizerDll_LO_Resource_ModelCode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pModelCode As Long) As Long
Declare Function afDigitizerDll_LO_Resource_ResourceString_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal resourceStringBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_LO_Resource_SerialNumber_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal SerialNumberBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_LO_Resource_SessionID_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSessionId As Long) As Long
Declare Function afDigitizerDll_LO_Resource_Temperature_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTemperature As Double) As Long

Declare Function afDigitizerDll_LO_Trigger_Mode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTriggerMode As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_Mode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal triggerMode As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_AddressedSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAddressedSource As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_AddressedSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal addressedSource As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTriggerSourceSingle As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal triggerSourceSingle As Long) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleStartChannel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSingleStartChannel As Integer) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleStartChannel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal singleStartChannel As Integer) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleStopChannel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSingleStopChannel As Integer) As Long
Declare Function afDigitizerDll_LO_Trigger_SingleStopChannel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal singleStopChannel As Integer) As Long

' methods
Declare Function afDigitizerDll_LO_Options_Enable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Password As Long) As Long
Declare Function afDigitizerDll_LO_Options_Disable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Password As Long) As Long
Declare Function afDigitizerDll_LO_Options_CheckFitted Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal OptionNumber As Long, ByRef pFitted As Long) As Long
Declare Function afDigitizerDll_LO_Options_Information Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Index As Long, ByRef pOptionNumber As Long, ByVal OptionDescriptionBuffer As String, ByVal bufferLen As Long) As Long

Declare Function afDigitizerDll_LO_Resource_FPGADescriptions Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef Numbers As Integer, ByVal Descriptions As String, ByRef pCount As Integer) As Long
Declare Function afDigitizerDll_LO_Resource_GetLastCalibrationDate Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef Year As Integer, ByRef Month As Integer, ByRef Day As Integer, ByRef Hour As Integer, ByRef Minutes As Integer, ByRef Seconds As Integer) As Long



'------------------------------------------------------------------------------------------------------
' LVDS
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_LVDS_AuxiliaryMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAuxiliaryMode As Long) As Long
Declare Function afDigitizerDll_LVDS_AuxiliaryMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal AuxiliaryMode As Long) As Long

Declare Function afDigitizerDll_LVDS_ClockEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pClockEnable As Long) As Long
Declare Function afDigitizerDll_LVDS_ClockEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal clockEnable As Long) As Long

Declare Function afDigitizerDll_LVDS_DataMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDataMode As Long) As Long
Declare Function afDigitizerDll_LVDS_DataMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal DataMode As Long) As Long
Declare Function afDigitizerDll_LVDS_DataDelay_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDataDelay As Double) As Long

Declare Function afDigitizerDll_LVDS_MarkerMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMarkerMode As Long) As Long
Declare Function afDigitizerDll_LVDS_MarkerMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal MarkerMode As Long) As Long


Declare Function afDigitizerDll_LVDS_SamplingRateModeAvailable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailable As Long) As Long
Declare Function afDigitizerDll_LVDS_SamplingRateMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMode As Long) As Long
Declare Function afDigitizerDll_LVDS_SamplingRateMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal mode As Long) As Long
Declare Function afDigitizerDll_LVDS_SamplingRate_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRate As Double) As Long

Declare Function afDigitizerDll_LVDS_ClockRate_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRate As Long) As Long
Declare Function afDigitizerDll_LVDS_ClockRate_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal rate As Long) As Long


'------------------------------------------------------------------------------------------------------
' Modulation
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Modulation_Mode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMode As Long) As Long
Declare Function afDigitizerDll_Modulation_Mode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal mode As Long) As Long

Declare Function afDigitizerDll_Modulation_DecimatedSamplingFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDecimatedSamplingFrequency As Double) As Long
Declare Function afDigitizerDll_Modulation_UndecimatedSamplingFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pUndecimatedSamplingFrequency As Double) As Long

Declare Function afDigitizerDll_Modulation_GenericDecimationRatio_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_GenericDecimationRatio_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal GenericDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_GenericDecimationRatioMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericDecimationRatioMin As Long) As Long
Declare Function afDigitizerDll_Modulation_GenericDecimationRatioMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericDecimationRatioMax As Long) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericSamplingFrequency As Double) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFrequency_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal GenericSamplingFrequency As Double) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFrequencyMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericSamplingFrequencyMax As Double) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFrequencyMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericSamplingFrequencyMin As Double) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFreqNumerator_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericSamplingFreqNumerator As Long) As Long
Declare Function afDigitizerDll_Modulation_GenericSamplingFreqDenominator_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGenericSamplingFreqDenominator As Long) As Long

Declare Function afDigitizerDll_Modulation_CDMA20001XDecimationRatio_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCDMA20001XDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_CDMA20001XDecimationRatio_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal CDMA20001XDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_CDMA20001XDecimationRatioMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCDMA20001XDecimationRatioMin As Long) As Long
Declare Function afDigitizerDll_Modulation_CDMA20001XDecimationRatioMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCDMA20001XDecimationRatioMax As Long) As Long

Declare Function afDigitizerDll_Modulation_Emu2319DecimationRatio_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEmu2319DecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_Emu2319DecimationRatio_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Emu2319DecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_Emu2319DecimationRatioMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEmu2319DecimationRatioMin As Long) As Long
Declare Function afDigitizerDll_Modulation_Emu2319DecimationRatioMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEmu2319DecimationRatioMax As Long) As Long

Declare Function afDigitizerDll_Modulation_GSMDecimationRatio_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGSMDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_GSMDecimationRatio_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal GSMDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_GSMDecimationRatioMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGSMDecimationRatioMin As Long) As Long
Declare Function afDigitizerDll_Modulation_GSMDecimationRatioMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pGSMDecimationRatioMax As Long) As Long

Declare Function afDigitizerDll_Modulation_UMTSDecimationRatio_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pUMTSDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_UMTSDecimationRatio_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal UMTSDecimationRatio As Long) As Long
Declare Function afDigitizerDll_Modulation_UMTSDecimationRatioMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pUMTSDecimationRatioMin As Long) As Long
Declare Function afDigitizerDll_Modulation_UMTSDecimationRatioMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pUMTSDecimationRatioMax As Long) As Long

' methods
Declare Function afDigitizerDll_Modulation_SetGenericSamplingFreqRatio Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numerator As Long, ByVal denominator As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF General
'------------------------------------------------------------------------------------------------------
' General Properties
Declare Function afDigitizerDll_RF_AutoFlatnessMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAutoFlatnessMode As Long) As Long
Declare Function afDigitizerDll_RF_AutoFlatnessMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal autoFlatnessMode As Long) As Long

Declare Function afDigitizerDll_RF_AutoTemperatureOptimization_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAutoTemperatureOptimization As Long) As Long
Declare Function afDigitizerDll_RF_AutoTemperatureOptimization_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal AutoTemperatureOptimization As Long) As Long

Declare Function afDigitizerDll_RF_CentreFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCenterFrequency As Double) As Long
Declare Function afDigitizerDll_RF_CentreFrequency_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal CenterFrequency As Double) As Long
Declare Function afDigitizerDll_RF_SetCentreFrequencyAndLOPosition Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal CenterFrequency As Double, ByVal LOPosition As Long) As Long
Declare Function afDigitizerDll_RF_CentreFrequencyLOAboveMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMaxCentreFreq As Double) As Long
Declare Function afDigitizerDll_RF_CentreFrequencyLOBelowMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMinCentreFreq As Double) As Long
Declare Function afDigitizerDll_RF_CentreFrequencyMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCenterFrequencyMax As Double) As Long
Declare Function afDigitizerDll_RF_CentreFrequencyMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCenterFrequencyMin As Double) As Long

Declare Function afDigitizerDll_RF_CurrentChannel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCurrentChannel As Integer) As Long
Declare Function afDigitizerDll_RF_CurrentChannel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal currentChannel As Integer) As Long

Declare Function afDigitizerDll_RF_DividedLOFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pDividedLOFrequency As Double) As Long

Declare Function afDigitizerDll_RF_ExternalReference_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pExternalReference As Long) As Long
Declare Function afDigitizerDll_RF_ExternalReference_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal externalReference As Long) As Long

Declare Function afDigitizerDll_RF_FrontEndMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pFrontEndMode As Long) As Long
Declare Function afDigitizerDll_RF_FrontEndMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal FrontEndMode As Long) As Long

Declare Function afDigitizerDll_RF_IFAttenuation_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_IFAttenuation_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal IFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_IFAttenuationMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFAttenuationMax As Long) As Long
Declare Function afDigitizerDll_RF_IFAttenuationMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFAttenuationMin As Long) As Long
Declare Function afDigitizerDll_RF_IFAttenuationStep_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFAttenuationStep As Long) As Long
Declare Function afDigitizerDll_RF_IFFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFFFrequency As Double) As Long
Declare Function afDigitizerDll_RF_IFFilterBypass_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFFilterBypass As Long) As Long
Declare Function afDigitizerDll_RF_IFFilterBypass_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal IFFilterBypass As Long) As Long
Declare Function afDigitizerDll_RF_IFInputLevel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_IFInputLevel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal IFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_IFInputLevelMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFInputLevelMax As Double) As Long
Declare Function afDigitizerDll_RF_IFInputLevelMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIFInputLevelMin As Double) As Long

Declare Function afDigitizerDll_RF_InputSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pInputSource As Long) As Long
Declare Function afDigitizerDll_RF_InputSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal InputSource As Long) As Long

Declare Function afDigitizerDll_RF_LevelCorrection_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLevelCorrection As Double) As Long

Declare Function afDigitizerDll_RF_LOFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOFrequency As Double) As Long
Declare Function afDigitizerDll_RF_LOOffset_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOOffset As Double) As Long
Declare Function afDigitizerDll_RF_LOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOPosition As Long) As Long
Declare Function afDigitizerDll_RF_LOPosition_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal LOPosition As Long) As Long


Declare Function afDigitizerDll_RF_UserLOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOPosition As Long) As Long
Declare Function afDigitizerDll_RF_UserLOPosition_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal LOPosition As Long) As Long

Declare Function afDigitizerDll_RF_ActualLOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOPosition As Long) As Long

Declare Function afDigitizerDll_RF_LOPositionMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pLOPositionMode As Long) As Long
Declare Function afDigitizerDll_RF_LOPositionMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal LOPositionMode As Long) As Long



Declare Function afDigitizerDll_RF_PreAmpEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEnable As Long) As Long
Declare Function afDigitizerDll_RF_PreAmpEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal enable As Long) As Long

Declare Function afDigitizerDll_RF_AutoPreAmpSelection_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEnable As Long) As Long
Declare Function afDigitizerDll_RF_AutoPreAmpSelection_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal enable As Long) As Long

Declare Function afDigitizerDll_RF_RemoveDCOffset_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRemoveDCOffset As Long) As Long
Declare Function afDigitizerDll_RF_RemoveDCOffset_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal removeDCOffset As Long) As Long

Declare Function afDigitizerDll_RF_Reference_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pReference As Long) As Long
Declare Function afDigitizerDll_RF_Reference_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Reference As Long) As Long
Declare Function afDigitizerDll_RF_ReferenceLocked_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef isLocked As Long) As Long

Declare Function afDigitizerDll_RF_RFAttenuation_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_RFAttenuation_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_RFAttenuationMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFAttenuationMax As Long) As Long
Declare Function afDigitizerDll_RF_RFAttenuationMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFAttenuationMin As Long) As Long
Declare Function afDigitizerDll_RF_RFAttenuationStep_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFAttenuationStep As Long) As Long
Declare Function afDigitizerDll_RF_RFInputLevel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_RFInputLevel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_RFInputLevelMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFInputLevelMax As Double) As Long
Declare Function afDigitizerDll_RF_RFInputLevelMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pRFInputLevelMin As Double) As Long

Declare Function afDigitizerDll_RF_SampleFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSampleFrequency As Double) As Long


Declare Function afDigitizerDll_RF_DitherEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEnable As Long) As Long
Declare Function afDigitizerDll_RF_DitherEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal enable As Long) As Long
Declare Function afDigitizerDll_RF_DitherAvailable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailable As Long) As Long



' RF Channel Properties
Declare Function afDigitizerDll_RF_Channel_CentreFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pCentreFrequency As Double) As Long
Declare Function afDigitizerDll_RF_Channel_CentreFrequency_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal pCentreFrequency As Double) As Long
Declare Function afDigitizerDll_RF_Channel_SetCentreFrequencyAndLOPosition Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal CenterFrequency As Double, ByVal LOPosition As Long) As Long

Declare Function afDigitizerDll_RF_Channel_FrontEndMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pFrontEndMode As Long) As Long
Declare Function afDigitizerDll_RF_Channel_FrontEndMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal FrontEndMode As Long) As Long

Declare Function afDigitizerDll_RF_Channel_IFInputLevel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pIFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_Channel_IFInputLevel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal IFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_Channel_IFInputLevelMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pIFInputLevelMax As Double) As Long
Declare Function afDigitizerDll_RF_Channel_IFInputLevelMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pIFInputLevelMin As Double) As Long
Declare Function afDigitizerDll_RF_Channel_IFAttenuation_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pIFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_Channel_IFAttenuation_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal IFAttenuation As Long) As Long

Declare Function afDigitizerDll_RF_Channel_LevelCorrection_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pLevelCorrection As Double) As Long

Declare Function afDigitizerDll_RF_Channel_LOFrequency_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pLOFrequency As Double) As Long
Declare Function afDigitizerDll_RF_Channel_LOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pLOPosition As Long) As Long
Declare Function afDigitizerDll_RF_Channel_LOPosition_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal LOPosition As Long) As Long
Declare Function afDigitizerDll_RF_Channel_UserLOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pLOPosition As Long) As Long
Declare Function afDigitizerDll_RF_Channel_UserLOPosition_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal LOPosition As Long) As Long
Declare Function afDigitizerDll_RF_Channel_ActualLOPosition_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pLOPosition As Long) As Long


Declare Function afDigitizerDll_RF_Channel_PreAmpEnable_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pEnable As Long) As Long
Declare Function afDigitizerDll_RF_Channel_PreAmpEnable_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal enable As Long) As Long

Declare Function afDigitizerDll_RF_Channel_RFInputLevel_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pRFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_Channel_RFInputLevel_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal RFInputLevel As Double) As Long
Declare Function afDigitizerDll_RF_Channel_RFInputLevelMax_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pRFInputLevelMax As Double) As Long
Declare Function afDigitizerDll_RF_Channel_RFInputLevelMin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pRFInputLevelMin As Double) As Long
Declare Function afDigitizerDll_RF_Channel_RFAttenuation_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByRef pRFAttenuation As Long) As Long
Declare Function afDigitizerDll_RF_Channel_RFAttenuation_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal channel As Integer, ByVal RFAttenuation As Long) As Long

' RF Options Properties
Declare Function afDigitizerDll_RF_Options_AvailableCount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pAvailableCount As Long) As Long

' RF Resource Properties
Declare Function afDigitizerDll_RF_Resource_FPGAConfiguration_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pFPGAConfig As Integer) As Long
Declare Function afDigitizerDll_RF_Resource_FPGAConfiguration_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal FPGAConfig As Integer) As Long
Declare Function afDigitizerDll_RF_Resource_FPGACount_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Integer) As Long

Declare Function afDigitizerDll_RF_Resource_IsActive_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsActive As Long) As Long
Declare Function afDigitizerDll_RF_Resource_IsPlugin_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIsPlugin As Long) As Long

Declare Function afDigitizerDll_RF_Resource_ModelCode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pModelCode As Long) As Long

Declare Function afDigitizerDll_RF_Resource_PluginName_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pluginNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_RF_Resource_PluginName_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal pluginNameBuffer As String) As Long

Declare Function afDigitizerDll_RF_Resource_ResourceString_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal resourceStringBuffer As String, ByVal bufferLen As Long) As Long

Declare Function afDigitizerDll_RF_Resource_SerialNumber_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal SerialNumberBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afDigitizerDll_RF_Resource_SessionID_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSessionId As Long) As Long

Declare Function afDigitizerDll_RF_Resource_Temperature_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTemperature As Double) As Long

' RF Routing Properties
Declare Function afDigitizerDll_RF_Routing_ScenarioListSize_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pScenarioListSize As Long) As Long



' Methods

' General methods
Declare Function afDigitizerDll_RF_OptimizeTemperatureCorrection Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_RF_GetBandwidth Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal centreFreq As Double, ByVal span As Double, ByVal flatness As Long, ByRef pBandwidth As Double) As Long
Declare Function afDigitizerDll_RF_GetRecommendedLOPosition Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal digitizerFreq As Double, ByVal signalFreq As Double, ByRef pLOPosition As Long) As Long
Declare Function afDigitizerDll_RF_GetHighSensitivitySettings Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal rfFreq As Double, ByVal peakLevel As Double, ByRef pRFAttenuationHS As Long, ByRef pIFAttenuationHS As Long, ByRef pPreAmplifierHS As Long) As Long

' RF Options methods
Declare Function afDigitizerDll_RF_Options_Enable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Password As Long) As Long
Declare Function afDigitizerDll_RF_Options_Disable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Password As Long) As Long
Declare Function afDigitizerDll_RF_Options_CheckFitted Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal OptionNumber As Long, ByRef pFitted As Long) As Long
Declare Function afDigitizerDll_RF_Options_Information Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Index As Long, ByRef pOptionNumber As Long, ByVal OptionDescriptionBuffer As String, ByVal bufferLen As Long) As Long

' RF Resource methods
Declare Function afDigitizerDll_RF_Resource_FPGADescriptions Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Numbers As Integer, ByVal Descriptions As String, ByRef pCount As Integer) As Long
Declare Function afDigitizerDll_RF_Resource_GetLastCalibrationDate Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef Year As Integer, ByRef Month As Integer, ByRef Day As Integer, ByRef Hour As Integer, ByRef Minutes As Integer, ByRef Seconds As Integer) As Long

' RF Routing Methods
Declare Function afDigitizerDll_RF_Routing_Reset Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long) As Long
Declare Function afDigitizerDll_RF_Routing_SetConnect Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal MatrixOutput As Long, ByVal MatrixInput As Long) As Long
Declare Function afDigitizerDll_RF_Routing_GetConnect Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal MatrixOutput As Long, ByRef pMatrixInput As Long) As Long
Declare Function afDigitizerDll_RF_Routing_SetOutputEnable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal MatrixOutput As Long, ByVal outputEnable As Long) As Long
Declare Function afDigitizerDll_RF_Routing_GetOutputEnable Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal MatrixOutput As Long, ByRef pOutputEnable As Long) As Long
Declare Function afDigitizerDll_RF_Routing_SetScenario Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RoutingScenario As Long) As Long
Declare Function afDigitizerDll_RF_Routing_AppendScenario Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RoutingScenario As Long) As Long
Declare Function afDigitizerDll_RF_Routing_RemoveScenario Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RoutingScenario As Long) As Long
Declare Function afDigitizerDll_RF_Routing_GetScenarioList Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef ScenarioList As Long, ByVal bufferLen As Long) As Long



'------------------------------------------------------------------------------------------------------
' Timer
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Timer_Advance_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTimer_Advance As Long) As Long
Declare Function afDigitizerDll_Timer_Advance_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Timer_Advance As Long) As Long

Declare Function afDigitizerDll_Timer_Period_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTimer_Period As Double) As Long
Declare Function afDigitizerDll_Timer_Period_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal Timer_Period As Double) As Long

Declare Function afDigitizerDll_Timer_SampleCounterMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSampleCounterMode As Long) As Long
Declare Function afDigitizerDll_Timer_SampleCounterMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal sampleCounterMode As Long) As Long


'------------------------------------------------------------------------------------------------------
' Trigger
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afDigitizerDll_Trigger_Count_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pCount As Long) As Long

Declare Function afDigitizerDll_Trigger_Detected_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTrigger_Detected As Long) As Long

Declare Function afDigitizerDll_Trigger_EdgeGatePolarity_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pEdgeGatePolarity As Long) As Long
Declare Function afDigitizerDll_Trigger_EdgeGatePolarity_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal edgeGatePolarity As Long) As Long

Declare Function afDigitizerDll_Trigger_HoldOff_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pHoldOff As Long) As Long
Declare Function afDigitizerDll_Trigger_HoldOff_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal holdOff As Long) As Long

Declare Function afDigitizerDll_Trigger_IntTriggerAbsTimeConst_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerAbsTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerAbsTimeConst_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerAbsTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerAbsThreshold_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerAbsThreshold As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerAbsThreshold_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerAbsThreshold As Double) As Long

Declare Function afDigitizerDll_Trigger_IntTriggerMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerMode As Long) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerMode As Long) As Long

Declare Function afDigitizerDll_Trigger_IntTriggerSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerSource As Long) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerSource As Long) As Long

Declare Function afDigitizerDll_Trigger_SetIntIQTriggerDigitalBandwidth Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal RequestBandwidth As Double, ByVal SelectionMode As Long, ByRef pAchievedBandwidth As Double) As Long


Declare Function afDigitizerDll_Trigger_IntTriggerRelFastTimeConst_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerRelFastTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerRelFastTimeConst_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerRelFastTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerRelSlowTimeConst_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerRelSlowTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerRelSlowTimeConst_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerRelSlowTimeConst As Double) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerRelThreshold_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pIntTriggerRelThreshold As Long) As Long
Declare Function afDigitizerDll_Trigger_IntTriggerRelThreshold_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal intTriggerRelThreshold As Long) As Long

Declare Function afDigitizerDll_Trigger_OffsetDelay_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pOffsetDelay As Long) As Long
Declare Function afDigitizerDll_Trigger_OffsetDelay_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal offsetDelay As Long) As Long

Declare Function afDigitizerDll_Trigger_PostGateTriggerSamples_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pPostGateTriggerSamples As Long) As Long
Declare Function afDigitizerDll_Trigger_PostGateTriggerSamples_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal postGateTriggerSamples As Long) As Long

Declare Function afDigitizerDll_Trigger_PreEdgeTriggerSamples_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pPreEdgeTriggerSamples As Long) As Long
Declare Function afDigitizerDll_Trigger_PreEdgeTriggerSamples_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal preEdgeTriggerSamples As Long) As Long

Declare Function afDigitizerDll_Trigger_Source_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTrigger_Source As Long) As Long
Declare Function afDigitizerDll_Trigger_Source_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal trigger_Source As Long) As Long

Declare Function afDigitizerDll_Trigger_SwTriggerMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pSwTriggerMode As Long) As Long
Declare Function afDigitizerDll_Trigger_SwTriggerMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal swTriggerMode As Long) As Long

Declare Function afDigitizerDll_Trigger_TType_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTType As Long) As Long
Declare Function afDigitizerDll_Trigger_TType_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal tType As Long) As Long

Declare Function afDigitizerDll_Trigger_UserReTrigSource_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pTrigger_Source As Long) As Long
Declare Function afDigitizerDll_Trigger_UserReTrigSource_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal trigger_Source As Long) As Long

Declare Function afDigitizerDll_Trigger_ReTrigSourceMode_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pMode As Long) As Long
Declare Function afDigitizerDll_Trigger_ReTrigSourceMode_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal mode As Long) As Long

Declare Function afDigitizerDll_Trigger_UserReTrigPolarity_Get Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByRef pPolarity As Long) As Long
Declare Function afDigitizerDll_Trigger_UserReTrigPolarity_Set Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal polarity As Long) As Long

' methods
Declare Function afDigitizerDll_Trigger_Arm Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal numberOfSamples As Long) As Long
Declare Function afDigitizerDll_Trigger_GetTriggerSampleNumber Lib "afDigitizerDll_32.dll" (ByVal digitizerId As Long, ByVal triggerNumber As Long, ByRef pSampleNumber As Long) As Long
