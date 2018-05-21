Attribute VB_Name = "afSigGenDll"
'= Aeroflex afSigGen Component ==================================================
'
' File         afSigGenDll.bas
'
' description  Interface exported by afSigGenDll
'
'================================================================================
'
' Copyright(c) 2000-2014, Aeroflex Ltd.
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

Public Enum afSigGenDll_mtModuleType_t
    afSigGenDll_mtAF3010 = &H3010
    afSigGenDll_mtAF3020 = &H3020
    afSigGenDll_mtAFSIGGEN = -1
    afSigGenDll_mtPlugin = &H0
End Enum



' list control
Public Enum afSigGenDll_liControl_t
    afSigGenDll_licRF = 0
    afSigGenDll_licRFARB = 1
End Enum


' list address source
Public Enum afSigGenDll_liAddressSource_t
    afSigGenDll_liasManual = 0
    afSigGenDll_liasExternal = 1
    afSigGenDll_liasCounter = 2
    afSigGenDll_liasExternalSerial = 3
End Enum


' list address output
Public Enum afSigGenDll_liAddressOut_t
    afSigGenDll_liaoImmediate = 0
    afSigGenDll_liaoDelayed = 1
End Enum

' list arb completion
Public Enum afSigGenDll_liArbCompletion_t
    afSigGenDll_liacOff = 0
    afSigGenDll_liacCurrent = 1
    afSigGenDll_liacAll = 2
End Enum


' list counter mode
Public Enum afSigGenDll_liCounterMode_t
    afSigGenDll_licmStep = 0
    afSigGenDll_licmOnce = 1
    afSigGenDll_licmRunNow = 2
    afSigGenDll_licmRunTrig = 3
End Enum

' list counter strobe
Public Enum afSigGenDll_liCounterStrobe_t
    afSigGenDll_licsExternal = 0
    afSigGenDll_licsTimer = 1
End Enum

' list strobe mode
Public Enum afSigGenDll_liStrobeMode_t
    afSigGenDll_lismManual = 0
    afSigGenDll_lismAuto = 1
End Enum


' List mode Training Levelling Mode
Public Enum afSigGenDll_lmTrainLevellingMode_t
    afSigGenDll_lmtmPeak = 0
    afSigGenDll_lmtmRMS = 1
End Enum



Public Enum afSigGenDll_lolbLoopBandwidth_t
    afSigGenDll_lolbNormal = 0
    afSigGenDll_lolbNarrow = 1
    afSigGenDll_lolbUnspecified = 2
End Enum


Public Enum afSigGenDll_lormReferenceMode_t
    afSigGenDll_lormOCXO = 0
    afSigGenDll_lormInternal = 1
    afSigGenDll_lormExternalDaisy = 2
    afSigGenDll_lormExternalTerminated = 3
End Enum


Public Enum afSigGenDll_lotmTriggerMode_t
    afSigGenDll_lotmNone = 0
    afSigGenDll_lotmAdvance = 1
    afSigGenDll_lotmToggle = 2
    afSigGenDll_lotmHop = 3
End Enum

Public Enum afSigGenDll_tssTriggerSourceSingle_t
    afSigGenDll_tssLBL0 = 0
    afSigGenDll_tssLBR0 = 1
    afSigGenDll_tssPTB0 = 2
    afSigGenDll_tssPXI_STAR = 3
    afSigGenDll_tssPTB1 = 4
    afSigGenDll_tssPTB2 = 5
    afSigGenDll_tssPTB3 = 6
    afSigGenDll_tssPTB4 = 7
    afSigGenDll_tssPTB5 = 8
    afSigGenDll_tssPTB6 = 9
    afSigGenDll_tssPTB7 = 10
    afSigGenDll_tssLBR1 = 11
    afSigGenDll_tssLBR2 = 12
    afSigGenDll_tssLBR3 = 13
    afSigGenDll_tssLBR4 = 14
    afSigGenDll_tssLBR5 = 15
    afSigGenDll_tssLBR6 = 16
    afSigGenDll_tssLBR7 = 17
    afSigGenDll_tssLBR8 = 18
    afSigGenDll_tssLBR9 = 19
    afSigGenDll_tssLBR10 = 20
    afSigGenDll_tssLBR11 = 21
    afSigGenDll_tssLBR12 = 22
End Enum

Public Enum afSigGenDll_hmetsHopModeExtTrigSrc_t
    afSigGenDll_hmetsPXI_STAR = 0
    afSigGenDll_hmetsLBR1 = 1
    afSigGenDll_hmetsLBR2 = 2
    afSigGenDll_hmetsLBR3 = 3
    afSigGenDll_hmetsLBR4 = 4
    afSigGenDll_hmetsLBR5 = 5
    afSigGenDll_hmetsLBR6 = 6
    afSigGenDll_hmetsLBR7 = 7
    afSigGenDll_hmetsLBR8 = 8
    afSigGenDll_hmetsLBR9 = 9
    afSigGenDll_hmetsLBR10 = 10
    afSigGenDll_hmetsLBR11 = 11
    afSigGenDll_hmetsLBR12 = 12
    afSigGenDll_hmetsSMB = 13
End Enum


Public Enum afSigGenDll_tsaTriggerSourceAddressed_t
    afSigGenDll_tsaNONE = 0
    afSigGenDll_tsaTRIG = 1
    afSigGenDll_tsaLBR = 2
    afSigGenDll_tsaSER = 3
End Enum


Public Enum afSigGenDll_msModulationSource_t
    afSigGenDll_msCW = 3
    afSigGenDll_msLVDS = 0
    afSigGenDll_msARB = 1
    afSigGenDll_msAM = 4
    afSigGenDll_msFM = 5
    afSigGenDll_msExtAnalog = 6
End Enum


Public Enum afSigGenDll_ibcIqBwCorrMode_t
    afSigGenDll_ibcOff = 0
    afSigGenDll_ibcManual = 1
End Enum


Public Enum afSigGenDll_lmLevelMode_t
    afSigGenDll_lmAuto = 0
    afSigGenDll_lmFrozen = 1
    afSigGenDll_lmPeak = 2
    afSigGenDll_lmRms = 3
End Enum


Public Enum afSigGenDll_ltLevelType_t
    afSigGenDll_ltPeak = 2
    afSigGenDll_ltRMS = 3
End Enum


Public Enum afSigGenDll_rmRoutingMatrix_t
    afSigGenDll_rmPXI_TRIG_0 = 0
    afSigGenDll_rmPXI_TRIG_1 = 1
    afSigGenDll_rmPXI_TRIG_2 = 2
    afSigGenDll_rmPXI_TRIG_3 = 3
    afSigGenDll_rmPXI_TRIG_4 = 4
    afSigGenDll_rmPXI_TRIG_5 = 5
    afSigGenDll_rmPXI_TRIG_6 = 6
    afSigGenDll_rmPXI_TRIG_7 = 7
    afSigGenDll_rmPXI_STAR = 8
    afSigGenDll_rmPXI_LBL_0 = 9
    afSigGenDll_rmPXI_LBL_1 = 10
    afSigGenDll_rmPXI_LBL_2 = 11
    afSigGenDll_rmPXI_LBL_3 = 12
    afSigGenDll_rmPXI_LBL_4 = 13
    afSigGenDll_rmPXI_LBL_5 = 14
    afSigGenDll_rmPXI_LBL_6 = 15
    afSigGenDll_rmPXI_LBL_7 = 16
    afSigGenDll_rmPXI_LBL_8 = 17
    afSigGenDll_rmPXI_LBL_9 = 18
    afSigGenDll_rmPXI_LBL_10 = 19
    afSigGenDll_rmPXI_LBL_11 = 20
    afSigGenDll_rmPXI_LBL_12 = 21
    afSigGenDll_rmLVDS_MARKER_1 = 22
    afSigGenDll_rmLVDS_MARKER_2 = 23
    afSigGenDll_rmLVDS_MARKER_3 = 24
    afSigGenDll_rmLVDS_MARKER_4 = 25
    afSigGenDll_rmLVDS_AUX_0 = 26
    afSigGenDll_rmLVDS_AUX_1 = 27
    afSigGenDll_rmLVDS_AUX_2 = 28
    afSigGenDll_rmLVDS_AUX_3 = 29
    afSigGenDll_rmLVDS_AUX_4 = 30
    afSigGenDll_rmLVDS_SPARE_0 = 31
    afSigGenDll_rmLVDS_SPARE_1 = 32
    afSigGenDll_rmLVDS_SPARE_2 = 33
    afSigGenDll_rmARB_MARKER_1 = 34
    afSigGenDll_rmARB_MARKER_2 = 35
    afSigGenDll_rmARB_MARKER_3 = 36
    afSigGenDll_rmARB_MARKER_4 = 37
    afSigGenDll_rmARB_TRIG = 38
    afSigGenDll_rmLA_OUT_0 = 39
    afSigGenDll_rmLA_OUT_1 = 40
    afSigGenDll_rmLA_OUT_2 = 41
    afSigGenDll_rmLA_OUT_3 = 42
    afSigGenDll_rmLA_OUT_4 = 43
    afSigGenDll_rmLA_OUT_5 = 44
    afSigGenDll_rmLA_OUT_6 = 45
    afSigGenDll_rmLA_OUT_7 = 46
    afSigGenDll_rmLSTB_OUT = 47
    afSigGenDll_rmLA_IN_0 = 48
    afSigGenDll_rmLA_IN_1 = 49
    afSigGenDll_rmLA_IN_2 = 50
    afSigGenDll_rmLA_IN_3 = 51
    afSigGenDll_rmLA_IN_4 = 52
    afSigGenDll_rmLA_IN_5 = 53
    afSigGenDll_rmLA_IN_6 = 54
    afSigGenDll_rmLA_IN_7 = 55
    afSigGenDll_rmRFOFF_EXT = 56
    afSigGenDll_rmMODOFF_EXT = 57
    afSigGenDll_rmFREEZE_EXT = 58
    afSigGenDll_rmGND = 59
    afSigGenDll_rmSeqStart = 60
    afSigGenDll_rmRfBlank = 61
    afSigGenDll_rmLSTB_IN = 62
    afSigGenDll_rmFRONT_SMB = 63
    afSigGenDll_rmSW = 64
    afSigGenDll_rmLA_SERIAL = 65
    afSigGenDll_rmTRIG_GATE_EN = 66
    afSigGenDll_rmTRIG_GATE_SIG = 67
    afSigGenDll_rmTRIG_GATE_OUT = 68
    afSigGenDll_rmGRP_SEQ_INSERT = 71
End Enum


Public Enum afSigGenDll_rsRoutingScenario_t
    afSigGenDll_rsNONE = 0
    afSigGenDll_rsLVDS_AUX_TO_PXI_LBL = 1
    afSigGenDll_rsLVDS_MKR1_TO_RFOFF = 2
    afSigGenDll_rsLVDS_MKR2_TO_RFOFF = 3
    afSigGenDll_rsLVDS_MKR3_TO_RFOFF = 4
    afSigGenDll_rsLVDS_MKR4_TO_RFOFF = 4
    afSigGenDll_rsLVDS_MKR1_TO_ARBTRIG = 5
    afSigGenDll_rsLVDS_MKR2_TO_ARBTRIG = 6
    afSigGenDll_rsLVDS_MKR3_TO_ARBTRIG = 7
    afSigGenDll_rsLVDS_MKR4_TO_ARBTRIG = 8
    afSigGenDll_rsPXI_STAR_TO_ARBTRIG = 9
    afSigGenDll_rsPXI_TRIG0_TO_ARBTRIG = 10
    afSigGenDll_rsPXI_TRIG1_TO_ARBTRIG = 11
    afSigGenDll_rsPXI_TRIG2_TO_ARBTRIG = 12
    afSigGenDll_rsPXI_TRIG3_TO_ARBTRIG = 13
    afSigGenDll_rsPXI_TRIG4_TO_ARBTRIG = 14
    afSigGenDll_rsPXI_TRIG5_TO_ARBTRIG = 15
    afSigGenDll_rsPXI_TRIG6_TO_ARBTRIG = 16
    afSigGenDll_rsPXI_TRIG7_TO_ARBTRIG = 17
End Enum


' Signal Generator Mode
Public Enum afSigGenDll_mMode_t
    afSigGenDll_mManual = 0
    afSigGenDll_mArbSeq = 1
    afSigGenDll_mHopping = 2
    afSigGenDll_mFull = 3
    afSigGenDll_mGroupSeq = 4
End Enum


Public Enum afSigGenDll_lcmLevelControlMode_t

    afSigGenDll_lcmAbsolute = 0
    afSigGenDll_lcmRelative = 1
End Enum


' Arb Sequencing/Hopping Repeat Mode
Public Enum afSigGenDll_rmRepeatMode_t

    afSigGenDll_rmSingle = 0
    afSigGenDll_rmNTimes = 1
    afSigGenDll_rmContinuous = 2
End Enum


' Arb Sequencing/Hopping Trigger Polarity
Public Enum afSigGenDll_tpTriggerPolarity_t
    afSigGenDll_tpNegative = 0
    afSigGenDll_tpPositive = 1
End Enum


' Arb Sequencing Trigger Type
Public Enum afSigGenDll_asTriggerType_t
    afSigGenDll_asSoftware = 0
    afSigGenDll_asExternal = 1
End Enum


' hopping Trigger Type
Public Enum afSigGenDll_hTriggerType_t
    afSigGenDll_hSoftware = 0
    afSigGenDll_hExternal = 1
    afSigGenDll_hManual = 2
End Enum


' Address source Type
Public Enum afSigGenDll_asAddressSource_t
    afSigGenDll_asInternal = 0
    afSigGenDll_asLVDS = 1
    afSigGenDll_asPXITriggerBus = 2
    afSigGenDll_asARBMarkers = 3
    afSigGenDll_asSerial = 4
End Enum

' hopping IntListStrobeSource Type
Public Enum afSigGenDll_ilssIntListStrobeSource_t
    afSigGenDll_ilssSoftware = 0
    afSigGenDll_ilssTimer = 1
End Enum



' 3025C/3026C Specific
Public Enum afSigGenDll_dcmDdsClockMode_t
    afSigGenDll_dcmFast = 0
    afSigGenDll_dcmLowNoise = 1
End Enum


' *****************************************************************************
'  Group Sequencing Types
' *****************************************************************************

Public Enum afSigGenDll_gsigGrpSeqInsAtGroup_t
    afSigGenDll_gsigAny = 0
    afSigGenDll_gsigPrimary = 3
    afSigGenDll_gsigAlternate = 2
End Enum





' *****************************************************************************
'  Enhanced ARB Control Types
' *****************************************************************************

Public Enum afSigGenDll_amArbControlMode_t
    afSigGenDll_amStandard = 0
    afSigGenDll_amEnhanced = 1
End Enum

Public Enum afSigGenDll_eaEnhArbTrigMode_t
    afSigGenDll_eatmGate = 0
    afSigGenDll_eatmStartOnly = 1
    afSigGenDll_eatmStartStop = 2
    afSigGenDll_eatmStartReTrig = 3
End Enum

Public Enum afSigGenDll_eaEnhArbExtTrigPolarity_t
    afSigGenDll_eaetPositiveEdge = 1
    afSigGenDll_eaetNegativeEdge = 2
    afSigGenDll_eaetAnyEdge = 3
End Enum


Public Enum afSigGenDll_eaEnhArbPlayTermination_t
    afSigGenDll_eaptImmediate = 0
    afSigGenDll_eaptAtEnd = 1
End Enum

Public Enum afSigGenDll_eaEnhArbPlayMode_t
    afSigGenDll_eaSingle = 0
    afSigGenDll_eaNTimes = 1
    afSigGenDll_eaContinuous = 2
End Enum

' Enhanced ARB Play status enumeration
Public Enum afSigGenDll_eaEnhArbRunStatus_t
    afSigGenDll_earsNotPlaying = 0
    afSigGenDll_earsWaitingForTrigger = 1
    afSigGenDll_earsPlaying = 2
    afSigGenDll_earsStopping = 3
End Enum


' ARB Trigger Sync Loopback Direction
Public Enum afSigGenDll_loopbackDirection_t
    afSigGenDll_master_primary_forward = 0
    afSigGenDll_master_secondary_forward = 1
    afSigGenDll_slave_forward = 2
    afSigGenDll_slave_reverse = 3
End Enum

' ARB Trigger sync. Branches in use on the Master Module
Public Enum afSigGenDll_master_branches_t
    afSigGenDll_primary_branch_only = &H1
    afSigGenDll_secondary_branch_only = &H2
    afSigGenDll_primary_and_secondary_branches = &H3
End Enum







'------------------------------------------------------------------------------------------------------
' Object eeprom Cache policy control
'------------------------------------------------------------------------------------------------------
Declare Function afSigGenDll_EepromCacheEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_EepromCacheEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pState As Long) As Long

Declare Function afSigGenDll_EepromCachePathLength_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLen As Long) As Long
Declare Function afSigGenDll_EepromCachePath_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal pathBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_EepromCachePath_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal buffer As String) As Long



'------------------------------------------------------------------------------------------------------
' Object creation and destruction
'------------------------------------------------------------------------------------------------------
' methods
Declare Function afSigGenDll_CreateObject Lib "afSigGenDll_32.dll" (ByRef pSigGenId As Long) As Long
Declare Function afSigGenDll_DestroyObject Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' Errors Information
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ErrorCode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pErrorCode As Long) As Long
Declare Function afSigGenDll_ErrorMessage_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal messageBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_ErrorSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pErrorSource As Long) As Long



'------------------------------------------------------------------------------------------------------
' General
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
' Properties - Synchronise RF channel switching speed between 3010 and 3020. Only applies if 3010 doesn't have Option 1 fitted
Declare Function afSigGenDll_LoRfSpeedSyncEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal enable As Long) As Long
Declare Function afSigGenDll_LoRfSpeedSyncEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pEnable As Long) As Long
' Properties - Select between Manual, Hopping and ArbSeq
Declare Function afSigGenDll_Mode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_Mode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long
' Properties - Select between Standard and Enhanced ARB Control
Declare Function afSigGenDll_ARBControlMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_ARBControlMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long


' Methods
Declare Function afSigGenDll_BootInstrument Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal LoResource As String, ByVal RfResource As String, ByVal LoIsPlugin As Long) As Long
Declare Function afSigGenDll_ClearErrors Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_CloseInstrument Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_GetVersion Lib "afSigGenDll_32.dll" (ByRef version As Long) As Long



'------------------------------------------------------------------------------------------------------
' Manual Mode Functions - Equivalent to a single channel
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_Manual_ModulationSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal src As Long) As Long
Declare Function afSigGenDll_Manual_ModulationSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSrc As Long) As Long
Declare Function afSigGenDll_Manual_AM_ModulationDepth_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal depth As Long) As Long
Declare Function afSigGenDll_Manual_AM_ModulationDepth_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDepth As Long) As Long
Declare Function afSigGenDll_Manual_AM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rate As Long) As Long
Declare Function afSigGenDll_Manual_AM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRate As Long) As Long
Declare Function afSigGenDll_Manual_FM_ModulationDeviation_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal deviation As Long) As Long
Declare Function afSigGenDll_Manual_FM_ModulationDeviation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDeviation As Long) As Long
Declare Function afSigGenDll_Manual_FM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rate As Long) As Long
Declare Function afSigGenDll_Manual_FM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRate As Long) As Long
Declare Function afSigGenDll_Manual_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal frequency As Double) As Long
Declare Function afSigGenDll_Manual_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_Manual_FrequencyMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_Manual_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal level As Double) As Long
Declare Function afSigGenDll_Manual_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_Manual_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_Manual_LevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLevelMode As Long) As Long
Declare Function afSigGenDll_Manual_LevelMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLevelMax As Double) As Long
Declare Function afSigGenDll_Manual_LevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLevelMax As Double) As Long
Declare Function afSigGenDll_Manual_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_Manual_RMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRMSdBc As Double) As Long
Declare Function afSigGenDll_Manual_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal ArbFile As String) As Long
Declare Function afSigGenDll_Manual_ArbFile_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbFileBuffer As String, ByVal arbFileBufLen As Long) As Long
Declare Function afSigGenDll_Manual_ArbFileNameLength_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pArbFileBufLen As Long) As Long
Declare Function afSigGenDll_Manual_ArbStopPlaying Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Manual_RfGating_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rfGating As Long) As Long
Declare Function afSigGenDll_Manual_RfGating_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRfGating As Long) As Long
Declare Function afSigGenDll_Manual_RfState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_Manual_RfState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pState As Long) As Long

' Methods
Declare Function afSigGenDll_Manual_Reset Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long


'------------------------------------------------------------------------------------------------------
' Hopping Mode Functions
'------------------------------------------------------------------------------------------------------
' General configuration properties
Declare Function afSigGenDll_Hopping_AM_ModulationDepth_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal depth As Long) As Long
Declare Function afSigGenDll_Hopping_AM_ModulationDepth_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDepth As Long) As Long
Declare Function afSigGenDll_Hopping_AM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rate As Long) As Long
Declare Function afSigGenDll_Hopping_AM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRate As Long) As Long
Declare Function afSigGenDll_Hopping_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbfileName As String) As Long
Declare Function afSigGenDll_Hopping_ArbFile_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbFileBuffer As String, ByVal arbFileBufLen As Long) As Long
Declare Function afSigGenDll_Hopping_ArbFileNameLength_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pArbFileNameLen As Long) As Long
Declare Function afSigGenDll_Hopping_BaseLevel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal baseLevel As Double) As Long
Declare Function afSigGenDll_Hopping_BaseLevel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pBaseLevel As Double) As Long
Declare Function afSigGenDll_Hopping_FM_ModulationDeviation_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal deviation As Long) As Long
Declare Function afSigGenDll_Hopping_FM_ModulationDeviation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDeviation As Long) As Long
Declare Function afSigGenDll_Hopping_FM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rate As Long) As Long
Declare Function afSigGenDll_Hopping_FM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRate As Long) As Long
Declare Function afSigGenDll_Hopping_LevelControlMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal levelControlMode As Long) As Long
Declare Function afSigGenDll_Hopping_LevelControlMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLevelControlMode As Long) As Long
Declare Function afSigGenDll_Hopping_ModulationSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal src As Long) As Long
Declare Function afSigGenDll_Hopping_ModulationSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSrc As Long) As Long
Declare Function afSigGenDll_Hopping_RfGating_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rfGating As Long) As Long
Declare Function afSigGenDll_Hopping_RfGating_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRfGating As Long) As Long
Declare Function afSigGenDll_Hopping_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_Hopping_RMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRMSdBc As Double) As Long

' List Iteration properties
Declare Function afSigGenDll_Hopping_RepeatCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal count As Integer) As Long
Declare Function afSigGenDll_Hopping_RepeatCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_Hopping_RepeatMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_Hopping_RepeatMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long

' List Address properties
Declare Function afSigGenDll_Hopping_AddressSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal addrSource As Long) As Long
Declare Function afSigGenDll_Hopping_AddressSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAddrSource As Long) As Long
Declare Function afSigGenDll_Hopping_ExportAddressToPXITrig_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal exportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_Hopping_ExportAddressToPXITrig_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_Hopping_SerialAddressSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal serAddrSource As Long) As Long
Declare Function afSigGenDll_Hopping_SerialAddressSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSerAddrSource As Long) As Long
Declare Function afSigGenDll_Hopping_StartAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal startAddress As Integer) As Long
Declare Function afSigGenDll_Hopping_StartAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStartAddress As Integer) As Long
Declare Function afSigGenDll_Hopping_StopAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal stopAddress As Integer) As Long
Declare Function afSigGenDll_Hopping_StopAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStopAddress As Integer) As Long

' List Strobe properties
Declare Function afSigGenDll_Hopping_IntListStrobeSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal source As Long) As Long
Declare Function afSigGenDll_Hopping_IntListStrobeSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSource As Long) As Long
Declare Function afSigGenDll_Hopping_LvdsListStrobe_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listStrobe As Long) As Long
Declare Function afSigGenDll_Hopping_LvdsListStrobe_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pListStrobe As Long) As Long

' Trigger configuration properties
Declare Function afSigGenDll_Hopping_ExtTriggerPolarity_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerPolarity As Long) As Long
Declare Function afSigGenDll_Hopping_ExtTriggerPolarity_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerPolarity As Long) As Long
Declare Function afSigGenDll_Hopping_ExtTriggerSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal extTriggerSource As Long) As Long
Declare Function afSigGenDll_Hopping_ExtTriggerSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExtTriggerSource As Long) As Long
Declare Function afSigGenDll_Hopping_TriggerType_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerType As Long) As Long
Declare Function afSigGenDll_Hopping_TriggerType_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerType As Long) As Long

' Properties applied to ALL channels
Declare Function afSigGenDll_Hopping_ArbModulationState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_DwellTime_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal dwellTime As Long) As Long
Declare Function afSigGenDll_Hopping_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal frequency As Double) As Long
Declare Function afSigGenDll_Hopping_FrequencyMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_Hopping_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal level As Double) As Long
Declare Function afSigGenDll_Hopping_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_Hopping_RelLevel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal relLevel As Double) As Long
Declare Function afSigGenDll_Hopping_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_RfState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_TrainLevellingMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_Hopping_TrainLevellingMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long

' Properties applied to INDIVIDUAL channels
Declare Function afSigGenDll_Hopping_Channel_ArbModulationState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_ArbModulationState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pState As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_DwellTime_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal dwellTime As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_DwellTime_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pDwellTime As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal frequency As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal level As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_LevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_LevelMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_LevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelMode As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_RelLevel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal relLevel As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_RelLevel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pRelLevel As Double) As Long
Declare Function afSigGenDll_Hopping_Channel_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_RfFrozenState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pState As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_RfState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_Hopping_Channel_RfState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pState As Long) As Long

' Methods
Declare Function afSigGenDll_Hopping_Abort Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Hopping_GotoAddress Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Address As Integer) As Long
Declare Function afSigGenDll_Hopping_Reset Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Hopping_Train Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Hopping_TriggerArm Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Hopping_TriggerNow Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long

' Status Properties
Declare Function afSigGenDll_Hopping_CurrentAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentAddress As Integer) As Long
Declare Function afSigGenDll_Hopping_CurrentRepeatCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentRepeatCount As Integer) As Long
Declare Function afSigGenDll_Hopping_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_Hopping_IsArmed_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsArmed As Long) As Long



'------------------------------------------------------------------------------------------------------
' ARB Sequencing Mode Functions
'------------------------------------------------------------------------------------------------------

' General configuration properties
Declare Function afSigGenDll_ArbSeq_ExportAddressToPXITrig_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal exportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_ArbSeq_ExportAddressToPXITrig_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_ArbSeq_RepeatMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_ArbSeq_RepeatMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long
Declare Function afSigGenDll_ArbSeq_RepeatCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal count As Integer) As Long
Declare Function afSigGenDll_ArbSeq_RepeatCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_ArbSeq_RfGating_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rfGating As Long) As Long
Declare Function afSigGenDll_ArbSeq_RfGating_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRfGating As Long) As Long
Declare Function afSigGenDll_ArbSeq_StartAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal startAddress As Integer) As Long
Declare Function afSigGenDll_ArbSeq_StartAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStartAddress As Integer) As Long
Declare Function afSigGenDll_ArbSeq_StopAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal stopAddress As Integer) As Long
Declare Function afSigGenDll_ArbSeq_StopAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStopAddress As Integer) As Long
Declare Function afSigGenDll_ArbSeq_TrainLevellingMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_ArbSeq_TrainLevellingMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long

' Trigger configuration properties
Declare Function afSigGenDll_ArbSeq_ExtTriggerPolarity_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerPolarity As Long) As Long
Declare Function afSigGenDll_ArbSeq_ExtTriggerPolarity_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerPolarity As Long) As Long
Declare Function afSigGenDll_ArbSeq_ExtTriggerSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal extTriggerSource As Long) As Long
Declare Function afSigGenDll_ArbSeq_ExtTriggerSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExtTriggerSource As Long) As Long
Declare Function afSigGenDll_ArbSeq_TriggerType_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerType As Long) As Long
Declare Function afSigGenDll_ArbSeq_TriggerType_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerType As Long) As Long

' Properties applied to ALL channels
Declare Function afSigGenDll_ArbSeq_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbfileName As String) As Long
Declare Function afSigGenDll_ArbSeq_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal frequency As Double) As Long
Declare Function afSigGenDll_ArbSeq_FrequencyMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_ArbSeq_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal level As Double) As Long
Declare Function afSigGenDll_ArbSeq_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_ArbSeq_PlayCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal playCount As Integer) As Long
Declare Function afSigGenDll_ArbSeq_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_ArbSeq_RfState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_ArbSeq_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RMSdBc As Double) As Long

' Properties applied to INDIVIDUAL channels
Declare Function afSigGenDll_ArbSeq_Channel_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal ArbFile As String) As Long
Declare Function afSigGenDll_ArbSeq_Channel_ArbFile_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal arbFileBuffer As String, ByVal arbFileBufLen As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_ArbFileNameLength_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pArbFileBufLen As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_PlayCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal playCount As Integer) As Long
Declare Function afSigGenDll_ArbSeq_Channel_PlayCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pPlayCount As Integer) As Long
Declare Function afSigGenDll_ArbSeq_Channel_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal frequency As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal level As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_LevelMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_LevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_LevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelMode As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RfFrozenState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pState As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pRMSdBc As Double) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RfState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_ArbSeq_Channel_RfState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pState As Long) As Long

' Methods
Declare Function afSigGenDll_ArbSeq_Abort Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ArbSeq_Reset Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ArbSeq_Train Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ArbSeq_TriggerArm Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ArbSeq_TriggerNow Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long

' Status Properties
Declare Function afSigGenDll_ArbSeq_CurrentAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentAddress As Integer) As Long
Declare Function afSigGenDll_ArbSeq_CurrentRepeatCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentRepeatCount As Integer) As Long
Declare Function afSigGenDll_ArbSeq_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_ArbSeq_IsArmed_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsArmed As Long) As Long



'------------------------------------------------------------------------------------------------------
' Group Sequencing Mode Functions
'------------------------------------------------------------------------------------------------------

' General configuration properties
Declare Function afSigGenDll_GroupSeq_ExportAddressToPXITrig_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal exportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_GroupSeq_ExportAddressToPXITrig_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExportAddressToPXITrig As Long) As Long
Declare Function afSigGenDll_GroupSeq_GroupCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal GroupCount As Integer) As Long
Declare Function afSigGenDll_GroupSeq_GroupCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pGroupCount As Integer) As Long
Declare Function afSigGenDll_GroupSeq_PrimaryGroup_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer) As Long
Declare Function afSigGenDll_GroupSeq_PrimaryGroup_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIndex As Integer) As Long
Declare Function afSigGenDll_GroupSeq_AlternateGroup_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer) As Long
Declare Function afSigGenDll_GroupSeq_AlternateGroup_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIndex As Integer) As Long
Declare Function afSigGenDll_GroupSeq_ToggleEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal enable As Long) As Long
Declare Function afSigGenDll_GroupSeq_ToggleEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pEnable As Long) As Long
Declare Function afSigGenDll_GroupSeq_StartAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal startAddress As Integer) As Long
Declare Function afSigGenDll_GroupSeq_StartAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStartAddress As Integer) As Long
Declare Function afSigGenDll_GroupSeq_StopAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal stopAddress As Integer) As Long
Declare Function afSigGenDll_GroupSeq_StopAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStopAddress As Integer) As Long
Declare Function afSigGenDll_GroupSeq_TrainLevellingMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_GroupSeq_TrainLevellingMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long
Declare Function afSigGenDll_GroupSeq_RfGating_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal rfGating As Long) As Long
Declare Function afSigGenDll_GroupSeq_RfGating_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRfGating As Long) As Long
Declare Function afSigGenDll_GroupSeq_FrequencyMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long

' Trigger configuration properties
Declare Function afSigGenDll_GroupSeq_ExtTriggerPolarity_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerPolarity As Long) As Long
Declare Function afSigGenDll_GroupSeq_ExtTriggerPolarity_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerPolarity As Long) As Long
Declare Function afSigGenDll_GroupSeq_ExtTriggerSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal extTriggerSource As Long) As Long
Declare Function afSigGenDll_GroupSeq_ExtTriggerSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExtTriggerSource As Long) As Long
Declare Function afSigGenDll_GroupSeq_TriggerType_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerType As Long) As Long
Declare Function afSigGenDll_GroupSeq_TriggerType_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerType As Long) As Long

' Properties applied to ALL channels in ALL Groups
Declare Function afSigGenDll_GroupSeq_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbfileName As String) As Long
Declare Function afSigGenDll_GroupSeq_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal frequency As Double) As Long
Declare Function afSigGenDll_GroupSeq_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal level As Double) As Long
Declare Function afSigGenDll_GroupSeq_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_GroupSeq_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal state As Long) As Long
Declare Function afSigGenDll_GroupSeq_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RMSdBc As Double) As Long

' Properties applied to ALL channels in an INDIVIDUAL Groups
Declare Function afSigGenDll_GroupSeq_Group_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal arbfileName As String) As Long
Declare Function afSigGenDll_GroupSeq_Group_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal frequency As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal level As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_NumActiveChannels_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByRef pCount As Integer) As Long

' Properties applied to INDIVIDUAL channels
Declare Function afSigGenDll_GroupSeq_Group_Channel_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal ArbFile As String) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_ArbFile_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal arbFileBuffer As String, ByVal arbFileBufLen As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_ArbFileNameLength_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pArbFileBufLen As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal frequency As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal level As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_LevelMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_LevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_LevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pLevelMode As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_RfFrozenState_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal state As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_RfFrozenState_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pState As Long) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_GroupSeq_Group_Channel_RMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal group As Integer, ByVal channel As Integer, ByRef pRMSdBc As Double) As Long

' Methods
Declare Function afSigGenDll_GroupSeq_Abort Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_GroupSeq_Reset Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_GroupSeq_Train Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_GroupSeq_TriggerArm Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_GroupSeq_TriggerNow Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long

Declare Function afSigGenDll_GroupSeq_SeqInsert Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal GroupId As Integer, ByVal SeqCount As Integer) As Long
Declare Function afSigGenDll_GroupSeq_SeqInsertAtIndex Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer, ByVal fromGroupId As Integer, ByVal SeqCount As Integer) As Long
Declare Function afSigGenDll_GroupSeq_SeqInsertAtGroup Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal ActiveGroup As Long, ByVal fromGroupId As Integer, ByVal SeqCount As Integer) As Long

Declare Function afSigGenDll_GroupSeq_SeqInsertIndefinite Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal IndefiniteGroupId As Integer) As Long
Declare Function afSigGenDll_GroupSeq_SeqInsertIndefiniteAtIndex Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer, ByVal IndefiniteGroupId As Integer) As Long
Declare Function afSigGenDll_GroupSeq_SeqInsertIndefiniteAtGroup Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal ActiveGroup As Long, ByVal IndefiniteGroupId As Integer) As Long
Declare Function afSigGenDll_GroupSeq_InsertIndefiniteCancel Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long

Declare Function afSigGenDll_GroupSeq_WaitOnNewSeqInsertReady Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal timeoutMillisecs As Long) As Long
Declare Function afSigGenDll_GroupSeq_WaitOnCurrSeqInsertComplete Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal timeoutMillisecs As Long) As Long
Declare Function afSigGenDll_GroupSeq_WaitOnInsertIndefiniteCancel Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal timeoutMillisecs As Long) As Long
Declare Function afSigGenDll_GroupSeq_WaitOnInsertIndefiniteExecuting Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal timeoutMillisecs As Long) As Long


' Status Properties
Declare Function afSigGenDll_GroupSeq_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_GroupSeq_IsArmed_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsArmed As Long) As Long
Declare Function afSigGenDll_GroupSeq_CurrSeqInsertComplete_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pState As Long) As Long
Declare Function afSigGenDll_GroupSeq_NewSeqInsertReady_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pState As Long) As Long
Declare Function afSigGenDll_GroupSeq_InsertIndefiniteExecuting_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pState As Long) As Long






'------------------------------------------------------------------------------------------------------
' ARB
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ARB_FilePlaying_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal fileNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_ARB_IsPlaying_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsPlaying As Long) As Long
Declare Function afSigGenDll_ARB_SingleShotMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal singleShotMode As Long) As Long
Declare Function afSigGenDll_ARB_SingleShotMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSingleShotMode As Long) As Long

' Methods
Declare Function afSigGenDll_ARB_StopPlaying Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ARB_SingleShotTrigger Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' ARB Catalogue
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ARB_Catalogue_FileCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFileCount As Integer) As Long

' Methods
Declare Function afSigGenDll_ARB_Catalogue_AddFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_ARB_Catalogue_DeleteAllFiles Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ARB_Catalogue_DeleteFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_ARB_Catalogue_FindFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_ARB_Catalogue_GetFileNameByIndex Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer, ByVal fileNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_ARB_Catalogue_GetFileSampleRate Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String, ByRef pSampleRate As Long) As Long
Declare Function afSigGenDll_ARB_Catalogue_PlayFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_ARB_Catalogue_ReloadAllFiles Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' ARB External Trigger
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ARB_ExternalTrigger_Enable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pEnabled As Long) As Long
Declare Function afSigGenDll_ARB_ExternalTrigger_Enable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal enabled As Long) As Long
Declare Function afSigGenDll_ARB_ExternalTrigger_Gated_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pGated As Long) As Long
Declare Function afSigGenDll_ARB_ExternalTrigger_Gated_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal gated As Long) As Long
Declare Function afSigGenDll_ARB_ExternalTrigger_NegativeEdge_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pNegativeEdge As Long) As Long
Declare Function afSigGenDll_ARB_ExternalTrigger_NegativeEdge_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal negativeEdge As Long) As Long



'------------------------------------------------------------------------------------------------------
' Enhanced ARB Control - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_EnhARB_PlayCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal count As Integer) As Long
Declare Function afSigGenDll_EnhARB_PlayCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_EnhARB_PlayMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_EnhARB_PlayMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long
Declare Function afSigGenDll_EnhARB_TerminationMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_EnhARB_TerminationMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long
Declare Function afSigGenDll_EnhARB_Status_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStatus As Long) As Long
Declare Function afSigGenDll_EnhARB_FilePlaying_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal fileNameBuffer As String, ByVal bufferLen As Long) As Long

' Methods
Declare Function afSigGenDll_EnhARB_AbortPlaying Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_EnhARB_StopPlaying Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' Enhanced ARB Control, Catalogue - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_EnhARB_Catalogue_FileCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFileCount As Integer) As Long

' Methods
Declare Function afSigGenDll_EnhARB_Catalogue_AddFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_DeleteAllFiles Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_DeleteFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_FindFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_GetFileNameByIndex Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Integer, ByVal fileNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_GetFileSampleRate Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String, ByRef pSampleRate As Long) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_PlayFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_EnhARB_Catalogue_ReloadAllFiles Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long


'------------------------------------------------------------------------------------------------------
' Enhanced ARB Control, External Trigger - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Delay_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal delay As Double) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Delay_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDelay As Double) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Enable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal enable As Long) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Enable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pEnable As Long) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Mode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerMode As Long) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Mode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerMode As Long) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Polarity_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal polarity As Long) As Long
Declare Function afSigGenDll_EnhARB_ExternalTrigger_Polarity_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pPolarity As Long) As Long


'------------------------------------------------------------------------------------------------------
' Calibrate Detector
'------------------------------------------------------------------------------------------------------
' Methods
Declare Function afSigGenDll_Calibrate_Detector_Zero Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' Calibrate AnalogIQin
'------------------------------------------------------------------------------------------------------
' Methods
Declare Function afSigGenDll_Calibrate_AnalogIQin_Inputs Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long


'------------------------------------------------------------------------------------------------------
' Calibrate DifferentialIQ
'------------------------------------------------------------------------------------------------------
' Methods
Declare Function afSigGenDll_Calibrate_DifferentialIQ_Outputs Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' Calibrate IQ
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_Calibrate_IQ_NumberOfBands_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pNumBands As Long) As Long

' Methods
Declare Function afSigGenDll_Calibrate_IQ_AllBands Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Calibrate_IQ_CurrentFrequency Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_Calibrate_IQ_GetBandInformation Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Band As Long, ByRef pStartFrequency As Double, ByRef pStopFrequency As Double) As Long
Declare Function afSigGenDll_Calibrate_IQ_RestoreBandedFromFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_Calibrate_IQ_RestoreSingleFromFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_Calibrate_IQ_SelectedBand Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Band As Long) As Long
Declare Function afSigGenDll_Calibrate_IQ_StoreBandedToFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long
Declare Function afSigGenDll_Calibrate_IQ_StoreSingleToFile Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FileName As String) As Long



'------------------------------------------------------------------------------------------------------
' AnalogIQin
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_AnalogIQin_Input50ohms_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAnalogIQinIs50Ohms As Long) As Long
Declare Function afSigGenDll_AnalogIQin_Input50ohms_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal analogIQinIs50Ohms As Long) As Long



'------------------------------------------------------------------------------------------------------
' DiffentialIQ Calibrate
'------------------------------------------------------------------------------------------------------
' Methods
Declare Function afSigGenDll_AnalogIQin_Calibrate_Inputs Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' DiffentialIQ
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_DifferentialIQ_Gain_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqGain As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_Gain_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqGain As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqLevel As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqLevel As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_LevelMin_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqLevel As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_LevelMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqLevel As Double) As Long

Declare Function afSigGenDll_DifferentialIQ_Modulation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIQModOn As Long) As Long
Declare Function afSigGenDll_DifferentialIQ_Modulation_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIQModOn As Long) As Long
Declare Function afSigGenDll_DifferentialIQ_OutputEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pOutputEnable As Long) As Long
Declare Function afSigGenDll_DifferentialIQ_OutputEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal outputEnable As Long) As Long
Declare Function afSigGenDll_DifferentialIQ_State_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqState As Long) As Long
Declare Function afSigGenDll_DifferentialIQ_State_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqState As Long) As Long



'------------------------------------------------------------------------------------------------------
' DiffentialIQ Calibrate
'------------------------------------------------------------------------------------------------------
' Methods
Declare Function afSigGenDll_DifferentialIQ_Calibrate_Outputs Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long



'------------------------------------------------------------------------------------------------------
' DiffentialIQ IChannel
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_DifferentialIQ_IChannel_Bias_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqIChanBias As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_IChannel_Bias_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqIChanBias As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_IChannel_Offset_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqIChanOffset As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_IChannel_Offset_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqIChanOffset As Double) As Long



'------------------------------------------------------------------------------------------------------
' DiffentialIQ QChannel
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_DifferentialIQ_QChannel_Bias_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqQChanBias As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_QChannel_Bias_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqQChanBias As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_QChannel_Offset_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDiffIqQChanOffset As Double) As Long
Declare Function afSigGenDll_DifferentialIQ_QChannel_Offset_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal diffIqQChanOffset As Double) As Long



'------------------------------------------------------------------------------------------------------
' ListMode
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ListMode_Available_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAvailable As Long) As Long
Declare Function afSigGenDll_ListMode_AddressOutputMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef plmaom As Long) As Long
Declare Function afSigGenDll_ListMode_AddressOutputMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal lmaom As Long) As Long
Declare Function afSigGenDll_ListMode_AddressSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pliAddSrc As Long) As Long
Declare Function afSigGenDll_ListMode_AddressSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal liAddSrc As Long) As Long
Declare Function afSigGenDll_ListMode_ArbCompletion_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pliArbCompletion As Long) As Long
Declare Function afSigGenDll_ListMode_ArbCompletion_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal liArbCompletion As Long) As Long
Declare Function afSigGenDll_ListMode_ArbLoadIfNeeded_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLoadFileIfNeeded As Long) As Long
Declare Function afSigGenDll_ListMode_ArbLoadIfNeeded_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal loadFileIfNeeded As Long) As Long
Declare Function afSigGenDll_ListMode_ArbSampleRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pArbSampleRate As Long) As Long
Declare Function afSigGenDll_ListMode_Control_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pListControl As Long) As Long
Declare Function afSigGenDll_ListMode_Control_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listControl As Long) As Long
Declare Function afSigGenDll_ListMode_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_ListMode_TrainLevellingMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_ListMode_TrainLevellingMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long

' List Iteration properties
Declare Function afSigGenDll_ListMode_RepeatCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal count As Integer) As Long
Declare Function afSigGenDll_ListMode_RepeatCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_ListMode_RepeatMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal mode As Long) As Long
Declare Function afSigGenDll_ListMode_RepeatMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMode As Long) As Long

' Methods
Declare Function afSigGenDll_ListMode_IssueSwTrig Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ListMode_PostTrainRfLevel Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_ListMode_PreTrainRfLevel Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long




'------------------------------------------------------------------------------------------------------
' ListMode Channel
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ListMode_Channel_ArbExplicitlySet_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByRef pArbExplicitlySet As Long) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbFile_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByVal fileNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbFile_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByVal FileName As String) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbModulationOn_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByRef pArbModulationOn As Long) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbModulationOn_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByVal arbModulationOn As Long) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbPlayCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_ListMode_Channel_ArbPlayCount_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByVal count As Integer) As Long
Declare Function afSigGenDll_ListMode_Channel_Dwellx10us_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByRef pDwellx10us As Long) As Long
Declare Function afSigGenDll_ListMode_Channel_Dwellx10us_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer, ByVal dwellx10us As Long) As Long

' Methods
Declare Function afSigGenDll_ListMode_Channel_TrainRfLevel Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal listAddress As Integer) As Long



'------------------------------------------------------------------------------------------------------
' ListMode Counter
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ListMode_Counter_Mode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCounterMode As Long) As Long
Declare Function afSigGenDll_ListMode_Counter_Mode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal counterMode As Long) As Long
Declare Function afSigGenDll_ListMode_Counter_StartAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStartAddress As Integer) As Long
Declare Function afSigGenDll_ListMode_Counter_StartAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal startAddress As Integer) As Long
Declare Function afSigGenDll_ListMode_Counter_StopAddress_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStopAddress As Integer) As Long
Declare Function afSigGenDll_ListMode_Counter_StopAddress_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal stopAddress As Integer) As Long
Declare Function afSigGenDll_ListMode_Counter_StrobeNegativeEdge_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStrobeNegEdge As Long) As Long
Declare Function afSigGenDll_ListMode_Counter_StrobeSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCounterStrobeSrc As Long) As Long
Declare Function afSigGenDll_ListMode_Counter_StrobeSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal counterStrobeSrc As Long) As Long



'------------------------------------------------------------------------------------------------------
' ListMode Strobe
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ListMode_Strobe_Mode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStrobeMode As Long) As Long
Declare Function afSigGenDll_ListMode_Strobe_Mode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal strobeMode As Long) As Long
Declare Function afSigGenDll_ListMode_Strobe_NegativeEdge_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pStrobeNegEdge As Long) As Long
Declare Function afSigGenDll_ListMode_Strobe_NegativeEdge_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal strobeNegEdge As Long) As Long



'------------------------------------------------------------------------------------------------------
' ListMode Timer
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_ListMode_Timer_Dwellx10us_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pDwellx10us As Long) As Long
Declare Function afSigGenDll_ListMode_Timer_Dwellx10us_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal dwellx10us As Long) As Long
Declare Function afSigGenDll_ListMode_Timer_TriggeredByLSTB_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggeredByLSTB As Long) As Long
Declare Function afSigGenDll_ListMode_Timer_TriggeredByLSTB_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggeredByLSTB As Long) As Long



'------------------------------------------------------------------------------------------------------
' LO
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_LO_LoopBandwidth_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLoopBandwidth As Long) As Long
Declare Function afSigGenDll_LO_LoopBandwidth_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal loopBandwidth As Long) As Long
Declare Function afSigGenDll_LO_Reference_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLoRefMode As Long) As Long
Declare Function afSigGenDll_LO_Reference_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal loRefMode As Long) As Long
Declare Function afSigGenDll_LO_ReferenceLocked_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef isLocked As Long) As Long



'------------------------------------------------------------------------------------------------------
' LO Options
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_LO_Options_AvailableCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAvailableCount As Long) As Long

' Methods
Declare Function afSigGenDll_LO_Options_CheckFitted Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal OptionNumber As Long, ByRef pFitted As Long) As Long
Declare Function afSigGenDll_LO_Options_Disable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Password As Long) As Long
Declare Function afSigGenDll_LO_Options_Enable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Password As Long) As Long
Declare Function afSigGenDll_LO_Options_Information Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Long, ByRef pOptionNumber As Long, ByVal OptionDescriptionBuffer As String, ByVal bufferLen As Long) As Long



'------------------------------------------------------------------------------------------------------
' LO Resource
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_LO_Resource_FPGAConfiguration_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFPGAConfig As Integer) As Long
Declare Function afSigGenDll_LO_Resource_FPGAConfiguration_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FPGAConfig As Integer) As Long
Declare Function afSigGenDll_LO_Resource_FPGACount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_LO_Resource_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_LO_Resource_IsPlugin_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsPlugin As Long) As Long
Declare Function afSigGenDll_LO_Resource_ModelCode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pModelCode As Long) As Long
Declare Function afSigGenDll_LO_Resource_PluginName_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal pluginNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_LO_Resource_PluginName_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal pluginNameBuffer As String) As Long
Declare Function afSigGenDll_LO_Resource_ResourceString_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal resourceStringBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_LO_Resource_SerialNumber_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal SerialNumberBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_LO_Resource_SessionID_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSessionId As Long) As Long

' Methods
Declare Function afSigGenDll_LO_Resource_FPGADescriptions Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef Numbers As Integer, ByVal Descriptions As String, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_LO_Resource_GetLastCalibrationDate Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef Year As Integer, ByRef Month As Integer, ByRef Day As Integer, ByRef Hour As Integer, ByRef Minutes As Integer, ByRef Seconds As Integer) As Long



'------------------------------------------------------------------------------------------------------
' LO Trigger
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_LO_Trigger_AddressedSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAddressedSource As Long) As Long
Declare Function afSigGenDll_LO_Trigger_AddressedSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal addressedSource As Long) As Long
Declare Function afSigGenDll_LO_Trigger_Mode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerMode As Long) As Long
Declare Function afSigGenDll_LO_Trigger_Mode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerMode As Long) As Long
Declare Function afSigGenDll_LO_Trigger_SingleSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pTriggerSourceSingle As Long) As Long
Declare Function afSigGenDll_LO_Trigger_SingleSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal triggerSourceSingle As Long) As Long
Declare Function afSigGenDll_LO_Trigger_SingleStartChannel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSingleStartChannel As Integer) As Long
Declare Function afSigGenDll_LO_Trigger_SingleStartChannel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal singleStartChannel As Integer) As Long
Declare Function afSigGenDll_LO_Trigger_SingleStopChannel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSingleStopChannel As Integer) As Long
Declare Function afSigGenDll_LO_Trigger_SingleStopChannel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal singleStopChannel As Integer) As Long



'------------------------------------------------------------------------------------------------------
' LVDS
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_LVDS_Interpolation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pInterpolation As Long) As Long
Declare Function afSigGenDll_LVDS_Ranged14BitData_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pRanged14BitData As Long) As Long
Declare Function afSigGenDll_LVDS_Ranged14BitData_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal ranged14BitData As Long) As Long
Declare Function afSigGenDll_LVDS_SampleRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSampleRate As Double) As Long
Declare Function afSigGenDll_LVDS_SampleRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal SampleRate As Double) As Long
Declare Function afSigGenDll_LVDS_UnsignedData_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pUnsignedData As Long) As Long
Declare Function afSigGenDll_LVDS_UnsignedData_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal unsignedData As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_AttenuatorHold_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAttenuatorHold As Long) As Long
Declare Function afSigGenDll_RF_AttenuatorHold_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal attenuatorHold As Long) As Long
Declare Function afSigGenDll_RF_CurrentChannel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentChannel As Integer) As Long
Declare Function afSigGenDll_RF_CurrentChannel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentChannel As Integer) As Long
Declare Function afSigGenDll_RF_CurrentFrequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentFrequency As Double) As Long
Declare Function afSigGenDll_RF_CurrentFrequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentFrequency As Double) As Long
Declare Function afSigGenDll_RF_CurrentGateRfOff_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentGateRfOff As Long) As Long
Declare Function afSigGenDll_RF_CurrentGateRfOff_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentGateRfOff As Long) As Long
Declare Function afSigGenDll_RF_CurrentLevel_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentLevel As Double) As Long
Declare Function afSigGenDll_RF_CurrentLevel_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentLevel As Double) As Long
Declare Function afSigGenDll_RF_CurrentLevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pActualLevel As Double) As Long
Declare Function afSigGenDll_RF_CurrentLevelClamped_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentLevelClamped As Long) As Long
Declare Function afSigGenDll_RF_CurrentLevelMaximum_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentLevelMaximum As Double) As Long
Declare Function afSigGenDll_RF_CurrentLevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentLevelMode As Long) As Long
Declare Function afSigGenDll_RF_CurrentLevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentLevelMode As Long) As Long
Declare Function afSigGenDll_RF_CurrentLevelType_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentLevelType As Long) As Long
Declare Function afSigGenDll_RF_CurrentOutputEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentOutputEnable As Long) As Long
Declare Function afSigGenDll_RF_CurrentOutputEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentOutputEnable As Long) As Long
Declare Function afSigGenDll_RF_CurrentRMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCurrentRMSdBc As Double) As Long
Declare Function afSigGenDll_RF_CurrentRMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal currentRMSdBc As Double) As Long
Declare Function afSigGenDll_RF_FrequencyMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequencyMax As Double) As Long
Declare Function afSigGenDll_RF_FrequencyMin_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequencyMin As Double) As Long
Declare Function afSigGenDll_RF_IQBwCorrGain_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIQBwCorrGain As Double) As Long
Declare Function afSigGenDll_RF_IQBwCorrGain_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal IQBwCorrGain As Double) As Long
Declare Function afSigGenDll_RF_IQBwCorrMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIQBwCorrMode As Long) As Long
Declare Function afSigGenDll_RF_IQBwCorrMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal IQBwCorrMode As Long) As Long
Declare Function afSigGenDll_RF_ModulationSource_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pModulationSource As Long) As Long
Declare Function afSigGenDll_RF_ModulationSource_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal modulationSource As Long) As Long
Declare Function afSigGenDll_RF_GainCalPresent_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pGainCalPresent As Long) As Long
Declare Function afSigGenDll_RF_CurrentLoFrequencyOverride_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal loFreq As Double) As Long
Declare Function afSigGenDll_RF_CurrentLoFrequencyOverride_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLoFreq As Double) As Long
Declare Function afSigGenDll_RF_CurrentLoFrequencyRequired_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pLoFreq As Double) As Long

Declare Function afSigGenDll_RF_ReferenceLocked_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef isLocked As Long) As Long

' Methods
Declare Function afSigGenDll_RF_GetArbFileCurrentLevelMaximum Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal arbfileName As String, ByRef pArbfileLevelMax As Double, ByRef pArbfileLevelType As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF AnalogModulation AM
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_AnalogModulation_AM_ModulationDepth_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAmDepth As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_AM_ModulationDepth_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal amDepth As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_AM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAmRate As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_AM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal amRate As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF AnalogModulation FM
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_AnalogModulation_FM_Deviation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFmDeviation As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_FM_Deviation_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal fmDeviation As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_FM_ModulationRate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFmRate As Long) As Long
Declare Function afSigGenDll_RF_AnalogModulation_FM_ModulationRate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal fmRate As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF Channel
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_Channel_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_RF_Channel_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal frequency As Double) As Long
Declare Function afSigGenDll_RF_Channel_GateRfOff_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pGateRfOff As Long) As Long
Declare Function afSigGenDll_RF_Channel_GateRfOff_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal gateRfOff As Long) As Long
Declare Function afSigGenDll_RF_Channel_Level_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevel As Double) As Long
Declare Function afSigGenDll_RF_Channel_Level_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal level As Double) As Long
Declare Function afSigGenDll_RF_Channel_LevelActual_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelActual As Double) As Long
Declare Function afSigGenDll_RF_Channel_LevelClamped_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelClamped As Long) As Long
Declare Function afSigGenDll_RF_Channel_LevelMaximum_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelMaximum As Double) As Long
Declare Function afSigGenDll_RF_Channel_LevelMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelMode As Long) As Long
Declare Function afSigGenDll_RF_Channel_LevelMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal levelMode As Long) As Long
Declare Function afSigGenDll_RF_Channel_LevelType_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLevelType As Long) As Long
Declare Function afSigGenDll_RF_Channel_OutputEnable_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pOutputEnable As Long) As Long
Declare Function afSigGenDll_RF_Channel_OutputEnable_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal outputEnable As Long) As Long
Declare Function afSigGenDll_RF_Channel_RMSdBc_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pRMSdBc As Double) As Long
Declare Function afSigGenDll_RF_Channel_RMSdBc_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal RMSdBc As Double) As Long
Declare Function afSigGenDll_RF_Channel_LoFrequencyOverride_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal loFreq As Double) As Long
Declare Function afSigGenDll_RF_Channel_LoFrequencyOverride_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLoFreq As Double) As Long
Declare Function afSigGenDll_RF_Channel_LoFrequencyRequired_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pLoFreq As Double) As Long

' Methods
Declare Function afSigGenDll_RF_Channel_GetArbFileLevelMaximum Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal arbfileName As String, ByRef pArbfileLevelMax As Double, ByRef pArbfileLevelType As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF Options
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_Options_AvailableCount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pAvailableCount As Long) As Long

' Methods
Declare Function afSigGenDll_RF_Options_CheckFitted Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal OptionNumber As Long, ByRef pFitted As Long) As Long
Declare Function afSigGenDll_RF_Options_Disable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Password As Long) As Long
Declare Function afSigGenDll_RF_Options_Enable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Password As Long) As Long
Declare Function afSigGenDll_RF_Options_Information Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal Index As Long, ByRef pOptionNumber As Long, ByVal OptionDescriptionBuffer As String, ByVal bufferLen As Long) As Long



'------------------------------------------------------------------------------------------------------
' RF Resource
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_Resource_FPGAConfiguration_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFPGAConfig As Integer) As Long
Declare Function afSigGenDll_RF_Resource_FPGAConfiguration_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal FPGAConfig As Integer) As Long
Declare Function afSigGenDll_RF_Resource_FPGACount_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_RF_Resource_IsActive_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsActive As Long) As Long
Declare Function afSigGenDll_RF_Resource_IsPlugin_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pIsPlugin As Long) As Long
Declare Function afSigGenDll_RF_Resource_ModelCode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pModelCode As Long) As Long
Declare Function afSigGenDll_RF_Resource_PluginName_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal pluginNameBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_RF_Resource_PluginName_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal pluginNameBuffer As String) As Long
Declare Function afSigGenDll_RF_Resource_ResourceString_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal resourceStringBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_RF_Resource_SerialNumber_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal SerialNumberBuffer As String, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_RF_Resource_SessionID_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSessionId As Long) As Long

' Methods
Declare Function afSigGenDll_RF_Resource_FPGADescriptions Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef Numbers As Integer, ByVal Descriptions As String, ByRef pCount As Integer) As Long
Declare Function afSigGenDll_RF_Resource_GetLastCalibrationDate Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef Year As Integer, ByRef Month As Integer, ByRef Day As Integer, ByRef Hour As Integer, ByRef Minutes As Integer, ByRef Seconds As Integer) As Long



'------------------------------------------------------------------------------------------------------
' RF Routing
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_RF_Routing_ScenarioListSize_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pScenarioListSize As Long) As Long

' Methods
Declare Function afSigGenDll_RF_Routing_AppendScenario Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RoutingScenario As Long) As Long
Declare Function afSigGenDll_RF_Routing_GetConnect Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal MatrixOutput As Long, ByRef pMatrixInput As Long) As Long
Declare Function afSigGenDll_RF_Routing_GetOutputEnable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal MatrixOutput As Long, ByRef pOutputEnable As Long) As Long
Declare Function afSigGenDll_RF_Routing_GetScenarioList Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef ScenarioList As Long, ByVal bufferLen As Long) As Long
Declare Function afSigGenDll_RF_Routing_RemoveScenario Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RoutingScenario As Long) As Long
Declare Function afSigGenDll_RF_Routing_Reset Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long) As Long
Declare Function afSigGenDll_RF_Routing_SetConnect Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal MatrixOutput As Long, ByVal MatrixInput As Long) As Long
Declare Function afSigGenDll_RF_Routing_SetOutputEnable Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal MatrixOutput As Long, ByVal outputEnable As Long) As Long
Declare Function afSigGenDll_RF_Routing_SetScenario Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal RoutingScenario As Long) As Long



'------------------------------------------------------------------------------------------------------
' VCO - 3020, 3020A & 3025 only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_VCO_ExternalReference_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pExternalReference As Long) As Long
Declare Function afSigGenDll_VCO_ExternalReference_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal externalReference As Long) As Long
Declare Function afSigGenDll_VCO_Frequency_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pFrequency As Double) As Long
Declare Function afSigGenDll_VCO_Frequency_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal frequency As Double) As Long
Declare Function afSigGenDll_VCO_Interpolation_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pInterpolation As Long) As Long



'------------------------------------------------------------------------------------------------------
' Generic Resampler - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_GenericResampler_Rate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal SampleRate As Double) As Long
Declare Function afSigGenDll_GenericResampler_Rate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pSampleRate As Double) As Long
Declare Function afSigGenDll_GenericResampler_SampleRateMax_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMaxSampleRate As Double) As Long
Declare Function afSigGenDll_GenericResampler_SampleRateMin_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pMinSampleRate As Double) As Long



'------------------------------------------------------------------------------------------------------
' Generic Resampler Channel - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_GenericResampler_Channel_Rate_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByVal SampleRate As Double) As Long
Declare Function afSigGenDll_GenericResampler_Channel_Rate_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal channel As Integer, ByRef pSampleRate As Double) As Long



'------------------------------------------------------------------------------------------------------
' DDS - C Variants only
'------------------------------------------------------------------------------------------------------
' Properties
Declare Function afSigGenDll_DDS_ClockMode_Set Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByVal clockMode As Long) As Long
Declare Function afSigGenDll_DDS_ClockMode_Get Lib "afSigGenDll_32.dll" (ByVal sigGenId As Long, ByRef pClockMode As Long) As Long
