Attribute VB_Name = "afGenericDll"
' This module contains the information needed for an application
' to use the Generic Analysis Library DLL.
' NB - 'afMeasLibDefs.bas' must be added to the project with this module.
Option Explicit

' =============================================================================================================
' Dll installed in system32 folder
' =============================================================================================================

'**** Type Definitions ****
' afGenericError - This enumeration defines the different error codes that
'                  can be returned by the functions defined in this DLL.
Public Enum afGenericError
    afGenericNotSupported = -16384
    afGenericInvalidWeighting = -16385
    afGenericInvalidDeemphasis = -16386
    afGenericInvalidNumAudioChannels = -16387
    afGenericInvalidAudioLevelUnit = -16388
    afGenericInvalidNumAudioHarmonics = -16389
    afGenericInvalidModType = -16390
    afGenericInvalidPresetToStandard = -16391
    afGenericInvalidNumSymbolsToAnalyse = -16392
    afGenericInvalidMeasurementsForAudioSignal = -16393
    afGenericFMChannelPowerNotSupported = -16394
    afGenericZeroInputSamples = -16395
    afGenericInvalidAudioLowerFrequencyLimit = -16396
    afGenericInvalidAudioUpperFrequencyLimit = -16397
    afGenericInvalidAudioFrequencyLimits = -16398
    afGenericInvalidSymbolOffset = -16399
    afGenericInvalidSyncPatternType = -16400
    afGenericInvalidSyncPatternSymbolOffset = -16401
    afGenericInvalidUserSyncPatternLength = -16402
    afGenericInvalidConstellationSymbolMapping = -16403
    afGenericInvalidSyncPatternLengthForModulation = -16404
    afGenericInvalidTransmitFilter = -16405
    afGenericInvalidMeasurementFilter = -16406
    afGenericInvalidChannelFilter = -16407
End Enum

' afGenericMeasurement - This enumeration defines the measurements provided by this DLL.
'                        Any combination of these can be passed into afGenericDll_Analyse.
Public Enum afGenericMeasurement
    afGenericMeasNotDefined = &H0
    afGenericMeasAudio = &H2
    afGenericMeasPower = &H4
    afGenericMeasCWFreq = &H8
    afGenericMeasLocateBurst = &H10
    afGenericMeasModAccuracy = &H20
    afGenericMeasCcdf = &H40

    ' Deprecated Values
    afGenericMeasFM = &H1       ' instead use afGenericMeasModAccuracy with ModulationType set to afGenericModFM
End Enum

' afGenericTraceData - This enumeration defines the different types of trace data that
'                      can be retrieved from this DLL (via afGenericDll_GetTraceData).
'                      after a measurement has completed.
Public Enum afGenericTraceData
    afGenericTraceIdealConstellation = 0
    afGenericTraceMeasConstellation = 1
    afGenericTraceEvmVsSymbol = 2
    afGenericTracePhaseErrorVsSymbol = 3
    afGenericTraceMagnitudeErrorVsSymbol = 4
    afGenericTraceCcdf = 5
    afGenericTraceBurstPowerVsTime = 6
    afGenericTraceAudioLeftSamplesVsTime = 7
    afGenericTraceAudioLeftSpectrum = 8
    afGenericTraceAudioRightSamplesVsTime = 9
    afGenericTraceAudioRightSpectrum = 10
    afGenericTraceFskRefVsSymbol = 11
    afGenericTraceFskMeasVsSymbol = 12
    afGenericTraceFskErrorVsSymbol = 13
    afGenericTraceEvmAltVsSymbol = 14
End Enum

' afGenericAnalysisMode - defines the different analysis modes provided by this DLL.
Public Enum afGenericAnalysisMode
    afGenericAnalysisModeAllIQ = 0
    afGenericAnalysisModeBurstIQ = 1
End Enum

' afGenericAudioWeighting - This enumeration defines the audio weighting that can be used.
Public Enum afGenericAudioWeighting
    afGenericAudioWeightingNone = 0
    afGenericAudioWeightingA = 1
    afGenericAudioWeightingC = 2
End Enum

' afGenericAudioLevelUnits - This enumeration defines audio-level measurement units.
Public Enum afGenericAudioLevelUnits
    afGenericAudioLevelUnitsdB = 0
    afGenericAudioLevelUnitsVolts = 1
End Enum

' afGenericAudioDeemphasisFilter - This enumeration defines the audio deemphasis filters that can be used.
Public Enum afGenericAudioDeemphasisFilter
    afGenericAudioDeemphasisFilterNone = 0
    afGenericAudioDeemphasisFilter50us = 1
    afGenericAudioDeemphasisFilter75us = 2
End Enum

' afGenericAudioChannel - This enumeration defines the audio channels that can be used.
Public Enum afGenericAudioChannel
    afGenericAudioChannelLeft = 0
    afGenericAudioChannelRight = 1
End Enum

' afGenericModType - This enumeration defines the types of modulation supported.
Public Enum afGenericModType
    afGenericModFM = 0
    afGenericModBpsk = 1
    afGenericModOQpsk = 2
    afGenericModQpsk = 3
    afGenericModDBpsk = 4
    afGenericModPiBy2DBpsk = 5
    afGenericModDQpsk = 6
    afGenericModPiBy4DQpsk = 7
    afGenericMod8psk = 8
    afGenericModD8psk = 9
    afGenericModPiBy8D8psk = 10
    afGenericMod8pskEdge = 11
    afGenericMod16Qam = 12
    afGenericMod32Qam = 13
    afGenericMod64Qam = 14
    afGenericMod128Qam = 15
    afGenericMod256Qam = 16
    afGenericMod512Qam = 17
    afGenericModMsk = 19
    afGenericMod2Fsk = 20
    afGenericMod4Fsk = 21
End Enum

' afGenericSyncPatternType - This enumeration defines the types of synchronisation patterns supported.
Public Enum afGenericSyncPatternType
    afGenericSyncPatternUser = 0
    afGenericSyncPatternEdgeTsc0 = 1
    afGenericSyncPatternEdgeTsc1 = 2
    afGenericSyncPatternEdgeTsc2 = 3
    afGenericSyncPatternEdgeTsc3 = 4
    afGenericSyncPatternEdgeTsc4 = 5
    afGenericSyncPatternEdgeTsc5 = 6
    afGenericSyncPatternEdgeTsc6 = 7
    afGenericSyncPatternEdgeTsc7 = 8
    afGenericSyncPatternTetraTsc0 = 9
    afGenericSyncPatternTetraTsc1 = 10
End Enum

' afGenericStdPresetType - This enumeration defines the types of preset to standard supported.
Public Enum afGenericStdPresetType
    afGenericStdPresetLrWPan780OQpsk = 0
    afGenericStdPresetLrWPan868Bpsk = 1
    afGenericStdPresetLrWPan868OQpsk = 2
    afGenericStdPresetLrWPan915Bpsk = 3
    afGenericStdPresetLrWPan915OQpsk = 4
    afGenericStdPresetLrWPan950Bpsk = 5
    afGenericStdPresetLrWPan2450OQpsk = 6
    afGenericStdPresetUmts = 7
    afGenericStdPresetTetra = 8
    afGenericStdPresetApco25 = 9
    afGenericStdPresetVdlMode2 = 10
    afGenericStdPresetVdlMode3 = 11
    afGenericStdPresetEdge = 12
    afGenericStdPresetVdlMode4 = 13
    afGenericStdPresetGsm = 14
    afGenericStdPresetBluetooth = 15

    ' Deprecated Values
    afGenericZigbee780OQpsk = 0
    afGenericZigbee868Bpsk = 1
    afGenericZigbee868OQpsk = 2
    afGenericZigbee915Bpsk = 3
    afGenericZigbee915OQpsk = 4
    afGenericZigbee950Bpsk = 5
    afGenericZigbee2450OQpsk = 6
End Enum


'**** Exported Functions ****

'** Methods **
Declare Function afGenericDll_CreateObject Lib "afGenericDll.dll" (ByRef ptrID As Long) As Long

Declare Function afGenericDll_DestroyObject Lib "afGenericDll.dll" (ByVal nID As Long) As Long

Declare Function afGenericDll_Analyse Lib "afGenericDll.dll" (ByVal nID As Long, _
                                                                  ByVal measurements As Long, _
                                                                  ByRef ptrIData As Single, _
                                                                  ByRef ptrQData As Single, _
                                                                  ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ

Declare Function afGenericDll_GetTraceDataLength Lib "afGenericDll.dll" (ByVal nID As Long, _
                                                                             ByVal traceType As Long, _
                                                                             ByRef ptrLength As Long) As Long

Declare Function afGenericDll_GetTraceData Lib "afGenericDll.dll" (ByVal nID As Long, _
                                                                       ByVal traceType As Long, _
                                                                       ByRef ptrX As Double, _
                                                                       ByRef ptrY As Double, _
                                                                       ByVal numPoints As Long) As Long
' ptrX must be an array of size numPoints
' ptrY must be an array of size numPoints

Declare Function afGenericDll_GetErrorMsgLength Lib "afGenericDll.dll" (ByVal nID As Long, _
                                                                            ByVal errorCode As Long, _
                                                                            ByRef ptrLength As Long) As Long

Declare Function afGenericDll_GetErrorMsg Lib "afGenericDll.dll" (ByVal nID As Long, _
                                                                      ByVal errorCode As Long, _
                                                                      ByVal ptrBuffer As String, _
                                                                      ByVal bufferSize As Long) As Long ' NB byval on the string

Declare Function afGenericDll_MaxErrorMsgLength_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrLength As Long) As Long

Declare Function afGenericDll_GetVersion Lib "afGenericDll.dll" (ByRef ptrVersion As Long) As Long
'* NOTE - ptrVersion must be a pointer to an array of 4 elements.*'



'* Configuration Methods and Properties *
Declare Function afGenericDll_GetMinSamplingFreq Lib "afGenericDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSampFreq As Double) As Long
Declare Function afGenericDll_GetRecommendedSamplingFreq Lib "afGenericDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSampFreq As Double) As Long
Declare Function afGenericDll_GetMinNumSamples Lib "afGenericDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrNumSamples As Long) As Long

Declare Function afGenericDll_SamplingFreq_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal sampFreq As Double) As Long
Declare Function afGenericDll_SamplingFreq_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSampFreq As Double) As Long

Declare Function afGenericDll_RfLevelCal_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal levelcal As Single) As Long
Declare Function afGenericDll_RfLevelCal_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrLevelCal As Single) As Long

Declare Function afGenericDll_AnalysisMode_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afGenericDll_AnalysisMode_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afGenericDll_PresetToStandard Lib "afGenericDll.dll" (ByVal nID As Long, ByVal stdPreset As Long) As Long

Declare Function afGenericDll_SymbolRate_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal symbolRate As Double) As Long
Declare Function afGenericDll_SymbolRate_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSymbolRate As Double) As Long

Declare Function afGenericDll_ModulationType_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long) As Long
Declare Function afGenericDll_ModulationType_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrModType As Long) As Long

Declare Function afGenericDll_NumBitsPerSymbol_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrNumBitsPerSymbol As Long) As Long

Declare Function afGenericDll_NumSymbolsToAnalyse_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal numSym As Long) As Long
Declare Function afGenericDll_NumSymbolsToAnalyse_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumSym As Long) As Long

Declare Function afGenericDll_SymbolOffset_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal numSym As Long) As Long
Declare Function afGenericDll_SymbolOffset_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumSym As Long) As Long

Declare Function afGenericDll_IQOriginOffsetRemoval_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal origOffRemoval As Long) As Long
Declare Function afGenericDll_IQOriginOffsetRemoval_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrOrigOffRemoval As Long) As Long

Declare Function afGenericDll_Oqpsk_Hs_Unity_Ref_Mag_Enable_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal oqpskHsUnityRefMagEnable As Long) As Long
Declare Function afGenericDll_Oqpsk_Hs_Unity_Ref_Mag_Enable_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrOqpskHsUnityRefMagEnable As Long) As Long

'* Filter Properties *
Declare Function afGenericDll_MeasurementFilter_Type_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal measFiltType As Long) As Long
Declare Function afGenericDll_MeasurementFilter_Type_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMeasFiltType As Long) As Long

Declare Function afGenericDll_MeasurementFilter_Alpha_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal measFiltAlpha As Single) As Long
Declare Function afGenericDll_MeasurementFilter_Alpha_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMeasFiltAlpha As Single) As Long

Declare Function afGenericDll_TransmitFilter_Type_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal txFiltType As Long) As Long
Declare Function afGenericDll_TransmitFilter_Type_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrTxFiltType As Long) As Long

Declare Function afGenericDll_TransmitFilter_Alpha_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal txFiltAlpha As Single) As Long
Declare Function afGenericDll_TransmitFilter_Alpha_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrTxFiltAlpha As Single) As Long

Declare Function afGenericDll_TransmitFilter_Bt_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal txFiltBT As Single) As Long
Declare Function afGenericDll_TransmitFilter_Bt_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrTxFiltBT As Single) As Long

Declare Function afGenericDll_ChannelFilter_Type_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal chFiltType As Long) As Long
Declare Function afGenericDll_ChannelFilter_Type_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrChFiltType As Long) As Long

Declare Function afGenericDll_ChannelFilter_Alpha_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal chFiltAlpha As Single) As Long
Declare Function afGenericDll_ChannelFilter_Alpha_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrChFiltAlpha As Single) As Long

Declare Function afGenericDll_ChannelFilter_Bt_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal chFiltBT As Single) As Long
Declare Function afGenericDll_ChannelFilter_Bt_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrChFiltBT As Single) As Long

'* Burst Properties *
Declare Function afGenericDll_Burst_Position_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal burstPos As Long) As Long
Declare Function afGenericDll_Burst_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBurstPos As Long) As Long

Declare Function afGenericDll_Burst_Length_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal burstLen As Long) As Long
Declare Function afGenericDll_Burst_Length_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBurstLen As Long) As Long

Declare Function afGenericDll_Burst_PreTriggerTime_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal preTriggerTime As Single) As Long
Declare Function afGenericDll_Burst_PreTriggerTime_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPreTriggerTime As Single) As Long

Declare Function afGenericDll_Burst_MinimumOnTime_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal minOnTime As Single) As Long
Declare Function afGenericDll_Burst_MinimumOnTime_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMinOnTime As Single) As Long

Declare Function afGenericDll_Burst_MinimumOffTime_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal minOffTime As Single) As Long
Declare Function afGenericDll_Burst_MinimumOffTime_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMinOffTime As Single) As Long

Declare Function afGenericDll_Burst_IntegrationTime_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal time As Single) As Long
Declare Function afGenericDll_Burst_IntegrationTime_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrTime As Single) As Long

Declare Function afGenericDll_Burst_IntegrationSkipTime_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal skipTime As Single) As Long
Declare Function afGenericDll_Burst_IntegrationSkipTime_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSkipTime As Single) As Long

Declare Function afGenericDll_Burst_ComparatorDelay_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal delay As Long) As Long
Declare Function afGenericDll_Burst_ComparatorDelay_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDelay As Long) As Long

Declare Function afGenericDll_Burst_RisingEdge_Threshold_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afGenericDll_Burst_RisingEdge_Threshold_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

Declare Function afGenericDll_Burst_FallingEdge_Threshold_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afGenericDll_Burst_FallingEdge_Threshold_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

'* Sync Pattern Configuration Methods and Properties *
Declare Function afGenericDll_SyncPattern_SetUserDefinedBits Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBits As Long, ByVal len_ As Long) As Long
' ptrBits must be an array of size len
Declare Function afGenericDll_SyncPattern_GetUserDefinedBits Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBits As Long, ByVal len_ As Long) As Long
' ptrBits must be an array of size len
Declare Function afGenericDll_SyncPattern_GetCurrentBits Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBits As Long, ByVal len_ As Long) As Long
' ptrBits must be an array of size len

Declare Function afGenericDll_SyncPattern_SearchEnabled_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal syncPattern_SearchEnabled As Long) As Long
Declare Function afGenericDll_SyncPattern_SearchEnabled_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPattern_SearchEnabled As Long) As Long

Declare Function afGenericDll_SyncPattern_SymbolOffset_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal syncPattern_SymbolOffset As Long) As Long
Declare Function afGenericDll_SyncPattern_SymbolOffset_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPattern_SymbolOffset As Long) As Long

Declare Function afGenericDll_SyncPattern_Type_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal syncPattern_Type As Long) As Long
Declare Function afGenericDll_SyncPattern_Type_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPattern_Type As Long) As Long

Declare Function afGenericDll_SyncPattern_NumBits_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPatternNumBits As Long) As Long

Declare Function afGenericDll_SyncPattern_NumUserDefinedBits_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPatternNumUserDefinedBits As Long) As Long

'* Constellation Methods and Properties *
Declare Function afGenericDll_Constellation_SetSymbolMapping Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrMapping As Long, ByVal len_ As Long) As Long
' ptrMapping must be an array of size len
Declare Function afGenericDll_Constellation_GetSymbolMapping Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrMapping As Long, ByVal len_ As Long) As Long
' ptrMapping must be an array of size len
Declare Function afGenericDll_Constellation_ResetSymbolMapping Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long) As Long
Declare Function afGenericDll_Constellation_GetCoordinates Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrI As Single, ByRef ptrQ As Single, ByVal len_ As Long) As Long
' ptrI must be an array of size len
' ptrQ must be an array of size len

Declare Function afGenericDll_Constellation_NumCoordinates_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrNumPoints As Long) As Long

Declare Function afGenericDll_Constellation_NumSymbolMappingPoints_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal modType As Long, ByRef ptrNumElem As Long) As Long

'* Audio Configuration Methods and Properties *
Declare Function afGenericDll_Audio_NumChannels_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal numChannels As Long) As Long
Declare Function afGenericDll_Audio_NumChannels_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumChannels As Long) As Long

Declare Function afGenericDll_Audio_Weighting_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal weighting As Long) As Long
Declare Function afGenericDll_Audio_Weighting_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrWeighting As Long) As Long

Declare Function afGenericDll_Audio_Deemphasis_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal deemphasis As Long) As Long
Declare Function afGenericDll_Audio_Deemphasis_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDeemphasis As Long) As Long

Declare Function afGenericDll_Audio_LevelUnits_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal levelUnits As Long) As Long
Declare Function afGenericDll_Audio_LevelUnits_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrLevelUnits As Long) As Long

Declare Function afGenericDll_Audio_NumHarmonicsForTHD_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal numHarmonicsForTHD As Long) As Long
Declare Function afGenericDll_Audio_NumHarmonicsForTHD_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumHarmonicsForTHD As Long) As Long

Declare Function afGenericDll_Audio_LowerFrequencyLimit_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal lowerFrequencyLimit As Double) As Long
Declare Function afGenericDll_Audio_LowerFrequencyLimit_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrLowerFrequencyLimit As Double) As Long

Declare Function afGenericDll_Audio_UpperFrequencyLimit_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal upperFrequencyLimit As Double) As Long
Declare Function afGenericDll_Audio_UpperFrequencyLimit_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrUpperFrequencyLimit As Double) As Long

'* Results Methods and Properties *
Declare Function afGenericDll_CWFreqOffset_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrCwFreq As Double) As Long

Declare Function afGenericDll_Evm_Rms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmRms As Single) As Long

Declare Function afGenericDll_Evm_Alt_Rms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmAltRms As Single) As Long
Declare Function afGenericDll_Evm_Peak_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmPeak As Single) As Long
Declare Function afGenericDll_Evm_Alt_Peak_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmAltPeak As Single) As Long
Declare Function afGenericDll_Evm_Peak_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmPeakPos As Long) As Long
Declare Function afGenericDll_Evm_Alt_Peak_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrEvmAltPeakPos As Long) As Long

Declare Function afGenericDll_PhaseError_Rms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPhaseErrorRms As Single) As Long
Declare Function afGenericDll_PhaseError_Peak_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPhaseErrorPeak As Single) As Long
Declare Function afGenericDll_PhaseError_Peak_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPhaseErrorPeakPos As Long) As Long

Declare Function afGenericDll_MagnitudeError_Rms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMagnitudeErrorRms As Single) As Long
Declare Function afGenericDll_MagnitudeError_Peak_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMagnitudeErrorPeak As Single) As Long
Declare Function afGenericDll_MagnitudeError_Peak_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrMagnitudeErrorPeakPos As Long) As Long

Declare Function afGenericDll_FreqError_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrFreqError As Single) As Long

Declare Function afGenericDll_IQOriginOffset_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrIQOriginOffset As Single) As Long
Declare Function afGenericDll_IQGainImbalance_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrIQGainImbalance As Single) As Long
Declare Function afGenericDll_IQSkew_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrIqSkew As Single) As Long
Declare Function afGenericDll_AveragePower_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrAveragePower As Single) As Long
Declare Function afGenericDll_PeakPower_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPeakPower As Single) As Long

Declare Function afGenericDll_FskError_Rms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrFskErrorRms As Double) As Long
Declare Function afGenericDll_FskError_Peak_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrFskErrorPeak As Double) As Long
Declare Function afGenericDll_FskError_Peak_Position_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrFskErrorPeakPos As Long) As Long
Declare Function afGenericDll_FskFrequencyDeviation_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrFskDeviation As Double) As Long

Declare Function afGenericDll_SyncPattern_Detected_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrSyncPatternDetected As Long) As Long
Declare Function afGenericDll_NumDemodSymbols_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumDemodSymbols As Long) As Long
Declare Function afGenericDll_NumDemodBits_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrNumDemodBits As Long) As Long
Declare Function afGenericDll_GetDemodBits Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBitsOut As Long, ByVal len_ As Long) As Long
' ptrBitsOut must be an array of size len
Declare Function afGenericDll_GetDemodSymbols Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrBitsOut As Long, ByVal len_ As Long) As Long
' ptrBitsOut must be an array of size len

'* Audio Results Methods and Properties *
Declare Function afGenericDll_Audio_Frequency_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrFrequency As Double) As Long
Declare Function afGenericDll_Audio_LevelRms_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrLevel As Single) As Long
Declare Function afGenericDll_Audio_Snr_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrSnr As Single) As Long
Declare Function afGenericDll_Audio_Sinad_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrSinad As Single) As Long
Declare Function afGenericDll_Audio_Thd_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrThd As Single) As Long
Declare Function afGenericDll_Audio_ThdPlusNoise_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrThdPlusNoise As Single) As Long
Declare Function afGenericDll_Audio_StereoIsolation_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByVal channel As Long, ByRef ptrStereoIsolation As Single) As Long

'* FM Results Methods and Properties *
Declare Function afGenericDll_FM_FrequencyDeviation_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDeviation As Double) As Long
Declare Function afGenericDll_FM_FrequencyDeviation_M_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDeviation As Double) As Long
Declare Function afGenericDll_FM_FrequencyDeviation_S_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDeviation As Double) As Long
Declare Function afGenericDll_FM_FrequencyDeviation_P_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrDeviation As Double) As Long

'* Deprecated Functions *

Declare Function afGenericDll_ReferenceFilter_Type_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal refFiltType As Long) As Long
Declare Function afGenericDll_ReferenceFilter_Type_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrRefFiltType As Long) As Long

Declare Function afGenericDll_ReferenceFilter_Alpha_Set Lib "afGenericDll.dll" (ByVal nID As Long, ByVal refFiltAlpha As Single) As Long
Declare Function afGenericDll_ReferenceFilter_Alpha_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrRefFiltAlpha As Single) As Long
Declare Function afGenericDll_Power_Get Lib "afGenericDll.dll" (ByVal nID As Long, ByRef ptrPower As Single) As Long
