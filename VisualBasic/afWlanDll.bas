Attribute VB_Name = "afWlanDll"
' This module contains the information needed for an application
' to use the Wireless LAN Analysis Library DLL.
' NB - 'afMeasLibDefs.bas' must be added to the project with this module.
Option Explicit

' =============================================================================================================
' Dll installed in system32 folder
' =============================================================================================================

'**** Type Definitions ****
' afWlanError - defines the different error codes that can be returned
'               by the functions defined in this DLL.
Public Enum afWlanError
    afWlanNoError = 0   ' afMeasNoError
    afWlanUnknownError = -1     ' afMeasUnknownError
    afWlanFailAllocMem = -2     ' afMeasFailAllocMem
    afWlanInvalidMemAddress = -3        ' afMeasInvalidMemAddress
    afWlanInvalidDllID = -4     ' afMeasInvalidDllID
    afWlanInvalidInputParameter = -5    ' afMeasInvalidInputParameter
    afWlanUnknownTraceType = -6 ' afMeasUnknownTraceType
    afWlanNoTraceDataAvailable = -7     ' afMeasNoTraceDataAvailable
    afWlanInvalidSamplingFreq = -8      ' afMeasInvalidSamplingFreq
    afWlanNoRisingEdge = -9     ' afMeasNoRisingEdge
    afWlanIncompleteBurst = -10 ' afMeasIncompleteBurst
    afWlanNoBurstDefined = -11  ' afMeasNoBurstDefined
    afWlanInsufficientIQ = -12  ' afMeasInsufficientIQ
    afWlanFailToSync = -13      ' afMeasFailToSync
    afWlanBufferSize = -16      ' afMeasInvalidBufferSize
    afWlanInvalidStitchSetup = -18      ' afMeasInvalidSpectrumStitchSetup
    afWlanInvalidDigitizerSpan = -19    ' afMeasInvalidDigitizerSpan
    afWlanInvalidMeasurementSpan = -20  ' afMeasInvalidMeasurementSpan
    afWlanInvalidStitchFreqIndex = -21  ' afMeasInvalidSpectrumStitchIndex
    afWlanInvalidSpectrum = -22 ' afMeasInvalidSpectrum
    afWlanInvalidSpectrumLength = -23   ' afMeasInvalidSpectrumLength
    afWlanInvalidSpectrumMaskDefinition = -24   ' afMeasInvalidSpectrumMaskDefinition
    afWlanDsssOfdmNotSupported = -16384
    afWlanDsss33MbpsNotSupported = -16385
    afWlanFailToDecodeDsssHeader = -16386
    afWlanFailToDecodeOfdmSignalField = -16387
    afWlanInvalidOfdmCarrierIndex = -16388
    afWlanSpaceTimeCodedSignalNotSupported = -16389
    afWlanInvalidAntennaNumber = -16390
    afWlanInvalidEqualizerLength = -16391
    afWlanInsufficientSymbolsToAnalyse = -16392
    afWlanUseAnalyseIQ = -16393
    afWlanUseDataRatekbps = -16394
    afWlanInvalidChannelBandwidth = -16395
    afWlanInvalidChannelOffset = -16396
    afWlanUseCombineSpectrumSegments = -16397
    afWlanInvalidSegmentSeparation = -16398
    afWlanInconsistentMIMOSettings = -16399
    afWlanInconsistentMIMOHeaderInfo = -16400
    afWlanDuplicateAntennaCapture = -16401
    afWlanInsufficientMIMOData = -16402
    afWlanInvalidStream = -16403
    afWlanInvalidAnalysisModeForMIMO = -16404
    afWlanInvalidNumberOfStreamsForMIMO = -16405
    afWlanInvalidSignalField = -16406
    afWlanInvalidBandwidthForSpectralMask = -16407
    afWlanReferenceFilenameNotProvided = -16408
    afWlanReferenceFileNotFound = -16409
    afWlanReferenceFileCrcFailed = -16410
    afWlanReferenceFileFormatInvalid = -16411
    afWlanReferenceFileModeInvalid = -16412
    afWlanReferenceFileMcsInvalid = -16413
    afWlanReferenceFileNssInvalid = -16414
    afWlanReferenceFileBwInvalid = -16415
    afWlanReferenceFileNumSymbolsInvalid = -16416
    afWlanReferenceFileNumPacketsInvalid = -16417
    afWlanReferenceFileParametersInconsistent = -16418
    afWlanReferenceFileParameterMismatchWithSignal = -16419
    afWlanReferenceFileInsufficientSymbols = -16420
    afWlanInsufficientPreBurstIQ = -16421
    afWlanCombinedResultNotAvailable = -16422
    afWlanInconsistentHeaderInfo = -16423
    afWlanCombinedResultNotSupported = -16424
    afWlanInvalid160MHzCaptureSetting = -16425
End Enum

' afWlanMeasurement - defines the measurements provided by this DLL. Any combination
'                     of these can be passed into the afWlanDll_Analyse function.
Public Enum afWlanMeasurement
    afWlanMeasNotDefined = &H0
    afWlanCWAveragePower = &H1
    afWlanCWFreq = &H2
    afWlanLocateBurst = &H4
    afWlanBurstPower = &H8
    afWlanModAccuracy = &H10
    afWlanMeasSpectrum = &H40
    afWlanMeasOccupiedBandwidth = &H80
    afWlanMeasSpectralMask = &H100
    afWlanMeasAdjChanPower = &H200
    afWlanMeasCcdf = &H400
    afWlanMeasMIMOStore = &H800

    ' Deprecated Values
    afWlanSpectrumAnalysis = &H20       ' instead use the individual spectral measurements above
End Enum

' afWlanTraceData - defines the types of trace data that can be retrieved
'                   from this DLL (via afWlanDll_GetTraceData) after a
'                   measurement has completed.
Public Enum afWlanTraceData
    afWlanSpectrum = 0
    afWlanPowerVsTime = 1
    afWlanConstellation = 2
    afWlanClockErrorVsTime = 6
    afWlanSpectralMask = 7
    afWlanDsssEvmVsChip = 3
    afWlanOfdmEvmVsCarrier = 3
    afWlanOfdmEvmVsSymbol = 4
    afWlanOfdmSpectralFlatness = 5
    afWlanOfdmSpectralFlatnessLowerLimit = 8
    afWlanOfdmSpectralFlatnessUpperLimit = 9
    afWlanCcdf = 10
    afWlanOfdmEvmRmsVsCarrier = 11
    afWlanStreamEvmVsCarrier = 12
    afWlanStreamEvmRmsVsCarrier = 13
    afWlanStreamEvmVsSymbol = 14
    afWlanStreamConstellation = 15
    afWlanChannelFrequencyResponse = 16
    afWlanMIMOMatrixConditionNumber = 17
    afWlanPreambleFrequencyError = 18
    afWlanReferenceConstellation = 19
    afWlanStreamReferenceConstellation = 20

    ' Deprecated Values
    afWlanErrorVectorMagnitude = 3      ' instead use afWlanDsssEvmVsChip or afWlanOfdmEvmVsCarrier
    afWlanOfdmSymbolEvm = 4     ' instead use afWlanOfdmEvmVsSymbol
End Enum

' afWlanAnalysisMode - defines the different 802.11 analysis modes.
Public Enum afWlanAnalysisMode
    afWlanAnalysisModeAuto11g = 0       ' Auto-detects between 802.11a/g (OFDM) and 802.11b/g (DSSS or DSSS-OFDM) signals
    afWlanAnalysisMode11b = 2   ' 802.11b/g (DSSS or DSSS-OFDM) signals
    afWlanAnalysisMode11nHT = 3 ' 802.11n HT (20/40MHz b/w Greenfield or Mixed Format) signals
    afWlanAnalysisMode11nNonHT = 4      ' 802.11n non-HT (20/40MHz b/w) signals
    afWlanAnalysisMode11acVHT = 5       ' 802.11ac VHT (20/40/80/160/80+80MHz b/w) signals
    afWlanAnalysisMode11acNonHT = 6     ' 802.11ac non-HT (20/40/80/160/80+80MHz b/w) signals
    afWlanAnalysisMode11aAllBw = 7      ' 802.11a/g (OFDM) signals (5/10/20MHz b/w)

    ' Deprecated Values
    afWlanAnalysisMode11a = 1   ' instead use afWlanAnalysisMode11aAllBw
End Enum

' afWlanChannelBandwidth - defines the different 802.11n channel bandwidths.
Public Enum afWlanChannelBandwidth
    afWlanChannelBandwidth20MHz = 0
    afWlanChannelBandwidth40MHz = 1
    afWlanChannelBandwidth80MHz = 2
    afWlanChannelBandwidth80plus80MHz = 3
    afWlanChannelBandwidth160MHz = 4
    afWlanChannelBandwidth5MHz = 5
    afWlanChannelBandwidth10MHz = 6
End Enum

' afWlanSpectrumAnalysisMode - defines the available spectrum analysis modes.
Public Enum afWlanSpectrumAnalysisMode
    afWlanSpectrumAnalysisGated = 0
    afWlanSpectrumAnalysisNonGated = 1
End Enum

' afWlanSpectralMaskType - defines the type of spectral mask to use for spectral
'                          mask measurements.
Public Enum afWlanSpectralMaskType
    afWlanSpectralMask11b = 1
    afWlanSpectralMaskUserDefined = 2
    afWlanSpectralMask11n = 3
    afWlanSpectralMask11ac = 4
    afWlanSpectralMask11n5GHzBand = 5
    afWlanSpectralMask11aAllBw = 6

    ' Deprecated Values
    afWlanSpectralMask11a = 0   ' instead use afWlanSpectralMask11aAllBw
End Enum

' afWlanBurstProfileMode - defines the modes available when measuring the burst
'                          profile rising and falling edge times.
Public Enum afWlanBurstProfileMode
    afWlanPeakPower = 0 ' afMeasBurstPeakPower
    afWlanAveragePower = 1      ' afMeasBurstAveragePower
End Enum

' afWlanDsssAnalysisMode - defines the modulation analysis modes available
'                          when analysing a WLAN DSSS signal.
Public Enum afWlanDsssAnalysisMode
    afWlanDsssAnalysisStd = 0
    afWlanDsssAnalysisLegacy = 1
End Enum

' afWlanSystemType - defines the different types of modulation system.
Public Enum afWlanSystemType
    afWlanSystemTypeNotDefined = -1
    afWlanOfdm = 0
    afWlanDsss = 1
    afWlanDsssOfdm = 2
End Enum

' afWlan11nHTFormat - defines the different 802.11n HT formats.
Public Enum afWlan11nHTFormat
    afWlan11nHTFormatNotDefined = -1
    afWlan11nHTFormatMixed = 0  ' 802.11n HT 20/40MHz Mixed Format
    afWlan11nHTFormatGreenfield = 1     ' 802.11n HT 20/40MHz Greenfield Format
End Enum

' afWlanModType - defines the different modulation types.
Public Enum afWlanModType
    afWlanModTypeNotDefined = -1
    afWlanDbpsk = 0
    afWlanDqpsk = 1
    afWlanCck = 2
    afWlanPbcc = 3
    afWlanBpsk = 4
    afWlanQpsk = 5
    afWlan16Qam = 6
    afWlan64Qam = 7
    afWlan256Qam = 8
End Enum

' afWlanFreqScale - defines the scaling of the frequency axis when getting the
'                   spectrum trace. For example, the frequency axis results
'                   could be in MHz.
Public Enum afWlanFreqScale
    afWlanFreqScale_Hertz = 0   ' afMeasFreqUnits_Hz
    afWlanFreqScale_kiloHertz = 1       ' afMeasFreqUnits_kHz
    afWlanFreqScale_MegaHertz = 2       ' afMeasFreqUnits_MHz
    afWlanFreqScale_GigaHertz = 3       ' afMeasFreqUnits_GHz
End Enum

' afWlanPassFailResult - defines the possible result values for pass/fail tests.
'                        This includes a 'not available' value when neither
'                        pass or fail could be determined.
Public Enum afWlanPassFailResult
    afWlanPass = 0      ' afMeasPass
    afWlanFail = -1     ' afMeasFail
    afWlanNotAvailable = -2     ' afMeasNotAvailable
End Enum

' afWlanOfdmEqMode - defines the equalisation Mode
Public Enum afWlanOfdmEqMode
    afWlanOfdmEqPreamble = 0    ' equalisation on Preamble only
    afWlanOfdmEqPreambleAndData = 1     ' equalisation on Preamble,Pilots and Data
End Enum

' afWlanVHTDataLengthSource - defines the field to use for detecting number of data symbols in
'                             case of VHT packets.
Public Enum afWlanVHTDataLengthSource
    afWlanAll = &H0
    afWlanLsig = &H1
    afWlanSigB = &H2
    afWlanBurstEdges = &H3
End Enum

' afWlanFrequencySegment - defines the Primary or Secondary segment
Public Enum afWlanFrequencySegment
    afWlanPrimary = 0   ' Primary segment
    afWlanSecondary = 1 ' Secondary segment
End Enum

' afWlanCapture - defines the lower, upper or combined in case of stitched 160MHz analysis
Public Enum afWlanCapture
    afWlanLower = 0     ' Lower segment
    afWlanUpper = 1     ' Upper segment
    afWlanCombined = 2  ' Combined (used for results only)
End Enum

' afWlanMultipleCaptureMode - defines the synchronisation of the input captures
Public Enum afWlanMultipleCaptureMode
    afWlanSequential = 0        ' Input Captures occurred sequentially
    afWlanConcurrent = 1        ' Input Captures occurred concurrently
End Enum


'**** Exported Functions ****

'** Methods **
Declare Function afWlanDll_CreateObject Lib "afWlanDll.dll" (ByRef ptrID As Long) As Long

Declare Function afWlanDll_DestroyObject Lib "afWlanDll.dll" (ByVal nID As Long) As Long

Declare Function afWlanDll_Analyse Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                  ByVal measurements As Long, _
                                                                  ByRef ptrIData As Single, _
                                                                  ByRef ptrQData As Single, _
                                                                  ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ

Declare Function afWlanDll_AnalyseIQ Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                    ByVal measurements As Long, _
                                                                    ByRef ptrIData As Single, _
                                                                    ByRef ptrQData As Single, _
                                                                    ByVal numIQ As Long, _
                                                                    ByVal analyseOffset As Long, _
                                                                    ByVal analyseNumber As Long, _
                                                                    ByVal tag As Double, _
                                                                    ByVal Key As Double) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ

Declare Function afWlanDll_AnalyseMIMO Lib "afWlanDll.dll" (ByRef ptrIDs As Long, _
                                                                  ByVal numIDs As Long) As Long
' ptrIDs must be an array of size numIDs

Declare Function afWlanDll_IQ_Subset_Key_Get Lib "afWlanDll.dll" (ByRef ptrIData As Single, _
                                                                            ByRef ptrQData As Single, _
                                                                            ByVal numIQ As Long, _
                                                                            ByVal tag As Double, _
                                                                            ByVal Key As Double, _
                                                                            ByVal subsetOffset As Long, _
                                                                            ByVal subsetNumber As Long, _
                                                                            ByRef newTag As Double, _
                                                                            ByRef newKey As Double) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ

Declare Function afWlanDll_IQ_MergedSet_Key_Get Lib "afWlanDll.dll" (ByRef ptrIData1 As Single, _
                                                                               ByRef ptrQData1 As Single, _
                                                                               ByVal numIQ1 As Long, _
                                                                               ByVal tag1 As Double, _
                                                                               ByVal key1 As Double, _
                                                                               ByRef ptrIData2 As Single, _
                                                                               ByRef ptrQData2 As Single, _
                                                                               ByVal numIQ2 As Long, _
                                                                               ByVal tag2 As Double, _
                                                                               ByVal key2 As Double, _
                                                                               ByRef newTag As Double, _
                                                                               ByRef newKey As Double) As Long
' ptrIData1 must be an array of size numIQ1
' ptrQData1 must be an array of size numIQ1
' ptrIData2 must be an array of size numIQ2
' ptrQData2 must be an array of size numIQ2

Declare Function afWlanDll_IQ_Key_Convert16 Lib "afWlanDll.dll" (ByRef data As Integer, _
                                                                           ByVal dataLength As Long, _
                                                                           ByVal tag As Double, _
                                                                           ByVal Key As Double, _
                                                                           ByRef iBuffer As Single, _
                                                                           ByRef qBuffer As Single, _
                                                                           ByVal numIQ As Long, _
                                                                           ByVal scale_ As Single, _
                                                                           ByRef newTag As Double, _
                                                                           ByRef newKey As Double) As Long
' data must be an array of size dataLength

Declare Function afWlanDll_IQ_Key_Convert32 Lib "afWlanDll.dll" (ByRef data As Long, _
                                                                           ByVal dataLength As Long, _
                                                                           ByVal tag As Double, _
                                                                           ByVal Key As Double, _
                                                                           ByRef iBuffer As Single, _
                                                                           ByRef qBuffer As Single, _
                                                                           ByVal numIQ As Long, _
                                                                           ByVal scale_ As Single, _
                                                                           ByRef newTag As Double, _
                                                                           ByRef newKey As Double) As Long
' data must be an array of size dataLength

Declare Function afWlanDll_GetTraceDataLength Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                             ByVal traceType As Long, _
                                                                             ByRef ptrLength As Long) As Long

Declare Function afWlanDll_GetTraceData Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                       ByVal traceType As Long, _
                                                                       ByRef ptrX As Double, _
                                                                       ByRef ptrY As Double, _
                                                                       ByVal numPoints As Long) As Long
' ptrX must be an array of size numPoints
' ptrY must be an array of size numPoints

Declare Function afWlanDll_GetErrorMsgLength Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                            ByVal errorCode As Long, _
                                                                            ByRef ptrLength As Long) As Long

Declare Function afWlanDll_GetErrorMsg Lib "afWlanDll.dll" (ByVal nID As Long, _
                                                                      ByVal errorCode As Long, _
                                                                      ByVal ptrBuffer As String, _
                                                                      ByVal bufferSize As Long) As Long ' NB byval on the string

Declare Function afWlanDll_MaxErrorMsgLength_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrLength As Long) As Long

Declare Function afWlanDll_GetVersion Lib "afWlanDll.dll" (ByRef ptrVersion As Long) As Long
'* NOTE - ptrVersion must be a pointer to an array of 4 elements.*'



'* Configuration Methods and Properties *
Declare Function afWlanDll_GetMinSamplingFreq Lib "afWlanDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSampFreq As Double) As Long
Declare Function afWlanDll_GetRecommendedSamplingFreq Lib "afWlanDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSampFreq As Double) As Long
Declare Function afWlanDll_GetMinMeasurementSpan Lib "afWlanDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSpan As Double) As Long

Declare Function afWlanDll_AutoDetectSignalFields_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal autoDetectFlag As Long) As Long
Declare Function afWlanDll_AutoDetectSignalFields_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAutoDetectFlag As Long) As Long

Declare Function afWlanDll_UserDefined_MCS_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal mcs As Long) As Long
Declare Function afWlanDll_UserDefined_MCS_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMcs As Long) As Long

Declare Function afWlanDll_UserDefined_ShortGI_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal shortGI As Long) As Long
Declare Function afWlanDll_UserDefined_ShortGI_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrShortGI As Long) As Long

Declare Function afWlanDll_UserDefined_Nss_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal nss As Long) As Long
Declare Function afWlanDll_UserDefined_Nss_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNss As Long) As Long

Declare Function afWlanDll_SamplingFreq_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal sampFreq As Double) As Long
Declare Function afWlanDll_SamplingFreq_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrSampFreq As Double) As Long

Declare Function afWlanDll_ChannelOffset_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal chanOffset As Long) As Long
Declare Function afWlanDll_ChannelOffset_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChanOffset As Long) As Long

Declare Function afWlanDll_CaptureOffset_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal captureOffset As Double) As Long
Declare Function afWlanDll_CaptureOffset_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCaptureOffset As Double) As Long

Declare Function afWlanDll_RfLevelCal_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal levelcal As Single) As Long
Declare Function afWlanDll_RfLevelCal_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrLevelCal As Single) As Long

Declare Function afWlanDll_AnalysisMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal anaMode As Long) As Long
Declare Function afWlanDll_AnalysisMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAnaMode As Long) As Long

Declare Function afWlanDll_ChannelBandwidth_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal chanBW As Long) As Long
Declare Function afWlanDll_ChannelBandwidth_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChanBW As Long) As Long

Declare Function afWlanDll_FrequencySegment_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal frequencySegment As Long) As Long
Declare Function afWlanDll_FrequencySegment_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFrequencySegment As Long) As Long
'* Clears results before performing stitched 160MHz analysis *
Declare Function afWlanDll_Stitched160MHz_Reset Lib "afWlanDll.dll" (ByVal nID As Long) As Long

Declare Function afWlanDll_Stitched160MHzMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal flag As Long) As Long
Declare Function afWlanDll_Stitched160MHzMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFlag As Long) As Long

Declare Function afWlanDll_Capture160MHz_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal captureSegment As Long) As Long
Declare Function afWlanDll_Capture160MHz_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCaptureSegment As Long) As Long

Declare Function afWlanDll_AntennaNumber_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal antNum As Long) As Long
Declare Function afWlanDll_AntennaNumber_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAntNum As Long) As Long

Declare Function afWlanDll_StreamNumber_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal strmNum As Long) As Long
Declare Function afWlanDll_StreamNumber_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStrmNum As Long) As Long

Declare Function afWlanDll_VHTDataLengthSource_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal field As Long) As Long
Declare Function afWlanDll_VHTDataLengthSource_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrField As Long) As Long

'* Spectrum *
Declare Function afWlanDll_DigitizerSpan_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal span As Double) As Long
Declare Function afWlanDll_DigitizerSpan_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrSpan As Double) As Long

Declare Function afWlanDll_MeasurementSpan_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal span As Double) As Long
Declare Function afWlanDll_MeasurementSpan_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrSpan As Double) As Long

Declare Function afWlanDll_Spectrum_AnalysisMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal anaMode As Long) As Long
Declare Function afWlanDll_Spectrum_AnalysisMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAnaMode As Long) As Long

Declare Function afWlanDll_Spectrum_NumStitches_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumStitches As Long) As Long

Declare Function afWlanDll_Spectrum_StitchIndex_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal Index As Long) As Long
Declare Function afWlanDll_Spectrum_StitchIndex_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrIndex As Long) As Long

Declare Function afWlanDll_Spectrum_StitchOffsetFreq_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreq As Double) As Long

Declare Function afWlanDll_Spectrum_FreqAxis_Centre_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal centreFreq As Double) As Long
Declare Function afWlanDll_Spectrum_FreqAxis_Centre_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCentreFreq As Double) As Long

Declare Function afWlanDll_Spectrum_FreqAxis_Scale_Units_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal freqUnits As Long) As Long
Declare Function afWlanDll_Spectrum_FreqAxis_Scale_Units_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreqUnits As Long) As Long

'* Spectral Mask *
Declare Function afWlanDll_SpectralMask_Type_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal maskType As Long) As Long
Declare Function afWlanDll_SpectralMask_Type_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMaskType As Long) As Long
Declare Function afWlanDll_SpectralMask_SetUserDefined Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMaskFreqs As Double, ByRef ptrMaskLevels As Single, ByVal numMaskPoints As Long) As Long
' ptrMaskFreqs must be an array of size numMaskPoints
' ptrMaskLevels must be an array of size numMaskPoints
'* Spectral Mask for 11ac 80+80 signals *
Declare Function afWlanDll_CombineSpectrumSegments Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPrimaryFreq As Double, ByRef ptrPrimaryPower As Double, ByVal primaryMaskPeakLevel As Single, ByRef ptrSecondaryFreq As Double, ByRef ptrSecondaryPower As Double, ByVal secondaryMaskPeakLevel As Single, ByVal numFreqPoints As Long) As Long
' ptrPrimaryFreq must be an array of size numFreqPoints
' ptrPrimaryPower must be an array of size numFreqPoints
' ptrSecondaryFreq must be an array of size numFreqPoints
' ptrSecondaryPower must be an array of size numFreqPoints
'* Resets and clears all stored MIMO data. *
Declare Function afWlanDll_MIMO_Reset Lib "afWlanDll.dll" (ByVal nID As Long) As Long

'* Burst Position and Length *
Declare Function afWlanDll_Burst_Position_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal burstPos As Long) As Long
Declare Function afWlanDll_Burst_Position_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrBurstPos As Long) As Long

Declare Function afWlanDll_Burst_Length_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal burstLen As Long) As Long
Declare Function afWlanDll_Burst_Length_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrBurstLen As Long) As Long

Declare Function afWlanDll_Burst_PreTriggerTime_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal preTriggerTime As Single) As Long
Declare Function afWlanDll_Burst_PreTriggerTime_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPreTriggerTime As Single) As Long

Declare Function afWlanDll_Burst_MinimumOnTime_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal minOnTime As Single) As Long
Declare Function afWlanDll_Burst_MinimumOnTime_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMinOnTime As Single) As Long

Declare Function afWlanDll_Burst_MinimumOffTime_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal minOffTime As Single) As Long
Declare Function afWlanDll_Burst_MinimumOffTime_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMinOffTime As Single) As Long

Declare Function afWlanDll_Burst_IntegrationTime_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal time As Single) As Long
Declare Function afWlanDll_Burst_IntegrationTime_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTime As Single) As Long

Declare Function afWlanDll_Burst_IntegrationSkipTime_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal skipTime As Single) As Long
Declare Function afWlanDll_Burst_IntegrationSkipTime_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrSkipTime As Single) As Long

Declare Function afWlanDll_Burst_ComparatorDelay_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal delay As Long) As Long
Declare Function afWlanDll_Burst_ComparatorDelay_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDelay As Long) As Long

Declare Function afWlanDll_Burst_RisingEdge_Threshold_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afWlanDll_Burst_RisingEdge_Threshold_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

Declare Function afWlanDll_Burst_FallingEdge_Threshold_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afWlanDll_Burst_FallingEdge_Threshold_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

Declare Function afWlanDll_Burst_DetectionMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afWlanDll_Burst_DetectionMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

'* Burst Profile *
Declare Function afWlanDll_BurstProfile_Mode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afWlanDll_BurstProfile_Mode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

'* Modulation Accuracy *
Declare Function afWlanDll_Dsss_AnalysisMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal dsssMode As Long) As Long
Declare Function afWlanDll_Dsss_AnalysisMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDsssMode As Long) As Long

'* Reference Filter *
Declare Function afWlanDll_Dsss_ReferenceFilter_Type_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal refDsssFilter As Long) As Long
Declare Function afWlanDll_Dsss_ReferenceFilter_Type_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRefDsssFilter As Long) As Long

'* Reference RRC filter parameter *
Declare Function afWlanDll_Dsss_ReferenceFilter_Alpha_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal refDsssFilterAlpha As Single) As Long
Declare Function afWlanDll_Dsss_ReferenceFilter_Alpha_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRefDsssFilterAlpha As Single) As Long

'* Reference Gaussian filter parameter *
Declare Function afWlanDll_Dsss_ReferenceFilter_BT_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal refDsssFilterBT As Single) As Long
Declare Function afWlanDll_Dsss_ReferenceFilter_BT_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRefDsssFilterBT As Single) As Long

'* Equalization Filter *
Declare Function afWlanDll_Dsss_EqualizationFilter_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal eqDsssFilterState As Long) As Long
Declare Function afWlanDll_Dsss_EqualizationFilter_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEqDsssFilterState As Long) As Long

'* Dsss Equalization filter parameter *
Declare Function afWlanDll_Dsss_EqualizationFilter_Length_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal eqDsssFilterLength As Long) As Long
Declare Function afWlanDll_Dsss_EqualizationFilter_Length_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEqDsssFilterLength As Long) As Long

Declare Function afWlanDll_CarrierIndex_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal carrierIndex As Long) As Long
Declare Function afWlanDll_CarrierIndex_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarrierIndex As Long) As Long

Declare Function afWlanDll_PilotTracking_Amplitude_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal trackAmplitude As Long) As Long
Declare Function afWlanDll_PilotTracking_Amplitude_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTrackAmplitude As Long) As Long

Declare Function afWlanDll_PilotTracking_Phase_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal trackPhase As Long) As Long
Declare Function afWlanDll_PilotTracking_Phase_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTrackPhase As Long) As Long

Declare Function afWlanDll_PilotTracking_Timing_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal trackTiming As Long) As Long
Declare Function afWlanDll_PilotTracking_Timing_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTrackTiming As Long) As Long

'* Spectral Flatness *
Declare Function afWlanDll_SpectralFlatness_Mode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afWlanDll_SpectralFlatness_Mode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afWlanDll_SpectralFlatness_Upper_SetUserLimits Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarriers As Long, ByRef ptrLimits As Single, ByVal numPoints As Long) As Long
' ptrCarriers must be an array of size numPoints
' ptrLimits must be an array of size numPoints
Declare Function afWlanDll_SpectralFlatness_Upper_GetUserLimits Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarriers As Long, ByRef ptrLimits As Single, ByVal numPoints As Long) As Long
' ptrCarriers must be an array of size numPoints
' ptrLimits must be an array of size numPoints

Declare Function afWlanDll_SpectralFlatness_Upper_NumUserPoints_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumPoints As Long) As Long

'* Equalization Mode *
Declare Function afWlanDll_Ofdm_EqualizationMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal equalizationMode As Long) As Long
Declare Function afWlanDll_Ofdm_EqualizationMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEqualizationMode As Long) As Long

Declare Function afWlanDll_SpectralFlatness_Lower_SetUserLimits Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarriers As Long, ByRef ptrLimits As Single, ByVal numPoints As Long) As Long
' ptrCarriers must be an array of size numPoints
' ptrLimits must be an array of size numPoints
Declare Function afWlanDll_SpectralFlatness_Lower_GetUserLimits Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarriers As Long, ByRef ptrLimits As Single, ByVal numPoints As Long) As Long
' ptrCarriers must be an array of size numPoints
' ptrLimits must be an array of size numPoints

Declare Function afWlanDll_SpectralFlatness_Lower_NumUserPoints_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumPoints As Long) As Long

Declare Function afWlanDll_NumSymbolsToAnalyse_AutoDetect_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal autoDetect As Long) As Long
Declare Function afWlanDll_NumSymbolsToAnalyse_AutoDetect_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAutoDetect As Long) As Long

Declare Function afWlanDll_NumSymbolsToAnalyse_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal numSymbols As Long) As Long
Declare Function afWlanDll_NumSymbolsToAnalyse_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumSymbols As Long) As Long

'* synchronisation between input captures *
Declare Function afWlanDll_MultipleCaptureMode_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal multipleCaptureMode As Long) As Long
Declare Function afWlanDll_MultipleCaptureMode_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMultipleCaptureMode As Long) As Long

'* Composite MIMO specific *
Declare Function afWlanDll_UseReferenceFile_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal useRefStrmSymbs As Long) As Long
Declare Function afWlanDll_UseReferenceFile_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrUseRefStrmSymbs As Long) As Long

Declare Function afWlanDll_SetReferenceFilename Lib "afWlanDll.dll" (ByVal nID As Long, ByVal ptrFilename As String) As Long

'* Results Methods and Properties *
Declare Function afWlanDll_CW_Freq_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCwFreq As Double) As Long
Declare Function afWlanDll_CW_AveragePower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPower As Single) As Long

'* Spectrum Analysis *
Declare Function afWlanDll_OccupiedBandwidth_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrOccBw As Double) As Long
Declare Function afWlanDll_SpectralMask_PassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_SpectralMask_FailFreq_Absolute_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreq As Double) As Long
Declare Function afWlanDll_SpectralMask_FailFreq_Relative_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreq As Double) As Long
Declare Function afWlanDll_SpectralMask_FailLevel_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrLevel As Single) As Long
Declare Function afWlanDll_SpectralMask_PeakLevel_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrLevel As Single) As Long
Declare Function afWlanDll_GetAcp Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAcpResults As Single, ByVal numChannels As Long) As Long
' ptrAcpResults must be an array of size numChannels

'* Burst Profile *
Declare Function afWlanDll_BurstProfile_PassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_BurstProfile_PeakPower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPower As Single) As Long
Declare Function afWlanDll_BurstProfile_AveragePower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPower As Single) As Long
Declare Function afWlanDll_BurstProfile_RisingEdge_PassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_BurstProfile_RisingEdge_Time_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTime As Double) As Long
Declare Function afWlanDll_BurstProfile_FallingEdge_PassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_BurstProfile_FallingEdge_Time_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrTime As Double) As Long

'* Modulation Accuracy - Common 802.11a/b/g/n/ac results *
Declare Function afWlanDll_SystemType_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrType As Long) As Long
Declare Function afWlanDll_ModulationType_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrModType As Long) As Long
Declare Function afWlanDll_DataRatekbps_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDataRatekbps As Long) As Long
Declare Function afWlanDll_NumberOfPsduBits_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumBits As Long) As Long
Declare Function afWlanDll_NumberOfPsduSymbols_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumSymbols As Long) As Long
Declare Function afWlanDll_FreqError_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreqError As Single) As Long
Declare Function afWlanDll_CarrierLeak_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCarrierLeak As Single) As Long
Declare Function afWlanDll_IQImbalance_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrIqImbalance As Single) As Long
Declare Function afWlanDll_IQSkew_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrIqSkew As Single) As Long
Declare Function afWlanDll_Evm_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEvm As Single) As Long
Declare Function afWlanDll_Rce_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRce As Single) As Long

'* 802.11b specific *
Declare Function afWlanDll_Evm_Peak_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPeakEvm As Single) As Long
Declare Function afWlanDll_Rce_Peak_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPeakRce As Single) As Long
Declare Function afWlanDll_ChipClkError_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrError As Single) As Long

'* 802.11a specific *
Declare Function afWlanDll_Evm_Carrier_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEvm As Single) As Long
Declare Function afWlanDll_Rce_Carrier_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRce As Single) As Long
Declare Function afWlanDll_Evm_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEvm As Single) As Long
Declare Function afWlanDll_Rce_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRce As Single) As Long
Declare Function afWlanDll_Evm_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrEvm As Single) As Long
Declare Function afWlanDll_Rce_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrRce As Single) As Long
Declare Function afWlanDll_SymbolClkError_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrError As Single) As Long
Declare Function afWlanDll_SpectralFlatness_OverallPassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_SpectralFlatness_UpperPassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afWlanDll_SpectralFlatness_LowerPassFail_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long

'* 802.11n specific *
Declare Function afWlanDll_HT_Format_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFormat As Long) As Long
'* 802.11n/ac specific *
Declare Function afWlanDll_MCS_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMcs As Long) As Long
Declare Function afWlanDll_NumAntennas_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumAntennas As Long) As Long
Declare Function afWlanDll_ShortGI_Detected_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrShortGI As Long) As Long
Declare Function afWlanDll_CrossPower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCrossPower As Single) As Long
Declare Function afWlanDll_Channel_Evm_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChEvmRms As Single) As Long
Declare Function afWlanDll_Channel_Evm_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChEvmDataCarrier As Single) As Long
Declare Function afWlanDll_Channel_Evm_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChEvmPilotCar As Single) As Long
Declare Function afWlanDll_Channel_Rce_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChRceRms As Single) As Long
Declare Function afWlanDll_Channel_Rce_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChRceDataCarrier As Single) As Long
Declare Function afWlanDll_Channel_Rce_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChRcePilotCarrier As Single) As Long
Declare Function afWlanDll_Channel_CrossPower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrChCrossPower As Single) As Long
Declare Function afWlanDll_Stream_Evm_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStEvmRms As Single) As Long
Declare Function afWlanDll_Stream_Evm_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStEvmDataCarrier As Single) As Long
Declare Function afWlanDll_Stream_Evm_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStEvmPilotCarrier As Single) As Long
Declare Function afWlanDll_Stream_Rce_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStRceRms As Single) As Long
Declare Function afWlanDll_Stream_Rce_DataCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStRceDataCarrier As Single) As Long
Declare Function afWlanDll_Stream_Rce_PilotCarriers_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStRcePilotCarrier As Single) As Long
Declare Function afWlanDll_Stream_Composite_Evm_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStmCmpsiteEvmRms As Single) As Long
Declare Function afWlanDll_Stream_Composite_Rce_Rms_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStmCmpsiteRceRms As Single) As Long
Declare Function afWlanDll_MeanFrequencyError_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreqError As Single) As Long
Declare Function afWlanDll_MeanSymbolClkError_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrSymClkError As Single) As Long
Declare Function afWlanDll_GetMIMOChannelMatrixElement_Cartesian Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrReal As Single, ByRef ptrImaginary As Single) As Long
Declare Function afWlanDll_GetMIMOChannelMatrixElement_Polar Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMagnitude As Single, ByRef ptrAngle As Single) As Long

'* Decoded header fields *
Declare Function afWlanDll_DecodedBits_Lsig_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_DecodedBits_HtSig1_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_DecodedBits_HtSig2_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_DecodedBits_VhtSigA1_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_DecodedBits_VhtSigA2_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_DecodedBits_VhtSigB_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDecodedBits As Long) As Long
Declare Function afWlanDll_Status_LsigParity_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStatus As Long) As Long
Declare Function afWlanDll_Status_HtSigCrc_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStatus As Long) As Long
Declare Function afWlanDll_Status_VhtSigACrc_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrStatus As Long) As Long

'* Deprecated Functions *
Declare Function afWlanDll_Spectrum_FreqAxisCentre_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal centreFreq As Double) As Long
Declare Function afWlanDll_Spectrum_FreqAxisCentre_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCentreFreq As Double) As Long
Declare Function afWlanDll_Spectrum_FreqAxisScale_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal freqUnits As Long) As Long
Declare Function afWlanDll_Spectrum_FreqAxisScale_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreqUnits As Long) As Long
Declare Function afWlanDll_BurstPosition_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal burstPos As Long) As Long
Declare Function afWlanDll_BurstPosition_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrBurstPos As Long) As Long
Declare Function afWlanDll_BurstLength_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal burstLen As Long) As Long
Declare Function afWlanDll_BurstLength_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrBurstLen As Long) As Long
Declare Function afWlanDll_CaptureThreshold_Set Lib "afWlanDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afWlanDll_CaptureThreshold_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long
Declare Function afWlanDll_CWFreq_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCwFreq As Double) As Long
Declare Function afWlanDll_CWAveragePower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPower As Single) As Long
Declare Function afWlanDll_SpectralMask_FailFreq_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrFreq As Double) As Long
'* NOTE - ptrAcpResults must be a pointer to an array of 5 elements *
Declare Function afWlanDll_GetAdjChanPower Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrAcpResults As Single) As Long
Declare Function afWlanDll_DataRate_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrDataRate As Long) As Long
Declare Function afWlanDll_Evm_Peak11bStd_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrPeakEvm As Single) As Long
Declare Function afWlanDll_HT_MCS_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrMcs As Long) As Long
Declare Function afWlanDll_HT_NumAntennas_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrNumAntennas As Long) As Long
Declare Function afWlanDll_HT_ShortGI_Detected_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrShortGI As Long) As Long
Declare Function afWlanDll_HT_CrossPower_Get Lib "afWlanDll.dll" (ByVal nID As Long, ByRef ptrCrossPower As Single) As Long
