Attribute VB_Name = "afSpectrumDll"
' This module contains the information needed for an application
' to use the Spectrum Analysis Library DLL.
' NB - 'afMeasLibDefs.bas' must be added to the project with this module.
Option Explicit

' =============================================================================================================
' Dll installed in system32 folder
' =============================================================================================================

'**** Type Definitions ****

' afSpectrumDllMeasurement - defines the measurements provided by this DLL. Any combination
'                            of these can be passed into afSpectrumDll_Analyse.
Public Enum afSpectrumDllMeasurement
    afSpectrumDllMeasNotDefined = &H0
    afSpectrumDllMeasLocateBurst = &H1
    afSpectrumDllMeasPowerVsFreq = &H2
    afSpectrumDllMeasMaskPowerVsFreq = &H4
    afSpectrumDllMeasAcp = &H8
    afSpectrumDllMeasOccupiedBW = &H10
    afSpectrumDllMeasPowerVsTime = &H20
    afSpectrumDllMeasFrequencyVsTime = &H40
    afSpectrumDllMeasPhaseVsTime = &H80
    afSpectrumDllMeasCcdf = &H100
End Enum

' afSpectrumDllTraceData - defines the different types of trace data that can be retrieved
'                          from this DLL (via afSpectrumDll_GetTraceData) after a
'                          measurement has completed.
Public Enum afSpectrumDllTraceData
    afSpectrumDllTracePowerVsFreq = 0
    afSpectrumDllTraceMaskPowerVsFreq = 1
    afSpectrumDllTraceMaskRefVsFreq = 2
    afSpectrumDllTraceMaskFailVsFreq = 3
    afSpectrumDllTracePowerVsTime = 4
    afSpectrumDllTraceFrequencyVsTime = 5
    afSpectrumDllTracePhaseVsTime = 6
    afSpectrumDllTraceCcdf = 7
    afSpectrumDllTraceNoiseMarker = 8
End Enum

' afSpectrumDllAnalysisMode - defines the different analysis modes provided by this DLL.
Public Enum afSpectrumDllAnalysisMode
    afSpectrumDllAnalysisModeAllIQ = 0
    afSpectrumDllAnalysisModeBurstIQ = 1
End Enum

' afSpectrumDllConfigurationModeType - defines the configuration modes provided by this DLL.
Public Enum afSpectrumDllConfigurationModeType
    afSpectrumDllConfigurationMode_FFTAnalyzer = 0
    afSpectrumDllConfigurationMode_SpectrumAnalyzer = 1
End Enum

' afSpectrumDllWindowType - defines the available spectrum window types.
Public Enum afSpectrumDllWindowType
    afSpectrumWindow_GaussianNoise = 0
    afSpectrumWindow_Gaussian3dB = 1
    afSpectrumWindow_BlackmanHarris = 2
    afSpectrumWindow_5pole = 3
    afSpectrumWindow_FlatTop = 4
End Enum

' afSpectrumDllDetectorModeType - defines the spectrum analysis detector modes provided by this DLL.
Public Enum afSpectrumDllDetectorModeType
    afSpectrumDllDetectorMode_MaxPeak = 1
    afSpectrumDllDetectorMode_MinPeak = 2
    afSpectrumDllDetectorMode_Rms = 3
    afSpectrumDllDetectorMode_Average = 4
    afSpectrumDllDetectorMode_Sample = 5
End Enum

' afSpectrumDllMaskRefLevelMode - defines the modes available to calculate
'                                 the spectrum mask reference level.
Public Enum afSpectrumDllMaskRefLevelMode
    afSpectrumDllMaskRefLevelChannelPower = 0
    afSpectrumDllMaskRefLevelPeakPower = 1
    afSpectrumDllMaskRefLevelUser = 2
End Enum

' afSpectrumDllAcpMode - defines the adjacent channel power measurement modes.
Public Enum afSpectrumDllAcpMode
    afSpectrumDllAcpMode_Auto = 0
    afSpectrumDllAcpMode_User = 1
End Enum

' afSpectrumDllPeakHoldEnabled - defines if the spectrum peak hold function is enabled or not.
Public Enum afSpectrumDllPeakHoldEnabled
    afSpectrumDllPeakHoldEnabled_False = 0
    afSpectrumDllPeakHoldEnabled_True = 1
End Enum


'**** Exported Functions ****

'** Methods **
Declare Function afSpectrumDll_CreateObject Lib "afSpectrumDll.dll" (ByRef ptrID As Long) As Long

Declare Function afSpectrumDll_DestroyObject Lib "afSpectrumDll.dll" (ByVal nID As Long) As Long

Declare Function afSpectrumDll_Analyse Lib "afSpectrumDll.dll" (ByVal nID As Long, _
                                                                  ByVal measurements As Long, _
                                                                  ByRef ptrIData As Single, _
                                                                  ByRef ptrQData As Single, _
                                                                  ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ

Declare Function afSpectrumDll_GetTraceDataLength Lib "afSpectrumDll.dll" (ByVal nID As Long, _
                                                                             ByVal traceType As Long, _
                                                                             ByRef ptrLength As Long) As Long

Declare Function afSpectrumDll_GetTraceData Lib "afSpectrumDll.dll" (ByVal nID As Long, _
                                                                       ByVal traceType As Long, _
                                                                       ByRef ptrX As Double, _
                                                                       ByRef ptrY As Double, _
                                                                       ByVal numPoints As Long) As Long
' ptrX must be an array of size numPoints
' ptrY must be an array of size numPoints

Declare Function afSpectrumDll_GetErrorMsgLength Lib "afSpectrumDll.dll" (ByVal nID As Long, _
                                                                            ByVal errorCode As Long, _
                                                                            ByRef ptrLength As Long) As Long

Declare Function afSpectrumDll_GetErrorMsg Lib "afSpectrumDll.dll" (ByVal nID As Long, _
                                                                      ByVal errorCode As Long, _
                                                                      ByVal ptrBuffer As String, _
                                                                      ByVal bufferSize As Long) As Long ' NB byval on the string

Declare Function afSpectrumDll_MaxErrorMsgLength_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLength As Long) As Long

Declare Function afSpectrumDll_GetVersion Lib "afSpectrumDll.dll" (ByRef ptrVersion As Long) As Long
'* NOTE - ptrVersion must be a pointer to an array of 4 elements.*'



'* Configuration Methods and Properties *
Declare Function afSpectrumDll_GetMinSamplingFreq Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSampFreq As Double) As Long
Declare Function afSpectrumDll_GetMinMeasurementSpan Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrSpan As Double) As Long

Declare Function afSpectrumDll_SamplingFreq_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal sampFreq As Double) As Long
Declare Function afSpectrumDll_SamplingFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrSampFreq As Double) As Long

Declare Function afSpectrumDll_RfLevelCal_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal levelcal As Single) As Long
Declare Function afSpectrumDll_RfLevelCal_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLevelCal As Single) As Long

Declare Function afSpectrumDll_DigitizerSpan_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal digSpan As Double) As Long
Declare Function afSpectrumDll_DigitizerSpan_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrDigSpan As Double) As Long

Declare Function afSpectrumDll_MeasurementSpan_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal span As Double) As Long
Declare Function afSpectrumDll_MeasurementSpan_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrSpan As Double) As Long

Declare Function afSpectrumDll_MeasurementRBW_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measBandwidth As Double) As Long
Declare Function afSpectrumDll_MeasurementRBW_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMeasBandwidth As Double) As Long

Declare Function afSpectrumDll_MeasurementVBW_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measBandwidth As Double) As Long
Declare Function afSpectrumDll_MeasurementVBW_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMeasBandwidth As Double) As Long

Declare Function afSpectrumDll_FFTSize_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal fftSize As Long) As Long
Declare Function afSpectrumDll_FFTSize_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFftSize As Long) As Long
Declare Function afSpectrumDll_RecommendedFFTSize_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrRecommendedFFTSize As Long) As Long
Declare Function afSpectrumDll_GetMinNumSamples Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measurements As Long, ByRef ptrNumMeasSamples As Long, ByRef ptrNumPreSamples As Long, ByRef ptrNumPostSamples As Long) As Long

Declare Function afSpectrumDll_FFTOverlap_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal fftOverlap As Single) As Long
Declare Function afSpectrumDll_FFTOverlap_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFftOverlap As Single) As Long

Declare Function afSpectrumDll_WindowType_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal windowType As Long) As Long
Declare Function afSpectrumDll_WindowType_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrWindowType As Long) As Long

Declare Function afSpectrumDll_AnalysisMode_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afSpectrumDll_AnalysisMode_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afSpectrumDll_DetectorMode_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afSpectrumDll_DetectorMode_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afSpectrumDll_ConfigurationMode_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afSpectrumDll_ConfigurationMode_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

'* Spectrum Stitching *
Declare Function afSpectrumDll_Spectrum_NumStitches_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumStitches As Long) As Long

Declare Function afSpectrumDll_Spectrum_StitchIndex_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal Index As Long) As Long
Declare Function afSpectrumDll_Spectrum_StitchIndex_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrIndex As Long) As Long

Declare Function afSpectrumDll_Spectrum_StitchOffsetFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPtrFreq As Double) As Long

Declare Function afSpectrumDll_Xaxis_CentreFreq_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal centreFreq As Double) As Long
Declare Function afSpectrumDll_Xaxis_CentreFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrCentreFreq As Double) As Long

Declare Function afSpectrumDll_Xaxis_Scaling_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal scaling As Long) As Long
Declare Function afSpectrumDll_Xaxis_Scaling_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrScaling As Long) As Long

'* Burst Position and Length *
Declare Function afSpectrumDll_Burst_Position_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal burstPos As Long) As Long
Declare Function afSpectrumDll_Burst_Position_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrBurstPos As Long) As Long

Declare Function afSpectrumDll_Burst_Length_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal burstLen As Long) As Long
Declare Function afSpectrumDll_Burst_Length_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrBurstLen As Long) As Long

'* Burst Detection *
Declare Function afSpectrumDll_Burst_RisingEdge_Threshold_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afSpectrumDll_Burst_RisingEdge_Threshold_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

Declare Function afSpectrumDll_Burst_FallingEdge_Threshold_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal threshold As Single) As Long
Declare Function afSpectrumDll_Burst_FallingEdge_Threshold_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrThreshold As Single) As Long

Declare Function afSpectrumDll_Burst_PreTriggerTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal preTriggerTime As Single) As Long
Declare Function afSpectrumDll_Burst_PreTriggerTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPreTriggerTime As Single) As Long

Declare Function afSpectrumDll_Burst_IntegrationTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal time As Single) As Long
Declare Function afSpectrumDll_Burst_IntegrationTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrTime As Single) As Long

Declare Function afSpectrumDll_Burst_IntegrationSkipTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal skipTime As Single) As Long
Declare Function afSpectrumDll_Burst_IntegrationSkipTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrSkipTime As Single) As Long

Declare Function afSpectrumDll_Burst_ComparatorDelay_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal delay As Long) As Long
Declare Function afSpectrumDll_Burst_ComparatorDelay_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrDelay As Long) As Long

Declare Function afSpectrumDll_Burst_MinimumOnTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal minOnTime As Single) As Long
Declare Function afSpectrumDll_Burst_MinimumOnTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMinOnTime As Single) As Long

Declare Function afSpectrumDll_Burst_MinimumOffTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal minOffTime As Single) As Long
Declare Function afSpectrumDll_Burst_MinimumOffTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMinOffTime As Single) As Long

'* Gating *
Declare Function afSpectrumDll_Gate_Enabled_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal enabled As Long) As Long
Declare Function afSpectrumDll_Gate_Enabled_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrEnabled As Long) As Long

Declare Function afSpectrumDll_Gate_Start_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal start As Long) As Long
Declare Function afSpectrumDll_Gate_Start_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrStart As Long) As Long

Declare Function afSpectrumDll_Gate_Length_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal length As Long) As Long
Declare Function afSpectrumDll_Gate_Length_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLength As Long) As Long

'* Peak Search *
Declare Function afSpectrumDll_Peak_SearchStart_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal start As Double) As Long
Declare Function afSpectrumDll_Peak_SearchStart_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrStart As Double) As Long

Declare Function afSpectrumDll_Peak_SearchStop_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal stop_ As Double) As Long
Declare Function afSpectrumDll_Peak_SearchStop_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrStop As Double) As Long

'* Spectrum Mask *
Declare Function afSpectrumDll_Mask_RefLevel_Mode_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afSpectrumDll_Mask_RefLevel_Mode_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afSpectrumDll_Mask_RefLevel_ChannelBW_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal channelBW As Double) As Long
Declare Function afSpectrumDll_Mask_RefLevel_ChannelBW_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrChannelBW As Double) As Long

Declare Function afSpectrumDll_Mask_RefLevel_FilterType_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal filter As Long) As Long
Declare Function afSpectrumDll_Mask_RefLevel_FilterType_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFilter As Long) As Long

Declare Function afSpectrumDll_Mask_RefLevel_FilterAlpha_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal alpha As Single) As Long
Declare Function afSpectrumDll_Mask_RefLevel_FilterAlpha_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrAlpha As Single) As Long

Declare Function afSpectrumDll_Mask_RefLevel_User_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal level As Single) As Long
Declare Function afSpectrumDll_Mask_RefLevel_User_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLevel As Single) As Long

Declare Function afSpectrumDll_Mask_SetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMaskFreqs As Double, ByRef ptrMaskLevels As Single, ByRef ptrMaskLevelUnits As Long, ByRef ptrMaskMeasBWs As Double, ByVal numMaskPoints As Long) As Long
' ptrMaskFreqs must be an array of size numMaskPoints
' ptrMaskLevels must be an array of size numMaskPoints
' ptrMaskLevelUnits must be an array of size numMaskPoints
' ptrMaskMeasBWs must be an array of size numMaskPoints
Declare Function afSpectrumDll_Mask_GetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMaskFreqs As Double, ByRef ptrMaskLevels As Single, ByRef ptrMaskLevelUnits As Long, ByRef ptrMaskMeasBWs As Double, ByVal numMaskPoints As Long) As Long
' ptrMaskFreqs must be an array of size numMaskPoints
' ptrMaskLevels must be an array of size numMaskPoints
' ptrMaskLevelUnits must be an array of size numMaskPoints
' ptrMaskMeasBWs must be an array of size numMaskPoints
Declare Function afSpectrumDll_Mask_NumUserPoints_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumPoints As Long) As Long
Declare Function afSpectrumDll_Mask_ResetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long) As Long

'* Occupied Bandwidth *
Declare Function afSpectrumDll_OccupiedBW_MeasWidth_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal measWidth As Double) As Long
Declare Function afSpectrumDll_OccupiedBW_MeasWidth_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMeasWidth As Double) As Long

Declare Function afSpectrumDll_OccupiedBW_Percentage_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal percentage As Single) As Long
Declare Function afSpectrumDll_OccupiedBW_Percentage_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPercentage As Single) As Long

'* ACP *
Declare Function afSpectrumDll_Acp_Mode_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal mode As Long) As Long
Declare Function afSpectrumDll_Acp_Mode_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrMode As Long) As Long

Declare Function afSpectrumDll_Acp_CentreFreq_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal centreFreq As Double) As Long
Declare Function afSpectrumDll_Acp_CentreFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrCentreFreq As Double) As Long

Declare Function afSpectrumDll_Acp_ChanWidth_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal dChanWidth As Double) As Long
Declare Function afSpectrumDll_Acp_ChanWidth_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrDChanWidth As Double) As Long

Declare Function afSpectrumDll_Acp_ChannelSpacing_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal channelSpacing As Double) As Long
Declare Function afSpectrumDll_Acp_ChannelSpacing_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrChannelSpacing As Double) As Long

Declare Function afSpectrumDll_Acp_NumChannels_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal numChannels As Long) As Long
Declare Function afSpectrumDll_Acp_NumChannels_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumChannels As Long) As Long

Declare Function afSpectrumDll_Acp_FilterType_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal filterType As Long) As Long
Declare Function afSpectrumDll_Acp_FilterType_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFilterType As Long) As Long

Declare Function afSpectrumDll_Acp_FilterAlpha_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal filterAlpha As Single) As Long
Declare Function afSpectrumDll_Acp_FilterAlpha_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFilterAlpha As Single) As Long

Declare Function afSpectrumDll_Acp_SetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrChanFreqs As Double, ByRef ptrChanBWs As Double, ByVal numChannels As Long) As Long
' ptrChanFreqs must be an array of size numChannels
' ptrChanBWs must be an array of size numChannels
Declare Function afSpectrumDll_Acp_GetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrChanFreqs As Double, ByRef ptrChanBWs As Double, ByVal numChannels As Long) As Long
' ptrChanFreqs must be an array of size numChannels
' ptrChanBWs must be an array of size numChannels
Declare Function afSpectrumDll_Acp_NumUserChannels_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumChannels As Long) As Long
Declare Function afSpectrumDll_Acp_ResetUserDefined Lib "afSpectrumDll.dll" (ByVal nID As Long) As Long

'* Zero Span *
Declare Function afSpectrumDll_ZeroSpan_TimeScaling_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal timeScaling As Long) As Long
Declare Function afSpectrumDll_ZeroSpan_TimeScaling_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrTimeScaling As Long) As Long

Declare Function afSpectrumDll_ZeroSpan_FrequencyScaling_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal frequencyScaling As Long) As Long
Declare Function afSpectrumDll_ZeroSpan_FrequencyScaling_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFrequencyScaling As Long) As Long

Declare Function afSpectrumDll_ZeroSpan_PhaseUnits_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal units As Long) As Long
Declare Function afSpectrumDll_ZeroSpan_PhaseUnits_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrUnits As Long) As Long

Declare Function afSpectrumDll_ZeroSpan_ReferenceTime_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal refTime As Double) As Long
Declare Function afSpectrumDll_ZeroSpan_ReferenceTime_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrRefTime As Double) As Long

'* Results Methods and Properties *
Declare Function afSpectrumDll_GetPowerAtFreq Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal dFrequencyOffset As Double, ByRef ptrPower As Single) As Long
Declare Function afSpectrumDll_GetNoiseMarkerPowerAtFreq Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal dFrequencyOffset As Double, ByRef ptrNoiseMarkerPower As Single) As Long

Declare Function afSpectrumDll_Peak_Find Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPeakFreq As Double, ByRef ptrPeakPower As Single) As Long
Declare Function afSpectrumDll_Peak_FindNext Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNextPeakFreq As Double, ByRef ptrNextPeakPower As Single) As Long

Declare Function afSpectrumDll_Mask_PassFail_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPassFail As Long) As Long
Declare Function afSpectrumDll_Mask_FailLevel_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFailLevel As Single) As Long
Declare Function afSpectrumDll_Mask_FailFreq_Absolute_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFailFreq As Double) As Long
Declare Function afSpectrumDll_Mask_FailFreq_Relative_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrFailFreq As Double) As Long
Declare Function afSpectrumDll_Mask_RefLevel_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLevel As Single) As Long

Declare Function afSpectrumDll_OccupiedBW_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrOccupiedBW As Double) As Long
Declare Function afSpectrumDll_OccupiedBW_UpperFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrUpperFreq As Double) As Long
Declare Function afSpectrumDll_OccupiedBW_LowerFreq_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrLowerFreq As Double) As Long

Declare Function afSpectrumDll_Acp_GetResults Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrAcpResults As Single, ByVal numResults As Long) As Long
' ptrAcpResults must be an array of size numResults
Declare Function afSpectrumDll_Acp_NumResults_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumResults As Long) As Long

'* Deprecated Functions *
Declare Function afSpectrumDll_ComputeSpectrum Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrIData As Single, ByRef ptrQData As Single, ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ
Declare Function afSpectrumDll_ComputeAcp Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrAcpResults As Single, ByVal numResults As Long) As Long
' ptrAcpResults must be an array of size numResults
Declare Function afSpectrumDll_ResetSpectrum Lib "afSpectrumDll.dll" (ByVal nID As Long) As Long
Declare Function afSpectrumDll_ZeroSpan_PowerVsTime Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrIData As Single, ByRef ptrQData As Single, ByRef ptrTime As Double, ByRef ptrPower As Single, ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ
' ptrTime must be an array of size numIQ
' ptrPower must be an array of size numIQ
Declare Function afSpectrumDll_ZeroSpan_FrequencyVsTime Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrIData As Single, ByRef ptrQData As Single, ByRef ptrTime As Double, ByRef ptrFrequency As Double, ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ
' ptrTime must be an array of size numIQ
' ptrFrequency must be an array of size numIQ
Declare Function afSpectrumDll_ZeroSpan_AveragePower Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrIData As Single, ByRef ptrQData As Single, ByRef ptrAvgPower As Single, ByVal numIQ As Long) As Long
' ptrIData must be an array of size numIQ
' ptrQData must be an array of size numIQ
Declare Function afSpectrumDll_NumAverage_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal numAverage As Long) As Long
Declare Function afSpectrumDll_NumAverage_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrNumAverage As Long) As Long
Declare Function afSpectrumDll_PeakHoldEnabled_Set Lib "afSpectrumDll.dll" (ByVal nID As Long, ByVal peakHoldEnabled As Long) As Long
Declare Function afSpectrumDll_PeakHoldEnabled_Get Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrPeakHoldEnabled As Long) As Long
Declare Function afSpectrumDll_SetUserDefinedAcpChannels Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrChanBandWidths As Double, ByRef ptrChanSpacings As Double, ByVal numChannels As Long) As Long
' ptrChanBandWidths must be an array of size numChannels
' ptrChanSpacings must be an array of size numChannels
Declare Function afSpectrumDll_GetTraceLength Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrTraceLen As Long) As Long
Declare Function afSpectrumDll_GetTrace Lib "afSpectrumDll.dll" (ByVal nID As Long, ByRef ptrX As Double, ByRef ptrY As Double, ByVal numPoints As Long) As Long
' ptrX must be an array of size numPoints
' ptrY must be an array of size numPoints
