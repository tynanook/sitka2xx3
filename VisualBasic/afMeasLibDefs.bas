Attribute VB_Name = "afMeasLibDefs"
' This header file contains all the common definitions
' for the Analysis Library DLLs.
Option Explicit

' afMeasFilterType - defines the supported filter types.
Public Enum afMeasFilterType
    afMeasFilterNone = 0
    afMeasFilterRc = 1
    afMeasFilterRrc = 2
    afMeasFilterHalfSine = 3
    afMeasFilterRect = 4
    afMeasFilterGaussianRect = 5
    afMeasFilterEdgeTx = 6
    afMeasFilterEdgeMeas = 7
    afMeasFilterGaussian = 5    ' Deprecated: instead use afMeasFilterGaussianRect
End Enum

' afMeasFreqUnits - allows the user to set the frequency axis units for the
'                   spectrum trace. For example using afMeasFreqUnits_MHz,
'                   the frequency axis results would be in MHz.
Public Enum afMeasFreqUnits
    afMeasFreqUnits_Hz = 0
    afMeasFreqUnits_kHz = 1
    afMeasFreqUnits_MHz = 2
    afMeasFreqUnits_GHz = 3
End Enum

' afMeasTimeUnits - allows the user to set the time axis units for a time-based
'                   trace. For example, the time axis could be in ms.
Public Enum afMeasTimeUnits
    afMeasTimeUnits_s = 0
    afMeasTimeUnits_ms = 1
    afMeasTimeUnits_us = 2
    afMeasTimeUnits_ns = 3
End Enum

' afMeasPhaseUnits - allows the user to set the phase axis units for a time-based
'                    trace. For example, the time axis could be in degrees.
Public Enum afMeasPhaseUnits
    afMeasPhaseUnits_Radians = 0
    afMeasPhaseUnits_Degrees = 1
End Enum

' afMeasPassFailResult - defines the possible result values for pass/fail tests.
'                        This includes a 'not available' value where neither
'                        pass or fail could be determined.
Public Enum afMeasPassFailResult
    afMeasPass = 0
    afMeasFail = -1
    afMeasNotAvailable = -2
End Enum

' afMeasBurstProfileMode - defines the modes available when measuring the burst
'                          profile rising and falling edge times.
Public Enum afMeasBurstProfileMode
    afMeasBurstPeakPower = 0
    afMeasBurstAveragePower = 1
End Enum

' afMeasSpectralFlatnessMode - defines the different OFDM spectral flatness modes.
Public Enum afMeasSpectralFlatnessMode
    afMeasSpectralFlatnessStandard = 0  ' use the limits defined in the specification
    afMeasSpectralFlatnessUser = 1      ' or allow the user to define their own limits.
End Enum

' afMeasBitPattern - defines the reference data pattern types
Public Enum afMeasBitPattern
    afMeasBitPatternAllOnes = 0 ' All ones data sequence
    afMeasBitPatternAllZeros = 1        ' All zeros data sequence
    afMeasBitPatternPN9 = 2     ' PN9 data sequence
    afMeasBitPatternPN15 = 3    ' PN15 data sequence
End Enum

' afMeasSpecMaskLevelUnits - allows the user to set the level units
'                            for a user-defined spectrum mask.
Public Enum afMeasSpecMaskLevelUnits
    afMeasSpecMaskLevelUnits_dBm = 0
    afMeasSpecMaskLevelUnits_dBr = 1
End Enum

' afMeasBurstDetectionMode - defines the different burst detection modes
Public Enum afMeasBurstDetectionMode
    afMeasBurstDetectionModeDefault = 0 ' Burst detection thresholds defined by user
    afMeasBurstDetectionModeAutoThreshold = 1   ' Auto-computes burst detection thresholds
End Enum


' afMeasError - defines the different error codes that can be returned by the
'               functions defined in the analysis library DLLs.
Public Enum afMeasError
    afMeasNoError = 0
    afMeasUnknownError = -1
    afMeasFailAllocMem = -2
    afMeasInvalidMemAddress = -3
    afMeasInvalidDllID = -4
    afMeasInvalidInputParameter = -5
    afMeasUnknownTraceType = -6
    afMeasNoTraceDataAvailable = -7
    afMeasInvalidSamplingFreq = -8
    afMeasNoRisingEdge = -9
    afMeasIncompleteBurst = -10
    afMeasNoBurstDefined = -11
    afMeasInsufficientIQ = -12
    afMeasFailToSync = -13
    afMeasNoResultAvailable = -14
    afMeasChannelIdentificationFailed = -15
    afMeasInvalidBufferSize = -16
    afMeasInvalidSpectrumStitchSetup = -18
    afMeasInvalidDigitizerSpan = -19
    afMeasInvalidMeasurementSpan = -20
    afMeasInvalidSpectrumStitchIndex = -21
    afMeasInvalidSpectrum = -22
    afMeasInvalidSpectrumLength = -23
    afMeasInvalidSpectrumMaskDefinition = -24
    afMeasInvalidSpectrumUnits = -25
    afMeasInvalidResolutionBW = -26
    afMeasInvalidOccupiedBWPercentage = -27
    afMeasInvalidAcpParams = -28
    afMeasInvalidSymbolRate = -29
    afMeasInvalidFFTSize = -30
    afMeasInvalidSpectrumWindow = -31
    afMeasInvalidSpectrumWindowLength = -32
    afMeasFailToAnalyseBurstProfile = -33
    afMeasInvalidSpectralFlatnessMode = -34
    afMeasInvalidSpectralFlatnessDefinition = -35
    afMeasFailToEstSymClkError = -36
    afMeasInvalidFilterType = -37
    afMeasInvalidFilterAlpha = -38
    afMeasInvalidFilterBT = -39
    afMeasInvalidNumChannels = -40
    afMeasInvalidChannelBW = -41
    afMeasInvalidChannelSpacing = -42
    afMeasInvalidNumSpectrumMaskPoints = -43
    afMeasInvalidSpectrumMaskFreq = -44
    afMeasInvalidSpectrumMaskLevel = -45
    afMeasInvalidSpectrumMaskLevelUnits = -46
    afMeasInvalidSpectrumMaskMeasBW = -47
    afMeasNoIQGateDefined = -48
    afMeasInvalidIQGateStart = -49
    afMeasInvalidIQGateLength = -50
    afMeasInvalidSubset = -51
    afMeasInvalidKey = -52
    afMeasDigitizerOptionRequired = -53
    afMeasInvalidVideoBW = -54
    afMeasInvalidChannelFreq = -55
    afMeasFunctionNotFound = -32767
End Enum

