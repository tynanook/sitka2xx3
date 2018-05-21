Attribute VB_Name = "TevAXRF"
Option Explicit


   Public Enum AXRF_CHANNEL
        AXRF_CH1
        AXRF_CH2
        AXRF_CH3
        AXRF_CH4
        AXRF_CH5
        AXRF_CH6
        AXRF_CH7
        AXRF_CH8
        AXRF_CH9
        AXRF_CH10
        AXRF_CH11
        AXRF_CH12
        AXRF_CH13
        AXRF_CH14
        AXRF_CH15
        AXRF_CH16
    End Enum

    Public Enum AXRF_GENERATOR
        AXRF_SRCA_GEN
        AXRF_SRCB_GEN
    End Enum


    Public Enum AXRF_ARRAY_TYPE
        AXRF_TIME_DOMAIN
        AXRF_FREQ_DOMAIN
    End Enum

    Public Enum AXRF_SPARAM_FORMAT
        AXRF_POLAR
        AXRF_LOG_POLAR
        AXRF_RECT
    End Enum

    Public Type NIComplexNumber
        real As Double
        imaginary As Double
    End Type


    Public Enum AXRF_RL_FORMAT
        AXRF_VSWR
        AXRF_RLDB
        AXRF_REFL
    End Enum

    Public Enum AXRF_CAL_MODE
        AXRF_CAL_SCALAR
        AXRF_CAL_VECTOR
        AXRF_CAL_NOISE
    End Enum

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Declare Function TevAXRF_Initialize Lib "TevAXRF.dll" () As Long
    Declare Function TevAXRF_InitializeSubSystem Lib "TevAXRF.dll" (ByVal subSystem As Long) As Long
    Declare Function TevAXRF_Close Lib "TevAXRF.dll" () As Long
    Declare Function TevAXRF_CloseSubSystem Lib "TevAXRF.dll" (ByVal subSystem As Long) As Long

    Declare Function TevAXRF_Isolate Lib "TevAXRF.dll" (ByVal channel As Long) As Long
    Declare Function TevAXRF_GateOn Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal generator As AXRF_GENERATOR) As Long
    Declare Function TevAXRF_GateOff Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal generator As AXRF_GENERATOR) As Long


    Declare Function TevAXRF_Source Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal sourceLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_SourceTwoTone Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal sourceLevel_1 As Double, ByVal frequency_1 As Double, ByVal sourceLevel_2 As Double, ByVal frequency_2 As Double) As Long
    Declare Function TevAXRF_SourceMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal sourceLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_SourceMultiChannelMultiLevel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByRef sourceArray As Double, ByVal frequency As Double) As Long

    Declare Function TevAXRF_MeasureSetup Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal measureLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_MeasureSetupMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal measureLevel As Double, ByVal frequency As Double) As Long


    Declare Function TevAXRF_MeasureTriggerArm Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal triggerSource As afDigitizerDll_tsTrigSource_t, ByVal edgeGatePolarity As afDigitizerDll_egpPolarity_t, ByVal timeout As Double) As Long
    Declare Function TevAXRF_MeasureIQTriggerArm Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal numberOfPoints As Long, ByVal triggerSource As afDigitizerDll_tsTrigSource_t, ByVal edgeGatePolarity As afDigitizerDll_egpPolarity_t, ByVal timeout As Double) As Long
    Declare Function TevAXRF_Measure Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL) As Double
    Declare Function TevAXRF_MeasureMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByRef resultsArray As Double) As Long
    Declare Function TevAXRF_MeasureSetupIQ Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal measureLevel As Double, ByVal frequency As Double, ByVal sampleFrequency As Double, ByVal rbw As Double, ByVal measurementSpan As Double) As Long

    Declare Function TevAXRF_MeasureSetupIQMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal measureLevel As Double, ByVal frequency As Double, ByVal sampleFrequency As Double, ByVal rbw As Double, ByVal measurementSpan As Double) As Long
    Declare Function TevAXRF_MeasureIQ Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByRef Result As Double) As Long
    Declare Function TevAXRF_MeasureIQMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByRef resultsArray As Double) As Long

    Declare Function TevAXRF_GetPowerAtFreqOffset Lib "TevAXRF.dll" (ByVal freqOffset As Double, ByRef power As Double) As Long
    Declare Function TevAXRF_GetPowerAtFreqOffsetMultiChannel Lib "TevAXRF.dll" (ByVal numChannels As Long, ByVal freqOffset As Double, ByRef powerArray As Double) As Long
    Declare Function TevAXRF_GetNumberOfIQTraceDataPoints Lib "TevAXRF.dll" () As Long
    Declare Function TevAXRF_GetNumberOfIQTraceDataPointsChannelIndex Lib "TevAXRF.dll" (ByVal channelIndex As Long) As Long
    Declare Function TevAXRF_GetIQTraceData Lib "TevAXRF.dll" (ByRef xData As Double, ByRef yData As Double) As Long
    Declare Function TevAXRF_GetIQTraceDataChannelIndex Lib "TevAXRF.dll" (ByVal channelIndex As Long, ByRef xData As Double, ByRef yData As Double) As Long

    Declare Function TevAXRF_MeasureArray Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByRef dataArray As Double, ByVal arrayType As AXRF_ARRAY_TYPE) As Double
    Declare Function TevAXRF_MeasureArrayMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByRef dataArray As Double, ByVal arrayType As AXRF_ARRAY_TYPE) As Long
    Declare Function TevAXRF_MeasureArrayIQ Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal numberOfPoints As Long, ByRef iDataArray As Single, ByRef iDataArray As Single) As Long

    Declare Function TevAXRF_MeasureSparametersSetup Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal sourceLevel As Double, ByVal measureLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_MeasureSparametersSetupMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal sourceLevel As Double, ByVal measureLevel As Double, ByVal frequency As Double) As Long

    Declare Function TevAXRF_MeasureSparameters Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal format As AXRF_SPARAM_FORMAT, ByRef spResults As NIComplexNumber) As Long
    Declare Function TevAXRF_MeasureSparametersMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal format As AXRF_SPARAM_FORMAT, ByRef spResults As NIComplexNumber) As Long
    Declare Function TevAXRF_MeasureSparametersArray Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal format As AXRF_SPARAM_FORMAT, ByRef spResults As NIComplexNumber, ByRef rawDataArray As Double) As Long
    Declare Function TevAXRF_MeasureSparametersArrayMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal format As AXRF_SPARAM_FORMAT, ByRef spResults As NIComplexNumber, ByRef rawDataArray As Double) As Long


    Declare Function TevAXRF_MeasureReturnLossSetup Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal sourceLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_MeasureReturnLossSetupMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal measureLevel As Double, ByVal frequency As Double) As Long
    Declare Function TevAXRF_MeasureReturnLoss Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal format As AXRF_RL_FORMAT, ByRef rlResults As NIComplexNumber) As Long
    Declare Function TevAXRF_MeasureReturnLossMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal format As AXRF_RL_FORMAT, ByRef rlResults As NIComplexNumber) As Long
    Declare Function TevAXRF_MeasureReturnLossArray Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal format As AXRF_RL_FORMAT, ByRef rlResults As NIComplexNumber, ByRef rawDataArray As Double) As Long
    Declare Function TevAXRF_MeasureReturnLossArrayMultiChannel Lib "TevAXRF.dll" (ByRef channelArray As AXRF_CHANNEL, ByVal numChannels As Long, ByVal format As AXRF_RL_FORMAT, ByRef rlResults As NIComplexNumber, ByRef rawDataArray As Double) As Long

    Declare Function TevAXRF_SetNoiseFigureSampleFrequency Lib "TevAXRF.dll" (ByVal sampleFrequency As Double) As Long
    Declare Function TevAXRF_MeasureNoiseFigureSetup Lib "TevAXRF.dll" (ByVal sourceChannel As AXRF_CHANNEL, ByVal measureChannel As AXRF_CHANNEL, ByVal frequency As Double, ByVal numberPoints As Long) As Long
    Declare Function TevAXRF_MeasureNoiseFigureYfactor Lib "TevAXRF.dll" (ByVal sourceChannel As AXRF_CHANNEL, ByVal measureChannel As AXRF_CHANNEL, ByVal aveCount As Long, ByRef noiseFigure As Double, ByRef noiseGain As Double) As Long
    Declare Function TevAXRF_SetIQSampleFrequency Lib "TevAXRF.dll" (ByVal measureChannel As AXRF_CHANNEL, ByVal sampleFrequency As Double) As Long
    Declare Function TevAXRF_SetFastPowerSampleFrequency Lib "TevAXRF.dll" (ByVal measureChannel As AXRF_CHANNEL, ByVal sampleFrequency As Double) As Long
    Declare Function TevAXRF_SetFastPowerMode Lib "TevAXRF.dll" (ByVal state As Long) As Long

    Declare Function TevAXRF_GetDigitizerHandle Lib "TevAXRF.dll" (ByVal subSystem As Long, ByRef digitizerId As Long) As Long
    Declare Function TevAXRF_GetSigGenHandle Lib "TevAXRF.dll" (ByVal subSystem As Long, ByRef sigGenId As Long) As Long
    Declare Function TevAXRF_GetSigGenBHandle Lib "TevAXRF.dll" (ByVal subSystem As Long, ByRef sigGenId As Long) As Long
    Declare Function TevAXRF_GetMeasureFactor Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByRef measureFactor As Double) As Long
    Declare Function TevAXRF_GetSourceFactor Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByRef sourceFactor As Double) As Long
    Declare Function TevAXRF_GetNoiseFactor Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal frequency As Double, ByRef enrFactor As Double, ByRef Pon As Double, ByRef Poff As Double) As Long
    Declare Function TevAXRF_IsCalibrationValid Lib "TevAXRF.dll" (ByVal subSystem As Long, ByVal calMode As AXRF_CAL_MODE, ByRef errorCode As Long) As Long
    Declare Function TevAXRF_LoadModulationFile Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal modulationFile As String) As Long
    Declare Function TevAXRF_UnloadModulationFile Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal modulationFile As String) As Long
    Declare Function TevAXRF_StartModulation Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal modulationFile As String) As Long
    Declare Function TevAXRF_StopModulation Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL) As Long
    Declare Function TevAXRF_ModulationTriggerArm Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal triggerSource As afSigGenDll_rmRoutingMatrix_t, ByVal Gate As Boolean, ByVal negativeEdge As Boolean) As Long
    Declare Function TevAXRF_SetMeasureSamples Lib "TevAXRF.dll" (ByVal samples As Long) As Long
    Declare Function TevAXRF_MeasureNoisePower Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal samples As Long, ByVal aveCount As Long) As Double
    Declare Function TevAXRF_ConnectNoiseSource Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal attenuation As Double) As Long
    Declare Function TevAXRF_DisconnectNoiseSource Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL) As Long
    Declare Function TevAXRF_SetNoiseSource Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByVal state As Boolean) As Long
    Declare Function TevAXRF_GetRawArray Lib "TevAXRF.dll" (ByVal channel As AXRF_CHANNEL, ByRef dataArray As Integer, ByVal numberOfPoints As Long) As Long

    Declare Function TevAXRF_CheckLocked Lib "TevAXRF.dll" (ByVal subSystem As Long) As Long


