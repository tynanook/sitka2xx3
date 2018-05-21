Attribute VB_Name = "VBT_AXRFBasic"
Option Explicit
'Public Function FindMyPLL_20M()
'Dim data(1023) As Double
'Dim x As Integer
'RFOUT = AXRF_CH2
'    With itl.Pins("clock").NI.RFSG
'
'        .ConfigureRF 19000000#, -50
'        .ConfigureOutputEnabled True
'        .ConfigureGenerationMode NIRFSG_CONSTS.Cw
'
'        .ConfigureGenerationMode NIRFSG_CONSTS.Cw
'        .Commit
'        .Initiate
'        .Abort
'    End With
'
'    Dim i As Integer
'    Dim power As Double
'    Dim GoodSignal(1) As Double
'    For i = 0 To 4000
'    itl.Pins("clock").NI.RFSG.ConfigureRF 18000000# + i * 1000#, 3
'    TheHdw.Wait 0.01
'    Call itl.Raw.AF.AXRF.MeasureSetup(RFOUT, -2, 20000000#)
'    power = itl.Raw.AF.AXRF.Measure(RFOUT)
'    Call itl.Raw.AF.AXRF.MeasureArray(AXRF_CH2, data, AXRF_FREQ_DOMAIN)
'    For x = 0 To 1023
'
'    If data(x) > -20 Then
'    GoodSignal(0) = power
'    GoodSignal(1) = 19000000# + i * 10000#
'    Call PlotDouble(data)
'    End If
'    Next x
'
'
'    TheExec.Datalog.WriteComment (19000000# + i * 10000# & "        " & power)
'    Next i
'
'    TheExec.Datalog.WriteComment (GoodSignal(1) & "       " & GoodSignal(0))
'
'End Function
'
''Public Function axrfWlanBasicTest() As Long
'
'    Dim SampleRate As Double
'    Dim NumSamples As Long
'    Dim freq As Double
'    Dim burstpower As Single
'
'    Dim meaureFactor As Double
'    Dim sourceFactor As Double
'    Dim sourceLevel As Double
'    Dim measureLevel As Double
'    Dim status As Long
'    Dim iData() As Single
'    Dim qData() As Single
'
'    Dim evmResult As Double
'    Dim avgPowerResult As Double
''    Dim WlanName As String
'    Dim attn_dB As Double
'
'
'    Dim Result As Double
'    Dim levelcal As Double
'    Dim Arr_data(1023) As Double
'
'
'    RFIN = AXRF_CH1
'    RFOUT = AXRF_CH2
'
'    SampleRate = 54000000# * 3
'    NumSamples = 93750
'    freq = 2400000000# ' 2412e6;
'    sourceLevel = -7.5 ' 'Pin start
'
'    '??????? sourceLevel = targetPout - dut_gain; //Pin start
'    measureLevel = -7 '+dut_gain-attn_dB
'
''    WlanName = "WlanBasic1"
'
''    ConfigureAnalysisWLAN WlanName, SampleRate, 0
'
'    With itl.Raw.AF.AXRF
'
'        .Source RFIN, sourceLevel, freq
'        .MeasureSetup RFOUT, measureLevel, freq
'
'        TheHdw.Wait 0.01
'
'        Result = .Measure(RFOUT)
'        .MeasureArray RFOUT, Arr_data, AXRF_FREQ_DOMAIN
'
'        status = .StartModulation(RFIN, TheBook.Path + "\Modulation\11apn9.aiq")
'
'        status = .SetIQSampleFrequency(AXRF_CH1, SampleRate)  'changed for new AXRF
'
'        TheHdw.Wait 0.01
'
'        .GetMeasureFactor RFOUT, levelcal
'
''        itl.RF.AF.WLAN.Analysis(WlanName).Configuration.RfLevelCal = levelcal + attn_dB ' //attn_dB = 20dB (external)
'
'        'Set up DUT
''        CaptureWLAN RFOUT, NumSamples, iData, qData
'
''        AnalyseWLAN WlanName, NumSamples, iData, qData, burstpower, evmResult, avgPowerResult
'
'        .StopModulation RFIN
'
'    End With
'
'    Dim iDSPWave As DspWave
'    Dim qDSPWave As DspWave
'
'    Set iDSPWave = GenJ750DSPWaveFromSingle(iData)
'    Set qDSPWave = GenJ750DSPWaveFromSingle(qData)
'
'    iDSPWave.Plot ("IDATA")
'    qDSPWave.Plot ("QDATA")
'
'End Function
'
''Public Sub ConfigureAnalysisWLAN(Name As String, SampleRate As Double, ByRef minsamples As Integer)
'
''    With itl.RF.AF.WLAN.Analysis(Name).Configuration
''        .SpectrumAnalysisMode = WlanSpectrumAnalysisMode_SpectrumAnalysisGated
''        .AnalysisMode = WlanAnalysisMode_AnalysisMode11a
'        .NumSymbolsToAnalyseAutoDetect = True
'        .SamplingFreq = SampleRate
'
'        'OFDM System Parameters
'        .PilotTrackingAmplitude = False
'        .PilotTrackingTiming = False
'        .PilotTrackingPhase = True
''        .SpectrumAnalysisMode = WlanSpectrumAnalysisMode_SpectrumAnalysisNonGated
''        .BurstProfileMode = WlanBurstProfileMode_PeakPower
'    End With
'
'End Sub

''Public Function CaptureWLAN(chan As AXRF_CHANNEL, NumOfSamples As Long, ByRef iData() As Single, ByRef qData() As Single) As Long
'
'    Dim status As Integer
'
'    'Allocate memory for IQ data
'    ReDim iData(NumOfSamples - 1)
'    ReDim qData(NumOfSamples - 1)
'
'    status = itl.Raw.AF.AXRF.MeasureArrayIQ(chan, NumOfSamples, iData, qData)
'
'End Function

''Public Function AnalyseWLAN(Name As String, NumSamples As Long, ByRef iData() As Single, ByRef qData() As Single, ByRef burstpower As Single, ByRef EVM As Double, ByRef averPower As Double) As Long
'
'    burstpower = -130#
'    Dim freqError As Single
'    Dim evmError As Single
'    Dim errorcode As Integer
'    Dim length As Integer
'    Dim Result As Double
'    Dim passfail As Integer
'    Dim bPassFail As Boolean
''    Dim nMeasurement As WlanMeasurement
'
'    freqError = 0
'    evmError = 0#
'
'    passfail = 0
'    errorcode = 0
'    length = 0
'
'    bPassFail = True ' 1=pass, 0=fail for datalogging
'
'    'Analyse
'
''    nMeasurement = WlanMeasurement_LocateBurst + WlanMeasurement_ModAccuracy + WlanMeasurement_BurstPower
'
''    With itl.RF.AF.WLAN.Analysis(Name)
'         .Analyse nMeasurement, iData, qData
'
'        '*-----------------*
'        '*-- Get results --*
'        '*-----------------*
'        '*-- Average Power --*
'
'        burstpower = .Results.BurstProfileAveragePower
'
'        '*-- Modulation Accuracy --*
'
'        evmError = .Results.EvmRms
'
'    End With
'
'    EVM = evmError
'    averPower = burstpower
'
'End Function



''        public void SetupDigitizerWLAN(double SAMPLERATE)
'        {
'            long status = 0;
'
'
'            //cerr << "Digitizer Hardware Setup" << endl;
'
'            status = afdig.Modulation.Mode_Set(afDigitizerDll_32Wrapper.afDigitizerDll_mmModulationMode_t.afDigitizerDll_mmGeneric);
'            status = afdig.Modulation.GenericDecimationRatio_Set(1);
''            //status += TevWRxDriver_GetDigitizerHandle (0, &digitizerWLAN);
'    //change generic re-sampling rate
'
'            status = afdig.Modulation.GenericSamplingFrequency_Set(SAMPLERATE);
'            status = afdig.Capture.IQ.Resolution_Set(afDigitizerDll_32Wrapper.afDigitizerDll_iqrIQResolution_t.afDigitizerDll_iqrAuto);
'            double f = 0.0;
'            status = afdig.Modulation.GenericSamplingFrequency_Get(ref f);
'        }










Public Function axrfZigbeeBasicTest() As Long

    Dim SampleRate As Double
    Dim NumSamples As Long
    Dim freq As Double
    Dim burstpower As Single
    
    Dim meaureFactor As Double
    Dim sourceFactor As Double
    Dim sourceLevel As Double
    Dim measureLevel As Double
    Dim status As Long
    Dim iData() As Single
    Dim qData() As Single
    
    Dim evmResult As Double
    Dim avgPowerResult As Double
    Dim ZigbeeName As String
    Dim attn_dB As Double
    
    Dim Arr_data(1023) As Double
    Dim Result As Double
    Dim levelcal As Double
    
'''    RFIN = AXRF_CH1
'''    RFOUT = AXRF_CH2
'''
'''    SampleRate = 100000000
'''    NumSamples = 320000
'''    freq = 2450000000# ' 2412e6;
'''    sourceLevel = -7.5 ' 'Pin start
'''
'''    '??????? sourceLevel = targetPout - dut_gain; //Pin start
'''    measureLevel = -6 '+dut_gain-attn_dB
'''
'''    ZigbeeName = "ZigbeeBasic1"
'''
'''    ConfigureAnalysisZigbee ZigbeeName, SampleRate, 0                'Removed For 34075 (Debug)
'''
'''    With itl.Raw.AF.AXRF
'''
'''        .Source RFIN, sourceLevel, freq
'''        .MeasureSetup RFOUT, measureLevel, freq
'''
'''        TheHdw.Wait 0.01
'''
'''        Result = .Measure(RFOUT)
'''        .MeasureArray RFOUT, Arr_data, AXRF_FREQ_DOMAIN
'''
'''        Call PlotDouble(Arr_data)
'''
'''        status = .StartModulation(RFIN, TheBook.Path + "\Modulation\Zigbee_250KHz_100Symbols.aiq")
'''
'''        status = .SetIQSampleFrequency(AXRF_CH1, SampleRate) 'changed for new AXRF
'''
'''
'''        TheHdw.Wait 0.01
'''
'''        .GetMeasureFactor RFOUT, levelcal
'''
''''        'itl.RF.AF.WLAN.Analysis(WlanName).Configuration.RfLevelCal = levelcal + attn_dB ' //attn_dB = 20dB (external)
'''        itl.RF.AF.Generic.Analysis(ZigbeeName).Configuration.RfLevelCal = 0 'levelcal
'''        'Set up DUT
'''        CaptureZigbee RFOUT, NumSamples, iData, qData
'''
'''        'AnalyseZigbee ZigbeeName, NumSamples, iData, qData, burstpower, evmResult, avgPowerResult 'type mismatch error -Nick
'''
'''        .StopModulation RFIN
'''
'''    End With
'''
'''    Dim iDSPWave As DspWave
'''    Dim qDSPWave As DspWave
'''
'''    Set iDSPWave = GenJ750DSPWaveFromSingle(iData)
'''    Set qDSPWave = GenJ750DSPWaveFromSingle(qData)
'''
'''    iDSPWave.Plot ("IDATA")
'''    qDSPWave.Plot ("QDATA")
 
     
End Function




