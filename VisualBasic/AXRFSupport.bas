Attribute VB_Name = "AXRFSupport"
Option Explicit
#Const Connected_to_J750 = True
#Const Connected_to_AXRF = True
#Const Debug_advanced = False


Public RFIN As AXRF_CHANNEL
Public RFOUT As AXRF_CHANNEL
Public RFOUT1 As AXRF_CHANNEL
Public RFOUT2 As AXRF_CHANNEL
Public RFOUT3 As AXRF_CHANNEL
Public RFOUT4 As AXRF_CHANNEL
Public RFOUT5 As AXRF_CHANNEL
Public RFOUT6 As AXRF_CHANNEL
Public RFOUT7 As AXRF_CHANNEL
Public RFOUT8 As AXRF_CHANNEL




'Public wlanNames As Object
Public Zigbee As Object

Public Function SetAXRFinTxMode(TxChannels() As AXRF_CHANNEL, ExpectedPower As Double, freq As Double)

    On Error GoTo errHandler
    
    #If Connected_to_AXRF Then
    
        Dim nSiteIndex As Long
        
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
        
            If TheExec.Sites.site(nSiteIndex).Active = True Then
            
                TevAXRF_MeasureSetup TxChannels(nSiteIndex), ExpectedPower, freq
                
            End If
            
        Next nSiteIndex
        
    #End If
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function SetAXRFinRxMode(RxChannels() As AXRF_CHANNEL, power As Double, freq As Double)

    On Error GoTo errHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.site(nSiteIndex).Active = True Then
                TevAXRF_Source RxChannels(nSiteIndex), power, freq

            End If
        Next nSiteIndex
    #End If
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MeasPowerAXRF(TxChannels() As AXRF_CHANNEL, MeasPower() As Double)
    On Error GoTo errHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long
        MeasPower(nSiteIndex) = -100
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.site(nSiteIndex).Active = True Then
                MeasPower(nSiteIndex) = TevAXRF_Measure(TxChannels(nSiteIndex))

            End If
        Next nSiteIndex
    #End If
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MeasDataAXRFandBasicDSP(TxChannels As AXRF_CHANNEL, MeasData() As Double, NumSamples As Integer, CapType As AXRF_ARRAY_TYPE, Optional PlotData As Boolean = False, Optional CalcMaxPower As Boolean = False, Optional plotName As String = "", Optional MaxPwr As Double, Optional FindIndexOfMaxPower As Boolean = False, Optional IndexofMaxPower As Double, Optional SumPowerBins As Boolean = False, Optional NumBinstoInclude As Long = 0, Optional SummedPower As Double) As Long
    On Error GoTo errHandler
    Dim nSiteIndex As Long
    Dim FFT As New DspWave
    Dim temp As Double
    Dim i As Double
    Dim MaxIndex As Double
    ReDim MeasData(NumSamples - 1)
    #If Connected_to_AXRF Then


        TevAXRF_MeasureArray TxChannels, MeasData(0), CapType
        If PlotData Or CalcMaxPower Or SumPowerBins Or FindIndexOfMaxPower Then
            FFT.data = MeasData

            If PlotData Then
                FFT.Plot plotName
            End If
            If CalcMaxPower Then
                FFT.CalcMinMax temp, MaxPwr, temp, temp
                TheExec.DataLog.WriteComment ("MaxPower -->>  " & MaxPwr)
            End If
            If FindIndexOfMaxPower Then
                FFT.CalcMinMax temp, temp, temp, IndexofMaxPower
                TheExec.DataLog.WriteComment ("MaxPowerIndex -->>  " & IndexofMaxPower)
            End If
            If SumPowerBins Then
                FFT.CalcMinMax temp, temp, temp, MaxIndex
                SummedPower = 0
                For i = -1 * NumBinstoInclude To NumBinstoInclude
                    SummedPower = SummedPower + MeasData(MaxIndex + i)
                Next i
            End If
        End If

    #End If
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function calcMaxFromAXRFData(TxChannels() As AXRF_CHANNEL, MeasData() As Double, NumSamples As Integer, CapType As AXRF_ARRAY_TYPE)
    On Error GoTo errHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long

        ReDim MeasData(TheExec.Sites.ExistingCount - 1, NumSamples)
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.site(nSiteIndex).Active = True Then
                TevAXRF_MeasureArray TxChannels(nSiteIndex), MeasData(0), CapType

            End If
        Next nSiteIndex
    #End If
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MeasDataAXRFandCalcMax(TxChannels As AXRF_CHANNEL, MeasData() As Double, NumSamples As Long, CapType As AXRF_ARRAY_TYPE, ByRef MaxPwr As Double, ByRef freqOffset As Double, Optional testXtalOffset As Boolean, Optional PlotData As Boolean = False, Optional plotName As String = "", Optional FindIndexOfMaxPower As Boolean = False, Optional IndexofMaxPower As Double, Optional SumPowerBins As Boolean = False, Optional NumBinstoInclude As Long = 0, Optional SummedPower As Double) As Long
    
On Error GoTo errHandler
    
    'Dim nSiteIndex As Long
    Dim FFT As New DspWave
    Dim temp As Double
    Dim i As Double
    Dim MaxIndex As Double
    
    ReDim MeasData(NumSamples - 1)
    
    #If Connected_to_AXRF Then

        TevAXRF_MeasureArray TxChannels, MeasData(0), CapType
        FFT.data = MeasData
        FFT.CalcMinMax temp, MaxPwr, temp, temp

        If PlotData Then
            FFT.Plot plotName
            TheExec.DataLog.WriteComment ("MaxPower -->>  " & MaxPwr)
        End If
        
        If FindIndexOfMaxPower Then
            FFT.CalcMinMax temp, temp, temp, IndexofMaxPower
            TheExec.DataLog.WriteComment ("MaxPowerIndex -->>  " & IndexofMaxPower)
        End If
        
        If SumPowerBins Then 'Sum +/- 1 bin surrounding Fc
            FFT.CalcMinMax temp, temp, temp, MaxIndex
            SummedPower = 0
            For i = -1 * NumBinstoInclude To NumBinstoInclude
                SummedPower = SummedPower + MeasData(MaxIndex + i)
            Next i
        End If
        
        If testXtalOffset Then
            freqOffset = 0#
            Dim capd As New DspWave, capax() As Double, ht() As Variant, cap As New DspWave
            Dim fs As Double, fres As Double, ifFrq As Double, ss As Long, hsize As Long
            Dim o As Double, phzo As Double, capr() As Double, caph() As Double, capi() As Double
            Dim phz() As Double, dp() As Double, pramp() As Double, j As Long
            ReDim capax(NumSamples * 2)
            TevAXRF_MeasureArray TxChannels, capax(0), AXRF_TIME_DOMAIN
            capd.data = capax
            hsize = 35
            ss = NumSamples
            fs = 250# * 1000000#
            fres = fs / ss
            ifFrq = fres * ss / 4#
            ht = Array(0.957143, 0.650478, -0.042857, 0.225208, -0.042857, 0.139465, _
            -0.042857, 0.10222, -0.042857, 0.081132, -0.042857, 0.06738, -0.042857, _
             0.05757, -0.042857, 0.050113, -0.042857, 0.044169, -0.042857, 0.039248, _
            -0.042857, 0.035044, -0.042857, 0.031356, -0.042857, 0.028045, -0.042857, _
             0.025009, -0.042857, 0.022171, -0.042857, 0.019471, -0.042857, 0.016857, _
            -0.042857, 0.014286, -0.042857, 0.011714, -0.042857, 0.009101, -0.042857, _
             0.006401, -0.042857, 0.003563, -0.042857, 0.000526, -0.042857, -0.002785, _
            -0.042857, -0.006473, -0.042857, -0.010676, -0.042857, -0.015598, -0.042857, _
            -0.021542, -0.042857, -0.028998, -0.042857, -0.038809, -0.042857, -0.05256, _
            -0.042857, -0.073648, -0.042857, -0.110894, -0.042857, -0.196637, -0.042857, _
            -0.621907)
            Set cap = capd.Select(0, 1, ss).Copy
            capr = cap.data
            ReDim caph(ss + 2 * hsize - 1)
            capr = cap.data
            For i = 0 To ss - 1
                For j = 0 To 2 * hsize - 1
                    caph(i + j) = caph(i + j) + capr(i) * ht(j)
                Next
            Next
            ReDim capi(ss - 1)
            For i = 0 To ss - 1
                capi(i) = caph(i + hsize)
            Next
            ReDim phz(ss)
            For i = 0 To ss - 1
                If capr(i) = 0# Then
                    capr(i) = 0.000000000001
                End If
                phz(i) = Atn(capi(i) / capr(i))
                If capr(i) < 0# Then
                    If capi(i) < 0# Then
                        phz(i) = phz(i) - 4# * Atn(1#)
                    ElseIf capi(i) > 0# Then
                        phz(i) = phz(i) + 4# * Atn(1#)
                    End If
                End If
            Next i
            phzo = phz(0)
            For i = 0 To ss - 2
                phz(i) = phz(i + 1) - phz(i)
                If phz(i) > 2# * Atn(1#) Then
                    phz(i) = phz(i) - 8# * Atn(1#)
                End If
                If phz(i) < -2# * Atn(1#) Then
                    phz(i) = phz(i) + 8# * Atn(1#)
                End If
            Next i
            ReDim pramp(ss - 1)
            pramp(0) = phzo
            For i = 1 To ss - 1
                pramp(i) = pramp(i - 1) + phz(i - 1)
            Next i
            ReDim dp(ss - 1 - 4 * hsize)
            For i = 0 To UBound(dp)
                dp(i) = (pramp(i + 1 + hsize) - pramp(i + hsize))
            Next i
            o = 0#
            For i = 0 To UBound(dp)
                o = o + dp(i)
            Next i
            freqOffset = (o * fs / (8# * Atn(1#) * (UBound(dp) + 1))) - ifFrq
        Else
            freqOffset = -99999
        End If

    #End If
    
    Exit Function
    
errHandler:

    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function CaptureAndPlotAXRF(CapChan As AXRF_CHANNEL, power As Double, freq As Double)
   

    Dim data(1023) As Double
    
    On Error GoTo errHandler
    
   
        TevAXRF_SetMeasureSamples 2048
        'TevAXRF_Source SrcChan, -50, 2450000000#
        TevAXRF_MeasureSetup CapChan, power, freq
        ' Power = TevAXRF_Measure(CapChan)
        'TheExecTevAXRF_DataLogTevAXRF_WriteComment ("Measured power from the LoopBack test--- " & Power)
        
        Stop
        'ty commented out next line 20180208
       ' TevAXRF_MeasureArray CapChan, data(0)(0), AXRF_FREQ_DOMAIN
''        power = PlotDouble(data)      ' Debug



    
     Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

Public Function cycle_power(InitVolt As Double, FinalVolt As Double, Optional FirstWait As Double = -1, Optional SecondWait As Double = -1)
    
    Dim PowerPinList As String
    Dim OriginalVoltages As Variant
    Dim NewVoltages As Variant
    Dim VdisParam As Double
    Dim chans() As Long
    Dim nchannels As Long
    Dim nsites As Long
    Dim errstr As String
    Dim i As Integer
    Dim FstWait As Double
    Dim ScnWait As Double


''    ' Were we called properly?
''    If argc < 3 Then
''        MsgBox "Error-ToggleVdd expected at least 3 arguments"
''        Exit Function
''    End If
    
    ' Recreate pinlist into single comma delimited string (needed for next step)
    PowerPinList = "VBAT"           'argv(2)
''    For i = 3 To argc - 1
''        PowerPinList = "," + PowerPinList + argv(i)
''    Next i
    
    ' Get list of channels
    
'    If tl_dm_GetChannelListForselectedSites(PowerPinList, TL_DPS_CHANNELTYPE, _
                chans, nchannels, nsites, errstr) <> TL_SUCCESS Then
'    End If
    Call TheExec.DataManager.GetChanListForSelectedSites(PowerPinList, chDPS, chans, nchannels, nsites, errstr)
    If errstr <> "" Then
        MsgBox "Error-ToggleVdd error return from tl_dm_GetChannelListForEnabledSites:" + _
                Chr$(13) + errstr
        Exit Function
    End If
    
    If SecondWait = -1 Then
        ScnWait = 0.001
    Else
        ScnWait = SecondWait
    End If
    
    If FirstWait = -1 Then
        FstWait = 0.001
    Else
        FstWait = FirstWait
    End If
    
''    ' Get original voltages
'''    Call tl_DpsGetPrimaryVoltages(chans, OriginalVoltages)
''    OriginalVoltages = InitVolt
    
    ' Set Init voltages
    If InitVolt > 0.001 Then
        VdisParam = InitVolt            'To Init Volt
    Else
        VdisParam = 0.001              'Do not set to Zero avoid Gound Bounce
    End If
    
    ReDim NewVoltages(nchannels)
    ReDim NewVoltages(nchannels - 1)
    For i = 0 To nchannels - 1
        'NewVoltages(i) = CDbl(val(argv(0)))
        NewVoltages(i) = VdisParam
    Next i
'    Call tl_DpsSetPrimaryVoltages(chans, NewVoltages)
    TheHdw.DPS.chans(chans).PrimaryVoltages = NewVoltages
'    Call tl_DpsSetOutputSourceChannels(chans, TL_DPS_PrimaryVoltage)
    TheHdw.DPS.chans(chans).OutputSource = dpsPrimaryVoltage
    
    TheHdw.wait (FstWait)
    
    ' Set Final voltages
    If FinalVolt > 0.001 Then
        VdisParam = FinalVolt           'To Final Volt
    Else
        VdisParam = 0.001              'Do not set to Zero avoid Gound Bounce
    End If
    
    ReDim NewVoltages(nchannels)
    ReDim NewVoltages(nchannels - 1)
    For i = 0 To nchannels - 1
        'NewVoltages(i) = CDbl(val(argv(0)))
        NewVoltages(i) = VdisParam
    Next i
'    Call tl_DpsSetPrimaryVoltages(chans, NewVoltages)
    TheHdw.DPS.chans(chans).PrimaryVoltages = NewVoltages
'    Call tl_DpsSetOutputSourceChannels(chans, TL_DPS_PrimaryVoltage)
    TheHdw.DPS.chans(chans).OutputSource = dpsPrimaryVoltage
    
    
    ' Wait
    'Call tl_wait(CDbl(val(argv(1))))
    'Call tl_wait(ResolveArgv(argv(1)))
'     Call TheHdw.Wait(ResolveArgv(argv(1)))
    
    TheHdw.wait (ScnWait)
'    TheHdw.Wait (0.01)
    
'    ' Restore original voltages
''    Call tl_DpsSetPrimaryVoltages(chans, OriginalVoltages)
'    TheHdw.DPS.chans(chans).PrimaryVoltages = OriginalVoltages
''     Call tl_DpsSetOutputSourceChannels(chans, TL_DPS_PrimaryVoltage)
'    TheHdw.DPS.chans(chans).OutputSource = dpsPrimaryVoltage
    
    TheHdw.wait (0.0001)
    
End Function

Public Function LoadZbData() As Long
    TevAXRF_LoadModulationFile RFIN, TheBook.path + ".\Modulation\Zigbee_250KHz_100Symbols.aiq"
End Function

Public Function CreateZigbeeAnalysisObjects() As Long

    Dim Key As Variant

    Set Zigbee = CreateObject("Scripting.Dictionary")

'    ' Add all of the names for WLAN Analysis here
    Zigbee.Add "ZigbeeBasic1", ""

    For Each Key In Zigbee.Keys
        itl.RF.AF.Generic.Analysis.Create Key
    Next Key

End Function
Public Function RemoveZigbeeAnalysisObjects() As Long

    Dim Key As Variant

    For Each Key In Zigbee.Keys
        itl.RF.AF.Generic.Analysis.Remove Key
    Next Key

End Function

Public Function Freq_Estimate_AXRF(CapChan As AXRF_CHANNEL, power As Double, freq As Double)
   

    Dim data(1023) As Double
    
    On Error GoTo errHandler
    
    Stop

        TevAXRF_SetMeasureSamples 2048
        'TevAXRF_Source SrcChan, -50, 2450000000#
        TevAXRF_MeasureSetup CapChan, power, freq
        ' Power = TevAXRF_Measure(CapChan)
        'TheExecTevAXRF_DataLogTevAXRF_WriteComment ("Measured power from the LoopBack test--- " & Power)
      '  TevAXRF_MeasureArray CapChan, data(0)(0), AXRF_FREQ_DOMAIN
        TevAXRF_MeasureArray CapChan, data(0), AXRF_FREQ_DOMAIN
''        power = PlotDouble(data)      ' Debug


    
     Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

Public Function check_AXRF_error_flag(argc As Long, argv() As String) As Long
    Dim TestStat As Long, testNum As Long, TestParm As Long
    Dim nSiteIndex As Long
    Dim axrffailflag As Double
 
    Site_Stat = TheExec.Sites.SelectFirst()
     AXRF_Error_Flag = False
       Do  'For each site loop
            nSiteIndex = TheExec.Sites.SelectedSite()
            If TheExec.Sites.site(nSiteIndex).Active = True Then
            
                If (AXRF_Error_Flag = True) Then
                    TestStat = logTestFail
                    TestParm = parmPass
                    axrffailflag = Initialize_status
                    TheExec.Sites.site(nSiteIndex).TestResult = siteFail
                Else
                    TestStat = logTestPass
                    TestParm = parmLow
                    axrffailflag = 0
                    TheExec.Sites.site(nSiteIndex).TestResult = sitePass
                End If
            End If
                
        Call TheExec.DataLog.WriteParametricResult(nSiteIndex, 101030, TestStat, TestParm, "NA", -1, 0, axrffailflag, 0, unitCustom, 0, unitNone, 0, "check_AXRF_Init", "")
        Loop While (TheExec.Sites.SelectNext(Site_Stat) <> loopDone)

Exit Function

End Function
