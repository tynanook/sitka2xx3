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

    On Error GoTo ErrHandler
    
    #If Connected_to_AXRF Then
    
        Dim nSiteIndex As Long
        
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
        
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
            
                itl.Raw.AF.AXRF.MeasureSetup TxChannels(nSiteIndex), ExpectedPower, freq
                
            End If
            
        Next nSiteIndex
        
    #End If
    
    Exit Function
    
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function SetAXRFinRxMode(RxChannels() As AXRF_CHANNEL, power As Double, freq As Double)

    On Error GoTo ErrHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                itl.Raw.AF.AXRF.Source RxChannels(nSiteIndex), power, freq

            End If
        Next nSiteIndex
    #End If
    Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function MeasPowerAXRF(TxChannels() As AXRF_CHANNEL, MeasPower() As Double)
    On Error GoTo ErrHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long
        MeasPower(nSiteIndex) = -100
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                MeasPower(nSiteIndex) = itl.Raw.AF.AXRF.Measure(TxChannels(nSiteIndex))

            End If
        Next nSiteIndex
    #End If
    Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MeasDataAXRFandBasicDSP(TxChannels As AXRF_CHANNEL, MeasData() As Double, NumSamples As Integer, CapType As AXRF_ARRAY_TYPE, Optional PlotData As Boolean = False, Optional CalcMaxPower As Boolean = False, Optional plotName As String = "", Optional MaxPwr As Double, Optional FindIndexOfMaxPower As Boolean = False, Optional IndexofMaxPower As Double, Optional SumPowerBins As Boolean = False, Optional NumBinstoInclude As Long = 0, Optional SummedPower As Double) As Long
    On Error GoTo ErrHandler
    Dim nSiteIndex As Long
    Dim FFT As New DspWave
    Dim temp As Double
    Dim i As Double
    Dim MaxIndex As Double
    ReDim MeasData(NumSamples - 1)
    #If Connected_to_AXRF Then


        itl.Raw.AF.AXRF.MeasureArray TxChannels, MeasData, CapType
        If PlotData Or CalcMaxPower Or SumPowerBins Or FindIndexOfMaxPower Then
            FFT.data = MeasData

            If PlotData Then
                FFT.Plot plotName
            End If
            If CalcMaxPower Then
                FFT.CalcMinMax temp, MaxPwr, temp, temp
                TheExec.Datalog.WriteComment ("MaxPower -->>  " & MaxPwr)
            End If
            If FindIndexOfMaxPower Then
                FFT.CalcMinMax temp, temp, temp, IndexofMaxPower
                TheExec.Datalog.WriteComment ("MaxPowerIndex -->>  " & IndexofMaxPower)
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
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function calcMaxFromAXRFData(TxChannels() As AXRF_CHANNEL, MeasData() As Double, NumSamples As Integer, CapType As AXRF_ARRAY_TYPE)
    On Error GoTo ErrHandler
    #If Connected_to_AXRF Then
        Dim nSiteIndex As Long

        ReDim MeasData(TheExec.Sites.ExistingCount - 1, NumSamples)
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                itl.Raw.AF.AXRF.MeasureArray TxChannels(nSiteIndex), MeasData, CapType

            End If
        Next nSiteIndex
    #End If
    Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MeasDataAXRFandCalcMax(TxChannels As AXRF_CHANNEL, MeasData() As Double, NumSamples As Long, CapType As AXRF_ARRAY_TYPE, ByRef MaxPwr As Double, Optional PlotData As Boolean = False, Optional plotName As String = "", Optional FindIndexOfMaxPower As Boolean = False, Optional IndexofMaxPower As Double, Optional SumPowerBins As Boolean = False, Optional NumBinstoInclude As Long = 0, Optional SummedPower As Double) As Long
    
On Error GoTo ErrHandler
    
    'Dim nSiteIndex As Long
    Dim FFT As New DspWave
    Dim temp As Double
    Dim i As Double
    Dim MaxIndex As Double
    
    ReDim MeasData(NumSamples - 1)
    
    #If Connected_to_AXRF Then

        itl.Raw.AF.AXRF.MeasureArray TxChannels, MeasData, CapType
        FFT.data = MeasData
        FFT.CalcMinMax temp, MaxPwr, temp, temp

        If PlotData Then
            FFT.Plot plotName
            TheExec.Datalog.WriteComment ("MaxPower -->>  " & MaxPwr)
        End If
        
        If FindIndexOfMaxPower Then
            FFT.CalcMinMax temp, temp, temp, IndexofMaxPower
            TheExec.Datalog.WriteComment ("MaxPowerIndex -->>  " & IndexofMaxPower)
        End If
        
        If SumPowerBins Then 'Sum +/- 1 bin surrounding Fc
            FFT.CalcMinMax temp, temp, temp, MaxIndex
            SummedPower = 0
            For i = -1 * NumBinstoInclude To NumBinstoInclude
                SummedPower = SummedPower + MeasData(MaxIndex + i)
            Next i
        End If

    #End If
    
    Exit Function
    
ErrHandler:

    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function CaptureAndPlotAXRF(CapChan As AXRF_CHANNEL, power As Double, freq As Double)
   

    Dim data(1023) As Double
    
    On Error GoTo ErrHandler
    
    
    With itl.Raw.AF.AXRF
        .SetMeasureSamples 2048
        '.Source SrcChan, -50, 2450000000#
        .MeasureSetup CapChan, power, freq
        ' Power = .Measure(CapChan)
        'TheExec.DataLog.WriteComment ("Measured power from the LoopBack test--- " & Power)
        .MeasureArray CapChan, data, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN
''        power = PlotDouble(data)      ' Debug


    End With
    
     Exit Function
ErrHandler:
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
    
    TheHdw.Wait (FstWait)
    
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
    
    TheHdw.Wait (ScnWait)
'    TheHdw.Wait (0.01)
    
'    ' Restore original voltages
''    Call tl_DpsSetPrimaryVoltages(chans, OriginalVoltages)
'    TheHdw.DPS.chans(chans).PrimaryVoltages = OriginalVoltages
''     Call tl_DpsSetOutputSourceChannels(chans, TL_DPS_PrimaryVoltage)
'    TheHdw.DPS.chans(chans).OutputSource = dpsPrimaryVoltage
    
    TheHdw.Wait (0.0001)
    
End Function

Public Function LoadZbData() As Long
    itl.Raw.AF.AXRF.LoadModulationFile RFIN, TheBook.path + ".\Modulation\Zigbee_250KHz_100Symbols.aiq"
End Function

Public Function CreateZigbeeAnalysisObjects() As Long

    Dim Key As Variant

    Set Zigbee = CreateObject("Scripting.Dictionary")

    ' Add all of the names for WLAN Analysis here
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
    
    On Error GoTo ErrHandler
    
    
    With itl.Raw.AF.AXRF
        .SetMeasureSamples 2048
        '.Source SrcChan, -50, 2450000000#
        .MeasureSetup CapChan, power, freq
        ' Power = .Measure(CapChan)
        'TheExec.DataLog.WriteComment ("Measured power from the LoopBack test--- " & Power)
        .MeasureArray CapChan, data, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN
''        power = PlotDouble(data)      ' Debug

    End With
    
     Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function
