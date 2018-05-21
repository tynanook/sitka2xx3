Attribute VB_Name = "VBT_RF"
Option Explicit

#Const Connected_to_J750 = True
#Const Connected_to_AXRF = True
#Const Debug_advanced = True

Public FIRSTLOAD As Boolean


Global Const DUTRefClkFreq As Double = 13560000# 'NOT USED FOR LoRa MODULES!

Public tx_path_db(7) As Double      'Dim for number of sites
Public coax_cable_db(7) As Double   'Dim for number of sites
Public SerialHramData As New Hram3kDataRdSer        'switch to one bit mode...



Public Function read_cal_factors() As Long              '(argc As Long, argv() As String) As Long

    'Public tx_path_db() As Double
    'Public rx_path_db() As Double
    'Public coax_cable_db() As Double
    
    'This function reads the RF_Cal_Factors worksheet to initialize global RF scalar calibration offsets. AXRF calibration is
    'performed with same cables and RF interface as the J750_AERO with Reid-Asman junction box interface and Pasternak coaxial cables.
    
    Dim nSiteIndex As Long
    
    On Error GoTo errHandler
    
     
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
    
        If TheExec.Sites.site(nSiteIndex).Active = True Then
        
            tx_path_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(3, 2 + nSiteIndex).value       'DIB TX path loss
                           
            coax_cable_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(4, 2 + nSiteIndex).value    'Cable loss from PXI to DIB
        
            'Debug.Print "TX_Path_dB = "; tx_path_db(nSiteIndex)
            'Debug.Print "Coax_Cable_dB = "; coax_cable_db(nSiteIndex)
        
        End If
        
    Next nSiteIndex
    
    Exit Function
    
errHandler:

    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into prouduction abort routine
    read_cal_factors = TL_ERROR

End Function

'This function will test a value against limits and set the appropriate Pass/Fail status based on the values.
'you must call the function from within a site loop.  Required inputs are Site, Value2Dlog, lo and hi limits.
'other inputs are optional.  Hint, you can use the pinname field as a unique identifier.
Public Sub sm_LogPassFail(site As Long, Value2Dlog As SiteDouble, LoLimit As Double, HiLimit As Double, _
    Optional PinName As String = "", Optional MeasUnit As UnitType = unitNone, _
    Optional ForceValue As Double = 0, Optional TestName As String = "", Optional fmtStr As String = "", _
    Optional measUnitStr As String = "", Optional scaleType_ As ScaleType = scaleNone, Optional testNum As Long = -1 _
    )

    If testNum = -1 Then
        testNum = TheExec.Sites.site(site).testnumber
    End If
    
    Select Case Value2Dlog(site)
        Case Is < LoLimit, Is > HiLimit
            Call TheExec.DataLog.WriteParametricResult(site, testNum, logTestFail, _
                 0, PinName, -1, LoLimit, Value2Dlog(site), HiLimit, MeasUnit, _
                ForceValue, unitNone, 0, TestName, measUnitStr, , scaleType_, fmtStr)
            TheExec.Sites.site(site).TestResult = siteFail
        Case Else
            Call TheExec.DataLog.WriteParametricResult(site, testNum, _
                logTestPass, 0, PinName, -1, LoLimit, Value2Dlog(site), HiLimit, MeasUnit, _
                ForceValue, unitNone, 0, TestName, measUnitStr, , scaleType_, fmtStr)
            TheExec.Sites.site(site).IncrementTestNumber                                'increment tnum
    End Select
    
End Sub

Public Function rn2483_tx868_cw(argc As Long, argv() As String) As Long

    Dim site As Variant
    Dim OOKDepthMode As Boolean
    Dim ExistingSiteCnt As Integer
    
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    Dim nSiteIndex As Long
    
    Dim MeasPower(1) As Double
    Dim MeasData() As Double

    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double          'Site Array
    Dim SumPower() As Double                'Site Array
    
    Dim MaxPowerTemp As Double
    Dim SumPowerTemp As Double
    Dim TestFreq As Double
    Dim oprVolt As Double
    
    Dim IndexMaxPower As New SiteDouble
    
    Dim MeasChans() As AXRF_CHANNEL         'Site Array

    Dim TxPower868 As New PinListData
    Dim I_TX868_CW As New PinListData

    Dim testXtalOffset As Boolean
    Dim FreqOffset_Hz As New SiteDouble
    Dim FreqOffset_Hz_temp As Double
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    ReDim SumPower(0 To ExistingSiteCnt - 1)
    
    On Error GoTo errHandler
    
    rn2483_tx868_cw = TL_SUCCESS
    
    'AXRF Channel assignments
    
    Dim LoLimit_I As Double
    Dim LoLimit_Tx As Double
    Dim HiLimit_I As Double
    Dim HiLimit_Tx As Double
    '--------Argument processing--------'
    LoLimit_I = argv(2)
    HiLimit_I = argv(3)
    LoLimit_Tx = argv(4)
    HiLimit_Tx = argv(5)
    '------- end of argument process -------'
    
    Select Case ExistingSiteCnt
        
    Case Is = 1
        MeasChans(0) = AXRF_CH1
        
    Case Is = 2
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        
    Case Is = 3
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        MeasChans(2) = AXRF_CH5
        
    Case Is = 4
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        MeasChans(2) = AXRF_CH5
        MeasChans(3) = AXRF_CH7
        
        
    Case Else
        MsgBox "Error in [rn2483_tx868_cw]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo errHandler
        
    End Select
    

    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    

    If argc < 2 Then
        MsgBox "Error - On rn2483_tx868_cw - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        TestFreq = argv(0)
        oprVolt = argv(1)
        
    End If
    
    testXtalOffset = True                   'set true if you want to test xtal offset

    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB
    
    Call TevAXRF_SetMeasureSamples(8192) 'Fres = 30.5176 kHz  (Fs = 250MHz, N=8192) NOTE: Fres chosen to bound TX freq
  
    TheHdw.wait 0.05
        
    Select Case TestFreq
    Case 868300000
        TxPower868.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        

    
    Case Else
        TxPower868.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        
        
    End Select


    TheHdw.wait (0.002)
    
        
        'Run pattern to put DUT into TX CW Mode @ 868 MHz, +15 dBm Power
        
    
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx868_cw").start ("start_tx_cw_on")
           
        
'Pattern loops after setting cpuA. VBT loops waiting for the pattern to set cpuA.

Flag_Loop:  'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
   
    Flags = TheHdw.Digital.Patgen.CpuFlags

        'Debug.Print "Flags ="; Flags
        
        If (Flags = 1) Then GoTo End_flag_loop 'cpuA set - DUT in TX CW mode
 
        If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set - DUT not in TX CW mode
  
End_flag_loop:
        
        'Measure TX Current
        
        ' DPS setup is inherently multi-site.
        
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.samples = 1
    End With
    
    
        'Measure Current
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX868_CW)
    

        'Measure RF Power
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1 'Site loop needed for AXRF

            If TheExec.Sites.site(nSiteIndex).Active = True Then
                
                TevAXRF_MeasureSetup MeasChans(nSiteIndex), 17, TestFreq  'set for +17dBm
                
                TheHdw.wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 4096, AXRF_FREQ_DOMAIN, MaxPowerTemp, FreqOffset_Hz_temp, testXtalOffset, False, "rf", False, False, False, 1, SumPowerTemp)  'True plots waveform
                
                FreqOffset_Hz(nSiteIndex) = FreqOffset_Hz_temp
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                SumPower(nSiteIndex) = SumPowerTemp
            
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                
                Select Case TestFreq
                Case 868300000
                    
                    TxPower868.pins("RFHOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                    TxPower868.pins("RFOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select


            End If
            
        Next nSiteIndex
        
        'Reset cpuA flag
        FlagsSet = 0
        FlagsClear = cpuA

        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) 'Pattern continues after cpuA reset.
        
        
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.
  
    'Run pattern to stop DUT transmitting
    
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx_cw_off").start ("start_tx_cw_off") 'all sites

        'TheHdw.Wait (0.1) 'avoids LVM Priming patgen RTE
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.

 'TheExec.DataLog.WriteComment ("=================== TX868_CW_PWR =====================")
 
    Select Case TestFreq
    
    Case 868300000
    
        TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If

'                Select Case TheExec.CurrentJob
'                Case "f1-prd-std-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower868, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
'
'                Case "f1-pgm-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, 0.035, 0.085, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
'
'                Case "q1-prd-std-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW_qc", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower868, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868_qc", , , , , , , , tlForceNone
'
'                Case Else
'
'                End Select

        
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.03, 0.09, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 17, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If
        
    End Select
 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

errHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower868.pins("RFHOUT").value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    
    Case 868300000
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If
  
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If

    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2483_tx868_cw = TL_ERROR
    
End Function







Public Function rn2483_tx433_cw(argc As Long, argv() As String) As Long

    'ty created 2018-01-09:
            ' assumption: TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx_cw_off").Load used in the tx868_cw function is also applicaple for the 433.
          ' assumption: using the 868 pattern is equivalent to 868 other than the freq setting; I just used the 868 pattern, and changed the freq from 868 to 433..
            ' asumption: Don't need to measure freq offet as the test is redunant and would only add unnecessary test time
            ' expanded limits in flow worksheet by +/- 2 dBm of the 868 version; used the 868 limits as the measured power was within a dB of its/that test's result.
                ' cleaned up the code a tad, but not much (from the 868 version it was copied from).
                ' i didn't test this function, but it should be able to be used in the both the 868 and 433 cw test, to keep this program DRY.
    
    Dim site As Variant
    Dim OOKDepthMode As Boolean
    Dim ExistingSiteCnt As Integer
    
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    Dim nSiteIndex As Long
    
    Dim MeasPower(1) As Double
    Dim MeasData() As Double

    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double          'Site Array
    Dim SumPower() As Double                'Site Array
    
    Dim MaxPowerTemp As Double
    Dim SumPowerTemp As Double
    Dim TestFreq As Double
    Dim oprVolt As Double
    
    Dim IndexMaxPower As New SiteDouble
    
    Dim MeasChans() As AXRF_CHANNEL         'Site Array

    Dim TxPower868 As New PinListData
    Dim TxPower433 As New PinListData
    Dim I_TX868_CW As New PinListData
    Dim I_TX433_CW As New PinListData

    Dim testXtalOffset As Boolean
    Dim FreqOffset_Hz As New SiteDouble
    Dim FreqOffset_Hz_temp As Double
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    ReDim SumPower(0 To ExistingSiteCnt - 1)
    
    On Error GoTo errHandler
    
    rn2483_tx433_cw = TL_SUCCESS
    
    'AXRF Channel assignments
    
    Dim LoLimit_I As Double
    Dim LoLimit_Tx As Double
    Dim HiLimit_I As Double
    Dim HiLimit_Tx As Double
    '--------Argument processing--------'
    LoLimit_I = argv(2)
    HiLimit_I = argv(3)
    LoLimit_Tx = argv(4)
    HiLimit_Tx = argv(5)
    
    If argc < 2 Then
        MsgBox "Error - On rn2483_tx433_cw - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        TestFreq = argv(0)
        oprVolt = argv(1)
    End If
    
    '------- end of argument process -------'
    
    Select Case ExistingSiteCnt
        
        Case Is = 1
            If TestFreq = 868300000 Then
                MeasChans(0) = AXRF_CH1
            ElseIf TestFreq = 433300000 Then
                MeasChans(0) = AXRF_CH2
            Else ' assumes hf path
                MeasChans(0) = AXRF_CH1
            End If
        Case Is = 2
            If TestFreq = 868300000 Then
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
            ElseIf TestFreq = 433300000 Then
                MeasChans(0) = AXRF_CH2
                MeasChans(1) = AXRF_CH4
            Else ' assumes hf path
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
            End If
        Case Is = 3
            If TestFreq = 868300000 Then
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
                MeasChans(2) = AXRF_CH5
            ElseIf TestFreq = 433300000 Then
                MeasChans(0) = AXRF_CH2
                MeasChans(1) = AXRF_CH4
                MeasChans(2) = AXRF_CH6
            Else ' assumes hf path
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
                MeasChans(2) = AXRF_CH5
            End If
        Case Is = 4
            If TestFreq = 868300000 Then
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
                MeasChans(2) = AXRF_CH5
                MeasChans(3) = AXRF_CH7
            ElseIf TestFreq = 433300000 Then
                MeasChans(0) = AXRF_CH2
                MeasChans(1) = AXRF_CH4
                MeasChans(2) = AXRF_CH6
                MeasChans(3) = AXRF_CH8
            Else ' assumes hf path
                MeasChans(0) = AXRF_CH1
                MeasChans(1) = AXRF_CH3
                MeasChans(2) = AXRF_CH5
                MeasChans(3) = AXRF_CH7
            End If
        Case Else
            MsgBox "Error in [rn2483_tx433_cw]" & vbCrLf & _
                   "Existnumber is not support by ITL", _
                   vbCritical + vbOKOnly, _
                   "Interpose Setup Error"
            GoTo errHandler
        
    End Select
    

    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
    testXtalOffset = False  'only due for 868 to save TT                 'set true if you want to test xtal offset

    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB
    
    Call TevAXRF_SetMeasureSamples(8192) 'Fres = 30.5176 kHz  (Fs = 250MHz, N=8192) NOTE: Fres chosen to bound TX freq
  
    TheHdw.wait 0.05
        
    Select Case TestFreq
    Case 868300000
        TxPower868.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        
    Case 433300000
        TxPower433.AddPin ("RFLOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower433.pins("RFLOUT").value(nSiteIndex) = -90
        Next nSiteIndex
    
    Case Else
    
        TxPower868.AddPin ("RFHOUT")
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        
    End Select


    TheHdw.wait (0.002)
    
        
    'Run pattern to put DUT into TX CW Mode @ 868 MHz, +15 dBm Power

    Select Case TestFreq
        Case 868300000
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx868_cw").start ("start_tx_cw_on")
        Case 433300000
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx433_cw").start ("start_tx_cw_on")
        Case Else
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx868_cw").start ("start_tx_cw_on")
    End Select
        
        
       
        
'Pattern loops after setting cpuA. VBT loops waiting for the pattern to set cpuA.

Flag_Loop:  'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
   
    Flags = TheHdw.Digital.Patgen.CpuFlags

        'Debug.Print "Flags ="; Flags
        
        If (Flags = 1) Then
            GoTo End_flag_loop 'cpuA set - DUT in TX CW mode
        End If
 
        If (Flags = 0) Then
            GoTo Flag_Loop 'cpuA not set - DUT not in TX CW mode
        End If
  
End_flag_loop:
        
        'Measure TX Current
        
        ' DPS setup is inherently multi-site.
        
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.samples = 1
    End With
    
    
    'Measure Current
   
    Select Case TestFreq
        Case 868300000
            Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX868_CW)
        Case 433300000
            Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX433_CW)
        Case Else
            Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX868_CW)
    End Select




        'Measure RF Power
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1 'Site loop needed for AXRF

            If TheExec.Sites.site(nSiteIndex).Active = True Then
                
                TevAXRF_MeasureSetup MeasChans(nSiteIndex), 17, TestFreq  'set for +17dBm
                
                TheHdw.wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 4096, AXRF_FREQ_DOMAIN, MaxPowerTemp, FreqOffset_Hz_temp, testXtalOffset, False, "rf", False, False, False, 1, SumPowerTemp)  'True plots waveform
                
                FreqOffset_Hz(nSiteIndex) = FreqOffset_Hz_temp
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                SumPower(nSiteIndex) = SumPowerTemp
            
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
            
                Select Case TestFreq
                    Case 868300000
                        TxPower868.pins("RFHOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                    Case 433300000
                        TxPower433.pins("RFLOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                    Case Else  'Dummy  for force fail purpose
                        TxPower868.pins("RFHOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                End Select

            End If
            
        Next nSiteIndex
        
        'Reset cpuA flag
        FlagsSet = 0
        FlagsClear = cpuA

        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) 'Pattern continues after cpuA reset.
        
        
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.
  
    'Run pattern to stop DUT transmitting
    
 '       TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx_cw_off").start ("start_tx_cw_off") 'all sites


        Select Case TestFreq
            Case 868300000
                TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx_cw_off").start ("start_tx_cw_off") 'all sites
            Case 433300000
                TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx_cw_off").start ("start_tx_cw_off") 'all sites
            Case Else  'Dummy  for force fail purpose
                TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx_cw_off").start ("start_tx_cw_off") 'all sites
        End Select




        'TheHdw.Wait (0.1) 'avoids LVM Priming patgen RTE
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.

 'TheExec.DataLog.WriteComment ("=================== TX868_CW_PWR =====================")
 
    Select Case TestFreq
    
    Case 868300000
    
        TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If

'                Select Case TheExec.CurrentJob
'                Case "f1-prd-std-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower433, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
'
'                Case "f1-pgm-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, 0.035, 0.085, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower433, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
'
'                Case "q1-prd-std-rn2483"
'                    TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW_qc", , , , , , , , tlForceNone
'                    TheExec.Flow.TestLimit TxPower433, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868_qc", , , , , , , , tlForceNone
'
'                Case Else
'
'                End Select


    Case 433300000

        TheExec.Flow.TestLimit I_TX433_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX433_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower433, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_433", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFLOUT", unitHz, tlForceNone, "FRQOFF_433")
                End If
            Next nSiteIndex
        End If
        
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.03, 0.09, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 17, , , , unitDb, "%2.1f", "TxPower_433", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If
        
    End Select
 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

errHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower433.pins("RFHOUT").value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    
    Case 868300000
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If
        
        
    Case 433300000
    
        TheExec.Flow.TestLimit I_TX433_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX433_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower433, 10, 16, , , , unitDb, "%2.1f", "TxPower_433", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFLOUT", unitHz, tlForceNone, "FRQOFF_433")
                End If
            Next nSiteIndex
        End If
  
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower433, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_868")
                End If
            Next nSiteIndex
        End If

    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2483_tx433_cw = TL_ERROR
    
End Function

Public Function rn2483_i_sleep(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for 2 sec via a UART command. During the 2 second window the VBAT current is measured and reported.

    Dim site As Variant
    
    Dim I_SLEEP As New PinListData
    
    Dim oprVolt As Double
    Dim dut_delay As Double
      
    Dim nSiteIndex As Long
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    
    Dim ExistingSiteCnt As Integer
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2483_i_sleep = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs

    If argc < 1 Then
        MsgBox "Error - On rn2483_i_sleep - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        oprVolt = argv(0) '3.3
        
    End If
    
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
TheExec.DataLog.WriteComment ("================= MEASURE I_SLEEP ====================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check Test Instance Parms for 3.3v
        
        
        'Disconnect GPIO_PINS
            TheHdw.pins("GPIO_PINS").InitState = chInitOff
            TheHdw.pins("GPIO_PINS").StartState = chStartOff
            
            TheHdw.pins("MISC_PIC_IOS").InitState = chInitOff
            TheHdw.pins("MISC_PIC_IOS").StartState = chStartOff
                      
         
            'Run pattern to put DUT into SLEEP Mode
   
        I_SLEEP.AddPin ("VBAT")
        
        
        'For nSiteIndex = 0 To ExistingSiteCnt - 1
  
            I_SLEEP.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value

        'Next nSiteIndex
        
        
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep").start ("start_i_sleep")
        
       
        'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
        'DUT is sleeping when cpuA = 1.
        
        
        
Flag_Loop:

   
Flags = TheHdw.Digital.Patgen.CpuFlags

'Debug.Print "Flags ="; Flags
        
 If (Flags = 1) Then GoTo End_flag_loop 'cpuA set in pattern (DUT should be asleep).
 
 If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set in pattern
 
 
End_flag_loop:
        
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1

        
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps500ua
            .CurrentLimit = 0.1
            TheHdw.DPS.samples = 1
        End With
            
            
            
                    'Measure Current
        Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps500ua, I_SLEEP)
        
    Next nSiteIndex
    
'Clear cpuA flag so pattern can continue

    FlagsSet = 0
    FlagsClear = cpuA

Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear)

    
    Call TheHdw.Digital.Patgen.HaltWait
    
    TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone
    
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit I_SLEEP, 0.00004, 0.0006, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
                
    Call TheHdw.Digital.Patgen.Halt
        
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_SLEEP.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
    I_SLEEP.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2483_i_sleep = TL_ERROR
    
    
End Function

Public Function rn2483_i_sleep_rev(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for up to 10 sec via a UART command, as recommended by HDC (Tibor Keller).
' During the 10 second window the VBAT current is measured and reported. The Upper Test Limit (UTL) is set approximately
' 6x higher than the guaranteed-by-design data sheet value due to the inferior DPS low current measurement capability.

'Sleep Current Test Methodology

'Steps performed across multiple sites (parallel):
    '1) Send UART command in pattern to put DUT to sleep.
    '2) Pattern sets cpuA flag when DUT is sleeping
    '3) Delay between 4-5 seconds.
    '4) Disconnect all DUT pins except VBAT and GND.
    '5) Take a number of single current readings (PinListData objects) separated by time.
    '6) Reconnect pins
    '7) Clear cpuA flag to allow pattern to finish
    
'Steps performed within site loop for active sites:
    '8) Take absolute value of current readings.
    '9) Find lowest value.
    
 'Final step performed outside of site loop:
    '10) Test site-minimum DUT sleep current values and datalog for all sites.

        

    Dim site As Variant
    
    Dim I_SLEEP As New PinListData
    Dim I_SLEEP_A As New PinListData
    Dim I_SLEEP_B As New PinListData
    Dim I_SLEEP_C As New PinListData
    Dim I_SLEEP_D As New PinListData
    Dim I_SLEEP_E As New PinListData
    Dim I_SLEEP_F As New PinListData
    Dim I_SLEEP_G As New PinListData
    Dim I_SLEEP_H As New PinListData
    Dim I_SLEEP_J As New PinListData
    Dim I_SLEEP_K As New PinListData
    
    
    
    
    
    
    Dim ExistingSiteCnt As Integer
    
    Dim i As Long
    Dim num_samps As Long
    Dim SiteStatus As Long
    Dim thisSite As Long
    'Dim i As Long
    
    Dim nSiteIndex As Long
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long

    Dim I_Sleep_Samples() As Double
    Dim Sleep_Current_Mean() As Double
    Dim Sleep_Current_Sum() As Double
    Dim Sleep_Current_Max() As Double
    Dim Sleep_Current_OK() As Double
    Dim Sleep_Current_Min() As Double
    
    Dim i_sleep_temp As Double
    Dim oprVolt As Double
    Dim SLEEP_DLY As Double
    Dim INTERVAL_DLY As Double
    
    Dim tmpA, tmpB, tmpC, tmpD, tmpE As Double
    Dim tmpF, tmpG, tmpH, tmpJ, tmpK As Double
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    If 0 Then Debug.Print "LoLimit = "; LoLimit ' 20170216 - ty added if 0
    If 0 Then Debug.Print "HiLimit = "; HiLimit ' 20170216 - ty added if 0
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    'Debug.Print "Existing Site Count = "; ExistingSiteCnt
    
    On Error GoTo errHandler
    
        rn2483_i_sleep_rev = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
    I_SLEEP.AddPin ("VBAT") 'add pin to PinListData Object
    I_SLEEP_A.AddPin ("VBAT")
    I_SLEEP_B.AddPin ("VBAT")
    I_SLEEP_C.AddPin ("VBAT")
    I_SLEEP_D.AddPin ("VBAT")
    I_SLEEP_E.AddPin ("VBAT")
    I_SLEEP_F.AddPin ("VBAT")
    I_SLEEP_G.AddPin ("VBAT")
    I_SLEEP_H.AddPin ("VBAT")
    I_SLEEP_J.AddPin ("VBAT")
    I_SLEEP_K.AddPin ("VBAT")
     

    If argc < 1 Then
        MsgBox "Error - On rn2483_i_sleep_rev - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        oprVolt = argv(0) '3.3
        
    End If
    
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
'TheExec.DataLog.WriteComment ("============================= MEASURE I_SLEEP_REV =====================================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check Test Instance Parms for 3.3v
        
        num_samps = 1      'Number of current samples taken per site
        
        SLEEP_DLY = 4.5            '4.5
        INTERVAL_DLY = 0.03
        
        
        ReDim Sleep_Current_Min(ExistingSiteCnt)
        
        'Disconnect GPIO_PINS
            TheHdw.pins("MODULE_IO").InitState = chInitOff
            TheHdw.pins("MODULE_IO").StartState = chStartOff
            
            TheHdw.pins("SDA").InitState = chInitOff
            TheHdw.pins("SDA").StartState = chStartOff
         

   
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize PinListData objects and variables
        
            I_SLEEP.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_A.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_B.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_C.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_D.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_E.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_F.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_G.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_H.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_J.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_K.pins("VBAT").value(nSiteIndex) = -99
            
            Sleep_Current_Min(nSiteIndex) = 0.00009999 'UTL = 11.11uA FAILING VALUE
            
        Next nSiteIndex
        
        
        
        
        'DPS Setup
        
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps50uA     'lowest DPS current range
            TheHdw.DPS.samples = num_samps
        End With
        
                TheHdw.wait (0.05)  'DPS settling

               TheHdw.Digital.Patgen.timeout = 30
            
            Call TheHdw.Digital.Patgen.HaltWait
            
            'Run pattern to put DUT into SLEEP Mode
            'TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Unload
            'TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Load
            
            'TheHdw.Wait (1)
                'Dim SlpWaitLoop As Long
                'SlpWaitLoop = 28800
                
                'TheExec.DataLog.WriteComment "SlpWait: " & SlpWaitLoop & " Vectors"
                'Call TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").ModifyVectorOperand("start_i_sleep", 304, SlpWaitLoop)
            
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Unload
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Load
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").start ("start_i_sleep")
        
       
                'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
                'DUT is sleeping when cpuA = 1.
              
Flag_Loop:


        Flags = TheHdw.Digital.Patgen.CpuFlags

        'Debug.Print "Flags ="; Flags

         If (Flags = 1) Then GoTo End_flag_loop 'cpuA set in pattern (DUT should be asleep).

         If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set in pattern

End_flag_loop:

        TheHdw.wait (SLEEP_DLY) 'Sleep Delay 'Critical to have sleep delay with pins connected!
        
        Call TheHdw.Digital.Relays.pins("MODULE_IO").disconnectPins 'disconnect pins for lowest sleep current
        Call TheHdw.Digital.Relays.pins("SDA").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_IO").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_TX").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_RX").disconnectPins
        Call TheHdw.Digital.Relays.pins("MCLR_nRESET").disconnectPins
        
        TheHdw.wait (INTERVAL_DLY)      'HW Delay
        
'Collect several measurements at all sites with interval delay

         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_A)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_B)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_C)
        
            TheHdw.wait (INTERVAL_DLY)
            
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_D)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_E)
         
            TheHdw.wait (INTERVAL_DLY)
         
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_F)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_G)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_H)
        
            TheHdw.wait (INTERVAL_DLY)
            
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_J)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_K)
            
 'All data for all sites have been collected at this point.
    
'Clear cpuA flag so pattern can continue

    FlagsSet = 0
    FlagsClear = cpuA
    
        'Re-connect pins previously disconnected
          
        Call TheHdw.Digital.Relays.pins("MODULE_IO").connectPins
        Call TheHdw.Digital.Relays.pins("SDA").connectPins
        Call TheHdw.Digital.Relays.pins("UART_IO").connectPins
        Call TheHdw.Digital.Relays.pins("UART_TX").connectPins
        Call TheHdw.Digital.Relays.pins("UART_RX").connectPins
        Call TheHdw.Digital.Relays.pins("MCLR_nRESET").connectPins
        
        TheHdw.wait (INTERVAL_DLY)      'HW Delay

    Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear)
    

    'Call TheHdw.Digital.Patgen.Halt  'For DEBUG to prevent pattern timeout. Comment out for faster test time.

        
' Loop through the active sites
    
With TheExec.Sites
    
''        siteStatus = .SelectFirst
''
''    Do While siteStatus <> loopDone
''        thisSite = .SelectedSite
    
        If (TheExec.Sites.site(0).Active) Then
          
            TheExec.Sites.SetOverride (0)       'Site 0 Active

            'Begin data processing for Site 0
            
                If 0 Then Debug.Print "Site "; 0  ' 20170216 - ty added if 0
                If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0
    
    
            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(0))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(0))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(0))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(0))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(0))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(0))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(0))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(0))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(0))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(0))
            
     
            If 0 Then  ' 20170216 - ty added if 0
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
    
            'Search for a sleep current minimum
        
            If tmpA < Sleep_Current_Min(0) Then 'search for minimum value
                Sleep_Current_Min(0) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpE
            End If

    
            I_SLEEP.pins("VBAT").value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            If 0 Then Debug.Print Sleep_Current_Min(0) ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(1)
            'Debug.Print Sleep_Current_Min(2)
            'Debug.Print Sleep_Current_Min(3)
            
           'End data processing for Site 0
        
        TheExec.Sites.RestoreFromOverride

     End If 'Site 0 active
     
     'Exit Do        'DEBUG
     
    
''    siteStatus = TheExec.Sites.SelectNext(siteStatus)  'DEBUG
''
''    If siteStatus > 1 Then
''        siteStatus = siteStatus - 1
''    End If
''
''
''    If siteStatus = loopDone Then Exit Do
     
        If (TheExec.Sites.site(1).Active) Then
          
            TheExec.Sites.SetOverride (1)   'Site 1 Active
            
            'Begin data processing for Site 1
        
            If 0 Then Debug.Print "Site "; 1 ' 20170216 - ty added if 0
            If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(1))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(1))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(1))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(1))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(1))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(1))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(1))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(1))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(1))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(1))
            
     
            If 0 Then ' 20170216 - ty added if 0
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
            'Search for sleep current minimum
        
            If tmpA < Sleep_Current_Min(1) Then 'search for minimum value
                Sleep_Current_Min(1) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpE
            End If

    
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            I_SLEEP.pins("VBAT").value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            If 0 Then Debug.Print Sleep_Current_Min(1) ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(2)
            'Debug.Print Sleep_Current_Min(3)
            
           'End data processing for Site 1


            TheExec.Sites.RestoreFromOverride
            
        End If 'Site 1 Active
      
      
''        siteStatus = TheExec.Sites.SelectNext(siteStatus)
''
''        If siteStatus > 2 Then
''            siteStatus = siteStatus - 1
''        End If
''
''        If siteStatus = loopDone Then Exit Do
        
        If (TheExec.Sites.site(2).Active) Then
          
            TheExec.Sites.SetOverride (2)   'Site 2 Active
            
            'Begin data processing for Site 2
        
            If 0 Then Debug.Print "Site "; 2 ' 20170216 - ty added if 0
            If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(2))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(2))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(2))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(2))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(2))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(2))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(2))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(2))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(2))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(2))
            
     
            If 0 Then  ' 20170216 - ty added if 0
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
    
            'Search for sleep current minimum
        
            If tmpA < Sleep_Current_Min(2) Then 'search for minimum value
                Sleep_Current_Min(2) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpE
            End If
    
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            I_SLEEP.pins("VBAT").value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            'Debug.Print Sleep_Current_Min(1)
            If 0 Then Debug.Print Sleep_Current_Min(2)  ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(3)


            TheExec.Sites.RestoreFromOverride
      
      
        End If 'Site 2 Active
      
      
''    siteStatus = TheExec.Sites.SelectNext(siteStatus)
''
''    If siteStatus > 3 Then
''            siteStatus = siteStatus - 1
''    End If
''
''    If siteStatus = loopDone Then Exit Do
        
        If (TheExec.Sites.site(3).Active) Then
          
            TheExec.Sites.SetOverride (3)   'Site 3 Active
        
            'Begin data processing for Site 3
        
            If 0 Then Debug.Print "Site "; 3 ' 20170216 - ty added if 0
            If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(3))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(3))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(3))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(3))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(3))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(3))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(3))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(3))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(3))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(3))
            
     
            If 0 Then  ' 20170216 - ty added if 0
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
    
            'Search for sleep current minimum
        
            If tmpA < Sleep_Current_Min(3) Then 'search for minimum value
                Sleep_Current_Min(3) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpE
            End If
            
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            I_SLEEP.pins("VBAT").value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            'Debug.Print Sleep_Current_Min(1)
            'Debug.Print Sleep_Current_Min(2)
            If 0 Then Debug.Print Sleep_Current_Min(3) ' 20170216 - ty added if 0



            TheExec.Sites.RestoreFromOverride
      
        End If 'Site 3 Active

    
''    siteStatus = TheExec.Sites.SelectNext(siteStatus)
''
''    If siteStatus = loopDone Then Exit Do
''
''    If (TheExec.Sites.ActiveCount = 0 Or TheExec.Sites.ActiveCount >= 3) Then Exit Do
''
''    siteStatus = .SelectNext(loopTop)
        
''   Loop
  
End With ' TheExec.Sites
      
      'Test to Limits and Datalog
      
            If (TheExec.CurrentJob = "f1-prd-std" Or TheExec.CurrentJob = "f1-prd-qtp") Then
    
                'TheExec.Flow.TestLimit I_SLEEP, 0.0000001, 0.00001, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV", , , , , , , , tlForceNone
    
            ElseIf (TheExec.CurrentJob = "q1-prd-std" Or TheExec.CurrentJob = "q1-prd-qtp") Then
    
                'TheExec.Flow.TestLimit I_SLEEP, 0.00004, 0.0006, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP_REV_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV_qc", , , , , , , , tlForceNone
                    
            End If


'        If TheExec.CurrentJob = "f1-prd-std-rn2483r101" Then
'
'            'TheExec.Flow.TestLimit I_SLEEP, 0.00000002, 0.00001, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV", , , , , , , , tlForceNone
'            TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2483r101" Then
'
'            'TheExec.Flow.TestLimit I_SLEEP, 0.00004, 0.0006, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP_REV_qc", , , , , , , , tlForceNone
'            TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.2f", "RN2483_I_SLEEP_REV_qc", , , , , , , , tlForceNone
'
'        End If
        
    Call TheHdw.Digital.Patgen.Halt
        
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_SLEEP.AddPin ("VBAT")
            
'      For nSiteIndex = 0 To ExistingSiteCnt - 1
'
'            I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
'
'      Next nSiteIndex

     
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP_REV", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    
    Call TheExec.ErrorLogMessage("Function Error: rn2483_i_sleep_rev")
    
    Call TheExec.ErrorReport
    
    rn2483_i_sleep_rev = TL_ERROR
    
    
End Function


Public Function rn2483_gpio_r1(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'The LoRa module has 14 gpio pins, GPIO0 - GPIO13. Each gpio pin is commanded via the UART to be set to a logical 1, tested in the pattern, then set to logical 0, then tested in the pattern.

    Dim site As Variant
    
      'Dim oprVolt As Double
      'Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2483_gpio_r1 = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
'TheExec.DataLog.WriteComment ("=================== GPIO CHECK ========================")
    
        'oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        'dut_delay = 0.1
            
                      'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
  
  
  Call TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483r1_gpio_full").Test(pfAlways, 0)
        
        
        TheHdw.wait (0.3)
        
    
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_gpio_r1")
    Call TheExec.ErrorReport
    rn2483_gpio_r1 = TL_ERROR
    
End Function

Public Function id_mfs(argc As Long, argv() As String) As Long

'Multisite LoRa module ID test.
'After a reset, a functional DUT sends the UART host its ID (and FW revision time and date).

'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received ID.
'If no ID is received, however, the pattern will time out with 100 forced fails.

'Some modules take more than 100msec to respond to system reset than others. A fast and slow response pattern is used to check which type of module
' is being tested. Passing the slow OR the fast pattern will pass the ID test by finding the start bit of the response.


    TheExec.DataLog.WriteComment ("==================  MODULE_ID  =================== ")
    
    If 0 Then
        Exit Function
    End If
    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    
    Dim ActiveSite() As Boolean
    Dim nSiteIndex As Long

    Dim ValidityCount(3) As Long
    Dim ValidityCountSlow(3) As Long 'Slow responding modules
    Dim ValidityCountFast(3) As Long 'Fast responding modules
    
    
    Dim patgen_fails(3) As Long
    Dim patgen_fails_slow(3) As Long
    Dim patgen_fails_fast(3) As Long
      
    Dim i As Long
    Dim fails_ids As Long
    Dim fails_idf As Long
    
    Dim loopstatus As Long
    Dim thisSite As Long
    Dim site_ii As Long
    
    Dim ID_Valid As New PinListData
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    Dim testNumbr As New SiteLong
    
    On Error GoTo errHandler
 
    id_mfs = TL_SUCCESS
 
    For i = 0 To TheExec.Sites.ExistingCount - 1
        testNumbr(i) = TheExec.Sites.site(i).testnumber
    Next i
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs

    ID_Valid.AddPin ("RFHOUT")
    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
        
        ID_Valid.pins("RFHOUT").value(nSiteIndex) = 0
        
           ValidityCountFast(nSiteIndex) = 0
           ValidityCountSlow(nSiteIndex) = 0
           
           patgen_fails_fast(nSiteIndex) = 0
           patgen_fails_slow(nSiteIndex) = 0
        
    Next nSiteIndex
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
    Call TheHdw.Digital.Patgen.HaltWait
                        
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
    loopstatus = TheExec.Sites.SelectFirst
    
    While loopstatus <> loopDone
    With TheExec.Sites
        thisSite = .SelectedSite
            
        If .site(thisSite).Active Then

            .SetOverride (thisSite)
        
            'Debug.Print "b4 failcount = " & CStr(TheHdw.Digital.Patgen.FailCount)
        
'            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_id_ver2.pat").Unload
'            TheHdw.wait 0.01
'            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_id_ver2.pat").Load
'            TheHdw.wait 0.01
            
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_id_ver2.pat").Run ("start_uart_id")
            Call TheHdw.Digital.Patgen.HaltWait
            
    
            If 0 Then
                Debug.Print "fail count(" & CStr(thisSite) & ") = " & CStr(TheHdw.Digital.Patgen.FailCount)
            End If

            Dim respRecvd As Boolean
            Dim numOfCharsRecvd As Long
            
            respRecvd = False
            numOfCharsRecvd = 0
      
            Dim responseStr As String
            
            Call getModuleId(thisSite, respRecvd, numOfCharsRecvd, responseStr)
            
            Dim Module_ As String
            Dim Firmware_ As String
            Dim Date_ As String
            Dim Time_ As String
            
            Dim Module_Expected As String
            Dim Firmware_Expected As String
            Dim Date_Expected As String
            Dim Time_Expected As String
            
            'this is where the limits for the test is determined for the Firmware version.
            'we need a production variable to read to improve this.... we don't want a hardcoded variable
            'but, insteand, we want to read a production variable and decide from a lookup table of firmware
            'optoins to determine the limits (to test against... I will need your help on this inthe next few weeks, but not now.)
            
            Dim moduleType_ As String
            moduleType_ = argv(3)
            
            If moduleType_ = "RN2483" Then
                Module_Expected = "RN2483"
                Firmware_Expected = "1.0.1"
                Date_Expected = "15Dec2015"
                Time_Expected = "09:38:09"
            ElseIf moduleType_ = "RN2903" Then
                Module_Expected = "RN2903"
                Firmware_Expected = "0.9.8"
                Date_Expected = "14Feb2017"
                Time_Expected = "20:17:03"
            Else
                Stop
            End If
            
            'RN2483 1.0.1 Dec 15 2015 09:38:09
            If respRecvd Then
                Module_ = Left(responseStr, InStr(1, responseStr, " ") - 1)
                Firmware_ = Mid(responseStr, InStr(1, responseStr, " ") + 1, InStr(InStr(1, responseStr, " ") + 1, responseStr, " ") - InStr(1, responseStr, " ") - 1)
                Date_ = Mid(responseStr, InStr(InStr(1, responseStr, " ") + 1, responseStr, " ") + 5, 2) & _
                        Mid(responseStr, InStr(InStr(1, responseStr, " ") + 1, responseStr, " ") + 1, 3) & _
                        Mid(responseStr, InStr(InStr(1, responseStr, " ") + 1, responseStr, " ") + 8, 4)
                Time_ = Mid(responseStr, InStr(InStr(1, responseStr, " ") + 1, responseStr, " ") + 13, 8)
            End If
            
            Dim moduleType As New SiteLong
            Dim moduleFW As New SiteLong
            Dim moduleDate As New SiteLong
            Dim moduleTime As New SiteLong
            
            If respRecvd Then
                If Module_Expected = Module_ Then
                    moduleType(thisSite) = 1
                Else
                    moduleType(thisSite) = 0
                End If
                
                If Firmware_Expected = Firmware_ Then
                    moduleFW(thisSite) = 1
                Else
                    moduleFW(thisSite) = 0
                End If
                
                If Date_Expected = Date_ Then
                    moduleDate(thisSite) = 1
                Else
                    moduleDate(thisSite) = 0
                End If
                
                If Time_Expected = Time_ Then
                    moduleTime(thisSite) = 1
                Else
                    moduleTime(thisSite) = 0
                End If
            Else  ' not received...
                    moduleType(thisSite) = -1
                    moduleFW(thisSite) = -1
                    moduleDate(thisSite) = -1
                    moduleTime(thisSite) = -1
            End If

            'making the limits so they always pass until we can fix this (read from the production variable)
            Call sm_LogPassFail(thisSite, moduleType, -1, 1, "UART_TX", unitNone, tlForceNone, "moduleType", "%2.1f", , scaleNoScaling, testNumbr(thisSite))
            Call sm_LogPassFail(thisSite, moduleFW, -1, 1, "UART_TX", unitNone, tlForceNone, "moduleFW", "%2.1f", "%", scaleNoScaling, testNumbr(thisSite) + 1)
            
            'we don't care about the data and time (at least for now)
            'Call sm_LogPassFail(thisSite, moduleDate, -1, 1, "UART_TX", unitNone, tlForceNone, "moduleDate", "%2.1f", "%", scaleNoScaling, testNumbr(thisSite) + 2)
            'Call sm_LogPassFail(thisSite, moduleTime, -1, 1, "UART_TX", unitNone, tlForceNone, "moduleTime", "%2.1f", "%", scaleNoScaling, testNumbr(thisSite) + 3)
    
            TheExec.Sites.RestoreFromOverride
            
        End If 'Site X active

        loopstatus = TheExec.Sites.SelectNext(loopstatus)
    End With
    Wend 'end WHILE loop
  
    Call TheHdw.Digital.Patgen.Halt

    For thisSite = 0 To TheExec.Sites.ExistingCount - 1
        If TheExec.Sites.site(thisSite).Active Then
        TheExec.DataLog.WriteComment ("  site(" & CStr(thisSite) & "): " & responseStr)
        End If
    Next thisSite
    'TheExec.DataLog.WriteComment ("======================================================= ")

    Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    Call TheHdw.Digital.Patgen.Halt
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_id_mfs")
    Call TheExec.ErrorReport
    id_mfs = TL_ERROR
    
End Function



Public Function fsk_pkt_rcv_m_rev(argc As Long, argv() As String) As Long

    'this is the main function...

    Dim SrcChans(3) As AXRF_CHANNEL

    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    
    Dim Gate As Long
    Dim Edge As Long
    Dim nSiteIndex As Long
    
    Dim fails As Long
    Dim pkt_fails As Long
    
    Dim SiteStatus As Long
    Dim thisSite As Long

    Dim PacketCount(3) As Long
    Dim patgen_fails(3) As Long
    Dim site_ii As Long
    Dim testNumbr As New SiteLong
    Dim pktCnt_ As New SiteDouble
    Dim pktErr_ As New SiteDouble
    
    Dim i As Long
    Dim pkt_sent_count As Long
    
    Dim PKTs_RCVd As New PinListData
    Dim pktRecvd As Boolean, pktErr As Double
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    On Error GoTo errHandler
    Dim theCase As String, patName As String, freqMhz As Double, tNameStr As String
    
    'theCase = "433"  ' or "433" or "915"
    theCase = argv(3)
    tNameStr = "fsk_pkt_rcv_m_rev"
    Select Case theCase
        Case "433"
            patName = "uart_rn2483_tx433_fsk_pkt_one_rev2"
            freqMhz = 433.3
        Case "868"
            patName = "uart_rn2483_tx868_fsk_pkt_one_rev2"
            freqMhz = 868.3
        Case "915"
            patName = "uart_rn2903_tx915_fsk_pkt_one_rev2"
            freqMhz = 915#
        Case Else
    
    End Select
    
 
    fsk_pkt_rcv_m_rev = TL_SUCCESS
 
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
        
    PKTs_RCVd.AddPin "RFHOUT"
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
        PKTs_RCVd.pins("RFHOUT").value(nSiteIndex) = 0
        patgen_fails(nSiteIndex) = 0
    Next nSiteIndex
     

    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB

    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
 
'    TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\" & patName & ".pat").Unload
'    TheHdw.wait 0.005
'    TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\" & patName & ".pat").Load
                     
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_ccitt.aiq"  'NFG      'Non-inverted CRC, MSB/LSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_swap_ccitt.aiq"       'Non-inverted CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_swap_ccitt.aiq"              'Inverted CCITT CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"           'Inverted CCITT CRC, MSB/LSB
    ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"
    
    If theCase = "433" Then
        SrcChans(0) = AXRF_CHANNEL_AXRF_CH2
        SrcChans(1) = AXRF_CHANNEL_AXRF_CH4
        SrcChans(2) = AXRF_CHANNEL_AXRF_CH6
        SrcChans(3) = AXRF_CHANNEL_AXRF_CH8
    Else
        SrcChans(0) = AXRF_CHANNEL_AXRF_CH1
        SrcChans(1) = AXRF_CHANNEL_AXRF_CH3
        SrcChans(2) = AXRF_CHANNEL_AXRF_CH5
        SrcChans(3) = AXRF_CHANNEL_AXRF_CH7
    End If
   
    'AXRF Modulation Trigger Arm Source Parameters
    Gate = 1  '1 = Modulation is ON for duration of HIGH trigger, or when the modulation ends; 0 =  Modulation starts w/Trig and runs continuously
    Edge = 0    'Positive edge
    For site_ii = 0 To TheExec.Sites.ExistingCount - 1
        testNumbr(site_ii) = TheExec.Sites.site(site_ii).testnumber
    Next site_ii
                 
    Call TheHdw.Digital.Patgen.HaltWait
    'TheHdw.Digital.Patgen.Threading = False
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
    ' Loop through all the active sites
    With TheExec.Sites
        SiteStatus = .SelectFirst
    
        Do While SiteStatus <> loopDone
            thisSite = .SelectedSite
            
                If (.site(thisSite).Active) Then
           
                    .SetOverride (thisSite)
            
                    'Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(thisSite), ModFilePath) 'Separate Loads for each site?
                    Call TevAXRF_LoadModulationFile(SrcChans(thisSite), ModFilePath) 'Separate Loads for each site?
                    TheHdw.wait (0.1)
            
                    'Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(thisSite), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, 0, Edge)  ' 0,63,0,0
                    If 0 Then
                    Call TevAXRF_ModulationTriggerArm(SrcChans(thisSite), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, 0, Edge)  ' 0,63,0,0
                    End If
                    'Call itl.Raw.AF.AXRF.StartModulation(SrcChans(thisSite), ModFilePath)
                    Call TevAXRF_StartModulation(SrcChans(thisSite), ModFilePath)
                    
                    TheHdw.wait (0.05)
                    '20180208 - all 4 MMT known bad parts (2903s) failed at -80dBm, which is expected.
                    '         - 6 known good (golden units) (2483s) were marginal at about -91dBm, but pass at -85dBm
                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
                    'Call itl.Raw.AF.AXRF.source(SrcChans(thisSite), -85, freqMhz * 1000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                    Call TevAXRF_Source(SrcChans(thisSite), -85, freqMhz * 1000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                
                    TheHdw.wait (0.005)
                    TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\" & patName & ".pat").Run ("start_fsk_pkt_one_rev")
                    Call TheHdw.Digital.Patgen.HaltWait
                    TheHdw.wait 0.005
                    
                    'the following will datalog the fail count.... if the failed count is 65535 - it means it went thru the loop and never saw a low.
                    '  that means, no signal (no data from DUT saying it received a signal).
                    ' if it is less than 65535, say 600, then it worked and we shuld have data to look at / process
                    If 0 Then
                        Debug.Print "fail count(" & CStr(thisSite) & ") = " & CStr(TheHdw.Digital.Patgen.FailCount)
                    End If
                    
                    pktRecvd = False
                    pktErr = 0#
                    'this isn't really a packet error test, but just looking at the ~50 characters transmitted and compare them to expected.
                    ' all we really care for the module is the Module actually received the signal from the AXRF and demodulated it properly...
                    '  ---> we just care about its functionality... to get warm fuzzy that the RX block is working
                    ' this is hte only RX tst we have...
                    Call getPktErr(thisSite, pktRecvd, pktErr)
                    If pktRecvd Then
                        PKTs_RCVd.AddPin("RFHOUT").value(thisSite) = 1  ' actually more than this... 5?
                        pktErr_(thisSite) = pktErr
                    Else
                        PKTs_RCVd.AddPin("RFHOUT").value(thisSite) = 0
                        pktErr_(thisSite) = -1#
                    End If

                    pktCnt_(thisSite) = PKTs_RCVd.AddPin("RFHOUT").value(thisSite)
                    Call sm_LogPassFail(thisSite, pktCnt_, LoLimit, HiLimit, "RFHIN", unitNone, tlForceNone, "PktCnt_" & theCase, "%2.1f", , scaleNoScaling, testNumbr(thisSite))
                    Call sm_LogPassFail(thisSite, pktErr_, 0, 6#, "RFHIN", unitCustom, tlForceNone, "PktErr_" & theCase, "%2.1f", "%", scaleNoScaling, testNumbr(thisSite) + 1)
                        
                    .RestoreFromOverride
                    SiteStatus = .SelectNext(SiteStatus)
                    
                End If 'Site X Active
            
                If SiteStatus = loopDone Then Exit Do

            'If (TheExec.Sites.ActiveCount = 0 Or TheExec.Sites.ActiveCount >= 3) Then Exit Do
                'siteStatus = .SelectNext(loopTop)
        Loop
          
    End With ' TheExec.Sites


    Call TheHdw.Digital.Patgen.Halt
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
        If TheExec.Sites.site(nSiteIndex).Active Then
            'Call itl.Raw.AF.AXRF.StopModulation(SrcChans(nSiteIndex))
            Call TevAXRF_StopModulation(SrcChans(nSiteIndex))
            TheHdw.wait 0.1
            'Call itl.Raw.AF.AXRF.UnloadModulationFile(SrcChans(nSiteIndex), ModFilePath)
            Call TevAXRF_UnloadModulationFile(SrcChans(nSiteIndex), ModFilePath)
            TheHdw.wait 0.1
        End If
    Next nSiteIndex
    Call SetAXRFinRxMode(SrcChans, -120, freqMhz * 1000000#) 'turn off RF source
    Call TheHdw.Digital.Patgen.Halt
    Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: " & tNameStr)
    Call TheExec.ErrorReport
    fsk_pkt_rcv_m_rev = TL_ERROR
    
End Function


Public Function rn2483_idle_current(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'Reset times vary for the modules, so two attempts are allowed to measure the idle current after reset.

    Dim site As Variant
    Dim I_IDLE As New PinListData
    
      Dim oprVolt As Double
      Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    

    
    
    
    On Error GoTo errHandler
    
        rn2483_idle_current = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
'TheExec.DataLog.WriteComment ("=================== MEASURE I_IDLE ===================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        dut_delay = 0.1
        

            
            'Disconnect GPIO_PINS
            TheHdw.pins("GPIO_PINS").InitState = chInitOff
            TheHdw.pins("GPIO_PINS").StartState = chStartOff
            
            
            'Disconnect MISC_PIC_IOS
            TheHdw.pins("MISC_PIC_IOS").InitState = chInitOff
            TheHdw.pins("MISC_PIC_IOS").StartState = chStartOff
            
                        'Disconnect UART_TX, UART_RX, SCL, SDA
            TheHdw.pins("UART_TX,UART_RX,SCL,SDA").InitState = chInitOff
            TheHdw.pins("UART_TX,UART_RX,SCL,SDA").StartState = chStartOff
            
            
            'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
            
        
         'Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
        TheHdw.wait (0.05)
        
        
                    'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
            
         TheHdw.wait (0.05)
         
                     'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
               
        
        I_IDLE.AddPin ("VBAT")
        
'  For nSiteIndex = 0 To ExistingSiteCnt - 1
'
'    I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
'
'  Next nSiteIndex
        
        TheHdw.wait (0.3)
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
     If TheExec.Sites.site(nSiteIndex).Active Then
        
        
        
        I_IDLE.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps100mA
            .CurrentLimit = 0.1
            TheHdw.DPS.samples = 1
            Call .MeasureCurrents(dps100mA, I_IDLE)
        End With
        
   
    
            If I_IDLE.pins("VBAT").value(nSiteIndex) > 0.007 Then
            
            TheHdw.wait (0.3)
            
                 With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.samples = 1
                    Call .MeasureCurrents(dps100mA, I_IDLE)
                End With
                
            End If 'failing first attempt
            
      End If 'Site active
    
        
    Next nSiteIndex
    
    TheExec.Flow.TestLimit I_IDLE, LoLimit, HiLimit, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE", , , , , , , , tlForceNone
    
        
''        Select Case TheExec.CurrentJob
''        Case "f1-prd-std-rn2483"
''            TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2483", , , , , , , , tlForceNone
''
''        Case "f1-pgm-rn2483"
''            TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2483", , , , , , , , tlForceNone
''
''        Case "q1-prd-std-rn2483"
''            TheExec.Flow.TestLimit I_IDLE, 0.001, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2483_qc", , , , , , , , tlForceNone
''
''        Case Else
''
''        End Select
   
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_IDLE.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
        I_IDLE.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2483", , , , , , , , tlForceNone

    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_idle_current")
    Call TheExec.ErrorReport
    rn2483_idle_current = TL_ERROR
    
End Function

Public Function rn2903_i_sleep(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for 2 sec via a UART command. During the 2 second window the VBAT current is measured and reported.

    Dim site As Variant
    
    Dim I_SLEEP As New PinListData
    
    Dim oprVolt As Double
    Dim dut_delay As Double
      
    Dim nSiteIndex As Long
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    
    Dim ExistingSiteCnt As Integer
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2903_i_sleep = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs

    If argc < 1 Then
        MsgBox "Error - On rn2483_i_sleep - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        oprVolt = argv(0) '3.3
        
    End If
    
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
TheExec.DataLog.WriteComment ("================== MEASURE I_SLEEP ==================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check Test Instance Parms for 3.3v
        
        
        'Disconnect GPIO_PINS
            TheHdw.pins("GPIO_PINS").InitState = chInitOff
            TheHdw.pins("GPIO_PINS").StartState = chStartOff
            
            TheHdw.pins("MISC_PIC_IOS").InitState = chInitOff
            TheHdw.pins("MISC_PIC_IOS").StartState = chStartOff
                      
         
            'Run pattern to put DUT into SLEEP Mode
   
        I_SLEEP.AddPin ("VBAT")
        
        
        'For nSiteIndex = 0 To ExistingSiteCnt - 1
  
            I_SLEEP.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value

        'Next nSiteIndex
        
        
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_sleep").start ("start_i_sleep")
        
       
        'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
        'DUT is sleeping when cpuA = 1.
        
        
        
Flag_Loop:

   
Flags = TheHdw.Digital.Patgen.CpuFlags

'Debug.Print "Flags ="; Flags
        
 If (Flags = 1) Then GoTo End_flag_loop 'cpuA set in pattern (DUT should be asleep).
 
 If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set in pattern
 
 
End_flag_loop:
        
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1

        
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps500ua
            .CurrentLimit = 0.1
            TheHdw.DPS.samples = 1
        End With
            
            
            
                    'Measure Current
        Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps500ua, I_SLEEP)
        
    Next nSiteIndex
    
'Clear cpuA flag so pattern can continue

    FlagsSet = 0
    FlagsClear = cpuA

Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear)

    
    Call TheHdw.Digital.Patgen.HaltWait
    
    TheExec.Flow.TestLimit I_SLEEP, LoLimit, HiLimit, , , scaleMicro, unitAmp, "%4.0f", "RN2903_I_SLEEP", , , , , , , , tlForceNone
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2903_I_SLEEP", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_SLEEP, 0.00004, 0.0006, , , scaleMicro, unitAmp, "%4.0f", "RN2903_I_SLEEP_qc", , , , , , , , tlForceNone
'
'        End If
        
    Call TheHdw.Digital.Patgen.Halt
        
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_SLEEP.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
    I_SLEEP.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2903_I_SLEEP", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2903_i_sleep = TL_ERROR
    
    
End Function
Public Function rn2903a_i_sleep_rev(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for up to 10 sec via a UART command, as recommended by HDC (Tibor Keller).
' During the 10 second window the VBAT current is measured and reported. The Upper Test Limit (UTL) is set approximately
' 6x higher than the guaranteed-by-design data sheet value due to the inferior DPS low current measurement capability.

'Sleep Current Test Methodology

'Steps performed across multiple sites (parallel):
    '1) Send UART command in pattern to put DUT to sleep.
    '2) Pattern sets cpuA flag when DUT is sleeping
    '3) Delay between 4-5 seconds.
    '4) Disconnect all DUT pins except VBAT and GND.
    '5) Take a number of single current readings (PinListData objects) separated by time.
    '6) Reconnect pins
    '7) Clear cpuA flag to allow pattern to finish
    
'Steps performed within site loop for active sites:
    '8) Take absolute value of current readings.
    '9) Find lowest value.
    
 'Final step performed outside of site loop:
    '10) Test site-minimum DUT sleep current values and datalog for all sites.

        

    Dim site As Variant
    
    Dim I_SLEEP As New PinListData
    Dim I_SLEEP_A As New PinListData
    Dim I_SLEEP_B As New PinListData
    Dim I_SLEEP_C As New PinListData
    Dim I_SLEEP_D As New PinListData
    Dim I_SLEEP_E As New PinListData
    Dim I_SLEEP_F As New PinListData
    Dim I_SLEEP_G As New PinListData
    Dim I_SLEEP_H As New PinListData
    Dim I_SLEEP_J As New PinListData
    Dim I_SLEEP_K As New PinListData
    
    
    
    
    
    
    Dim ExistingSiteCnt As Integer
    
    Dim i As Long
    Dim num_samps As Long
    Dim SiteStatus As Long
    Dim thisSite As Long
    'Dim i As Long
    
    Dim nSiteIndex As Long
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long

    Dim I_Sleep_Samples() As Double
    Dim Sleep_Current_Mean() As Double
    Dim Sleep_Current_Sum() As Double
    Dim Sleep_Current_Max() As Double
    Dim Sleep_Current_OK() As Double
    Dim Sleep_Current_Min() As Double
    
    Dim i_sleep_temp As Double
    Dim oprVolt As Double
    Dim SLEEP_DLY As Double
    Dim INTERVAL_DLY As Double
    
    Dim tmpA, tmpB, tmpC, tmpD, tmpE As Double 'First pass
    Dim tmpF, tmpG, tmpH, tmpJ, tmpK As Double  'Second pass (if needed)
    
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    If 0 Then Debug.Print "Existing Site Count = "; ExistingSiteCnt ' 20170216 - ty added if 0
    
    On Error GoTo errHandler
    
        rn2903a_i_sleep_rev = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
    I_SLEEP.AddPin ("VBAT") 'add pin to PinListData Object
    I_SLEEP_A.AddPin ("VBAT")
    I_SLEEP_B.AddPin ("VBAT")
    I_SLEEP_C.AddPin ("VBAT")
    I_SLEEP_D.AddPin ("VBAT")
    I_SLEEP_E.AddPin ("VBAT")
    I_SLEEP_F.AddPin ("VBAT")
    I_SLEEP_G.AddPin ("VBAT")
    I_SLEEP_H.AddPin ("VBAT")
    I_SLEEP_J.AddPin ("VBAT")
    I_SLEEP_K.AddPin ("VBAT")
     

    If argc < 1 Then
        MsgBox "Error - On rn2903a_i_sleep_rev - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        oprVolt = argv(0) '3.3
        
    End If
    
    TheHdw.Digital.Patgen.ThreadingForActiveSites = False
    
'TheExec.DataLog.WriteComment ("============================= MEASURE I_SLEEP_REV =====================================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check Test Instance Parms for 3.3v
        
        num_samps = 1      'Number of current samples taken per site
        
        SLEEP_DLY = 4.5
        INTERVAL_DLY = 0.1
        
        
        ReDim Sleep_Current_Min(ExistingSiteCnt)
        
        'Disconnect GPIO_PINS
            TheHdw.pins("MODULE_IO").InitState = chInitOff
            TheHdw.pins("MODULE_IO").StartState = chStartOff
            
            TheHdw.pins("SDA").InitState = chInitOff
            TheHdw.pins("SDA").StartState = chStartOff
         

   
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize PinListData objects and variables
        
            I_SLEEP.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_A.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_B.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_C.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_D.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_E.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_F.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_G.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_H.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_J.pins("VBAT").value(nSiteIndex) = -99
            I_SLEEP_K.pins("VBAT").value(nSiteIndex) = -99
            
            Sleep_Current_Min(nSiteIndex) = 0.00009999 'UTL = 11.1uA FAILING VALUE
            
        Next nSiteIndex
        
        
        
        
        'DPS Setup
        
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps50uA        'lowest DPS current range
            TheHdw.DPS.samples = num_samps
        End With
        
                TheHdw.wait (0.05) 'DPS settling

               TheHdw.Digital.Patgen.timeout = 30
            
            Call TheHdw.Digital.Patgen.HaltWait
            
            'Run pattern to put DUT into SLEEP Mode
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Unload
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").Load
            TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_sleep_revised").start ("start_i_sleep")
        
       
                'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
                'DUT is sleeping when cpuA = 1.
              
Flag_Loop:
        
           
        Flags = TheHdw.Digital.Patgen.CpuFlags
        
        'Debug.Print "Flags ="; Flags
                
         If (Flags = 1) Then GoTo End_flag_loop 'cpuA set in pattern (DUT should be asleep).
         
         If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set in pattern
                  
End_flag_loop:

        TheHdw.wait (SLEEP_DLY) 'Sleep Delay 'Critical to have sleep delay with pins connected!
        
        Call TheHdw.Digital.Relays.pins("MODULE_IO").disconnectPins 'disconnect pins for lowest sleep current
        Call TheHdw.Digital.Relays.pins("SDA").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_IO").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_TX").disconnectPins
        Call TheHdw.Digital.Relays.pins("UART_RX").disconnectPins
        Call TheHdw.Digital.Relays.pins("MCLR_nRESET").disconnectPins
        
        TheHdw.wait (INTERVAL_DLY)
        
'Collect several measurements at all sites with interval delay

        Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_A)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_B)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_C)
        
            TheHdw.wait (INTERVAL_DLY)
            
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_D)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_E)
         
            TheHdw.wait (INTERVAL_DLY)
         
        Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_F)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_G)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_H)
        
            TheHdw.wait (INTERVAL_DLY)
            
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_J)
        
            TheHdw.wait (INTERVAL_DLY)
        
         Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_SLEEP_K)
            
 'All data for all sites have been collected at this point.
    
'Clear cpuA flag so pattern can continue

    FlagsSet = 0
    FlagsClear = cpuA
    
        'Re-connect pins previously disconnected
          
        Call TheHdw.Digital.Relays.pins("MODULE_IO").connectPins
        Call TheHdw.Digital.Relays.pins("SDA").connectPins
        Call TheHdw.Digital.Relays.pins("UART_IO").connectPins
        Call TheHdw.Digital.Relays.pins("UART_TX").connectPins
        Call TheHdw.Digital.Relays.pins("UART_RX").connectPins
        Call TheHdw.Digital.Relays.pins("MCLR_nRESET").connectPins
        
        TheHdw.wait (INTERVAL_DLY)

    Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear)
    

    'Call TheHdw.Digital.Patgen.Halt  'For DEBUG to prevent pattern timeout. Comment out for faster test time.

        
' Loop through the active sites
    
With TheExec.Sites
    
'        siteStatus = .SelectFirst
'
'    Do While siteStatus <> loopDone
'        thisSite = .SelectedSite
    
        If (TheExec.Sites.site(0).Active) Then
          
            TheExec.Sites.SetOverride (0)       'Site 0 Active

            'Begin data processing for Site 0
            
                If 0 Then Debug.Print "Site "; 0 ' 20170216 - ty added if 0
                If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0
    
    
            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(0))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(0))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(0))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(0))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(0))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(0))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(0))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(0))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(0))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(0))
            
     
            If 0 Then
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
            
            'Search the first 5 values for a sleep current minimum
        
            If tmpA < Sleep_Current_Min(0) Then 'search for minimum value
                Sleep_Current_Min(0) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(0) Then
                Sleep_Current_Min(0) = tmpE
            End If
    
            I_SLEEP.pins("VBAT").value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            If 0 Then Debug.Print Sleep_Current_Min(0)  ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(1)
            'Debug.Print Sleep_Current_Min(2)
            'Debug.Print Sleep_Current_Min(3)
            
           'End data processing for Site 0
        
        TheExec.Sites.RestoreFromOverride

     End If 'Site 0 active
     
     'Exit Do        'DEBUG
     
    
'    siteStatus = TheExec.Sites.SelectNext(siteStatus)  'DEBUG
'
'    If siteStatus > 1 Then
'        siteStatus = siteStatus - 1
'    End If
'
'
'    If siteStatus = loopDone Then Exit Do
     
        If (TheExec.Sites.site(1).Active) Then
          
            TheExec.Sites.SetOverride (1)   'Site 1 Active
            
            'Begin data processing for Site 1
        
            If 0 Then Debug.Print "Site "; 1 ' 20170216 - ty added if 0
            If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(1))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(1))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(1))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(1))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(1))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(1))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(1))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(1))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(1))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(1))
            
     
            If 0 Then
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
            'Search the first 5 values for sleep current minimum
        
            If tmpA < Sleep_Current_Min(1) Then 'search for minimum value
                Sleep_Current_Min(1) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(1) Then
                Sleep_Current_Min(1) = tmpE
            End If
    
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            I_SLEEP.pins("VBAT").value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            If 0 Then Debug.Print Sleep_Current_Min(1)  ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(2)
            'Debug.Print Sleep_Current_Min(3)
            
           'End data processing for Site 1


            TheExec.Sites.RestoreFromOverride
            
        End If 'Site 1 Active
      
      
'        siteStatus = TheExec.Sites.SelectNext(siteStatus)
'
'        If siteStatus > 2 Then
'            siteStatus = siteStatus - 1
'        End If
'
'        If siteStatus = loopDone Then Exit Do
        
        If (TheExec.Sites.site(2).Active) Then
          
            TheExec.Sites.SetOverride (2)   'Site 2 Active
            
            'Begin data processing for Site 2
        
            If 0 Then Debug.Print "Site "; 2 ' 20170216 - ty added if 0
            If 0 Then Debug.Print "Sleep Current Samples..." ' 20170216 - ty added if 0


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(2))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(2))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(2))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(2))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(2))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(2))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(2))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(2))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(2))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(2))
            
     
            If 0 Then  ' 20170216 - ty added if 0
            Debug.Print "Current Samples..."
            Debug.Print "A = "; tmpA
            Debug.Print "B = "; tmpB
            Debug.Print "C = "; tmpC
            Debug.Print "D = "; tmpD
            Debug.Print "E = "; tmpE
            Debug.Print "F = "; tmpF
            Debug.Print "G = "; tmpG
            Debug.Print "H = "; tmpH
            Debug.Print "J = "; tmpJ
            Debug.Print "K = "; tmpK
            End If
    
            'Search the first 5 values for sleep current minimum
        
            If tmpA < Sleep_Current_Min(2) Then 'search for minimum value
                Sleep_Current_Min(2) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(2) Then
                Sleep_Current_Min(2) = tmpE
            End If
    
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            I_SLEEP.pins("VBAT").value(2) = Sleep_Current_Min(2) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            'Debug.Print Sleep_Current_Min(1)
            If 0 Then Debug.Print Sleep_Current_Min(2) ' 20170216 - ty added if 0
            'Debug.Print Sleep_Current_Min(3)


            TheExec.Sites.RestoreFromOverride
      
      
        End If 'Site 2 Active
      
      
'    siteStatus = TheExec.Sites.SelectNext(siteStatus)
'
'    If siteStatus > 3 Then
'            siteStatus = siteStatus - 1
'    End If
'
'    If siteStatus = loopDone Then Exit Do
        
        If (TheExec.Sites.site(3).Active) Then
          
            TheExec.Sites.SetOverride (3)   'Site 3 Active
        
            'Begin data processing for Site 3
        
            If (0) Then ' ty added 20170216
                Debug.Print "Site "; 3
                Debug.Print "Sleep Current Samples..."
            End If


            'Extract absolute value of site measurements assigned to local variables.
        
            tmpA = Abs(I_SLEEP_A.pins("VBAT").value(3))
            
            tmpB = Abs(I_SLEEP_B.pins("VBAT").value(3))
            
            tmpC = Abs(I_SLEEP_C.pins("VBAT").value(3))
            
            tmpD = Abs(I_SLEEP_D.pins("VBAT").value(3))
            
            tmpE = Abs(I_SLEEP_E.pins("VBAT").value(3))
            
            tmpF = Abs(I_SLEEP_F.pins("VBAT").value(3))
            
            tmpG = Abs(I_SLEEP_G.pins("VBAT").value(3))
            
            tmpH = Abs(I_SLEEP_H.pins("VBAT").value(3))
            
            tmpJ = Abs(I_SLEEP_J.pins("VBAT").value(3))
            
            tmpK = Abs(I_SLEEP_K.pins("VBAT").value(3))
            
     
            If (0) Then ' ty added 20170216
                Debug.Print "Current Samples..."
                Debug.Print "A = "; tmpA
                Debug.Print "B = "; tmpB
                Debug.Print "C = "; tmpC
                Debug.Print "D = "; tmpD
                Debug.Print "E = "; tmpE
                Debug.Print "F = "; tmpF
                Debug.Print "G = "; tmpG
                Debug.Print "H = "; tmpH
                Debug.Print "J = "; tmpJ
                Debug.Print "K = "; tmpK
            End If
    
    
            'Search the first 5 values for sleep current minimum
        
            If tmpA < Sleep_Current_Min(3) Then 'search for minimum value
                Sleep_Current_Min(3) = tmpA
            End If
            
            If tmpB < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpB
            End If
            
            If tmpC < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpC
            End If
            
            If tmpD < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpD
            End If
            
            If tmpE < Sleep_Current_Min(3) Then
                Sleep_Current_Min(3) = tmpE
            End If
    
            'I_SLEEP.pins("VBAT").Value(0) = Sleep_Current_Min(0) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(1) = Sleep_Current_Min(1) 'DEBUG
            'I_SLEEP.pins("VBAT").Value(2) = Sleep_Current_Min(2) 'DEBUG
            I_SLEEP.pins("VBAT").value(3) = Sleep_Current_Min(3) 'DEBUG
            
            'Debug.Print Sleep_Current_Min(0)
            'Debug.Print Sleep_Current_Min(1)
            'Debug.Print Sleep_Current_Min(2)
            If 0 Then Debug.Print Sleep_Current_Min(3)  ' 20170216 - ty added if 0



            TheExec.Sites.RestoreFromOverride
      
        End If 'Site 3 Active

    
'    siteStatus = TheExec.Sites.SelectNext(siteStatus)
'
'    If siteStatus = loopDone Then Exit Do
'
'    If (TheExec.Sites.ActiveCount = 0 Or TheExec.Sites.ActiveCount >= 3) Then Exit Do
'
'    siteStatus = .SelectNext(loopTop)
'
'   Loop
  
End With ' TheExec.Sites
      
      'Test to Limits and Datalog
      
        If ((TheExec.CurrentJob = "f1-prd-std") Or (TheExec.CurrentJob = "f1-prd-qtp")) Then
        
            TheExec.Flow.TestLimit I_SLEEP, 0.0000001, 0.00001, , , scaleMicro, unitAmp, "%4.2f", "RN2903A_I_SLEEP_REV", , , , , , , , tlForceNone
              
        ElseIf ((TheExec.CurrentJob = "q1-prd-std") Or (TheExec.CurrentJob = "q1-prd-qtp")) Then

            TheExec.Flow.TestLimit I_SLEEP, 0.0000001, 0.000011, , , scaleMicro, unitAmp, "%4.2f", "RN2903A_I_SLEEP_REV_qc", , , , , , , , tlForceNone
     
        End If
        
    Call TheHdw.Digital.Patgen.Halt
        
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_SLEEP.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
            I_SLEEP.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

     
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2903A_I_SLEEP_REV", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    
    Call TheExec.ErrorLogMessage("Function Error: rn2903a_i_sleep_rev")
    
    Call TheExec.ErrorReport
    
    rn2903a_i_sleep_rev = TL_ERROR
    
    
End Function



Public Function rn2903_idle_current(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'Reset times vary for the modules, so two attempts are allowed to measure the idle current after reset.

    Dim site As Variant
    Dim I_IDLE As New PinListData
    
      Dim oprVolt As Double
      Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
    On Error GoTo errHandler
    
        rn2903_idle_current = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
TheExec.DataLog.WriteComment ("=================== MEASURE I_IDLE ==================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        dut_delay = 0.1
        

            
            'Disconnect GPIO_PINS
            TheHdw.pins("GPIO_PINS").InitState = chInitOff
            TheHdw.pins("GPIO_PINS").StartState = chStartOff
            
            
            'Disconnect MISC_PIC_IOS
            TheHdw.pins("MISC_PIC_IOS").InitState = chInitOff
            TheHdw.pins("MISC_PIC_IOS").StartState = chStartOff
            
                        'Disconnect UART_TX, UART_RX, SCL, SDA
            TheHdw.pins("UART_TX,UART_RX,SCL,SDA").InitState = chInitOff
            TheHdw.pins("UART_TX,UART_RX,SCL,SDA").StartState = chStartOff
            
            
            'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
            
        
         'Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
        TheHdw.wait (0.05)
        
        
                    'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
            
         TheHdw.wait (0.05)
         
                     'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
        
        I_IDLE.AddPin ("VBAT")
        
'  For nSiteIndex = 0 To ExistingSiteCnt - 1
'
'    I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
'
'  Next nSiteIndex
        
        TheHdw.wait (0.3)
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
     If TheExec.Sites.site(nSiteIndex).Active Then
        
        
        
        I_IDLE.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps100mA
            .CurrentLimit = 0.1
            TheHdw.DPS.samples = 1
            Call .MeasureCurrents(dps100mA, I_IDLE)
        End With
        
   
    
            If I_IDLE.pins("VBAT").value(nSiteIndex) > 0.007 Then
            
            TheHdw.wait (0.3)
            
                 With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.samples = 1
                    Call .MeasureCurrents(dps100mA, I_IDLE)
                End With
                
            End If 'failing first attempt
            
      End If 'Site active
        
    Next nSiteIndex
    
    TheExec.Flow.TestLimit I_IDLE, LoLimit, HiLimit, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2903", , , , , , , , tlForceNone
    
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2903", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_IDLE, 0.001, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2903_qc", , , , , , , , tlForceNone
'
'        End If
    
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    I_IDLE.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
        I_IDLE.pins("VBAT").value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2903", , , , , , , , tlForceNone

    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_idle_current")
    Call TheExec.ErrorReport
    rn2903_idle_current = TL_ERROR
    
End Function

Public Function rn2903_tx915_cw(argc As Long, argv() As String) As Long

    Dim site As Variant
    Dim OOKDepthMode As Boolean
    Dim ExistingSiteCnt As Integer
    
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    Dim nSiteIndex As Long
    
    Dim MeasPower(1) As Double
    Dim MeasData() As Double

    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double          'Site Array
    Dim SumPower() As Double                'Site Array
    
    Dim MaxPowerTemp As Double
    Dim SumPowerTemp As Double
    Dim TestFreq As Double
    Dim oprVolt As Double
    
    Dim IndexMaxPower As New SiteDouble
    
    Dim MeasChans() As AXRF_CHANNEL         'Site Array

    Dim TxPower915 As New PinListData
    Dim I_TX915_CW As New PinListData

    Dim testXtalOffset As Boolean
    Dim FreqOffset_Hz As New SiteDouble
    Dim FreqOffset_Hz_temp As Double
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    ReDim SumPower(0 To ExistingSiteCnt - 1)
    
    Dim LoLimit_I As Double
    Dim LoLimit_Tx As Double
    Dim HiLimit_I As Double
    Dim HiLimit_Tx As Double
    '--------Argument processing--------'
    LoLimit_I = argv(2)
    HiLimit_I = argv(3)
    LoLimit_Tx = argv(4)
    HiLimit_Tx = argv(5)
    '------- end of argument process -------'
    
    On Error GoTo errHandler
    
        rn2903_tx915_cw = TL_SUCCESS
    
    'AXRF Channel assignments
    
    Select Case ExistingSiteCnt
        
    Case Is = 1
        MeasChans(0) = AXRF_CH1
        
    Case Is = 2
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        
    Case Is = 3
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        MeasChans(2) = AXRF_CH5
        
    Case Is = 4
        MeasChans(0) = AXRF_CH1
        MeasChans(1) = AXRF_CH3
        MeasChans(2) = AXRF_CH5
        MeasChans(3) = AXRF_CH7
        
        
    Case Else
        MsgBox "Error in [rn2903_tx915_cw]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo errHandler
        
    End Select
    

    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    

    If argc < 2 Then
        MsgBox "Error - On rn2903_tx915_cw - Wrong Argument Assigned", , "Error"
        GoTo errHandler
    Else
        TestFreq = argv(0)
        oprVolt = argv(1)
        
    End If
    
    testXtalOffset = True                   'set true if you want to test xtal offset

    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB
    
    Call TevAXRF_SetMeasureSamples(8192) 'Fres = 30.5176 kHz  (Fs = 250MHz, N=8192) NOTE: Fres chosen to bound TX freq
  
    TheHdw.wait 0.05
        
    Select Case TestFreq
    Case 915000000
        TxPower915.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower915.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        

    
    Case Else
        TxPower915.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower915.pins("RFHOUT").value(nSiteIndex) = -90
        Next nSiteIndex
        
        
    End Select


    TheHdw.wait (0.002)
    
        
        'Run pattern to put DUT into TX CW Mode @ 915 MHz, +20 dBm Power
        
    
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_tx915_cw").start ("start_tx_cw_on")
           
        
'Pattern loops after setting cpuA. VBT loops waiting for the pattern to set cpuA.

Flag_Loop:  'Check for cpuA flag from pattern 'Flags: cpuA = 1 when set and 0 when cleared
   
    Flags = TheHdw.Digital.Patgen.CpuFlags

        'Debug.Print "Flags ="; Flags
        
        If (Flags = 1) Then GoTo End_flag_loop 'cpuA set - DUT in TX CW mode
 
        If (Flags = 0) Then GoTo Flag_Loop 'cpuA not set - DUT not in TX CW mode
  
End_flag_loop:
        
        'Measure TX Current
        
        ' DPS setup is inherently multi-site.
        
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.samples = 1
    End With
    
    
        'Measure Current
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX915_CW)
    

        'Measure RF Power
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1 'Site loop needed for AXRF

            If TheExec.Sites.site(nSiteIndex).Active = True Then
                
                TevAXRF_MeasureSetup MeasChans(nSiteIndex), 20, TestFreq  'set for +20dBm
                
                TheHdw.wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 4096, AXRF_FREQ_DOMAIN, MaxPowerTemp, FreqOffset_Hz_temp, testXtalOffset, False, "rf", False, False, False, 1, SumPowerTemp)  'True plots waveform
                
                FreqOffset_Hz(nSiteIndex) = FreqOffset_Hz_temp
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                SumPower(nSiteIndex) = SumPowerTemp
            
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                Select Case TestFreq
                Case 915000000
                    
                    TxPower915.pins("RFHOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                
                    TxPower915.pins("RFOUT").value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select

            End If
            
        Next nSiteIndex
        
        'Reset cpuA flag
        FlagsSet = 0
        FlagsClear = cpuA

        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) 'Pattern continues after cpuA reset.
        
        
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.
  
    'Run patt Iern to stop DUT transmitting
    
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_tx_cw_off").start ("start_tx_cw_off") 'all sites

        'TheHdw.Wait (0.1) 'avoids LVM Priming patgen RTE
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.

 'TheExec.DataLog.WriteComment ("==================== TX915_CW_PWR ===================")
 
    Select Case TestFreq
    
    Case 915000000
    
        TheExec.Flow.TestLimit I_TX915_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_915")
                End If
            Next nSiteIndex
        End If
        
        
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_TX915_CW, 0.065, 0.11, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
'            TheExec.Flow.TestLimit TxPower915, 13, 19, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit I_TX915_CW, 0.064, 0.111, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW_qc", , , , , , , , tlForceNone
'            TheExec.Flow.TestLimit TxPower915, 12.5, 20, , , , unitDb, "%2.1f", "TxPower_915_qc", , , , , , , , tlForceNone
'
'         End If
        
        
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX915_CW, 0.065, 0.11, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, 13, 19, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_915")
                End If
            Next nSiteIndex
        End If
        
    End Select

 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

errHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower915.pins("RFHOUT").value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 915000000
    
        TheExec.Flow.TestLimit I_TX915_CW, 0.065, 0.11, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, 13, 19, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_915")
                End If
            Next nSiteIndex
        End If
            
    Case Else      'Dummy for force fail purpose
    
        TheExec.Flow.TestLimit I_TX915_CW, 0.064, 0.111, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, 12.5, 20, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
        If testXtalOffset Then
            For nSiteIndex = 0 To ExistingSiteCnt - 1
                If TheExec.Sites.site(nSiteIndex).Active = True Then
                    Call sm_LogPassFail(nSiteIndex, FreqOffset_Hz, -100000000, 100000000, "RFHOUT", unitHz, tlForceNone, "FRQOFF_915")
                End If
            Next nSiteIndex
        End If
        
    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2903_tx915_cw = TL_ERROR
    
End Function

Function hc_prime(argc As Long, argv() As String) As Long
    SerialHramData.BusSize = 1                            ' program bus size
    SerialHramData.BusPinNames = "UART_TX"             ' probram bus pins
    SerialHramData.CaptureSize = CInt(argv(0))            ' program number of capture cycles
    SerialHramData.PrimeCaptureSTV                            ' setup the hram for capture
    hc_prime = TL_SUCCESS
End Function

Function hc_RdSerial(argc As Long, argv() As String) As Long

Dim Loop1 As Integer
Dim Loop2 As Integer
Dim cycle As Integer
Dim bitpos As Integer
Dim tempnum As Integer
Dim regnum As Integer
Dim BitsPerWord As Integer
Dim lngCalcCrc As Long
Dim dblCalcCrc As Double


Dim PrintReadback As Boolean
Dim BadCap As Boolean

Dim site As Long

Dim CurrentDate As Variant
Dim CurrentTime As Variant

Dim NumWords As Integer


    BitsPerWord = CInt(argv(0))
    NumWords = CInt(argv(1))


    ' set raw data readback printout according to debug enable word
    If TheExec.EnableWord("DumpRawHram") Then
        PrintReadback = True   ' set to true to print out raw data
    Else
        PrintReadback = False
    End If
    
    ' Readback the Hram data
    If TheHdw.Digital.HRAM.Size >= SerialHramData.CaptureSize Then
        SerialHramData.ReadHRAM PrintReadback
        BadCap = False
    Else
        BadCap = True
    End If

    ' Translate raw data to the Device registers
    If TheExec.Sites.SelectFirst <> loopDone Then
        Do
            site = TheExec.Sites.SelectedSite  'Get the site
            
            cycle = 0
            bitpos = 0
            regnum = 0
            tempnum = 0

            For Loop2 = 0 To NumWords - 1
                For Loop1 = 0 To BitsPerWord - 1
                    tempnum = (SerialHramData.HramData(site, cycle) * 2 ^ (Loop1 Mod (4))) + tempnum
                    If (Loop1 Mod (4)) = 3 Then
                        tempnum = 0
                    End If
                    cycle = cycle + 1
                Next Loop1
            Next Loop2
            

         Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    End If
                   
    hc_RdSerial = TL_SUCCESS
End Function







Public Function cBin2Dec(BinNum As String) As Long
Dim numbits As Integer
Dim i As Integer
Dim asum As Long

    numbits = Len(BinNum)
    asum = 0
    For i = 1 To numbits
        asum = asum + Mid(BinNum, i, 1) * 2 ^ (numbits - i)
    Next
    cBin2Dec = asum
End Function

Public Sub getModuleId(ByVal thisSite As Long, _
    ByRef respRecvd As Boolean, ByRef numChars As Long, _
    ByRef responseStr As String)

    Dim arrayLocation() As Long
    Dim arrayLength() As Long
    Dim arrayBits() As Long
    Dim bitSize As Long
    Dim Words() As String
    Dim nWords As Long
    Dim dataReceivedFlag As Boolean
    Dim xoffset As Long
    Dim xstart As Long
    Dim nDataWords  As Long
    Dim nCorrectDataWords As Long
    Dim percentError As Double
    Dim dataValid As Boolean
    dataValid = True
    
    dataReceivedFlag = True
    ReDim testArray(2552 - 1) As Long
    For bitSize = 0 To 2552 - 1
        If SerialHramData.CaptureSize > bitSize Then
            testArray(bitSize) = SerialHramData.HramData(thisSite, bitSize)
        Else
            If bitSize < 1 Then
                dataReceivedFlag = False
                dataValid = False
                ReDim testArray(0) As Long
                testArray(0) = -1
            Else
                bitSize = bitSize - 1
                ReDim Preserve testArray(bitSize - 1) As Long
                If bitSize < 70 Then
                    dataReceivedFlag = False
                    dataValid = False
                End If
            End If
            Exit For
        End If
    Next bitSize

    Dim nWordsExpected As Long
    nWordsExpected = 40
    
    If dataReceivedFlag Then
        Call getWords(testArray(), nWordsExpected, Words(), nWords, dataValid, thisSite)
    End If
    
    'this cool section of code puts hex into ascii
    responseStr = ""
    Dim xi As Long
    For xi = 0 To nWords - 1
        responseStr = responseStr & ChrW(CInt("&H" & Words(xi)))
    Next xi

    If dataValid Then
        If TheHdw.Digital.Patgen.FailCount = 0 Then
            dataReceivedFlag = False
        ElseIf UBound(Words) < 7 Then
            dataReceivedFlag = False
        Else
            Dim m As Long
            If Words(0) <> "52" Or Words(1) <> "4E" Then
                dataReceivedFlag = False
            End If
        End If
        respRecvd = dataReceivedFlag
    Else
        respRecvd = False
    End If

End Sub

Public Sub getWords(ByRef lcap() As Long, nWordsExpected As Long, ByRef Words() As String, ByRef nWords As Long, _
  ByRef dataValid As Boolean, ByVal thisSite As Long)

    Dim i As Long
    Dim i_start As Long
    Dim dataExists As Boolean
    Dim j As Long, k As Long
    Dim bitCapRaw As Double
    Dim bitCapMajority As Long
    Dim bitsWordCapRawAll() As Long
    Dim bitsWordCapData() As Long
    Dim stopBitFound As Boolean
    Dim StartTransFound As Boolean
    Dim m As Long, n As Long
    Dim EndOfDataReached As Boolean
    Dim iwords As Long
    Dim nBitsRaw As Long, nBitsData As Long
    
    EndOfDataReached = False

    '===================================================================
    '===================================================================
    ' get first start sample (data bit, not start bit)
    dataExists = False
    For i = 0 To 70 'UBound(lcap)
        If lcap(i) = 1 Then
            dataExists = True
            i_start = i Mod 5
            Exit For
        End If
    Next i
    If dataExists = False Then
        'return result fail packet rec
        'Stop
        dataValid = False
        Exit Sub
    End If

    '===================================================================
    '===================================================================
    'get first word
    ReDim Preserve bitsWordCapRawAll(0) As Long
    bitsWordCapRawAll(0) = 0  ' if got to here, then start bit exists, manually add it to array of bits captured
    nBitsRaw = 1                     ' for to next start at 2nd bit
    i = i_start
    For k = nBitsRaw To 9   ' step through first databits (8), and the stop bit (1) = 9 bits or (9*5=40 cap samples)
        bitCapRaw = 0
        For j = 0 To 4
            bitCapRaw = bitCapRaw + lcap(i)
            i = i + 1
        Next j
        bitCapRaw = bitCapRaw / 5
        bitCapMajority = CLng(Round(bitCapRaw, 0))
        ReDim Preserve bitsWordCapRawAll(nBitsRaw) As Long
        bitsWordCapRawAll(nBitsRaw) = bitCapMajority                 ' this is call cap bits (start bit, 8 data, & stop)
        nBitsRaw = nBitsRaw + 1
    Next k
    
    'confirm stop bit found
    stopBitFound = True 'assume yea
    If bitsWordCapRawAll(nBitsRaw - 1) = 0 Then     ' look at stop bit that was just captured, must be 1 or data no longer synced to uart data stream from dut
       ' stopBitFound = False
        'Stop
        'word array will be null/empty
        dataValid = False
        Exit Sub
    End If

        
    '===================================================================
    '===================================================================
    'get all bits, not last word captured may be partial and will be found when stop bit exists, but next start bit not found

    
    nWords = nWordsExpected ' expected number of words    '35, but added buffer
    iwords = 1  '(processed first capture word above)
    
    For iwords = 1 To nWords - 1
    
        '--------------------------------
        ' find start bit
        '-------------------------------
        StartTransFound = False
        stopBitFound = False
        For m = 0 To 9  ' look at cap samples from stop to start, if start not found, then end of uart data
            n = i - 5
            If lcap(n) = 1 Then
                stopBitFound = True
                Exit For
            End If
            n = n + 1
        Next m
        If stopBitFound = False Then
            'Stop  ' this should not occur
            Exit Sub ' return error?
        Else
            For m = m To 9  ' look at cap samples from stop to start, if start not found, then end of uart data
                If lcap(n) = 0 Then
                    StartTransFound = True
                    i_start = n
                    Exit For
                End If
                n = n + 1
            Next m
        End If
        If StartTransFound = False Then
            'Stop  ' this will occur when not more data (all 1s), possible that last word is partial capture
            EndOfDataReached = True
            nWords = iwords - 1
            Exit For
            'Exit Sub ' return error?
        End If
        
        
        '-----------------------------
        ' get captured bits, raw
        '------------------------------
       
        i = i_start
        For k = 0 To 9   ' step through first databits (8), and the stop bit (1) = 9 bits or (9*5=40 cap samples)
            bitCapRaw = 0
            For j = 0 To 4
                bitCapRaw = bitCapRaw + lcap(i)
                i = i + 1
            Next j
            bitCapRaw = bitCapRaw / 5
            bitCapMajority = CLng(Round(bitCapRaw, 0))
            ReDim Preserve bitsWordCapRawAll(nBitsRaw) As Long
            bitsWordCapRawAll(nBitsRaw) = bitCapMajority                 ' this is call cap bits (start bit, 8 data, & stop)
            If (UBound(lcap) - i) < 5 Then
                'Stop
                EndOfDataReached = True
                nWords = iwords - 1
                i = i - 1
                Exit For
            End If
            nBitsRaw = nBitsRaw + 1
        Next k
        
        'confirm stop bit found
        If k = 10 Then
            stopBitFound = True 'assume yea
            If bitsWordCapRawAll(nBitsRaw - 1) = 0 Then     ' look at stop bit that was just captured, must be 1 or data no longer synced to uart data stream from dut
                stopBitFound = False
                'Stop
                Exit Sub
            End If
        Else
            If EndOfDataReached = True Then
                Exit For
            End If
        End If
        
        
        'this cool section of code puts hex into ascii

            
    Next iwords
    
    If iwords < 2 Then  ' nWords = 0
        dataValid = False
        Exit Sub
    End If
    
    If EndOfDataReached = False Then
        'Stop ' this shouldn't occur
        Exit Sub
    End If
    
    ReDim bitsWordCapData(nBitsRaw - 1) As Long
    Dim iBits As Long
    iBits = 0
    For iwords = 0 To nWords + 1 'nWords + 1, for partial
        For i = 1 To 8
            If (iwords * 10 + i) > nBitsRaw Then
                'Stop
                Exit For
                iBits = iBits - 1
            End If
            bitsWordCapData(iBits) = bitsWordCapRawAll(iwords * 10 + i)
            iBits = iBits + 1
        Next i
    Next iwords
    
    Dim PartialBits As Long
    PartialBits = iBits Mod 8
    
    Dim bitsWordsBin() As String
    Dim bitsWordsHex() As String
    Dim bitsWordsBin_lsByte() As String
    Dim bitsWordsBin_msByte() As String
    Dim bitsWordsHex_lsByte() As String
    Dim bitsWordsHex_msByte() As String
   
    'diregards final partial bit, if exists
    ReDim bitsWordsBin(nWords - 1) As String
    ReDim bitsWordsHex(nWords - 1) As String
    ReDim bitsWordsBin_lsByte(nWords - 1) As String
    ReDim bitsWordsBin_msByte(nWords - 1) As String
    ReDim bitsWordsHex_lsByte(nWords - 1) As String
    ReDim bitsWordsHex_msByte(nWords - 1) As String
    ReDim Words(nWords - 1) As String
    iBits = 0
    For iwords = 0 To nWords - 1
        bitsWordsBin(iwords) = ""
        bitsWordsHex_lsByte(iwords) = ""
        bitsWordsHex_msByte(iwords) = ""
        For i = 7 To 0 Step -1
            If i > 3 Then
                bitsWordsBin_lsByte(iwords) = bitsWordsBin_lsByte(iwords) & CStr(bitsWordCapData(iwords * 8 + i))
            Else
                bitsWordsBin_msByte(iwords) = bitsWordsBin_msByte(iwords) & CStr(bitsWordCapData(iwords * 8 + i))
            End If
            iBits = iBits + 1
        Next i
        bitsWordsBin(iwords) = bitsWordsBin_lsByte(iwords) & bitsWordsBin_msByte(iwords)
        bitsWordsHex_lsByte(iwords) = Hex(cBin2Dec(bitsWordsBin_lsByte(iwords)))
        bitsWordsHex_msByte(iwords) = Hex(cBin2Dec(bitsWordsBin_msByte(iwords)))
        bitsWordsHex(iwords) = bitsWordsHex_lsByte(iwords) & bitsWordsHex_msByte(iwords)
        Words(iwords) = bitsWordsHex(iwords)
        
        If False Then
        'radio_rx  3C3C3C6D696C726F636869702D6C6F72613E3E3E
        Dim responseStr As String
        If iwords = 0 Then responseStr = ""
        Dim xi As Long
            responseStr = responseStr & ChrW(CInt("&H" & Words(iwords)))
        End If
        
    Next iwords
    
    If False Then
        Debug.Print "  dut response site(" & CStr(thisSite) & ") = " & responseStr
    End If
Exit Sub

End Sub


Public Sub getPktErr(ByVal thisSite As Long, ByRef pktRecvd As Boolean, ByRef pktErr As Double)

    Dim arrayLocation() As Long
    Dim arrayLength() As Long
    Dim arrayBits() As Long
    Dim bitSize As Long
    Dim Words() As String
    Dim nWords As Long
    Dim expectedResponse As Variant, expectedData As Variant
    Dim packetReceivedFlag As Boolean
    Dim xoffset As Long
    Dim xstart As Long
    Dim nDataWords  As Long
    Dim nCorrectDataWords As Long
    Dim percentError As Double
    Dim dataValid As Boolean
    dataValid = True
    'here is the data, what it should be... the first is the dut saying it has the RX data.
    expectedResponse = Array("72", "61", "64", "69", "6F", "5F", "72", "78", "20", "20")
    expectedData = Array("33", "43", "33", "43", "33", "43", "36", "44", "36", "39", "36", "43", "37", "32", "36", "46", "36", "33", "36", "38", "36", "39", "37", "30", "32", "44", "36", "43", "36", "46", "37", "32", "36", "31", "33", "45", "33", "45", "33", "45", "0D", "0A")
    ReDim testArray(2552 - 1) As Long
    For bitSize = 0 To 2552 - 1
        testArray(bitSize) = SerialHramData.HramData(thisSite, bitSize)
    Next bitSize

    If TheHdw.Digital.Patgen.FailCount > 65500 Then
        pktRecvd = False
        pktErr = -1
    Else
    
        Dim nWordsExpected As Long
        nWordsExpected = 52
        
        If packetReceivedFlag Or True Then
            Call getWords(testArray(), nWordsExpected, Words(), nWords, dataValid, thisSite)
            
            If False Then
        
                Dim iwords As Long
                For iwords = 0 To nWords - 1
                'radio_rx  3C3C3C6D696C726F636869702D6C6F72613E3E3E
                Dim responseStr As String
                If iwords = 0 Then responseStr = ""
                Dim xi As Long
                    responseStr = responseStr & ChrW(CInt("&H" & Words(iwords)))
           
                Next iwords
    
                Debug.Print "  dut response site(" & CStr(thisSite) & ") = " & responseStr
            End If
            
        End If
        
       If dataValid Then packetReceivedFlag = True
        
        If dataValid Then
            If TheHdw.Digital.Patgen.FailCount = 0 Then
                packetReceivedFlag = False
            ElseIf Len(Join(Words)) > 0 Then
                If UBound(Words) < 7 Then
                    packetReceivedFlag = False
                Else
                    Dim m As Long
                    For m = 0 To 7
                        If Words(m) <> expectedResponse(m) Then
                            packetReceivedFlag = False
                        End If
                    Next m
                End If
            Else
                'Stop
                packetReceivedFlag = False
                Exit Sub
            End If
            If packetReceivedFlag = True Then
                If Words(9) = "20" Then
                    xoffset = 10
                Else
                    xoffset = 9
                End If
                nDataWords = nWords - xoffset
                nCorrectDataWords = 0
                For m = 0 To nDataWords - 1
                    If Words(m + xoffset) = expectedData(m) Then
                        nCorrectDataWords = nCorrectDataWords + 1
                    End If
                Next m
                percentError = 100# * (1 - (nCorrectDataWords / nDataWords))
            End If
    
    
    
            pktRecvd = packetReceivedFlag
            pktErr = percentError
    
        Else
            pktRecvd = False
            pktErr = -1  'neg 1 means the signal wasn't recvd so no PER calculated... my menthod for info only.
        End If
    End If
    
End Sub


Public Function rn2483_gpio(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'The LoRa module has 14 gpio pins, GPIO0 - GPIO13. Each gpio pin is commanded via the UART to be set to a logical 1, tested in the pattern, then set to logical 0, then tested in the pattern.

    Dim site As Variant
    
      'Dim oprVolt As Double
      'Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2483_gpio = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
TheExec.DataLog.WriteComment ("==================== GPIO CHECK ======================")
    
        'oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        'dut_delay = 0.1
            
                      'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
  
  
  Call TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_gpio_full").Test(pfAlways, 0)
        
        
        TheHdw.wait (0.3)
        
    
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_gpio")
    Call TheExec.ErrorReport
    rn2483_gpio = TL_ERROR
    
End Function

Public Function rn2903_gpio(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'The LoRa module has 14 gpio pins, GPIO0 - GPIO13. Each gpio pin is commanded via the UART to be set to a logical 1, tested in the pattern, then set to logical 0, then tested in the pattern.

    Dim site As Variant
    
      'Dim oprVolt As Double
      'Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2903_gpio = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
'TheExec.DataLog.WriteComment ("===================== GPIO CHECK ====================")
    
        'oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        'dut_delay = 0.1
            
                      'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
  
  
  Call TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_gpio_full").Test(pfAlways, 0)
        
        
        TheHdw.wait (0.3)
        
    
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_gpio")
    Call TheExec.ErrorReport
    rn2903_gpio = TL_ERROR
    
End Function



