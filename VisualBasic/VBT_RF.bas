Attribute VB_Name = "VBT_RF"
Option Explicit

#Const Connected_to_J750 = True
#Const Connected_to_AXRF = True
#Const Debug_advanced = True

Public FIRSTLOAD As Boolean


Global Const DUTRefClkFreq As Double = 13560000# 'NOT USED FOR LoRa MODULES!

Public tx_path_db(7) As Double      'Dim for number of sites
Public coax_cable_db(7) As Double   'Dim for number of sites



Public Function read_cal_factors() As Long              '(argc As Long, argv() As String) As Long

    'Public tx_path_db() As Double
    'Public rx_path_db() As Double
    'Public coax_cable_db() As Double
    
    'This function reads the RF_Cal_Factors worksheet to initialize global RF scalar calibration offsets. AXRF calibration is
    'performed with same cables and RF interface as the J750_AERO with Reid-Asman junction box interface and Pasternak coaxial cables.
    
    Dim nSiteIndex As Long
    
    On Error GoTo errHandler
    
     
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
    
        If TheExec.Sites.Site(nSiteIndex).Active = True Then
        
            tx_path_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(3, 2 + nSiteIndex).Value       'DIB TX path loss
                           
            coax_cable_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(4, 2 + nSiteIndex).Value    'Cable loss from PXI to DIB
        
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


Public Function rn2483_tx868_cw(argc As Long, argv() As String) As Long

    Dim Site As Variant
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
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        
    Case Is = 2
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        
    Case Is = 3
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        
    Case Is = 4
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
        
        
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
    


    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB
    
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(8192) 'Fres = 30.5176 kHz  (Fs = 250MHz, N=8192) NOTE: Fres chosen to bound TX freq
  
        TheHdw.Wait 0.05
        
    Select Case TestFreq
    Case 868300000
        TxPower868.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        

    
    Case Else
        TxPower868.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFHOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        
    End Select


    TheHdw.Wait (0.002)
    
        
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
        TheHdw.DPS.Samples = 1
    End With
    
    
        'Measure Current
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX868_CW)
    

        'Measure RF Power
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1 'Site loop needed for AXRF

            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(nSiteIndex), 17, TestFreq  'set for +17dBm
                
                TheHdw.Wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 4096, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", False, False, False, 1, SumPowerTemp) 'True plots waveform
                
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                SumPower(nSiteIndex) = SumPowerTemp
            
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                
                Select Case TestFreq
                Case 868300000
                    
                    TxPower868.pins("RFHOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                    TxPower868.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
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

 TheExec.DataLog.WriteComment ("=================== TX868_CW_PWR =====================")
 
    Select Case TestFreq
    
    Case 868300000
    
                    TheExec.Flow.TestLimit I_TX868_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
                    TheExec.Flow.TestLimit TxPower868, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
                
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

    End Select

 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

errHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower868.pins("RFHOUT").Value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 868300000
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone
  
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, "%2.2f", "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, "%2.1f", "TxPower_868", , , , , , , , tlForceNone

    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2483_tx868_cw = TL_ERROR
    
End Function





Public Function rn2483_i_sleep(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for 2 sec via a UART command. During the 2 second window the VBAT current is measured and reported.

    Dim Site As Variant
    
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
  
            I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value

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
            TheHdw.DPS.Samples = 1
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
      
    I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2483_i_sleep = TL_ERROR
    
    
End Function



Public Function rn2483_id_mfs(argc As Long, argv() As String) As Long

'Multisite LoRa module ID test. After a reset, a functional DUT sends the UART host its ID (and FW revision time and date).

'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received ID. If no ID is received, however,
'the pattern will time out with 100 forced fails.

'Some modules take more than 100msec to respond to system reset than others. A fast and slow response pattern is used to check which type of module
' is being tested. Passing the slow OR the fast pattern will pass the ID test by finding the start bit of the response.


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
    
    Dim ID_Valid As New PinListData
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
 On Error GoTo errHandler
 
    rn2483_id_mfs = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
     
    

    ID_Valid.AddPin ("RFHOUT")
    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
        
        ID_Valid.pins("RFHOUT").Value(nSiteIndex) = 0
        
           ValidityCountFast(nSiteIndex) = 0
           ValidityCountSlow(nSiteIndex) = 0
           
           patgen_fails_fast(nSiteIndex) = 0
           patgen_fails_slow(nSiteIndex) = 0
           
           
        
    Next nSiteIndex
    
'    If Right(ActiveWorkbook.Path, 1) = "\" Then
'        xTPPath = ActiveWorkbook.Path
'    Else
'        xTPPath = ActiveWorkbook.Path & "\"
'    End If
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
'TheHdw.Wait (0.2) 'Wait for DUT POR to complete.

Call TheHdw.Digital.Patgen.HaltWait
                    
TheHdw.Digital.Patgen.ThreadingForActiveSites = False

'Serial Loop

loopstatus = TheExec.Sites.SelectFirst
                     
                        
While loopstatus <> loopDone

      If TheExec.Sites.Site(0).Active Then
      
        TheExec.Sites.SetOverride (0)
      
    'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                patgen_fails_slow(0) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 0"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_fast.pat").Run ("start_uart_id_fast")
            
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                patgen_fails_fast(0) = TheHdw.Digital.Patgen.FailCount
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(0) = 0 Or patgen_fails_slow(0) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(0) = 0
                    ValidityCountSlow(0) = 0
                
                Else
                       
                    ValidityCountFast(0) = Int(55 - patgen_fails_fast(0))
                       
                    ValidityCountSlow(0) = Int(55 - patgen_fails_slow(0))
                    
                End If
                            
                            
                If (ValidityCountFast(0) > 0) Or (ValidityCountSlow(0) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(0) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(0) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      End If 'Site 0 active
      
      
      If TheExec.Sites.Site(1).Active Then
      
        TheExec.Sites.SetOverride (1)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(1) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 1"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(1) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(1) = 0 Or patgen_fails_slow(1) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(1) = 0
                    ValidityCountSlow(1) = 0
                
                Else
                       
                    ValidityCountFast(1) = Int(55 - patgen_fails_fast(1))
                       
                    ValidityCountSlow(1) = Int(55 - patgen_fails_slow(1))
                
                End If
                            
                            
                If (ValidityCountFast(1) > 0) Or (ValidityCountSlow(1) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(1) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(1) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 1 Active
      
      If TheExec.Sites.Site(2).Active Then
      
        TheExec.Sites.SetOverride (2)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(2) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 2"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(2) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(2) = 0 Or patgen_fails_slow(2) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(2) = 0
                    ValidityCountSlow(2) = 0
                
                Else
                       
                    ValidityCountFast(2) = Int(55 - patgen_fails_fast(2))
                       
                    ValidityCountSlow(2) = Int(55 - patgen_fails_slow(2))
                
                End If
                            
                            
                If (ValidityCountFast(2) > 0) Or (ValidityCountSlow(2) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(2) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(2) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 2 Active

      If TheExec.Sites.Site(3).Active Then
      
        TheExec.Sites.SetOverride (3)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(3) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 3"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(3) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(3) = 0 Or patgen_fails_slow(3) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(3) = 0
                    ValidityCountSlow(3) = 0
                
                Else
                       
                    ValidityCountFast(3) = Int(55 - patgen_fails_fast(3))
                       
                    ValidityCountSlow(3) = Int(55 - patgen_fails_slow(3))
                
                End If
                            
                            
                If (ValidityCountFast(3) > 0) Or (ValidityCountSlow(3) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(3) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(3) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 3 Active
      
     loopstatus = TheExec.Sites.SelectNext(loopstatus)
    
Wend 'end WHILE loop
  
  'End Serial Loop

    Call TheHdw.Digital.Patgen.Halt
    

    TheExec.DataLog.WriteComment ("==================  READ_MODULE_ID  =================== ")
    
    TheExec.Flow.TestLimit ID_Valid, LoLimit, HiLimit, , , , unitNone, "%2.1f", "ID", , , , , , , , tlForceNone
    
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit ID_Valid, 0.5, 1.5, , , , unitNone, "%2.1f", "ID", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit ID_Valid, 0.5, 1.5, , , , unitNone, "%2.1f", "ID", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit ID_Valid, 0.4, 1.6, , , , unitNone, "%2.1f", "ID_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
   
        Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    Call TheHdw.Digital.Patgen.Halt
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_id_mfs")
    Call TheExec.ErrorReport
    rn2483_id_mfs = TL_ERROR
    
End Function



Public Function rn2483_fsk_pkt_rcv_m_rev(argc As Long, argv() As String) As Long

'Multisite Packets Received Test using digital channel triggering of 3025C Modulation Source. This function is NOT a PER test! When CRC is Off in the module radio driver,
'assuming received signal strength is sufficient to recognize PREAMBLE and SYNC, packet data will be sent from DUT FIFO to the host on UART_TX.
'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received packets. However,
'when the DUT cannot receive the packet, no FIFO data will be sent to the UART host, and the pattern will time out with 100 forced fails.
'The MATCH LOOP does not work when a packet loop count (>1) is used inside the pattern, therefore a VBT FOR LOOP is coded for multiple packets.
' The AXRF Modulation Source is triggered by the digital pattern using MW_SRC_TRIG. Note: TDR calibration MUST be performed, or else this test will FAIL!
'For TTR, only one packet is sent to each device.

    Dim SrcChans(3) As AXRF_CHANNEL

    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    
    Dim Gate As Long
    Dim Edge As Long
    Dim nSiteIndex As Long
    
    Dim fails As Long
    Dim pkt_fails As Long
    
    Dim siteStatus As Long
    Dim thisSite As Long

    Dim PacketCount(3) As Long
    Dim patgen_fails(3) As Long
    
    Dim i As Long
    Dim pkt_sent_count As Long
    
    Dim PKTs_RCVd As New PinListData
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
 On Error GoTo errHandler
 
    rn2483_fsk_pkt_rcv_m_rev = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
        
        PKTs_RCVd.AddPin ("RFHOUT")
        
        
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
           
        PKTs_RCVd.pins("RFHOUT").Value(nSiteIndex) = 0
        
        patgen_fails(nSiteIndex) = 0
        
    Next nSiteIndex
     

        Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB

    
'    If Right(ActiveWorkbook.Path, 1) = "\" Then
'        xTPPath = ActiveWorkbook.Path
'    Else
'        xTPPath = ActiveWorkbook.Path & "\"
'    End If
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_ccitt.aiq"  'NFG      'Non-inverted CRC, MSB/LSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_swap_ccitt.aiq"       'Non-inverted CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_swap_ccitt.aiq"              'Inverted CCITT CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"           'Inverted CCITT CRC, MSB/LSB
    
    ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"
    
    
    SrcChans(0) = AXRF_CHANNEL_AXRF_CH1
    SrcChans(1) = AXRF_CHANNEL_AXRF_CH3
    SrcChans(2) = AXRF_CHANNEL_AXRF_CH5
    SrcChans(3) = AXRF_CHANNEL_AXRF_CH7
   
    'AXRF Modulation Trigger Arm Source Parameters

     Gate = 1  '1 = Modulation is ON for duration of HIGH trigger, or when the modulation ends; 0 =  Modulation starts w/Trig and runs continuously
    
     Edge = 0    'Positive edge
     
Call TheHdw.Digital.Patgen.HaltWait

'TheHdw.Digital.Patgen.Threading = False
                    
TheHdw.Digital.Patgen.ThreadingForActiveSites = False

'Serial Site Loop

TheExec.DataLog.WriteComment ("===================== PKTs_RCVd ======================")

' Loop through all the active sites
With TheExec.Sites
  siteStatus = .SelectFirst

Do While siteStatus <> loopDone
    thisSite = .SelectedSite

      If (TheExec.Sites.Site(0).Active) Then
      
        TheExec.Sites.SetOverride (0)

            'Working Hardware Triggered Modulation Code
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(0), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(0), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(0), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            'Setting is ~2 dB above highest passing threshhold for functional DUTS.
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(0), -85, 868300000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
 
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
                     
'                    With TheHdw.pins("UART_CTS,MW_SRC_TRIG") 'Initialize Logic Analyzer Trigger
'                        .InitState = chInitLo
'                        .StartState = chStartLo
'                    End With
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
        
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
        
                        'Debug.Print "Site = 0"
'                        Debug.Print "Pkts = "; pkt_sent_count
'
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                  If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(0) = 0
                        
                   Else
                   
                        patgen_fails(0) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(0) > 0 Then
            
                             PKTs_RCVd.AddPin("RFHOUT").Value(0) = PacketCount(0) + 1
                        
                        Else
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(0) = PacketCount(0)
                        
                        End If
                   
                   End If
                
        
            'Next pkt_sent_count
            
     'Datalog site here
     
     TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
               
        TheExec.Sites.RestoreFromOverride
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
     
    End If 'Site 0 Active
    

    
    
    If siteStatus = loopDone Then Exit Do
    
    
    If TheExec.Sites.Site(1).Active Then
      
        TheExec.Sites.SetOverride (1)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(1), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(1), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(1), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(1), -85, 868300000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 1"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(1) = 0
                        
                    Else
                   
                        patgen_fails(1) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(1) > 0 Then  'And pkt_sent_count > 0
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(1) = PacketCount(1) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(1) = PacketCount(1)
                            
                        End If
                        
                     End If
        
            'Next pkt_sent_count
            
        
        'Datalog site here
        TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
        
        TheExec.Sites.RestoreFromOverride
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 1 Active
    
    
    
    
    If siteStatus = loopDone Then Exit Do
    
    If TheExec.Sites.Site(2).Active Then
      
        TheExec.Sites.SetOverride (2)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(2), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(2), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(2), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(2), -85, 868300000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
                     
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 2"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(2) = 0
                        
                    Else
                   
                        patgen_fails(2) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(2) > 0 Then 'And pkt_sent_count > 0 Then
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(2) = PacketCount(2) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(2) = PacketCount(2)
                            
                        End If
                        
                   End If
                        
        
            'Next pkt_sent_count
            
        
        'Datalog site here
        TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
        
        TheExec.Sites.RestoreFromOverride
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 2 Active

    
    
    
    If siteStatus = loopDone Then Exit Do
    
   If TheExec.Sites.Site(3).Active Then
      
        TheExec.Sites.SetOverride (3)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(3), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(3), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(3), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(3), -85, 868300000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 3"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(3) = 0
                        
                    Else
                   
                        patgen_fails(3) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(3) > 0 Then  'And pkt_sent_count > 0 Then
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(3) = PacketCount(3) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(3) = PacketCount(3)
                            
                        End If
                        
                    End If
        
            'Next pkt_sent_count
            
        
        'Datalog site here
        TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        Select Case TheExec.CurrentJob
'            Case "f1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "f1-pgm-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'            Case "q1-prd-std-rn2483"
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'            Case Else
'
'        End Select
        
        TheExec.Sites.RestoreFromOverride
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 3 Active

    
    
    
    If siteStatus = loopDone Then Exit Do
    
    If (TheExec.Sites.ActiveCount = 0 Or TheExec.Sites.ActiveCount >= 3) Then Exit Do
        siteStatus = .SelectNext(loopTop)
    
    
Loop
  
End With ' TheExec.Sites

'End Serial Site Loop

    Call TheHdw.Digital.Patgen.Halt
    
    'Cleanup

    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
        If TheExec.Sites.Site(nSiteIndex).Active Then
            Call itl.Raw.AF.AXRF.StopModulation(SrcChans(nSiteIndex))
            Call itl.Raw.AF.AXRF.UnloadModulationFile(SrcChans(nSiteIndex), ModFilePath)
        End If
        
    Next nSiteIndex

        Call SetAXRFinRxMode(SrcChans, -120, 868300000#) 'turn off RF source
    
        Call TheHdw.Digital.Patgen.Halt
    
        Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_fsk_pkt_rcv_m_rev")
    Call TheExec.ErrorReport
    rn2483_fsk_pkt_rcv_m_rev = TL_ERROR
    
End Function


Public Function rn2483_idle_current(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'Reset times vary for the modules, so two attempts are allowed to measure the idle current after reset.

    Dim Site As Variant
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
    
TheExec.DataLog.WriteComment ("=================== MEASURE I_IDLE ===================")
    
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
    
        TheHdw.Wait (0.05)
        
        
                    'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
            
         TheHdw.Wait (0.05)
         
                     'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
               
        
        I_IDLE.AddPin ("VBAT")
        
'  For nSiteIndex = 0 To ExistingSiteCnt - 1
'
'    I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
'
'  Next nSiteIndex
        
        TheHdw.Wait (0.3)
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
     If TheExec.Sites.Site(nSiteIndex).Active Then
        
        
        
        I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps100mA
            .CurrentLimit = 0.1
            TheHdw.DPS.Samples = 1
            Call .MeasureCurrents(dps100mA, I_IDLE)
        End With
        
   
    
            If I_IDLE.pins("VBAT").Value(nSiteIndex) > 0.007 Then
            
            TheHdw.Wait (0.3)
            
                 With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.Samples = 1
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
      
        I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2483", , , , , , , , tlForceNone

    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_idle_current")
    Call TheExec.ErrorReport
    rn2483_idle_current = TL_ERROR
    
End Function

Public Function rn2903_fsk_pkt_rcv_m_rev(argc As Long, argv() As String) As Long

'Multisite Packets Received Test using digital channel triggering of 3025C Modulation Source. This function is NOT a PER test! When CRC is Off in the module radio driver,
'assuming received signal strength is sufficient to recognize PREAMBLE and SYNC, packet data will be sent from DUT FIFO to the host on UART_TX.
'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received packets. However,
'when the DUT cannot receive the packet, no FIFO data will be sent to the UART host, and the pattern will time out with 100 forced fails.
'The MATCH LOOP does not work when a packet loop count (>1) is used inside the pattern, therefore a VBT FOR LOOP is coded for multiple packets.
' The AXRF Modulation Source is triggered by the digital pattern using MW_SRC_TRIG. Note: TDR calibration MUST be performed, or else this test will FAIL!
'For TTR, only one packet is sent to each device.

    Dim SrcChans(3) As AXRF_CHANNEL

    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    
    Dim Gate As Long
    Dim Edge As Long
    Dim nSiteIndex As Long
    
    Dim fails As Long
    Dim pkt_fails As Long
    
    Dim siteStatus As Long
    Dim thisSite As Long

    Dim PacketCount(3) As Long
    Dim patgen_fails(3) As Long
    
    Dim i As Long
    Dim pkt_sent_count As Long
    
    Dim PKTs_RCVd As New PinListData
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
 On Error GoTo errHandler
 
    rn2903_fsk_pkt_rcv_m_rev = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
        
        PKTs_RCVd.AddPin ("RFHOUT")
        
        
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
           
        PKTs_RCVd.pins("RFHOUT").Value(nSiteIndex) = 0
        
        patgen_fails(nSiteIndex) = 0
        
    Next nSiteIndex
     

        Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB

    
'    If Right(ActiveWorkbook.Path, 1) = "\" Then
'        xTPPath = ActiveWorkbook.Path
'    Else
'        xTPPath = ActiveWorkbook.Path & "\"
'    End If
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_ccitt.aiq"  'NFG      'Non-inverted CRC, MSB/LSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_noninv_swap_ccitt.aiq"       'Non-inverted CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_swap_ccitt.aiq"              'Inverted CCITT CRC, LSB/MSB
    'ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"           'Inverted CCITT CRC, MSB/LSB
    
    ModFilePath = xTPPath & "\patterns\fsk_31b_ccitt_gauss03.aiq"
    
    
    SrcChans(0) = AXRF_CHANNEL_AXRF_CH1
    SrcChans(1) = AXRF_CHANNEL_AXRF_CH3
    SrcChans(2) = AXRF_CHANNEL_AXRF_CH5
    SrcChans(3) = AXRF_CHANNEL_AXRF_CH7
   
    'AXRF Modulation Trigger Arm Source Parameters

     Gate = 1  '1 = Modulation is ON for duration of HIGH trigger, or when the modulation ends; 0 =  Modulation starts w/Trig and runs continuously
    
     Edge = 0    'Positive edge
     
Call TheHdw.Digital.Patgen.HaltWait

'TheHdw.Digital.Patgen.Threading = False
                    
TheHdw.Digital.Patgen.ThreadingForActiveSites = False

'Serial Site Loop

TheExec.DataLog.WriteComment ("====================== PKTs_RCVd ====================")

' Loop through all the active sites
With TheExec.Sites
  siteStatus = .SelectFirst

Do While siteStatus <> loopDone
    thisSite = .SelectedSite

      If (TheExec.Sites.Site(0).Active) Then
      
        TheExec.Sites.SetOverride (0)

            'Working Hardware Triggered Modulation Code
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(0), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(0), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(0), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            'Setting is ~2 dB above highest passing threshhold for functional DUTS.
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(0), -85, 915000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
 
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
                     
'                    With TheHdw.pins("UART_CTS,MW_SRC_TRIG") 'Initialize Logic Analyzer Trigger
'                        .InitState = chInitLo
'                        .StartState = chStartLo
'                    End With
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
        
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_tx915_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
        
                        'Debug.Print "Site = 0"
'                        Debug.Print "Pkts = "; pkt_sent_count
'
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                  If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(0) = 0
                        
                   Else
                   
                        patgen_fails(0) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(0) > 0 Then
            
                             PKTs_RCVd.AddPin("RFHOUT").Value(0) = PacketCount(0) + 1
                        
                        Else
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(0) = PacketCount(0)
                        
                        End If
                   
                   End If
                
        
            'Next pkt_sent_count
            
     'Datalog site here
     
    TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'             TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'        End If
                
        TheExec.Sites.RestoreFromOverride
        
    siteStatus = TheExec.Sites.SelectNext(siteStatus)
        
     
    End If 'Site 0 Active
    

    'siteStatus = TheExec.Sites.SelectNext(siteStatus)  nueng
    
    If siteStatus = loopDone Then Exit Do
    
    
    If TheExec.Sites.Site(1).Active Then
      
        TheExec.Sites.SetOverride (1)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(1), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(1), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(1), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(1), -85, 915000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_tx915_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 1"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(1) = 0
                        
                    Else
                   
                        patgen_fails(1) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(1) > 0 Then  'And pkt_sent_count > 0
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(1) = PacketCount(1) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(1) = PacketCount(1)
                            
                        End If
                        
                     End If
        
            'Next pkt_sent_count
            
        
        'Datalog site here
        
    TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'        End If
        
        TheExec.Sites.RestoreFromOverride
        
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 1 Active
    
    
    'siteStatus = TheExec.Sites.SelectNext(siteStatus)  nueng
    
    If siteStatus = loopDone Then Exit Do
    
    If TheExec.Sites.Site(2).Active Then
      
        TheExec.Sites.SetOverride (2)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(2), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(2), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(2), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(2), -85, 915000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
                     
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_tx915_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 2"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(2) = 0
                        
                    Else
                   
                        patgen_fails(2) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(2) > 0 Then 'And pkt_sent_count > 0 Then
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(2) = PacketCount(2) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(2) = PacketCount(2)
                            
                        End If
                        
                   End If
                        
        
            'Next pkt_sent_count
            
        
        'Datalog site here
                    
    TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'        End If
        
        TheExec.Sites.RestoreFromOverride
        
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 2 Active

    
    'siteStatus = TheExec.Sites.SelectNext(siteStatus) nueng
    
    If siteStatus = loopDone Then Exit Do
    
   If TheExec.Sites.Site(3).Active Then
      
        TheExec.Sites.SetOverride (3)
    
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(3), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(3), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(3), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(3), -85, 915000000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
        
            'For pkt_sent_count = 1 To 5 '5 packets sent
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_tx915_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                        'Debug.Print "Site = 3"
                        'Debug.Print "Pkts = "; pkt_sent_count
                    
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        'Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                    If TheHdw.Digital.Patgen.FailCount = 0 Then  'Open socket workaround
                   
                        patgen_fails(3) = 0
                        
                    Else
                   
                        patgen_fails(3) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(3) > 0 Then  'And pkt_sent_count > 0 Then
                        
                             PKTs_RCVd.AddPin("RFHOUT").Value(3) = PacketCount(3) + 1
                        
                        Else
                        
                            PKTs_RCVd.AddPin("RFHOUT").Value(3) = PacketCount(3)
                            
                        End If
                        
                    End If
        
            'Next pkt_sent_count
            
        
        'Datalog site here
        
    TheExec.Flow.TestLimit PKTs_RCVd, LoLimit, HiLimit, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
        
'        If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.5, 1.5, , , , unitNone, "%2.1f", "PktCnt", , , , , , , , tlForceNone
'
'        ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'            TheExec.Flow.TestLimit PKTs_RCVd, 0.4, 1.6, , , , unitNone, "%2.1f", "PktCnt_qc", , , , , , , , tlForceNone
'
'        End If
        
        TheExec.Sites.RestoreFromOverride
        siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    End If 'Site 3 Active

    
    'siteStatus = TheExec.Sites.SelectNext(siteStatus)
    
    'If siteStatus = loopDone Then Exit Do
    
    'If (TheExec.Sites.ActiveCount = 0 Or TheExec.Sites.ActiveCount >= 3) Then Exit Do
        'siteStatus = .SelectNext(loopTop)
    
    
Loop
  
End With ' TheExec.Sites

'End Serial Site Loop

    Call TheHdw.Digital.Patgen.Halt
    
    'Cleanup

    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
        If TheExec.Sites.Site(nSiteIndex).Active Then
            Call itl.Raw.AF.AXRF.StopModulation(SrcChans(nSiteIndex))
            Call itl.Raw.AF.AXRF.UnloadModulationFile(SrcChans(nSiteIndex), ModFilePath)
        End If
        
    Next nSiteIndex

        Call SetAXRFinRxMode(SrcChans, -120, 915000000#) 'turn off RF source
    
        Call TheHdw.Digital.Patgen.Halt
    
        Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_fsk_pkt_rcv_m_rev")
    Call TheExec.ErrorReport
    rn2903_fsk_pkt_rcv_m_rev = TL_ERROR
    
End Function

Public Function rn2903_i_sleep(argc As Long, argv() As String) As Long

'The DUT is commanded to sleep for 2 sec via a UART command. During the 2 second window the VBAT current is measured and reported.

    Dim Site As Variant
    
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
  
            I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value

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
            TheHdw.DPS.Samples = 1
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
      
    I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, 0.00005, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2903_I_SLEEP", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2903_i_sleep = TL_ERROR
    
    
End Function

Public Function rn2903_id_mfs(argc As Long, argv() As String) As Long

'Multisite LoRa module ID test. After a reset, a functional DUT sends the UART host its ID (and FW revision time and date).

'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received ID. If no ID is received, however,
'the pattern will time out with 100 forced fails.

'Some modules take more than 100msec to respond to system reset than others. A fast and slow response pattern is used to check which type of module
' is being tested. Passing the slow OR the fast pattern will pass the ID test by finding the start bit of the response.


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
    
    Dim ID_Valid As New PinListData
    
    Dim LoLimit As Double
    Dim HiLimit As Double
    '--------Argument processing--------'
    LoLimit = argv(1)
    HiLimit = argv(2)
    '------- end of argument process -------'
    
 On Error GoTo errHandler
 
    rn2903_id_mfs = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
     
    

    ID_Valid.AddPin ("RFHOUT")
    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid variables
        
        ID_Valid.pins("RFHOUT").Value(nSiteIndex) = 0
        
           ValidityCountFast(nSiteIndex) = 0
           ValidityCountSlow(nSiteIndex) = 0
           
           patgen_fails_fast(nSiteIndex) = 0
           patgen_fails_slow(nSiteIndex) = 0
           
           
        
    Next nSiteIndex
    
'    If Right(ActiveWorkbook.Path, 1) = "\" Then
'        xTPPath = ActiveWorkbook.Path
'    Else
'        xTPPath = ActiveWorkbook.Path & "\"
'    End If
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
'TheHdw.Wait (0.2) 'Wait for DUT POR to complete.

Call TheHdw.Digital.Patgen.HaltWait
                    
TheHdw.Digital.Patgen.ThreadingForActiveSites = False

'Serial Loop

loopstatus = TheExec.Sites.SelectFirst
                     
                        
While loopstatus <> loopDone

      If TheExec.Sites.Site(0).Active Then
      
        TheExec.Sites.SetOverride (0)
      
    'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                patgen_fails_slow(0) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 0"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_fast.pat").Run ("start_uart_id_fast")
            
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                patgen_fails_fast(0) = TheHdw.Digital.Patgen.FailCount
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(0) = 0 Or patgen_fails_slow(0) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(0) = 0
                    ValidityCountSlow(0) = 0
                
                Else
                       
                    ValidityCountFast(0) = Int(55 - patgen_fails_fast(0))
                       
                    ValidityCountSlow(0) = Int(55 - patgen_fails_slow(0))
                    
                End If
                            
                            
                If (ValidityCountFast(0) > 0) Or (ValidityCountSlow(0) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(0) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(0) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      End If 'Site 0 active
      
      
      If TheExec.Sites.Site(1).Active Then
      
        TheExec.Sites.SetOverride (1)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(1) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 1"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(1) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(1) = 0 Or patgen_fails_slow(1) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(1) = 0
                    ValidityCountSlow(1) = 0
                
                Else
                       
                    ValidityCountFast(1) = Int(55 - patgen_fails_fast(1))
                       
                    ValidityCountSlow(1) = Int(55 - patgen_fails_slow(1))
                
                End If
                            
                            
                If (ValidityCountFast(1) > 0) Or (ValidityCountSlow(1) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(1) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(1) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 1 Active
      
      If TheExec.Sites.Site(2).Active Then
      
        TheExec.Sites.SetOverride (2)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(2) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 2"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(2) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(2) = 0 Or patgen_fails_slow(2) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(2) = 0
                    ValidityCountSlow(2) = 0
                
                Else
                       
                    ValidityCountFast(2) = Int(55 - patgen_fails_fast(2))
                       
                    ValidityCountSlow(2) = Int(55 - patgen_fails_slow(2))
                
                End If
                            
                            
                If (ValidityCountFast(2) > 0) Or (ValidityCountSlow(2) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(2) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(2) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 2 Active

      If TheExec.Sites.Site(3).Active Then
      
        TheExec.Sites.SetOverride (3)
      
      'Site specific code here
                    'Run ID patterns
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_slow.pat").Run ("start_uart_id_slow")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                    
                    patgen_fails_slow(3) = TheHdw.Digital.Patgen.FailCount
                        
                        Debug.Print "Site = 3"
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                    
            TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2903_id_fast.pat").Run ("start_uart_id_fast")
                    
            Call TheHdw.Digital.Patgen.HaltWait
                  
                    patgen_fails_fast(3) = TheHdw.Digital.Patgen.FailCount
                        
                        
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                            
                       'FailCount Interpretation
                If patgen_fails_fast(3) = 0 Or patgen_fails_slow(3) = 0 Then 'Open socket workaround
                
                    ValidityCountFast(3) = 0
                    ValidityCountSlow(3) = 0
                
                Else
                       
                    ValidityCountFast(3) = Int(55 - patgen_fails_fast(3))
                       
                    ValidityCountSlow(3) = Int(55 - patgen_fails_slow(3))
                
                End If
                            
                            
                If (ValidityCountFast(3) > 0) Or (ValidityCountSlow(3) > 0) Then
                            
                    ID_Valid.AddPin("RFHOUT").Value(3) = 1
                            
                Else
                            
                    ID_Valid.AddPin("RFHOUT").Value(3) = 0
                                
    
                End If
                    
                TheExec.Sites.RestoreFromOverride
      
      
      End If 'Site 3 Active
      
     loopstatus = TheExec.Sites.SelectNext(loopstatus)
    
Wend 'end WHILE loop
  
  'End Serial Loop

    Call TheHdw.Digital.Patgen.Halt
    

    TheExec.DataLog.WriteComment ("==================  READ_MODULE_ID  =================")
    
    TheExec.Flow.TestLimit ID_Valid, LoLimit, HiLimit, , , , unitNone, "%2.1f", "ID", , , , , , , , tlForceNone
    
'    If TheExec.CurrentJob = "f1-prd-std-rn2903" Then
'
'        TheExec.Flow.TestLimit ID_Valid, 0.5, 1.5, , , , unitNone, "%2.1f", "ID", , , , , , , , tlForceNone
'
'    ElseIf TheExec.CurrentJob = "q1-prd-std-rn2903" Then
'
'        TheExec.Flow.TestLimit ID_Valid, 0.4, 1.6, , , , unitNone, "%2.1f", "ID_qc", , , , , , , , tlForceNone
'
'    End If
      
   
        Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

errHandler:

    Call TheHdw.Digital.Patgen.Halt
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_id_mfs")
    Call TheExec.ErrorReport
    rn2903_id_mfs = TL_ERROR

    
End Function

Public Function rn2903_idle_current(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'Reset times vary for the modules, so two attempts are allowed to measure the idle current after reset.

    Dim Site As Variant
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
    
        TheHdw.Wait (0.05)
        
        
                    'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
            
         TheHdw.Wait (0.05)
         
                     'RESET INACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitHi
            TheHdw.pins("MCLR_nRESET").StartState = chStartHi
        
        I_IDLE.AddPin ("VBAT")
        
'  For nSiteIndex = 0 To ExistingSiteCnt - 1
'
'    I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
'
'  Next nSiteIndex
        
        TheHdw.Wait (0.3)
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
     If TheExec.Sites.Site(nSiteIndex).Active Then
        
        
        
        I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps100mA
            .CurrentLimit = 0.1
            TheHdw.DPS.Samples = 1
            Call .MeasureCurrents(dps100mA, I_IDLE)
        End With
        
   
    
            If I_IDLE.pins("VBAT").Value(nSiteIndex) > 0.007 Then
            
            TheHdw.Wait (0.3)
            
                 With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.Samples = 1
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
      
        I_IDLE.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_IDLE, 0.002, 0.007, , , scaleMilli, unitAmp, "%2.2f", "I_IDLE_RN2903", , , , , , , , tlForceNone

    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_idle_current")
    Call TheExec.ErrorReport
    rn2903_idle_current = TL_ERROR
    
End Function

Public Function rn2903_tx915_cw(argc As Long, argv() As String) As Long

    Dim Site As Variant
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
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        
    Case Is = 2
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        
    Case Is = 3
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        
    Case Is = 4
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
        
        
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
    


    Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB
    
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(8192) 'Fres = 30.5176 kHz  (Fs = 250MHz, N=8192) NOTE: Fres chosen to bound TX freq
  
        TheHdw.Wait 0.05
        
    Select Case TestFreq
    Case 915000000
        TxPower915.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower915.pins("RFHOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        

    
    Case Else
        TxPower915.AddPin ("RFHOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower915.pins("RFHOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        
    End Select


    TheHdw.Wait (0.002)
    
        
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
        TheHdw.DPS.Samples = 1
    End With
    
    
        'Measure Current
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, I_TX915_CW)
    

        'Measure RF Power
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1 'Site loop needed for AXRF

            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(nSiteIndex), 20, TestFreq  'set for +20dBm
                
                TheHdw.Wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 4096, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", False, False, False, 1, SumPowerTemp) 'True plots waveform
                
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                SumPower(nSiteIndex) = SumPowerTemp
            
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                
                Select Case TestFreq
                Case 915000000
                    
                    TxPower915.pins("RFHOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                    TxPower915.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select


            End If
            
        Next nSiteIndex
        
        'Reset cpuA flag
        FlagsSet = 0
        FlagsClear = cpuA

        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) 'Pattern continues after cpuA reset.
        
        
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.
  
    'Run pattern to stop DUT transmitting
    
        TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_tx_cw_off").start ("start_tx_cw_off") 'all sites

        'TheHdw.Wait (0.1) 'avoids LVM Priming patgen RTE
        Call TheHdw.Digital.Patgen.HaltWait 'Wait for pattern to halt.

 TheExec.DataLog.WriteComment ("==================== TX915_CW_PWR ===================")
 
    Select Case TestFreq
    
    Case 915000000
    
    TheExec.Flow.TestLimit I_TX915_CW, LoLimit_I, HiLimit_I, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
    TheExec.Flow.TestLimit TxPower915, LoLimit_Tx, HiLimit_Tx, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
        
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

    End Select

 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

errHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower915.pins("RFHOUT").Value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 915000000
        TheExec.Flow.TestLimit I_TX915_CW, 0.065, 0.11, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, 13, 19, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone
  
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit I_TX915_CW, 0.064, 0.111, , , scaleMilli, unitAmp, "%2.2f", "I_TX915_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower915, 12.5, 20, , , , unitDb, "%2.1f", "TxPower_915", , , , , , , , tlForceNone

    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    rn2903_tx915_cw = TL_ERROR
    
End Function

Public Function rn2483_gpio(argc As Long, argv() As String) As Long

'Previous template version of this test did not bin properly for multi-site operation, so the test was written in VBT.
'The LoRa module has 14 gpio pins, GPIO0 - GPIO13. Each gpio pin is commanded via the UART to be set to a logical 1, tested in the pattern, then set to logical 0, then tested in the pattern.

    Dim Site As Variant
    
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
        
        
        TheHdw.Wait (0.3)
        
    
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

    Dim Site As Variant
    
      'Dim oprVolt As Double
      'Dim dut_delay As Double
      
      Dim nSiteIndex As Long
      
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo errHandler
    
        rn2903_gpio = TL_SUCCESS
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    
TheExec.DataLog.WriteComment ("===================== GPIO CHECK ====================")
    
        'oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        'dut_delay = 0.1
            
                      'RESET ACTIVE
            TheHdw.pins("MCLR_nRESET").InitState = chInitLo
            TheHdw.pins("MCLR_nRESET").StartState = chStartLo
  
  
  Call TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2903_gpio_full").Test(pfAlways, 0)
        
        
        TheHdw.Wait (0.3)
        
    
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

errHandler:


    If AbortTest Then Exit Function Else Resume Next
    
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2903_gpio")
    Call TheExec.ErrorReport
    rn2903_gpio = TL_ERROR
    
End Function


