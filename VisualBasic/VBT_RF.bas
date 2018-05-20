Attribute VB_Name = "VBT_RF"
Option Explicit

#Const Connected_to_J750 = True
#Const Connected_to_AXRF = True
#Const Debug_advanced = True

Public FIRSTLOAD As Boolean


Global Const DUTRefClkFreq As Double = 13560000# 'MHz

Public tx_path_db(7) As Double      'Dim for number of sites
Public coax_cable_db(7) As Double   'Dim for number of sites



Public Function read_cal_factors() As Long              '(argc As Long, argv() As String) As Long

    'Public tx_path_db() As Double
    'Public rx_path_db() As Double
    'Public coax_cable_db() As Double
    
    'This function reads the RF_Cal_Factors worksheet to initialize global RF scalar calibration offsets
    
    Dim nSiteIndex As Long
    
    On Error GoTo ErrHandler
    
     
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
    
        If TheExec.Sites.Site(nSiteIndex).Active = True Then
        
            tx_path_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(3, 2 + nSiteIndex).Value       'DIB TX path loss
                           
            coax_cable_db(nSiteIndex) = Worksheets("RF_Cal_Factors").Cells(4, 2 + nSiteIndex).Value    'Cable loss from PXI to DIB
        
            'Debug.Print "TX_Path_dB = "; tx_path_db(nSiteIndex)
            'Debug.Print "Coax_Cable_dB = "; coax_cable_db(nSiteIndex)
        
        End If
        
    Next nSiteIndex
    
    Exit Function
    
ErrHandler:

    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into prouduction abort routine
    read_cal_factors = TL_ERROR

End Function



Public Function rf_power_meas_t39a(argc As Long, argv() As String) As Long

    Dim Site As Variant

    On Error GoTo ErrHandler

    'DUT/DIB Setup:
    'Step -1: set axrf in TX mode
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount

    Dim MeasChans() As AXRF_CHANNEL         'Site Array
    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double           'Site Array
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    
    'AXRF Channel assign across Sites
    Select Case ExistingSiteCnt                  'added for ITL v1.4.7
'    Case Is >= 4
'        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
'        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
'        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
'        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
'
'    Case 8
'        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
'        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
'        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
'        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
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
        
    Case Is = 5
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        
    Case Is = 6
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        
    Case Is = 7
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        
    Case Is = 8
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
    Case Else
        MsgBox "Error in [rf_power_meas_3407x]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo ErrHandler
        
    End Select
    
    Dim TxPower868 As New PinListData
    Dim TxPower434 As New PinListData
    Dim IDD_TX As New PinListData
    Dim TX_Off As New PinListData
    Dim OOKDepth As New PinListData
    Dim I_Toff As New PinListData
    
    Dim MeasPower(1) As Double
    Dim MeasData() As Double
    Dim nSiteIndex As Long
    
    Dim IndexMaxPower As New SiteDouble
    Dim MaxPowerTemp As Double
    
    Dim TestFreq As Double
    Dim DATAPinLevel As Integer
    Dim OOKDepthMode As Boolean
    Dim OnSCK_BIAS As Boolean

    Dim oprVolt As Double

    If argc < 4 Then
        MsgBox "Error - On Rf_power_meas_t39a - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
    
'''        'The Variable transfering is not work - Need Debug more now use argv array directly
        TestFreq = argv(0)              ' What is testing Freq?
        DATAPinLevel = argv(1)          ' What is data being use to test?
        OOKDepthMode = argv(2)          ' Add OOK testing?
        OnSCK_BIAS = argv(3)            ' ON SCK_BIAS?
        oprVolt = ResolveArgv(argv(4))  ' Operating Voltage
        
    End If
    
    Select Case TestFreq
    Case 868300000
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO1 @ +1.8V =============================")
    
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState0 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
         
    Case 433920000
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO2 @ +3.3V =============================")
        
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState1 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
        
        Call Ref_Clock_On(DUTRefClkFreq)  'DEBUG NICK 11192014
        
    Case Else
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO1 =============================")
    
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState0 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
    
    End Select

    'TheHdw.Utility.Pins("rlyXTAL").State = utilBitState1 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
    
    Call read_cal_factors   'RF Calibration Offsets
    
    'Call Ref_Clock_On(DUTRefClkFreq)
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(2048)
    'Note:  All reg_writes are (address, data)

    'Setup for TXConfig1
    
    TheHdw.Wait 0.05
    '=================================================================================================================================
    
    Select Case TestFreq
    Case 868300000
        
        TxPower868.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
            TxPower868.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
        
        TheHdw.pins("SCK_BIAS").InitState = chInitOff    'No SCK_BIAS For weak PullUp
        TheHdw.pins("SCK").InitState = chInitLo         'SCK Setup for 868.30 MHZ
        
    Case 433920000
    
        TxPower434.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower434.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
        
        TheHdw.pins("SCK_BIAS").InitState = chInitOff    'No SCK_BIAS For weak PullUp
        TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 433.92 MHZ
        
    Case Else
        
'        Just Made Dummy Tests to be force fail.
        TxPower868.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
        
        TheHdw.pins("SCK_BIAS").InitState = chInitOff    'No SCK_BIAS For weak PullUp
        TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 868.30 MHZ
        
    End Select
        
    If argv(3) = "1" Then
        TheHdw.pins("SCK_BIAS").InitState = chInitLo     'On  SCK_BIAS For weak PullUp
    Else
        TheHdw.pins("SCK_BIAS").InitState = chInitOff    'Off SCK_BIAS For weak PullUp
    End If
        
    'Pre Force Fail
    
    'Trigger Capture for each site, wait 1ms between captures
    'Setup DUT in Mode 1
    
    TheHdw.Wait (0.002)
    
    Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
    Select Case DATAPinLevel
    Case 0
        TheHdw.pins("DATA").InitState = chInitLo            'DATA Lo = No TX Signal
    
    Case 1
        TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal
    
    Case Else   'Dummy
        TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal
        
    End Select
    
    TheHdw.Wait (0.002)
    
    '#If Debug_advanced Then

        For nSiteIndex = 0 To ExistingSiteCnt - 1
        
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(nSiteIndex), 10, TestFreq
                
                TheHdw.Wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 1024, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                              
                Select Case TestFreq
                Case 868300000
                    TxPower868.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                Case 433920000
                    TxPower434.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                Case Else       'Dummy for force fail purpose
                    TxPower868.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select
                               
                ' For OOK Test
                If argv(2) = "1" Then
                    
                    TheHdw.pins("DATA").InitState = chInitLo           'DATA Lo = No TX Signal
                    TheHdw.Wait (0.001)
                    
                    Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 1024, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                    
                    UncalMaxPower(nSiteIndex) = MaxPowerTemp
                    MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                    
                    TX_Off.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                    
                    Select Case TestFreq
                    Case 868300000
                        OOKDepth.pins("RFOUT").Value(nSiteIndex) = TxPower868.pins("RFOUT").Value(nSiteIndex) - TX_Off.pins("RFOUT").Value(nSiteIndex)
                        
                    Case 433920000
                        OOKDepth.pins("RFOUT").Value(nSiteIndex) = TxPower434.pins("RFOUT").Value(nSiteIndex) - TX_Off.pins("RFOUT").Value(nSiteIndex)
                        
                    Case Else
                        OOKDepth.pins("RFOUT").Value(nSiteIndex) = TxPower868.pins("RFOUT").Value(nSiteIndex) - TX_Off.pins("RFOUT").Value(nSiteIndex)
                        
                    End Select
                    
                    TheHdw.pins("DATA").InitState = chInitHi           'DATA hi = TX Signal
                    TheHdw.Wait (0.001)
                    
                End If
                
            End If
            
        Next nSiteIndex

    ' Perform DPS set up and Measuring
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.Samples = 1
        TheHdw.Wait 0.01          'Settling Time is here
        Call .MeasureCurrents(dps100mA, IDD_TX)
    End With
    
    Select Case TestFreq
    Case 868300000
        TheExec.Flow.TestLimit TxPower868, -10, 0, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_868", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , unitDb, , "OOKDepth_868", , , , , , , , tlForceNone
        
    Case 433920000
        TheExec.Flow.TestLimit TxPower434, 5, 15, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_434", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, , "OOKDepth_434", , , , , , , , tlForceNone
        
        ' ============================= I_Toff_2mS portion ==================================================
    
        TheHdw.pins("SCK").InitState = chInitLo     'Set For T_OFF 2 MS
        TheHdw.pins("DATA").InitState = chInitHi     'Set For T_OFF 2 MS
        
        TheHdw.pins("SCK").StartState = chStartLo
        TheHdw.pins("DATA").StartState = chStartHi
        
        ' Perform DPS set up
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps50uA
            .CurrentLimit = 0.1
            TheHdw.DPS.Samples = 1
        End With
           
        TheHdw.Wait (0.005)      ' Wait For Device Ready
           
        TheHdw.pins("SCK").InitState = chInitLo     'Set For T_OFF 2 MS
        TheHdw.pins("DATA").InitState = chInitLo     'Set For T_OFF 2 MS
        
        TheHdw.pins("SCK").StartState = chStartLo   'Set to Sleep
        TheHdw.pins("DATA").StartState = chStartLo  'Set to Sleep
        
        TheHdw.Wait (0.03 + 0.002)     ' Settling Time 2mS as Spec metion     '30mS to compensate delay discharge
        
        'Make Current Measurement
        Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_Toff)
           
        TheExec.Flow.TestLimit I_Toff, -0.00000001, 0.0000011, , , scaleNano, unitAmp, , "I_TOFF_2mS", , , , , , , , tlForceNone
           
        '=========================================================================================
        
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit TxPower868, -10, 0, , , , unitDb, , "TxPower_NA", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_NA", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, , "OOKDepth_NA", , , , , , , , tlForceNone
        
    End Select
    
''    TheHdw.Pins("SCK").InitState = chInitLo           'Remain for Toff_2mS test
''    TheHdw.Pins("DATA").InitState = chInitLo          'Remain for Toff_2mS test
    TheHdw.Wait 0.00001
''    TheHdw.Pins("SCK").InitState = chInitOff          'Remain for Toff_2mS test
''    TheHdw.Pins("DATA").InitState = chInitOff         'Remain for Toff_2mS test
    
    Exit Function

ErrHandler:
    
''    TheHdw.Pins("SCK").InitState = chInitLo
''    TheHdw.Pins("DATA").InitState = chInitLo
''    TheHdw.Wait 0.00001
''    TheHdw.Pins("SCK").InitState = chInitOff
''    TheHdw.Pins("DATA").InitState = chInitOff
''    TheHdw.Pins("SCK_BIAS").InitState = chInitOff
    
    Select Case TestFreq
    Case 868300000
        TxPower868.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
        
        TheExec.Flow.TestLimit TxPower868, -6, 6, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_868", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, , "OOKDepth_868", , , , , , , , tlForceNone
        
    Case 433920000
        TxPower434.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower434.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
    
        TheExec.Flow.TestLimit TxPower434, 5, 15, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_434", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, , "OOKDepth_434", , , , , , , , tlForceNone
        
    Case Else      'Dummy for force fail purpose
        TxPower868.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower868.pins("RFOUT").Value(nSiteIndex) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(nSiteIndex) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(nSiteIndex) = 100
        Next nSiteIndex
        
        TheExec.Flow.TestLimit TxPower868, -6, 6, , , , unitDb, , "TxPower_NA", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit IDD_TX, -0.001, 0.03, , , scaleMilli, unitAmp, , "IDD_TX_NA", , , , , , , , tlForceNone
        If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, , "OOKDepth_NA", , , , , , , , tlForceNone
        
    End Select

    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function forced_tx_mode_t39a(argc As Long, argv() As String) As Long

    Dim Site As Variant

    On Error GoTo ErrHandler
    
    TheExec.Datalog.WriteComment ("============================= MEASURE FORCED TX MODE ===========================")
    'DUT/DIB Setup:
    'Step -1: set axrf in TX mode
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    Dim MeasChans() As AXRF_CHANNEL         'Site Array
    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double          'Site Array
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    
    Dim I_Toff As New PinListData
       
    'AXRF Channel assign across Sites
    Select Case ExistingSiteCnt
'    Case Is >= 4
'        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
'        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
'        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
'        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
'
'    Case 8
'        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
'        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
'        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
'        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
    Case Is = 1                                 'added for ITL v1.4.7
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
        
    Case Is = 5
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        
    Case Is = 6
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        
    Case Is = 7
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        
    Case Is = 8
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
    Case Else
        MsgBox "Error in [forced_tx_mode_3407x]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo ErrHandler
        
    End Select
    
    Dim TxPower390 As New PinListData   'freq for forced tx mode set by pattern

    Dim IDD_TX As New PinListData

    Dim MeasPower(1) As Double
    Dim MeasData() As Double
    Dim nSiteIndex As Long
    
    Dim IndexMaxPower As New SiteDouble
    Dim MaxPowerTemp As Double
    
    Dim TestFreq As Double
    Dim DATAPinLevel As Integer
    Dim OOKDepthMode As Boolean
    
    Dim oprVolt As Double

    If argc < 2 Then
        MsgBox "Error - On forced_tx_mode_3407x - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
        TestFreq = argv(0)              ' What is testing Freq?
        DATAPinLevel = argv(1)          ' What is data being use to test?
        oprVolt = 3.3
        
    End If

    TheHdw.Utility.pins("rlyXTAL").State = utilBitState1    '1 = NI-6652 Ref clock source; 0 = DIB Crystal
    
    Call read_cal_factors                   'RF Calibration Offsets
    
    Call Ref_Clock_On(DUTRefClkFreq)        'Reference Clock Setip to 26 MHz
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(2048)
    'Note:  All reg_writes are (address, data)

    'Setup for TXConfig1
    
        TheHdw.Wait 0.05
    '=================================================================================================================================
    Select Case TestFreq
    Case 390000000
        TxPower390.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower390.pins("RFOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 433.92 MHZ
    
    Case Else
        TxPower390.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower390.pins("RFOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 433.92 MHZ
        
    End Select

    'Pre Force Fail
    
    'Trigger Capture for each site, wait 1ms between captures
    'Setup DUT in Mode 1
    
    TheHdw.Wait (0.002)
    
    Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
    Select Case DATAPinLevel

    Case 0
        TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal at 433.92 MHz (not measured)
        
        TheHdw.Wait (0.002)
        
        TheHdw.pins("DATA").InitState = chInitLo            'Stop TX at 433.92 MHz
        
        'Run pattern to put DUT into Forced TX Mode @ 390 MHz, +10 dBm Power
    
        TheHdw.Digital.Patterns.Pat("./Patterns/TX_ON_390.PAT").start ("Forced_TX_Start")
        
        TheHdw.Wait (0.01)
        
    Case Else   'Dummy  For force fail purpose
        TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal
        
        TheHdw.Wait (0.002)
        
        TheHdw.pins("DATA").InitState = chInitLo            'Stop TX at 433.92 MHz
        
        'Run pattern to put DUT into Forced TX Mode @ 390 MHz, +10 dBm Power
    
        TheHdw.Digital.Patterns.Pat("./Patterns/TX_ON_390.PAT").start ("Forced_TX_Start")
        
        TheHdw.Wait (0.01)
        
    End Select
    
    TheHdw.Wait (0.002)
    
    '#If Debug_advanced Then

        For nSiteIndex = 0 To ExistingSiteCnt - 1
        
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(nSiteIndex), 10, TestFreq
                
                TheHdw.Wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 1024, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                'MaxPower.Pins("RFOUT").Value(nSiteIndex) = MaxPowerTemp + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                Select Case TestFreq
                Case 390000000
                    TxPower390.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                    TxPower390.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select

            End If
            
        Next nSiteIndex

    'Run pattern to stop DUT transmitting
                
'    TheHdw.Digital.Patterns.Pat("./Patterns/TX_OFF_434.PAT").Start ("Force_TX_OFF_Start")
    
    Select Case TestFreq
    Case 390000000
        TheExec.Flow.TestLimit TxPower390, 5, 15, , , , unitDb, , "TxPower_390", , , , , , , , tlForceNone
        
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit TxPower390, 5, 15, , , , unitDb, , "TxPower_390", , , , , , , , tlForceNone

    End Select

    ' ============================= I_Toff_20mS portion ==================================================

    TheHdw.pins("SCK").InitState = chInitLo     'Set For T_OFF 20 MS
    TheHdw.pins("DATA").InitState = chInitLo     'Set For T_OFF 20 MS
    
    TheHdw.pins("SCK").StartState = chStartLo
    TheHdw.pins("DATA").StartState = chStartLo
    
    ' Perform DPS set up
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps50uA
        .CurrentLimit = 0.1
        TheHdw.DPS.Samples = 1
    End With
    
'    Call cycle_power(0, 3.3, 0.1, 0.01)
     
    TheHdw.Digital.Patterns.Pat("./Patterns/T39A_writeDA_Toff_delay_20ms.pat").start ("")
       
    TheHdw.Wait (0.005)      ' Wait For Device Ready
       
    TheHdw.pins("SCK").InitState = chInitLo     'Set For T_OFF 20 MS
    TheHdw.pins("DATA").InitState = chInitLo     'Set For T_OFF 20 MS
    
    TheHdw.pins("SCK").StartState = chStartLo   'Set to Sleep
    TheHdw.pins("DATA").StartState = chStartLo  'Set to Sleep
    
    TheHdw.Wait (0.03 + 0.02)      ' Settling Time 20mS as Spec metion     '30mS to compensate delay discharge
    
    'Make Current Measurement
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps50uA, I_Toff)
       
    TheExec.Flow.TestLimit I_Toff, -0.00000001, 0.0000011, , , scaleNano, unitAmp, , "I_TOFF_20mS", , , , , , , , tlForceNone
       
    '=========================================================================================
    
    Exit Function

ErrHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower390.pins("RFOUT").Value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 390000000
        TheExec.Flow.TestLimit TxPower390, 5, 15, , , , unitDb, , "TxPower_390", , , , , , , , tlForceNone
  
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit TxPower390, 5, 15, , , , unitDb, , "TxPower_390", , , , , , , , tlForceNone

    End Select
    
    TheHdw.Digital.Patterns.Pat("./Patterns/TX_OFF_434.PAT").start ("Force_TX_OFF_Start")
    
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function rf_power_meas_t48a(argc As Long, argv() As String) As Long

'RF Tests contained in this function for T48A products:

    'PWRNGO_434_10DBM :
        'TX power @ 433.92 MHz and +10dBm
    'IDD_TX_434:
        'DC current on VBAT (+1.8V) while DUT is in TX mode @ 433.92 MHz and +10dBm. Uses 26MHz crystal. Uses SCK 20K pullup resistor.
        
    'PWRNGO_868_10DBM :
        'TX power @ 868.30 MHz and +10dBm
    'IDD_TX_868:
        'DC current on VBAT(+3.3V) while DUT is in TX mode @ 868.30 MHz and +10dBm. Uses DDS 26 MHz reference.



    Dim Site As Variant
    
    Dim TxPower868 As New PinListData
    Dim TxPower434 As New PinListData
    Dim IDD_TX_434 As New PinListData
    Dim IDD_TX_868 As New PinListData
    
    Dim IDD_TX As New PinListData
    Dim TX_Off As New PinListData
    Dim OOKDepth As New PinListData
    Dim I_Toff As New PinListData
    
    Dim MeasPower(1) As Double
    Dim MeasData() As Double
    Dim nSiteIndex As Long
    Dim td_samples As Long
    Dim fft_samples As Long
    
    Dim IndexMaxPower As New SiteDouble
    Dim MaxPowerTemp As Double
    
    Dim TestFreq As Double
    Dim DATAPinLevel As Integer
    Dim OOKDepthMode As Boolean
    Dim OnSCK_BIAS As Boolean

    Dim oprVolt As Double
    
    Dim ExistingSiteCnt As Integer
    

    Dim MeasChans() As AXRF_CHANNEL         'Site Array
    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double           'Site Array
    
    On Error GoTo ErrHandler
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    
    'AXRF Channel assignment across Sites:
    
    Select Case ExistingSiteCnt                  'changed for ITL v1.4.7

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

    Case Is = 5
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7

        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2

    Case Is = 6

        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7

        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4

    Case Is = 7

        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7

        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        
    Case Is = 8
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
    Case Else
        MsgBox "Error in [rf_power_meas_t48a]" & vbCrLf & _
               "Site number is not supported by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo ErrHandler

    End Select
    


    If argc < 4 Then
        MsgBox "Error - On Rf_power_meas_t48a - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
    

        TestFreq = argv(0)              ' What is testing Freq?
        DATAPinLevel = argv(1)          ' What is data being use to test?
        OOKDepthMode = argv(2)          ' Add OOK testing?
        OnSCK_BIAS = argv(3)            ' ON SCK_BIAS?
        oprVolt = ResolveArgv(argv(4))  ' Operating Voltage
        
    End If
    
    Select Case TestFreq
    
    Case 433920000
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO_434 @ +1.8V =============================")
        
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState0 '1 = NI-6652 Ref clock source; 0 = DIB Crystal; DUT PLL operates from site 26MHz crystal
        
        
        
    Case 868300000
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO_868 @ +3.3V =============================")
    
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState1 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
        
        Call Ref_Clock_On(DUTRefClkFreq) 'NI-6652 supplied 26MHz DDS reference clock for DUT
         

        
    Case Else
    
        TheExec.Datalog.WriteComment ("============================= MEASURE TX POWERnGO_434 =============================")
    
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState0 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
    
    End Select

    
    Call read_cal_factors   'RF Calibration Offsets from worksheet
    
    'Digitizer Samples
    td_samples = 2048   'Time Domain Samples [Default is 2048], must be an integer power of 2
    fft_samples = td_samples / 2 'Frequency Domain Samples   Center (tuned) Frequency is in FFT Bin (fft_samples/2 -1)
    
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(td_samples)
    
    
    
    'Setup for PWRNGO_434_10DBM (from FLEX code)
        'Turn off VBAT, wait 20msec
        'Disconnect DATA, SCK, SCK_BIAS in order
        'Set DATA to Logic 0
        'Set SCK to Logic 0
        'Connect SCK, DATA
        'Set SCK to Logic 1, wait 3 msec
        'Set DATA to Logic 1
        'Power up VBAT at +1.8V, wait 10msec
        'Set Data to Logic 0, wait 1msec
        'Set Data to Logic 1, wait 100msec
        'Measure TX power at 433.92MHz
        'Measure IDD_TX_434
        
    'Setup for PWRNGO_868_10DBM (from FLEX code)
        'Turn off VBAT, wait 20msec
        'Disconnect DATA, SCK, SCK_BIAS in order
        'Power OFF DUT
        'Set SCK to Logic 0
        'Set DATA to Logic 0
        'Connect DATA, SCK
        'wait 50ms
        'Set DATA to Logic 1
        'Set SCK to Logic 0, wait 100msec
        'Power up VBAT at +3.3V
        'Set DATA to Logic 0
        'Power up VBAT at +3.3V, wait 10msec
        'Measure TX power at 868.30MHz
        'Measure IDD_TX_868
        

    
    '=================================================================================================================================
    
    Select Case TestFreq

        
    Case 433920000 'Setup for PWRNGO_434_10DBM (from FLEX code)
    
        oprVolt = 1.8  'Minimum Data Sheet VBAT
    
        'Disconnect DATA, SCK, SCK_BIAS in order
            TheHdw.pins("DATA").InitState = chInitOff
            TheHdw.pins("SCK").InitState = chInitOff
            TheHdw.pins("SCK_BIAS").InitState = chInitOff
            
        'Set DATA to Logic 0
            TheHdw.pins("DATA").InitState = chInitLo
            
            If argv(3) = "1" Then       'SCK_BIAS uses 20K weak pullup resistor
                'Set SCK_BIAS to Logic 0
                    TheHdw.pins("SCK_BIAS").InitState = chInitLo
                    
                'Set SCK_BIAS to Logic 1, wait 3msec
                    TheHdw.pins("SCK_BIAS").InitState = chInitHi
                    TheHdw.Wait 0.003
            Else
                'Set SCK to Logic 0
                    TheHdw.pins("SCK").InitState = chInitLo
                      
                'Set SCK to Logic 1, wait 3msec
                    TheHdw.pins("SCK").InitState = chInitHi
                    TheHdw.Wait 0.003
            End If
             
        'Set DATA to Logic 1
            TheHdw.pins("DATA").InitState = chInitHi
            

        TxPower434.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")       'Added if OOK Depth is selected in test instance
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
If TheExec.Sites.SelectFirst <> loopDone Then

    Do
        Site = TheExec.Sites.SelectedSite
            
            TxPower434.pins("RFOUT").Value(Site) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(Site) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(Site) = 100
            
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    
End If

        
        Call cycle_power(0.001, oprVolt, 0.01, 0.01)   'DUT should be in TX mode after this line is excuted
        
    Case 868300000
    
        oprVolt = 3.3  'Nominal Data Sheet VBAT
        
        
        'Disconnect DATA, SCK, SCK_BIAS in order
            TheHdw.pins("DATA").InitState = chInitOff
            TheHdw.pins("SCK").InitState = chInitOff
            TheHdw.pins("SCK_BIAS").InitState = chInitOff
            
        'Set DATA to Logic 0
            TheHdw.pins("DATA").InitState = chInitLo
            
            
            If argv(3) = "1" Then       'SCK_BIAS uses 20K weak pullup resistor
                'Set SCK_BIAS to Logic 0
                    TheHdw.pins("SCK_BIAS").InitState = chInitLo
            Else
                'Set SCK to Logic 0
                    TheHdw.pins("SCK").InitState = chInitLo
            End If
            
                'Set DATA to Logic 1
            TheHdw.pins("DATA").InitState = chInitHi
            
                
        
        TxPower868.AddPin ("RFOUT")
        If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
        If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")

If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
        
            TxPower868.pins("RFOUT").Value(Site) = -90
            If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(Site) = 90
            If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(Site) = 100

    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    
End If
        Call cycle_power(0.001, oprVolt, 0.01, 0.01)   'DUT should be in TX mode after this  line is excuted
        
    Case Else
        

        
    End Select

    
        TheHdw.Wait (0.002)
    
    Select Case DATAPinLevel
    
        Case 0
            TheHdw.pins("DATA").InitState = chInitLo            'DATA Lo = No TX Signal
        
        Case 1
            TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal
        
        Case Else   'Dummy
            TheHdw.pins("DATA").InitState = chInitHi            'DATA Hi = TX Signal
        
    End Select
    
        TheHdw.Wait (0.002)
        
'RF Power Measurement Site Loop
        
If TheExec.Sites.SelectFirst <> loopDone Then

    Do
        Site = TheExec.Sites.SelectedSite
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, TestFreq
                
                    TheHdw.Wait (0.01)      'RF MUX Speed depended.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                
                UncalMaxPower(Site) = MaxPowerTemp
                
                MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                Select Case TestFreq
                
                    Case 433920000
                        TxPower434.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
            
                    Case 868300000
                        TxPower868.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
            
    
                    Case Else       'Dummy for force fail purpose
                        TxPower868.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
        
                End Select
                               
                ' For OOK Test
                If argv(2) = "1" Then
                    
                    TheHdw.pins("DATA").InitState = chInitLo           'DATA Low = TX OFF
                        TheHdw.Wait (0.001)
                    
                    Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots data
                    
                    UncalMaxPower(Site) = MaxPowerTemp
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                    
                    TX_Off.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                    
                    Select Case TestFreq
                    
                       Case 433920000
                           OOKDepth.pins("RFOUT").Value(Site) = TxPower434.pins("RFOUT").Value(Site) - TX_Off.pins("RFOUT").Value(Site)
    
                       Case 868300000       '868MHz Power-N-Go mode is FSK only. There is no OOK mode!
                           'OOKDepth.pins("RFOUT").Value(site) = TxPower868.pins("RFOUT").Value(site) - TX_Off.pins("RFOUT").Value(site)
                           
                          
                       Case Else
                           OOKDepth.pins("RFOUT").Value(Site) = TxPower434.pins("RFOUT").Value(Site) - TX_Off.pins("RFOUT").Value(Site)
                           
                    End Select
                    
                    TheHdw.pins("DATA").InitState = chInitHi           'DATA hi = TX Signal
                        TheHdw.Wait (0.001)
                    
                End If
                
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    
End If
        

' Perform DPS setup for Measuring IDD_TX
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.Samples = 1
        TheHdw.Wait 0.01          'Settling Time is here
        Call .MeasureCurrents(dps100mA, IDD_TX)
    End With
    
    Select Case TestFreq
    
    Case 433920000
    
        If TheExec.CurrentJob = "f1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit IDD_TX, -0.002, 0.003, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower434, 0, 7, , , , unitDb, "%4.1f", "TxPower_434", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, "%4.1f", "OOKDepth_434", , , , , , , , tlForceNone
        
        ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit IDD_TX, -0.002, 0.005, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434_qc", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower434, 0, 8, , , , unitDb, "%4.1f", "TxPower_434_qc", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 39, 60, , , , unitDb, "%4.1f", "OOKDepth_434_qc", , , , , , , , tlForceNone
        
        End If
    
    
    
    Case 868300000
    
    
        If TheExec.CurrentJob = "f1-prd-std-t48a" Then
    
            TheExec.Flow.TestLimit IDD_TX, 0.01, 0.02, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_868", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower868, 7, 13, , , , unitDb, "%4.1f", "TxPower_868", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , unitDb, "%4.1f", "OOKDepth_868", , , , , , , , tlForceNone
        
        ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit IDD_TX, 0.01, 0.021, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_868_qc", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower868, 5, 13, , , , unitDb, "%4.1f", "TxPower_868_qc", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 39, 60, , , unitDb, "%4.1f", "OOKDepth_868_qc", , , , , , , , tlForceNone
         
        End If
        
        
        
        
    Case Else      'Dummy for force fail purpose
    
    
        If TheExec.CurrentJob = "f1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit IDD_TX, -0.002, 0.003, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower434, 0, 7, , , , unitDb, "%4.1f", "TxPower_434", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, "%4.1f", "OOKDepth_434", , , , , , , , tlForceNone
        
        ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit IDD_TX, -0.002, 0.005, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434_qc", , , , , , , , tlForceNone
            TheExec.Flow.TestLimit TxPower434, 0, 8, , , , unitDb, "%4.1f", "TxPower_434_qc", , , , , , , , tlForceNone
            
            If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 39, 60, , , , unitDb, "%4.1f", "OOKDepth_434_qc", , , , , , , , tlForceNone
        
        End If
    
    
    End Select
    
    
    Exit Function

ErrHandler:
    
    
    Select Case TestFreq
    
    Case 433920000
                TxPower434.AddPin ("RFOUT")
                If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
                If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
        If TheExec.Sites.SelectFirst <> loopDone Then
        
            Do
                Site = TheExec.Sites.SelectedSite
                
            
                    TxPower434.pins("RFOUT").Value(Site) = -90
                    If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(Site) = 90
                    If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(Site) = 100
                
            Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
            
        End If
                TheExec.Flow.TestLimit TxPower434, 0, 7, , , , unitDb, "%4.1", "TxPower_434", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit IDD_TX, -0.002, 0.003, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434", , , , , , , , tlForceNone
                If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, "%4.1f", "OOKDepth_434", , , , , , , , tlForceNone
            
    
    Case 868300000
                TxPower868.AddPin ("RFOUT")
                If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
                If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
                
        If TheExec.Sites.SelectFirst <> loopDone Then
        
            Do
                Site = TheExec.Sites.SelectedSite
                '
                    TxPower868.pins("RFOUT").Value(Site) = -90
                    If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(Site) = 90
                    If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(Site) = 100
        
             Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
            
        End If
                TheExec.Flow.TestLimit TxPower868, 7, 13, , , , unitDb, "%4.1f", "TxPower_868", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit IDD_TX, 0.01, 0.02, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_868", , , , , , , , tlForceNone
                If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, "%4.1f", "OOKDepth_868", , , , , , , , tlForceNone
                
        
    Case Else      'Dummy for force fail purpose
                TxPower868.AddPin ("RFOUT")
                If argv(2) = "1" Then TX_Off.AddPin ("RFOUT")
                If argv(2) = "1" Then OOKDepth.AddPin ("RFOUT")
        
            If TheExec.Sites.SelectFirst <> loopDone Then
        
            Do
                Site = TheExec.Sites.SelectedSite
                
                    TxPower868.pins("RFOUT").Value(Site) = -90
                    If argv(2) = "1" Then TX_Off.pins("RFOUT").Value(Site) = 90
                    If argv(2) = "1" Then OOKDepth.pins("RFOUT").Value(Site) = 100
            
             Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
            
        End If
        
              TheExec.Flow.TestLimit TxPower434, 0, 7, , , , unitDb, "%4.1f", "TxPower_434", , , , , , , , tlForceNone
              TheExec.Flow.TestLimit IDD_TX, -0.002, 0.003, , , scaleMilli, unitAmp, "%5.3f", "IDD_TX_434", , , , , , , , tlForceNone
              If argv(2) = "1" Then TheExec.Flow.TestLimit OOKDepth, 40, 60, , , , unitDb, "%4.1f", "OOKDepth_434", , , , , , , , tlForceNone
        

    End Select

    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function sleep_current_t48a(argc As Long, argv() As String) As Long

    Dim Site As Variant
    Dim I_SLEEP As New PinListData
    
      Dim oprVolt As Double
      Dim dut_delay As Double
      
      Dim nSiteIndex As Long
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo ErrHandler
    
TheExec.Datalog.WriteComment ("============================= MEASURE I_SLEEP ==============================================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        dut_delay = 0.025
        
        'Disconnect DATA, SCK, SCK_BIAS in order
            TheHdw.pins("DATA").InitState = chInitOff
            TheHdw.pins("SCK").InitState = chInitOff
            TheHdw.pins("SCK_BIAS").InitState = chInitOff
        
        
        Call cycle_power(0.001, 0.001, 0.01, 0.01)
        
        'Set control pins to 0
        
            TheHdw.pins("DATA").InitState = chInitLo
            TheHdw.pins("SCK").InitState = chInitLo
            
        
         Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
        
        I_SLEEP.AddPin ("VBAT")
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
    I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value

  Next nSiteIndex
        
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1
  
        
    TheHdw.Wait dut_delay
    
    
        With TheHdw.DPS.pins("VBAT")
            .ClearLatchedCurrentLimit
            .ClearOverCurrentLimit
            .CurrentRange = dps50uA
            .CurrentLimit = 0.1
            TheHdw.DPS.Samples = 3
            Call .MeasureCurrents(dps50uA, I_SLEEP)
        End With
        
    Next nSiteIndex
    
        If TheExec.CurrentJob = "f1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit I_SLEEP, -0.0000001, 0.000001, , , scaleNano, unitAmp, "%4.0f", "I_SLEEP_T48A", , , , , , , , tlForceNone
              
        ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit I_SLEEP, -0.00000011, 0.0000011, , , scaleNano, unitAmp, "%4.0f", "I_SLEEP_T48A_qc", , , , , , , , tlForceNone
        
        End If
    
    
    Exit Function
    

ErrHandler:


    I_SLEEP.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
        I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, -0.0000001, 0.000001, , , scaleNano, unitAmp, "%4.0f", "I_SLEEP_T48A", , , , , , , , tlForceNone

    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function rf_adv_tx_mode_t48a(argc As Long, argv() As String) As Long

'Advanced Mode RF Tests for T48A products:

    'After Power & Go startup at 868MHz, the DUT is programmed for 8 different TX frequencies via control words written from a pattern.
    'PatGen-VBT handshaking is used to communicate between the pattern and the VBT. The following frequencies are programmed, and RF power measured:
    '
    '869.85 MHz
    '868.95 MHz
    '868.65 MHz
    '868.30 MHz
    '864.00 MHz
    '433.92 MHz
    '433.42 MHz
    '418.00 MHz
    
    'After each frequency is programmed into the DUT, the DUT is placed into TX mode with +10dBm output.
    'The pattern will set the CPU Flag and wait for the interposer VBT function to set up the Aeroflex and perform the RF power measurement. The interposer
    'function will then cause the pattern to continue by clearing the CPU Flag A so the pattern can program the next frequency.
    'The pattern will halt after the final power measurement at the last programmed frequency.



Dim Site As Variant

Dim ExistingSiteCnt As Integer

Dim MeasChans() As AXRF_CHANNEL         'Site Array
Dim MeasFactor As Double
Dim MaxPowerToSubstract() As Double     'Site Array
Dim UncalMaxPower() As Double           'Site Array

Dim DA_Freq(8) As Double        'Advanced Mode RF Frequencies

Dim TxPowerAdvMode_DA1 As New PinListData   '869.85 MHz
Dim TxPowerAdvMode_DA2 As New PinListData   '868.95 MHz
Dim TxPowerAdvMode_DA3 As New PinListData   '868.65 MHz
Dim TxPowerAdvMode_DA4 As New PinListData   '868.30 MHz

Dim TxPowerAdvMode_DA5 As New PinListData   '864.00 MHz
Dim TxPowerAdvMode_DA6 As New PinListData   '433.92 MHz
Dim TxPowerAdvMode_DA7 As New PinListData   '433.42 MHz
Dim TxPowerAdvMode_DA8 As New PinListData   '418.00 MHz



Dim MeasPower(1) As Double
Dim MeasData() As Double

Dim IndexMaxPower As Double
Dim MaxPowerTemp As Double

Dim TestFreq As Double
Dim DATAPinLevel As Integer

Dim oprVolt As Double

Dim CpuFlagBit As Long
Dim Flags As Long
Dim FlagsSet As Long
Dim FlagsClear As Long
Dim td_samples As Long
Dim fft_samples As Long


Dim dut_delay As Double
Dim rf_mux_delay As Double


 On Error GoTo ErrHandler
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    
 'AXRF Channel assignment across Sites:
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        

    TheExec.Datalog.WriteComment ("============================= MEASURE ADVANCED MODE TX POWER @ +3.3V ===========================")
    
   
    TheHdw.Utility.pins("rlyXTAL").State = utilBitState1 '1 = NI-6652 Ref clock source; 0 = DIB Crystal
        
    Call Ref_Clock_On(DUTRefClkFreq) 'NI-6652 supplied 26MHz DDS reference clock for DUT
    
    Call read_cal_factors   'RF Calibration Offsets from worksheet
       
    'DUT and Instrument Delays
             
    dut_delay = 0.1
    rf_mux_delay = 0.01    'DEBUG was 0.01
    
    'Digitizer Samples
        td_samples = 2048   'Time Domain Samples [Default is 2048], must be an integer power of 2
        fft_samples = td_samples / 2 'Frequency Domain Samples   Center (tuned) Frequency is in FFT Bin (fft_samples/2 -1)
         
        Call itl.Raw.AF.AXRF.SetMeasureSamples(td_samples) 'AXRF sample size
    
        oprVolt = 3.3  'Nominal Data Sheet VBAT
       
  
       
    'Identify and assign 8 T48A Frequencies in pattern:
     
     DA_Freq(1) = 869850000#
     DA_Freq(2) = 868950000#
     DA_Freq(3) = 868650000#
     DA_Freq(4) = 868300000#
     DA_Freq(5) = 864000000#
     DA_Freq(6) = 433920000#
     DA_Freq(7) = 433420000#
     DA_Freq(8) = 418000000#
             
   'Bind PinListData Objects to DUT TX Pin
        
    TxPowerAdvMode_DA1.AddPin ("RFOUT")
    TxPowerAdvMode_DA2.AddPin ("RFOUT")
    TxPowerAdvMode_DA3.AddPin ("RFOUT")
    TxPowerAdvMode_DA4.AddPin ("RFOUT")
    
    TxPowerAdvMode_DA5.AddPin ("RFOUT")
    TxPowerAdvMode_DA6.AddPin ("RFOUT")
    TxPowerAdvMode_DA7.AddPin ("RFOUT")
    TxPowerAdvMode_DA8.AddPin ("RFOUT")
    
    
    'Init VBT flags
    FlagsSet = 0                            'Pattern sets cpuA. VBT does not set cpuA.
    FlagsClear = cpuA + cpuB + cpuC + cpuD   'VBT asserts all FlagsClear.
        
  'Setup PowernGo at 868MHz before Advanced Mode
    
    'Disconnect DATA, SCK, SCK_BIAS in order
        TheHdw.pins("DATA").InitState = chInitOff
        TheHdw.pins("SCK").InitState = chInitOff
        TheHdw.pins("SCK_BIAS").InitState = chInitOff

    'Set DATA to Logic 0
        TheHdw.pins("DATA").InitState = chInitLo

    'Set SCK to Logic 0
        TheHdw.pins("SCK").InitState = chInitLo

    'Set DATA to Logic 1
        TheHdw.pins("DATA").InitState = chInitHi
            
    Call cycle_power(0.001, oprVolt, 0.01, 0.01)
        

    'Set DATA to Logic 0
        TheHdw.pins("DATA").InitState = chInitLo


If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
        
        TxPowerAdvMode_DA1.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA2.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA3.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA4.pins("RFOUT").Value(Site) = -99.9
                    
        TxPowerAdvMode_DA5.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA6.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA7.pins("RFOUT").Value(Site) = -99.9
        TxPowerAdvMode_DA8.pins("RFOUT").Value(Site) = -99.9
        
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    
End If
                 
        'Initial 868.30 MHz PowerNGo TX state

        'Wait for cpuA set in pattern
           Do

            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA

           Loop While (CpuFlagBit = 0)
            
            
        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
'================================================================================================
        
        'Begin 869.85 MHz TX Measure

            
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
           
           Loop While (CpuFlagBit = 0)
                
                
 If TheExec.Sites.SelectFirst <> loopDone Then

    Do
        Site = TheExec.Sites.SelectedSite
        
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(1)  '869.85 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                    
                    TheHdw.Wait rf_mux_delay
               
                    
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA1.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
            
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If



            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
        
        'End 869.85 MHz TX Measure
        
'================================================================================================

        'Begin 868.95 MHz TX Measure
            
            
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(2)  '868.95 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                 
                 TheHdw.Wait rf_mux_delay
                 
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA2.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
        

            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
        
         'End 868.95 MHz TX Measure

'================================================================================================

        'Begin 868.65 MHz TX Measure
         
         
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(3)  '868.65 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                    
                    TheHdw.Wait rf_mux_delay
                
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA3.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
        
            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
         
         'End 868.65 MHz TX Measure
         
'================================================================================================
         
         'Begin 868.30 MHz TX Measure

         
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(4)  '868.30 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                    
                    TheHdw.Wait rf_mux_delay
                    
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA4.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
    Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
       
  
            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
        
         'End 868.30 MHz TX Measure
         
'================================================================================================
         
         'Begin 864.00 MHz TX Measure
            
         
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(5)  '864.00 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                    
                    TheHdw.Wait rf_mux_delay
                
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA5.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
     Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If

            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
        
         'End 864.00 MHz TX Measure
         
'================================================================================================

        'Begin 433.92 MHz TX Measure

                     
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
               
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(6)  '433.92 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                    
                TheHdw.Wait rf_mux_delay
                
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA6.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
     Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
        

            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
        
         'End 433.92 MHz TX Measure
         
'================================================================================================
         
        'Begin 433.42 MHz TX Measure

                
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(7)  '433.42 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                   
                    TheHdw.Wait rf_mux_delay
                 
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA7.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
     Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
        

            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
    
        
         'End 433.42 MHz TX Measure
         
'================================================================================================
         
        'Begin 418.00 MHz TX Measure

                    
        'Wait for cpuA set in pattern
           Do
           
            'Get cpu Flags state (cpuA)
                Flags = TheHdw.Digital.Patgen.CpuFlags
                CpuFlagBit = cpuA
                
           Loop While (CpuFlagBit = 0)
                
                
If TheExec.Sites.SelectFirst <> loopDone Then          'Use this site loop construct!

    Do
        Site = TheExec.Sites.SelectedSite
                
            
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(Site), 10, DA_Freq(8)  '418.00 MHz
                
                    TheHdw.Wait rf_mux_delay      'RF MUX Speed dependent.
                
                Call MeasDataAXRFandCalcMax(MeasChans(Site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                'Call MeasDataAXRFandCalcMax(MeasChans(site), MeasData, fft_samples, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf", True, IndexMaxPower, False, 0, 0)
                 
                    TheHdw.Wait rf_mux_delay
                 
                    UncalMaxPower(Site) = MaxPowerTemp
                
                    MaxPowerToSubstract(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
                              
                TxPowerAdvMode_DA8.pins("RFOUT").Value(Site) = UncalMaxPower(Site) + (coax_cable_db(Site) + tx_path_db(Site))
    
    
     Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone

End If
            
            Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
    
    
        
         'End 418.00 MHz TX Measure
         
 '================================================================================================
    
      TheHdw.Wait dut_delay 'DUT delay
      TheHdw.Wait rf_mux_delay 'RF Mux dependent
        

            

        Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
        
  'Test and Datalog All
  
         If TheExec.CurrentJob = "f1-prd-std-t48a" Then
            
        
                TheExec.Flow.TestLimit TxPowerAdvMode_DA1, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8699", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA2, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8689", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA3, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8687", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA4, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8683", , , , , , , , tlForceNone
                
                TheExec.Flow.TestLimit TxPowerAdvMode_DA5, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8640", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA6, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4339", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA7, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4334", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA8, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4180", , , , , , , , tlForceNone
        
        ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
                TheExec.Flow.TestLimit TxPowerAdvMode_DA1, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8699_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA2, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8689_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA3, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8687_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA4, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8683_qc", , , , , , , , tlForceNone
                
                TheExec.Flow.TestLimit TxPowerAdvMode_DA5, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8640_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA6, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4339_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA7, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4334_qc", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA8, 6, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4180_qc", , , , , , , , tlForceNone
        
        End If
    
    Exit Function

ErrHandler:
            
                Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear) ' Clearing cpuA allows the pattern to proceed.
                
                TheExec.Flow.TestLimit TxPowerAdvMode_DA1, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8699", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA2, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8689", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA3, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8687", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA4, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8683", , , , , , , , tlForceNone
                
                TheExec.Flow.TestLimit TxPowerAdvMode_DA5, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_8640", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA6, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4339", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA7, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4334", , , , , , , , tlForceNone
                TheExec.Flow.TestLimit TxPowerAdvMode_DA8, 7, 13, , , , unitDb, "%4.1f", "TxAdvPwr_4180", , , , , , , , tlForceNone

    
        
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function rf_tx_off_20ms_t48a(argc As Long, argv() As String) As Long

'Default 20ms transmit off time to sleep mode is tested as a current difference measurement (I_TX_ON - I_TX_OFF). The DUT is powered up in Power-N-Go mode at 433.92MHz OOK mode with DATA = High.
'SCK (or SCK_BIAS) = High
'DATA is set High.
'DUT is powered up. (VBAT = +3.3V)
'DUT is in 433.92 MHz OOK mode with TX ON. VBAT current is measured. (IDD_TX_ON)
'After a 20ms delay, TX current is measured again (should be <1 mA). (IDD_TX_OFF)
'The difference between IDD_TX_ON and IDD_TX_OFF is tested (should be close to IDD_TX_ON) to show that the DUT timer has timed out, taking DUT out of TX mode by 20Ms after DATA goes Low.



    Dim Site As Variant

    
    Dim IDD_TX_ON As New PinListData
    Dim IDD_TX_OFF As New PinListData
    Dim I_TOFFT_20MS As New PinListData
    
    Dim OnSCK_BIAS As Boolean

    Dim oprVolt As Double

    
On Error GoTo ErrHandler
    
    

    If argc < 2 Then
        MsgBox "Error - On rf_tx_off_20ms_t48a - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
    
        OnSCK_BIAS = argv(0)            ' ON SCK_BIAS(1) or SCK (SCK_BIAS(0))
        oprVolt = ResolveArgv(argv(1))  ' Operating Voltage
        
    End If
    
    
        TheExec.Datalog.WriteComment ("============================= MEASURE I_TX_OFF_20MS =============================")
        
        TheHdw.Utility.pins("rlyXTAL").State = utilBitState1 '1 = NI-6652 Ref clock source; 0 = DIB Crystal; DUT PLL operates from site 26MHz crystal
        
    
    oprVolt = 3.3  'Nominal Data Sheet VBAT
    
    'Disconnect DATA, SCK, SCK_BIAS in order
        TheHdw.pins("DATA").InitState = chInitOff
        TheHdw.pins("SCK").InitState = chInitOff
        TheHdw.pins("SCK_BIAS").InitState = chInitOff
        
    'Set DATA to Logic 0
        TheHdw.pins("DATA").InitState = chInitLo
        
        
            If argv(0) = "1" Then       'SCK_BIAS uses 20K weak pullup resistor
            
                'Set SCK_BIAS to Logic 0
                    TheHdw.pins("SCK_BIAS").InitState = chInitLo
                    
                'Set SCK_BIAS to Logic 1, wait 3msec
                    TheHdw.pins("SCK_BIAS").InitState = chInitHi
                    'TheHdw.Wait 0.003
            Else
                'Set SCK to Logic 0
                    TheHdw.pins("SCK").InitState = chInitLo
                      
                'Set SCK to Logic 1, wait 3msec
                    TheHdw.pins("SCK").InitState = chInitHi
                    'TheHdw.Wait 0.003
            End If
         
    'Set DATA to Logic 1
        TheHdw.pins("DATA").InitState = chInitHi

    
    Call cycle_power(0.001, oprVolt, 0.01, 0.01)   'DUT should be in TX mode after this line is excuted
        
     'Initialize failing default test variables
        
        If TheExec.Sites.SelectFirst <> loopDone Then
        
            Do
                Site = TheExec.Sites.SelectedSite
                
                    IDD_TX_ON.AddPin("VBAT").Value(Site) = -0.999999
                    IDD_TX_OFF.AddPin("VBAT").Value(Site) = 0.888888
                    I_TOFFT_20MS.AddPin("VBAT").Value(Site) = -0.777777
                    
            Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
        
        End If
                
                 ' Measure IDD_TX_ON
                 
                With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.Samples = 1
                    TheHdw.Wait 0.01          'Settling Time is here
                    Call .MeasureCurrents(dps100mA, IDD_TX_ON)
                End With
                
            
                TheHdw.pins("DATA").InitState = chInitLo            'DATA Lo = No TX Signal
                
                    TheHdw.Wait (0.01)    '50% of TS Delay is incorporated into DPS measurements
                    
                    
                With TheHdw.DPS.pins("VBAT")
                    .ClearLatchedCurrentLimit
                    .ClearOverCurrentLimit
                    .CurrentRange = dps100mA
                    .CurrentLimit = 0.1
                    TheHdw.DPS.Samples = 1
                    TheHdw.Wait 0.01          'Settling Time is here
                    Call .MeasureCurrents(dps100mA, IDD_TX_OFF)
                End With
                        
        
          If TheExec.Sites.SelectFirst <> loopDone Then
        
            Do
                Site = TheExec.Sites.SelectedSite
                
                I_TOFFT_20MS.AddPin("VBAT").Value(Site) = IDD_TX_ON.AddPin("VBAT").Value(Site) - IDD_TX_OFF.AddPin("VBAT").Value(Site)
                
                If ((IDD_TX_ON.AddPin("VBAT").Value(Site)) >= 0.013 And (IDD_TX_OFF.AddPin("VBAT").Value(Site) <= 0.001)) Then 'Check for valid IDD_TX_ON and IDD_TX_OFF
                    
                    'I_TOFFT_20MS.AddPin("VBAT").Value(site) = IDD_TX_ON.AddPin("VBAT").Value(site) - IDD_TX_OFF.AddPin("VBAT").Value(site)
                    
                    'TheExec.Flow.TestLimit I_TOFFT_20MS, 0.013, 0.02, , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS", , , , , , , , tlForcePass
                
                    'TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , , , , , , , , , , , , tlForcePass
                    
                    If TheExec.CurrentJob = "f1-prd-std-t48a" Then
                    
                        TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS", , , , , , , , tlForcePass
                
                    ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
                    
                         TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS_qc", , , , , , , , tlForcePass
                   
                    End If
                    
                Else
                
                    'TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , , , , , , , , , , , , tlForceFail
                    
                    If TheExec.CurrentJob = "f1-prd-std-t48a" Then
                    
                        TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS", , , , , , , , tlForceFail
                
                    ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
                    
                        TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS_qc", , , , , , , , tlForceFail
                    
                    End If
                    
                End If
                
                    'Debug.Print "IDD_TX_ON = "; IDD_TX_ON.AddPin("VBAT").Value(site); " for Site "; (site)
                    'Debug.Print "IDD_TX_OFF = "; IDD_TX_OFF.AddPin("VBAT").Value(site); " for Site "; (site)
                    'Debug.Print "I_TOFFT_20MS = "; I_TOFFT_20MS.AddPin("VBAT").Value(site); " for Site "; (site)
                    
            Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
        
        End If
    

        'TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS", , , , , , , , tlForcePass
    
    
    Exit Function

ErrHandler:
    

        'TheExec.Flow.TestLimit I_TOFFT_20MS, 0.013, 0.02, , , scaleMilli, unitAmp, "%5.3f", "I_TOFFT_20MS", , , , , , , , tlForceNone
    TheExec.Flow.TestLimit I_TOFFT_20MS, , , , , , , , , , , , , , , , tlForceFail

    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function tx_config_mrf34ta(argc As Long, argv() As String) As Long

    Dim Site As Variant

    On Error GoTo ErrHandler
    
    TheExec.Datalog.WriteComment ("============================= CONFIG & TX MODE ===========================")
    'DUT/DIB Setup:
    'Step -1: set axrf in TX mode
    Dim ExistingSiteCnt As Integer
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    Dim MeasChans() As AXRF_CHANNEL         'Site Array
    Dim MeasFactor As Double
    Dim MaxPowerToSubstract() As Double     'Site Array
    Dim UncalMaxPower() As Double          'Site Array
    
    ReDim MeasChans(0 To ExistingSiteCnt - 1)
    ReDim MaxPowerToSubstract(0 To ExistingSiteCnt - 1)
    ReDim UncalMaxPower(0 To ExistingSiteCnt - 1)
    
    'Dim IDD_TX As New PinListData
       
    'AXRF Channel assign across Sites
    Select Case ExistingSiteCnt
        
    Case Is = 1                                 'added for ITL v1.4.7
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
        
    Case Is = 5
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        
    Case Is = 6
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        
    Case Is = 7
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        
    Case Is = 8
    
        MeasChans(0) = AXRF_CHANNEL_AXRF_CH1
        MeasChans(1) = AXRF_CHANNEL_AXRF_CH3
        MeasChans(2) = AXRF_CHANNEL_AXRF_CH5
        MeasChans(3) = AXRF_CHANNEL_AXRF_CH7
    
        MeasChans(4) = AXRF_CHANNEL_AXRF_CH2
        MeasChans(5) = AXRF_CHANNEL_AXRF_CH4
        MeasChans(6) = AXRF_CHANNEL_AXRF_CH6
        MeasChans(7) = AXRF_CHANNEL_AXRF_CH8
        
    Case Else
        MsgBox "Error in [tx_config_mrf34ta]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo ErrHandler
        
    End Select
    
    Dim TxPower434 As New PinListData

    Dim IDD_TX As New PinListData

    Dim MeasPower() As Double '(1) DEBUG
    Dim MeasData() As Double
    Dim nSiteIndex As Long
    
    Dim IndexMaxPower As New SiteDouble
    Dim MaxPowerTemp As Double
    
    Dim TestFreq As Double
    Dim DATAPinLevel As Integer
    Dim OOKDepthMode As Boolean
    
    Dim oprVolt As Double

    If argc < 2 Then
        MsgBox "Error - tx_config_mrf34ta - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
        TestFreq = argv(0)              ' What is testing Freq?
        DATAPinLevel = argv(1)          ' What is data being use to test?
        oprVolt = argv(2)
        
    End If

    TheHdw.Utility.pins("XTAL_SW").State = utilBitState1   '1 = NI-6652 Ref clock source; 0 = DIB Crystal
    
    Call read_cal_factors                   'RF Calibration Offsets
    
    Call Ref_Clock_On(DUTRefClkFreq)        'Reference Clock Set to 13.56 MHz
    
    Call itl.Raw.AF.AXRF.SetMeasureSamples(2048)
    
        TheHdw.Wait 0.05
    '=================================================================================================================================
    Select Case TestFreq
    Case 434000000
        TxPower434.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower434.pins("RFOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        'TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 433.92 MHZ
    
    Case Else
        TxPower434.AddPin ("RFOUT")
        
        For nSiteIndex = 0 To ExistingSiteCnt - 1
            TxPower434.pins("RFOUT").Value(nSiteIndex) = -90
        Next nSiteIndex
        
        'TheHdw.pins("SCK").InitState = chInitHi         'SCK Setup for 433.92 MHZ
        
    End Select

    'Pre Force Fail
    
    'Trigger Capture for each site, wait 1ms between captures
    'Setup DUT in Mode 1
    
    TheHdw.Wait (0.002)
    
    'Call cycle_power(0.001, oprVolt, 0.01, 0.01)
    
    Select Case DATAPinLevel

    Case 0
        
        'TheHdw.Wait (0#)
        

    
        TheHdw.Digital.Patterns.Pat("./Patterns/CFG_TX_ON.PAT").start ("Start")
        
        TheHdw.Wait (0.25)  'DEBUG Critical for TW101 TX
        
        TheHdw.pins("SDI").InitState = chInitHi            'SDI Hi = TX Signal at 433.92 MHz (not measured)
        
        
        'TheHdw.pins("SDI").InitState = chInitLo            'Stop TX at 433.92 MHz
        
        'Run pattern to put DUT into TX Mode @ 434 MHz, +10 dBm Power
        
        
        
        TheHdw.Wait (0.01)  'DEBUG TUNE
        
    Case Else   'Dummy  For force fail purpose

        
        'Run pattern to put DUT into TX Mode @ 434 MHz, +0 dBm Power
    
        'TheHdw.Digital.Patterns.pat("./Patterns/CFG_TX_ON.PAT").start ("Start")
        
        'TheHdw.pins("SDI").InitState = chInitHi            'DATA Hi = TX Signal
        
        'TheHdw.Wait (0.002)
        
        'TheHdw.pins("SDI").InitState = chInitLo            'Stop TX at 433.92 MHz
        
        'TheHdw.Wait (0.01)
        
    End Select
    
    TheHdw.Wait (0.002)
    
    '#If Debug_advanced Then

        For nSiteIndex = 0 To ExistingSiteCnt - 1
        
            If TheExec.Sites.Site(nSiteIndex).Active = True Then
                
                itl.Raw.AF.AXRF.MeasureSetup MeasChans(nSiteIndex), 10, TestFreq
                
                TheHdw.Wait (0.2)      'RF MUX Speed depended. DEBUG
                
                Call MeasDataAXRFandCalcMax(MeasChans(nSiteIndex), MeasData, 1024, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN, MaxPowerTemp, False, "rf")  'True plots waveform
                
                UncalMaxPower(nSiteIndex) = MaxPowerTemp
                
                MaxPowerToSubstract(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                'MaxPower.Pins("RFOUT").Value(nSiteIndex) = MaxPowerTemp + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
                
                Select Case TestFreq
                Case 434000000
                    TxPower434.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))

                Case Else       'Dummy  for force fail purpose
                    TxPower434.pins("RFOUT").Value(nSiteIndex) = UncalMaxPower(nSiteIndex) + (coax_cable_db(nSiteIndex) + tx_path_db(nSiteIndex))
        
                End Select

            End If
            
        Next nSiteIndex

    'Run pattern to stop DUT transmitting
                
    'TheHdw.Digital.Patterns.pat("./Patterns/TX_OFF_IDLE.PAT").start ("Start")
    
    Select Case TestFreq
    Case 434000000
        TheExec.Flow.TestLimit TxPower434, 0, 10, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone
        
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit TxPower434, 0, 10, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone

    End Select

    ' ============================= I_TX MEASURE ==================================================

    
    ' Perform DPS set up
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps100mA
        .CurrentLimit = 0.1
        TheHdw.DPS.Samples = 1
    End With
    
'    Call cycle_power(0, 3.3, 0.1, 0.01)
     
    'TheHdw.Digital.Patterns.pat("./Patterns/CFG_TX_ON.PAT").start ("Start")
       
    'TheHdw.Wait (0.005)      ' Wait For Device Ready
       
    'TheHdw.pins("SCK").InitState = chInitLo
    'TheHdw.pins("SDI").InitState = chInitLo
    
    'TheHdw.pins("SCK").StartState = chStartHi   'Set to Sleep
    'TheHdw.pins("SDI").StartState = chStartHi   'Set to Sleep
    
    'TheHdw.Wait (0.03 + 0.02)      ' Settling Time 20mS as Spec metion     '30mS to compensate delay discharge
    
    'Make Current Measurement
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps100mA, IDD_TX)
       
    TheExec.Flow.TestLimit IDD_TX, 0.01, 0.025, , , scaleMilli, unitAmp, , "I_TX", , , , , , , , tlForceNone
       
    '=========================================================================================
        'Run pattern to stop DUT transmitting
                
    'TheHdw.Digital.Patterns.pat("./Patterns/TX_OFF_IDLE.PAT").start ("Start")
    
    Exit Function

ErrHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower434.pins("RFOUT").Value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 434000000
        TheExec.Flow.TestLimit TxPower434, 0, 10, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone
  
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit TxPower434, 5, 15, , , , unitDb, , "TxPower_434", , , , , , , , tlForceNone

    End Select
    
   'TheHdw.Digital.Patterns.pat("./Patterns/TX_OFF_IDLE.PAT").start ("Start")
    
    If AbortTest Then Exit Function Else Resume Next
    
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
    
    On Error GoTo ErrHandler
    
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
        MsgBox "Error in [rn2483_tx868_cw]" & vbCrLf & _
               "Existnumber is not support by ITL", _
               vbCritical + vbOKOnly, _
               "Interpose Setup Error"
        GoTo ErrHandler
        
    End Select
    

    Call enable_store_inactive_sites 'For Pass/Fail LEDs
    

    If argc < 2 Then
        MsgBox "Error - On rn2483_tx868_cw - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
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
        
        
        
                
        
        
'Pattern loops after setting cpuA. VBT loops waiting for the pattern to set.

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

 TheExec.Datalog.WriteComment ("==================  TX868_CW_PWR =================== ")
 
    Select Case TestFreq
    
    Case 868300000
        TheExec.Flow.TestLimit I_TX868_CW, 0.035, 0.085, , , scaleMilli, unitAmp, , "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone
        
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit I_TX868_CW, 0.03, 0.09, , , scaleMilli, unitAmp, , "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 17, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone

    End Select



 
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function

ErrHandler:
    
    For nSiteIndex = 0 To ExistingSiteCnt - 1
        TxPower868.pins("RFHOUT").Value(nSiteIndex) = -90
    Next nSiteIndex
        
    Select Case TestFreq
    Case 868300000
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, , "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone
  
    Case Else      'Dummy for force fail purpose
        TheExec.Flow.TestLimit I_TX868_CW, 0.02, 0.06, , , scaleMilli, unitAmp, , "I_TX868_CW", , , , , , , , tlForceNone
        TheExec.Flow.TestLimit TxPower868, 10, 16, , , , unitDb, , "TxPower_868", , , , , , , , tlForceNone

    End Select

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function rn2483_fsk_pkt_rcv(argc As Long, argv() As String) As Long

'Multisite Packets Received Test using digital channel triggering of 3025C Modulation Source. This function is NOT a PER test! When CRC is Off in the module radio driver,
'assuming received signal strength is sufficient to recognize PREAMBLE and SYNC, packet data will be sent from DUT FIFO to the host on UART_TX.
'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received packets. However,
'when the DUT cannot receive the packet, no FIFO data will be sent to the UART host, and the pattern will time out with 100 forced fails.
'The MATCH LOOP does not work when a packet loop count (>1) is used inside the pattern, therefore a VBT FOR LOOP is coded for multiple packets.
' The AXRF Modulation Source is triggered by the digital pattern using MW_SRC_TRIG.

    Dim SrcChans(3) As AXRF_CHANNEL

    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    
    Dim Gate As Long
    Dim Edge As Long
    Dim nSiteIndex As Long

    Dim PacketCount(3) As Long
    Dim patgen_fails(3) As Long
    
    Dim i As Long
    Dim pkt_sent_count As Long
    
    Dim PKTs_RCVd As New PinListData
    
    
 On Error GoTo ErrHandler
 
    rn2483_fsk_pkt_rcv = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
     

        Call read_cal_factors                   'RF Calibration Offsets Note: AXRF calibration performed with same coax cables and RF junction boxes as production AXRF with DIB

    
    PKTs_RCVd.AddPin ("RFHOUT")
    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize packets received
        
        PKTs_RCVd.pins("RFHOUT").Value(nSiteIndex) = 0
        
    Next nSiteIndex
    
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
     

    
For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Site Loop

    If TheExec.Sites.Site(nSiteIndex).Active Then
        
       patgen_fails(nSiteIndex) = 0
       
      
        
 
        '#If Connected_to_AXRF Then
        
            'Working Hardware Triggered Modulation Code
            
            'If FIRSTLOAD = True Then
            
                Call itl.Raw.AF.AXRF.LoadModulationFile(SrcChans(nSiteIndex), ModFilePath) 'Separate Loads for each site?
                TheHdw.Wait (0.1)
                'FIRSTLOAD = False
                'Debug.Print ModFilePath
                
            'End If
            
                Call itl.Raw.AF.AXRF.ModulationTriggerArm(SrcChans(nSiteIndex), afSigGenDll_rmRoutingMatrix_t_afSigGenDll_rmFRONT_SMB, Gate, Edge)
                
                
                Call itl.Raw.AF.AXRF.StartModulation(SrcChans(nSiteIndex), ModFilePath)
                
                TheHdw.Wait (0.05)
                
            
                
                Call itl.Raw.AF.AXRF.Source(SrcChans(nSiteIndex), -85, 868300000#) 'Assumes AXRF calibration performed with DIB cables and AXRF interface junction box.
                                                                                    'Setting is ~2 dB above highest passing threshhold for functional DUTS.
    
        '#End If
        
               For pkt_sent_count = 1 To 5 '5 packets sent
                     
                    With TheHdw.pins("UART_CTS,MW_SRC_TRIG") 'Initialize Logic Analyzer Trigger
                        .InitState = chInitLo
                        .StartState = chStartLo
                    End With
            
                    TheHdw.Wait (0.01)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
        
                
                
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one_rev.pat").Run ("start_fsk_pkt_one_rev")
                
        
                    'TRAP HERE for pattern debug'
                    
                    Call TheHdw.Digital.Patgen.HaltWait
        
                'TheHdw.Wait 0.01
        
                        Debug.Print "Site = "; nSiteIndex
                        Debug.Print "Pkts = "; pkt_sent_count
                        Dim pkt_fails As Long
                        pkt_fails = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_Pkts = "; pkt_fails
                        
                   'PKTs_RCVd
                   
                   patgen_fails(nSiteIndex) = Int(50 - TheHdw.Digital.Patgen.FailCount) 'Allow 50 Match Loop FailCount
                   
                        If patgen_fails(nSiteIndex) > 0 Then
                        
                             PacketCount(nSiteIndex) = PacketCount(nSiteIndex) + 1
                        
                        Else
                        
                             PacketCount(nSiteIndex) = PacketCount(nSiteIndex)  'Pattern forces 100 FailCount for no packet received (timeout)
                        
                        End If
                   
                   
                   PKTs_RCVd.AddPin("RFHOUT").Value(nSiteIndex) = PacketCount(nSiteIndex)
        
            Next pkt_sent_count
    
    End If 'Sites Active
    

    
Next nSiteIndex 'Site Loop
    

    TheExec.Datalog.WriteComment ("================== PKTs_RCVd =================== ")
   
    TheExec.Flow.TestLimit PKTs_RCVd, 4.5, 5.5, , , , unitNone, , "PktCnt", , , , , , , , tlForceNone
    
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

ErrHandler:

    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_fsk_pkt_rcv")
    Call TheExec.ErrorReport
    rn2483_fsk_pkt_rcv = TL_ERROR
    
End Function

Public Function rn2483_id(argc As Long, argv() As String) As Long

'Multisite LoRa module ID test. After a reset, a functional DUT sends the UART host its ID (and FW revision time and date).

'Because of the MATCH LOOP used in the pattern, there will be some pattern FailCounts for correctly received ID. If no ID is received, however,
'the pattern will time out with 100 forced fails.

'Some modules take more than 100msec to respond to system reset than others. A fast and slow response pattern is used to check which type of module
' is being tested. Passing the slow OR the fast pattern will pass the ID test by finding the start bit of the response.


    Dim print_var As Long
    
    Dim ModFilePath As String
    Dim xTPPath As String
    

    Dim nSiteIndex As Long

    Dim ValidityCount(3) As Long
    Dim ValidityCountSlow(3) As Long 'Slow responding modules
    Dim ValidityCountFast(3) As Long 'Fast responding modules
    
    
    Dim patgen_fails(3) As Long
    Dim patgen_fails_slow(3) As Long
    Dim patgen_fails_fast(3) As Long
    
    
    Dim i As Long

    
    Dim ID_Valid As New PinListData
    
    
 On Error GoTo ErrHandler
 
    rn2483_id = TL_SUCCESS
 
    
        Call enable_store_inactive_sites 'For Pass/Fail LEDs
     


    ID_Valid.AddPin ("RFHOUT")
    
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Initialize ID_Valid
        
        ID_Valid.pins("RFHOUT").Value(nSiteIndex) = 0
        
    Next nSiteIndex
    
'    If Right(ActiveWorkbook.Path, 1) = "\" Then
'        xTPPath = ActiveWorkbook.Path
'    Else
'        xTPPath = ActiveWorkbook.Path & "\"
'    End If
    
    xTPPath = "D:\LoRa"
    xTPPath = ActiveWorkbook.path
    
TheHdw.Wait (0.2) 'Wait for DUT POR to complete.
    
For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1  'Site Loop used because pattern MATCH LOOP will be unique for each site.

    If TheExec.Sites.Site(nSiteIndex).Active Then
        
       'ValidityCount(nSiteIndex) = 0

       ValidityCountFast(nSiteIndex) = 0
       ValidityCountSlow(nSiteIndex) = 0
       
       patgen_fails_fast(nSiteIndex) = 0
       patgen_fails_slow(nSiteIndex) = 0
                 
'                    With TheHdw.pins("UART_CTS") 'Initialize Logic Analyzer Trigger
'                        .InitState = chInitLo
'                        .StartState = chStartLo
'                    End With
'
'                    TheHdw.Wait (0.03)
            
            
                    'If TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").IsPatLoaded = memNone Then
                    'TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_tx868_fsk_pkt_one.pat").Load
                    'End If
        
                
                
                'Run ID patterns
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_slow.pat").Run ("start_uart_id_slow")
                
                Call TheHdw.Digital.Patgen.HaltWait
                
                    patgen_fails_slow(nSiteIndex) = TheHdw.Digital.Patgen.FailCount
                    
                        Debug.Print "Site = "; nSiteIndex
                        Dim fails_ids As Long
                        fails_ids = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Slow = "; fails_ids
                
                TheHdw.Digital.Patterns.Pat(xTPPath & "\patterns\uart_rn2483_id_fast.pat").Run ("start_uart_id_fast")
        
                    'TRAP HERE for pattern debug'
                
                Call TheHdw.Digital.Patgen.HaltWait
              
                    patgen_fails_fast(nSiteIndex) = TheHdw.Digital.Patgen.FailCount
                    
                        Debug.Print "Site = "; nSiteIndex
                        Dim fails_idf As Long
                        fails_idf = TheHdw.Digital.Patgen.FailCount
                        Debug.Print "FailCount_ID_Fast = "; fails_idf
                        
                   'FailCount Interpretation
                   
                   ValidityCountFast(nSiteIndex) = Int(55 - patgen_fails_fast(nSiteIndex))
                   
                   ValidityCountSlow(nSiteIndex) = Int(55 - patgen_fails_slow(nSiteIndex))
                        
                        
                        If (ValidityCountFast(nSiteIndex) > 0) Or (ValidityCountSlow(nSiteIndex) > 0) Then
                        
                            ID_Valid.AddPin("RFHOUT").Value(nSiteIndex) = 1
                        
                        Else
                        
                            ID_Valid.AddPin("RFHOUT").Value(nSiteIndex) = 0
                            

                        End If
                        

                   
                   

    
    End If 'Sites Active
    

    
Next nSiteIndex 'Site Loop

    'Call TheHdw.Digital.Patgen.Halt
    

    TheExec.Datalog.WriteComment ("==================  READ_MODULE_ID  =================== ")
   
    TheExec.Flow.TestLimit ID_Valid, 0.5, 1.5, , , , unitNone, , "ID", , , , , , , , tlForceNone
    
    'Cleanup

    
        'Call TheHdw.Digital.Patgen.Halt
    
        Call disable_inactive_sites 'For Pass/Fail LEDs
 
    Exit Function

ErrHandler:

    Call TheHdw.Digital.Patgen.Halt
    On Error GoTo 0
    Call TheExec.ErrorLogMessage("Function Error: rn2483_id")
    Call TheExec.ErrorReport
    rn2483_id = TL_ERROR
    
End Function

Public Function rn2483_i_sleep(argc As Long, argv() As String) As Long

    Dim Site As Variant
    
    Dim I_SLEEP As New PinListData
    
    Dim oprVolt As Double
    Dim dut_delay As Double
      
    Dim nSiteIndex As Long
    Dim Flags As Long
    Dim FlagsSet As Long, FlagsClear As Long
    
    Dim ExistingSiteCnt As Integer
    
    ExistingSiteCnt = TheExec.Sites.ExistingCount
    
    On Error GoTo ErrHandler
    
    Call enable_store_inactive_sites 'For Pass/Fail LEDs

    If argc < 1 Then
        MsgBox "Error - On rn2483_i_sleep - Wrong Argument Assigned", , "Error"
        GoTo ErrHandler
    Else
        oprVolt = argv(0) '3.3
        
    End If
    
TheExec.Datalog.WriteComment ("============================= MEASURE I_SLEEP =====================================")
    
        oprVolt = ResolveArgv(argv(0))  ' Operating Voltage - check TI Parms
        
        'dut_delay = 0.025
        
        'TheHdw.Wait 0.5 'DEBUG
        
        
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
        
        
        
Flag_Loop:

   
Flags = TheHdw.Digital.Patgen.CpuFlags

'Debug.Print "Flags ="; Flags
        
 If (Flags = 1) Then GoTo End_flag_loop
 
 If (Flags = 0) Then GoTo Flag_Loop
 
 
End_flag_loop:
        
        
  For nSiteIndex = 0 To ExistingSiteCnt - 1


        
    With TheHdw.DPS.pins("VBAT")
        .ClearLatchedCurrentLimit
        .ClearOverCurrentLimit
        .CurrentRange = dps500ua
        .CurrentLimit = 0.1
        TheHdw.DPS.Samples = 1
    End With
        
        'TheHdw.Wait (0.03) 'CRITICAL
        
                'Measure Current
    Call TheHdw.DPS.pins("VBAT").MeasureCurrents(dps500ua, I_SLEEP)
        
    Next nSiteIndex
    


    FlagsSet = 0
    FlagsClear = cpuA

Call TheHdw.Digital.Patgen.Continue(FlagsSet, FlagsClear)

    
    Call TheHdw.Digital.Patgen.HaltWait
    
        'If TheExec.CurrentJob = "f1-prd-std-t48a" Then
        
            TheExec.Flow.TestLimit I_SLEEP, 0.00001, 0.0005, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone
              
        'ElseIf TheExec.CurrentJob = "q1-prd-std-t48a" Then
        
            'TheExec.Flow.TestLimit I_SLEEP, -0.00000011, 0.0000011, , , scaleMicro, unitAmp, "%4.0f", "RN2483_I_SLEEP_qc", , , , , , , , tlForceNone
        
        'End If
        
    Call TheHdw.Digital.Patgen.Halt
        
    Call disable_inactive_sites 'For Pass/Fail LEDs
    
    Exit Function
    

ErrHandler:


    I_SLEEP.AddPin ("VBAT")
            
      For nSiteIndex = 0 To ExistingSiteCnt - 1
      
    I_SLEEP.pins("VBAT").Value(nSiteIndex) = 9999 'Failing initialization value
    
      Next nSiteIndex

        
         TheExec.Flow.TestLimit I_SLEEP, -0.0000001, 0.000001, , , scaleNano, unitAmp, "%4.0f", "RN2483_I_SLEEP", , , , , , , , tlForceNone

    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

