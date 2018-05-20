Attribute VB_Name = "Exec_IP_Module"
Option Explicit

Public AbortTest As Boolean
Public test_i As New DspWave
Public test_q As New DspWave

Public Const MaxSites = 32

Public lngSitesExisting As Long
Public lngSitesStarting As Long
Public FIRSTRUN As Boolean
Public TxMTX(16, MaxSites) As Double

'Public gSuppression_C_hi(MaxSites) As Double
'Public gSuppression_SB_hi(MaxSites) As Double
'Public gSuppression_C_lo(MaxSites) As Double
'Public gSuppression_SB_lo(MaxSites) As Double


' This module contains empty Exec Interpose functions (see online help
' for details).  These are here for convenience and are completely optional.
' It is not necessary to delete them if they are not being used, nor is it
' necessary that they exist in the program.


' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
'===========================================================
'Function OnTesterInitialized()
'    On Error GoTo errHandler
'
'    ' Put code here
'
'
'    Exit Function
'errHandler:
'    ' OnTesterInitialized executes before TheExec is even established so nothing
'    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
'    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
'    ' need to be debugged carefully.
'    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
'           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
'End Function
'============================================================

'' Immediately at the conclusion of the load process.
'' Do not program test system hardware from this function.
'Function OnProgramLoaded()
'
'    On Error GoTo errHandler
'
'    'ITLValidated = False        'Default For first Validate - PT Workaround
'    'ITLStarted = False          'Default For first Validate - PT Workaround
'
'    'Call AddDatalogButtons                       'load Andrew Datalog button
'
'
'
'    ' Put code here
'    ''    thehdw.ExtUtility.EnableDIBLoopCheck = True
'
'    ' Enable a more detailed validation of timing and levels:
'    ''    Call thehdw.Digital.Patterns.Pat("").EnableExtendedValidation(True)
'
'    Exit Function
'errHandler:
'    MsgBox "Error encountered in Exec Interpose Function OnProgramLoaded" + vbCrLf + _
'           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
'End Function
'=============================================================

' Immediately at the conclusion of the validate process. Called only if validation succeeds.
'Function RFOnProgramValidated_T39A()
'
'    Dim i As Long, loopStatus As Long
'
'    Dim site As Long
'
'    On Error GoTo errHandler
'
'    AbortTest = True
'
'    Call ITLOnProgramValidated          'Self Check inside. If Done once, no more execute
'
'    TheHdw.DIB.powerOn = True
'
'    TheHdw.Digital.ACCalExcludePins ("RF_HW_Dig_Trig")
'
'    '''' Need to verify if you need these files fo J750
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\TX_loop.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\TX_loop_no_match.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\TX_burst.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\VHFAC_ReSync_00.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\VHFAC_ReSync_02.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\VHFAC_ReSync_04.PAT").Load
'    ''    thehdw.Digital.Patterns.Pat(".\Patts\IDD_IPD_pop.PAT").Load
'
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_check_chipversion_V1A.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_read_applicationbits.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_read_testbits.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_read_freqchangebits.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_recovery_command.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_writeDA_Toff_delay_2ms.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\T39A_writeDA_Toff_delay_20ms.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\TX_ON_390.PAT").Load
'        TheHdw.Digital.Patterns.pat(".\patterns\TX_OFF_434.PAT").Load
'
'    FIRSTRUN = True
'
'    TheExec.Datalog.Setup.LotSetup.PartType = TheExec.CurrentPart
'    TheExec.Datalog.Setup.DatalogSetup.HeaderEveryRun = True
'    TheExec.Datalog.ApplySetup
'
'    Exit Function
'errHandler:
'    MsgBox "Error encountered in Exec Interpose Function RFOnProgramValidated" + vbCrLf + _
'           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
'End Function

'==============================================================================
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)

Function OnProgramStarted()
    On Error GoTo ErrHandler
    
    Call ITLOnProgramStarted        'Self Check inside. If Done once, no more execute
    
'    CreateZigbeeAnalysisObjects
    ' Put code here
    lngSitesStarting = TheExec.Sites.StartingCount
    
    Dim xSite As Long
    
    ReDim sites_active(TheExec.Sites.ExistingCount - 1)
    ReDim sites_inactive(TheExec.Sites.ExistingCount - 1)
    ReDim sites_tested(TheExec.Sites.ExistingCount - 1)
    
    For xSite = 0 To TheExec.Sites.ExistingCount - 1
        If TheExec.Sites.Site(xSite).Active Then
            sites_tested(xSite) = False
            sites_inactive(xSite) = False
            sites_active(xSite) = True
        Else
            sites_tested(xSite) = False
            sites_inactive(xSite) = True
            sites_active(xSite) = False
        End If
    Next xSite

    Exit Function
ErrHandler:
    MsgBox "Error encountered in Exec Interpose Function OnProgramStarted" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function
'================================================================================
' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.

Function OnProgramEnded()

    On Error GoTo ErrHandler


Dim led_level_high As Double
Dim led_level_low As Double

'Debug.Print "OnProgramEnded..."
'This function will set FAIL LEDs based on logic generated within the test flow.

led_level_high = 4.99
led_level_low = 0.01



Call TheExec.Sites.SetAllActive(True)     'Activate all sites

If TheExec.Sites.Site(0).Active = True Then
    
     If ((site0_failed = True Or Passing_Site_Flag = False)) Then
                
                TheHdw.PPMU.pins("RED2_ON").Connect
                TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.Wait (LED_PULSE)
                TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED2_ON").Disconnect
                           
     End If
     
End If
    
If TheExec.Sites.Site(1).Active = True Then

    If ((site1_failed = True Or Passing_Site_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED3_ON").Connect
                TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.Wait (LED_PULSE)
                TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED3_ON").Disconnect
    
    End If
    
End If

If TheExec.Sites.Site(2).Active = True Then

    If ((site2_failed = True Or Passing_Site_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED4_ON").Connect
                TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.Wait (LED_PULSE)
                TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED4_ON").Disconnect
    
    
    End If
    
End If

If TheExec.Sites.Site(3).Active = True Then
    
    If ((site3_failed = True Or Passing_Site_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED1_ON").Connect
                TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.Wait (LED_PULSE)
                TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED1_ON").Disconnect
    
    End If
    
End If

        TheExec.Sites.SetAllActive (False) 'Deactivate all sites


    Exit Function
    
ErrHandler:
    MsgBox "Error encountered in Exec Interpose Function OnProgramEnded" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function



'==============================================================================
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()

    On Error GoTo ErrHandler
'
'    ' Put code here
'
    Exit Function
ErrHandler:
    MsgBox "Error encountered in Exec Interpose Function OnTDRCalibrated" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function
'===============================================================================
Function RFOnProgramValidated_TW101() 'MRF34TA

   Dim i As Long, loopstatus As Long

    Dim Site As Long

    On Error GoTo ErrHandler
    
    AbortTest = True
    
    Call ITLOnProgramValidated          'Self Check inside. If Done once, no more execute
    
    TheHdw.DIB.powerOn = True
    
    'TheHdw.Digital.ACCalExcludePins ("RF_HW_Dig_Trig")
    TheHdw.Digital.ACCalExcludePins ("SPI_EN,INTERIOR_TP_PINS,LED_PINS,MW_TRIG_PINS")   'TW101
    
        TheHdw.Digital.Patterns.Pat(".\patterns\CFG_TX_ON.PAT").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\TX_OFF_IDLE.PAT").Load


    FIRSTRUN = True
    
    TheExec.Datalog.setup.LotSetup.PartType = TheExec.CurrentPart
    TheExec.Datalog.setup.DatalogSetup.HeaderEveryRun = True
    TheExec.Datalog.ApplySetup

    Exit Function
    
ErrHandler:
    MsgBox "Error encountered in Exec Interpose Function RFOnProgramValidated_TW101" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description

End Function
'===============================================================================
Function RFOnProgramValidated_LoRa() 'RN2483

   Dim i As Long, loopstatus As Long

    Dim Site As Long

    On Error GoTo ErrHandler
    
    AbortTest = True
    
    Call ITLOnProgramValidated          'Self Check inside. If Done once, no more execute
    
    TheHdw.DIB.powerOn = True
    
    
    TheHdw.Digital.ACCalExcludePins ("SPI_EN,INTERIOR_TP_PINS,LED_PINS,MW_TRIG_PINS")   'TW101
    
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_id_slow").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_id_fast").Load
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_sleep").Load
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_cw").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx_cw_off").Load
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_per_fsk_crc_on").Load
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_per_fsk_crc_off").Load
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_nocrc_pkt_rcv").Load
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_fsk_pkt_trig").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_fsk_pkt_one").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_fsk_pkt_one_rev").Load
        
    FIRSTRUN = True
    FIRSTLOAD = True
    
    
    TheExec.Datalog.setup.LotSetup.PartType = TheExec.CurrentPart
    TheExec.Datalog.setup.DatalogSetup.HeaderEveryRun = True
    TheExec.Datalog.ApplySetup

    Exit Function
    
ErrHandler:
    MsgBox "Error encountered in Exec Interpose Function RFOnProgramValidated_LoRa" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description

End Function
