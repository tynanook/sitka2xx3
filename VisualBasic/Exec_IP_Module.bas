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
    On Error GoTo errHandler
    
    Dim lockedStatus As Long
    Dim LogInName As String
    
    LogInName = Application.UserName
    
'    If (LogInName = "finalop") Then
'        TheExec.Datalog.Setup.LotSetup.TestMode = TL_LOTPRODMODE
'    Else
'        TheExec.Datalog.Setup.LotSetup.TestMode = TL_LOTENGMODE
'    End If

    lockedStatus = -1
         
    'Initiate AXRF Lock status check for primary AXRF subsystem
    lockedStatus = TevAXRF_CheckLocked(0)
  
    'Evaluate if both source generator and digitizer is locked to the 10 MHz clock
    If (lockedStatus <> 0) Or (Initialize_status <> 0) Then
        AXRF_Error_Flag = True
    Else
        Call niSync_init("PXI44::15::INSTR", True, True, Dev1)  'Initialize NI-6652 and create NI-Sync Driver session
    End If
'    CreateZigbeeAnalysisObjects
    ' Put code here
    lngSitesStarting = TheExec.Sites.StartingCount

    Exit Function
errHandler:
    MsgBox "Error encountered in Exec Interpose Function OnProgramStarted" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function
'================================================================================
' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.

Function OnProgramEnded()


On Error GoTo errHandler


Dim led_level_high As Double
Dim led_level_low As Double

'Debug.Print "OnProgramEnded..."
'This function will set FAIL LEDs based on logic generated within the test flow.

led_level_high = 4.99
led_level_low = 0.01

    'Reset all attributesd of the NI Clock
    Call niSync_reset(Dev1)
    
    'Close Ni-6652 and close all handles
    Call niSync_close(Dev1)   ' Closes the NI-Sync I/O session and destroys all its attributes from Dev1 handle 'MJD. 020516


'DEBUG 07022015 Passing_Site0-3_Flag logic added here. Must validate Passing Site Flag with at least one of the individual passing sites.

Call TheExec.Sites.SetAllActive(True)     'Activate all sites

If TheExec.Sites.site(0).Active = True Then
    
     If (sites_tested(0) = True And (site0_failed = True Or Passing_Site0_Flag = False)) Then
     
     
                
                TheHdw.PPMU.pins("RED1_ON").Connect
                TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.wait (LED_PULSE)
                TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED1_ON").Disconnect
                If 0 Then Debug.Print "Site 0 FAILED" ' 20170216 - ty added if 0
                           
     End If
     
End If
    
If TheExec.Sites.site(1).Active = True Then

    If (sites_tested(1) = True And (site1_failed = True Or Passing_Site1_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED2_ON").Connect
                TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.wait (LED_PULSE)
                TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED2_ON").Disconnect
                If (0) Then Debug.Print "Site 1 FAILED"  ' 20170216 - ty added if 0
    
    End If
    
End If

If TheExec.Sites.site(2).Active = True Then

    If (sites_tested(2) = True And (site2_failed = True Or Passing_Site2_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED3_ON").Connect
                TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.wait (LED_PULSE)
                TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED3_ON").Disconnect
                If 0 Then Debug.Print "Site 2 FAILED" ' 20170216 - ty added if 0
    
    
    End If
    
End If

If TheExec.Sites.site(3).Active = True Then
    
    If (sites_tested(3) = True And (site3_failed = True Or Passing_Site3_Flag = False)) Then
    
                TheHdw.PPMU.pins("RED4_ON").Connect
                TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_high
                TheHdw.wait (LED_PULSE)
                TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_low
                TheHdw.PPMU.pins("RED4_ON").Disconnect
                If 0 Then Debug.Print "Site 3 FAILED" ' 20170216 - ty added if 0
    
    End If
    
End If

        TheExec.Sites.SetAllActive (False) 'Deactivate all sites


    Exit Function
    
errHandler:
    MsgBox "Error encountered in Exec Interpose Function OnProgramEnded" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function



'==============================================================================
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()

    On Error GoTo errHandler
'
'    ' Put code here
'
    Exit Function
errHandler:
    MsgBox "Error encountered in Exec Interpose Function OnTDRCalibrated" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description
End Function
'===============================================================================
Function RFOnProgramValidated_TW101() 'MRF34TA

   Dim i As Long, loopstatus As Long

    Dim site As Long

    On Error GoTo errHandler
    
    AbortTest = True
    
    Call ITLOnProgramValidated          'Self Check inside. If Done once, no more execute
    
    TheHdw.DIB.powerOn = True
    
    'TheHdw.Digital.ACCalExcludePins ("RF_HW_Dig_Trig")
    TheHdw.Digital.ACCalExcludePins ("SPI_EN,INTERIOR_TP_PINS,LED_PINS,MW_TRIG_PINS")   'TW101
    
        TheHdw.Digital.Patterns.Pat(".\patterns\CFG_TX_ON.PAT").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\TX_OFF_IDLE.PAT").Load


    FIRSTRUN = True
    
    TheExec.DataLog.Setup.LotSetup.PartType = TheExec.CurrentPart
    TheExec.DataLog.Setup.DatalogSetUp.HeaderEveryRun = True
    TheExec.DataLog.ApplySetup

    Exit Function
    
errHandler:
    MsgBox "Error encountered in Exec Interpose Function RFOnProgramValidated_TW101" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description

End Function
'===============================================================================
Function RFOnProgramValidated_LoRa() 'RN2483

   Dim i As Long, loopstatus As Long
   Dim site As Long
    
    On Error GoTo errHandler
    
    AbortTest = True
    Initialize_status = -1
    ReferenceTime = 0

    On Error GoTo errHandler
    
    'Close all AXRF handles first prior initializion
    Call TevAXRF_Close
    
    ReferenceTime = TheExec.Timer           ' Initiate timer
  
    Initialize_status = TevAXRF_Initialize  ' Initializes the AXRF susbsystem
    
    ElapsedTime = TheExec.Timer(ReferenceTime) 'Check elapsed time of initialization

    'Evaluate if AXRF initialization is successful
    'An error code returned bu the API idnetifies any AXRF Module
    'exhibiting a problem during initialization.
    'Refer to AXRF RF subsystem and Autocal Unit User Manual
    
    If (Initialize_status <> 0) Or (ElapsedTime > 20) Then
        AXRFInitialized = False
        AXRF_Error_Flag = True
    Else
        AXRF_Error_Flag = False
        AXRFInitialized = True
    End If
    
 
    TheHdw.DIB.powerOn = True
    
    TheHdw.wait (0.1)
    
    'Call init_leds
    
    
    TheHdw.Digital.ACCalExcludePins ("LED_PINS,MW_TRIG_PINS")   'RN2903
    
    
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2903_tx915_fsk_pkt_one_rev2").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_fsk_pkt_one_rev2").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx433_fsk_pkt_one_rev2").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_id_ver2").Load
        
        
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_gpio_full").Load
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2903_sleep").Load
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2903_tx915_cw").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2903_tx_cw_off").Load


        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_gpio_full").Load
        'TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483r1_gpio_full").Load 'FW Revision 1.0.0
        
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_sleep").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_sleep_revised").Load
        
        
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx868_cw").Load
        TheHdw.Digital.Patterns.Pat(".\patterns\uart_rn2483_tx_cw_off").Load

   TheHdw.Digital.Patterns.Pat("./patterns/uart_rn2483_tx433_cw").Load ' ("start_tx_cw_on")  'TS added 2017-01-10 for 433 MHz TX power test
   
     TheHdw.Digital.Patterns.Pat(".\patterns\icsp_subr").Load
        
        
    FIRSTRUN = True
    FIRSTLOAD = True
    
    
    
    TheExec.DataLog.Setup.LotSetup.PartType = TheExec.CurrentPart
    TheExec.DataLog.Setup.DatalogSetUp.HeaderEveryRun = True
    TheExec.DataLog.ApplySetup
    
    

    Exit Function
    
errHandler:
    MsgBox "Error encountered in Exec Interpose Function RFOnProgramValidated_LoRa" + vbCrLf + _
           "VBT Error # " + Trim(str(err.Number)) + ": " + err.Description

End Function

