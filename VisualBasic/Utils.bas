Attribute VB_Name = "Utils"
Option Explicit

'---------------------------------------------------------------------------------------------------------
' Utils
'
' Author      : Unknown
' Date        : Unknown
' Description : Unknown
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
' Rev     Date          Engineer         Description
' <0>     06/29/2009    M. Hudiani       Created set_hram_fail as a debugging tool.
'
'---------------------------------------------------------------------------------------------------------

Private ret_Trig As TrigType
Private ret_Capt As CaptType
Private ret_WaitForEvent As Boolean
Private ret_PreTrigCycleCnt As Long
Private ret_StopOnFull As Boolean
Private ret_CompressRepeats As Boolean

'Public Function ResolveArgv(s As String) As Double
'    Dim s_variant As Variant
'
'    ' IsNumeric requires Variant argument
'    s_variant = s
'
'    ' Grab value based on whether a literal of variable was passed
'    If (IsNumeric(s_variant)) Then
'        ResolveArgv = CDbl(Val(s))
'    Else
'        ResolveArgv = resolve_spec(s)
'    End If
'End Function

Public Function set_hram_fail(argc As Long, argv() As String) As Long
    
    Call TheHdw.Digital.HRAM.SetTrigger(trigFail, False, 0, True)
    Call TheHdw.Digital.HRAM.SetCapture(captFail, False)
    
End Function

Public Function set_hram_stv(argc As Long, argv() As String) As Long
    
    Call TheExec.Flow.CallFuncWithArgs("save_hram", "")
    
    Call TheHdw.Digital.HRAM.SetTrigger(trigSTV, False, 0, False)
    Call TheHdw.Digital.HRAM.SetCapture(captSTV, False)
    
End Function

Public Function save_hram(argc As Long, argv() As String) As Long

    ' Made standalone function so that it could be called from within VB in
    ' the case where one bit mode is used.

    Call TheHdw.Digital.HRAM.GetTrigger(ret_Trig, ret_WaitForEvent, ret_PreTrigCycleCnt, ret_StopOnFull)
    Call TheHdw.Digital.HRAM.GetCapture(ret_Capt, ret_CompressRepeats)

End Function

Public Function restore_hram(argc As Long, argv() As String) As Long

    Call TheHdw.Digital.HRAM.SetTrigger(ret_Trig, ret_WaitForEvent, ret_PreTrigCycleCnt, ret_StopOnFull)
    Call TheHdw.Digital.HRAM.SetCapture(ret_Capt, ret_CompressRepeats)

End Function

 
' [=============================================================================]
' [ DEVICE :                                                                    ]
' [ MASK NO:                                                                    ]
' [ SCOPE  : Generic, specifically written for the mid-range core and sram      ]
' [          voltage disturb test.  This routine will work when the following   ]
' [          sequence is needed:                                                ]
' [             1. Execute one vector section of pattern set                    ]
' [             2. Enter this VB function via post pattern interpose call       ]
' [                a. read and save current Vdd voltage                         ]
' [                b. change power supplies to passed in voltage                ]
' [                c. wait time specified in arg(1)                             ]
' [                d. restore Vdd to original voltage                           ]
' [                e. exit function                                             ]
' [             3. If another pattern left, goto step 1                         ]
' [=============================================================================]
' [                                                                             ]
' [                   MICROCHIP TECHNOLOGY INC.                                 ]
' [                   2355 WEST CHANDLER BLVD.                                  ]
' [                   CHANDLER AZ 85224-6199                                    ]
' [                   (602) 963-7373                                            ]
' [                                                                             ]
' [================= Copyright Statement =======================================]
' [                                                                             ]
' [   THIS PROGRAM AND ITS VECTORS ARE  PROPERTY OF Microchip Technology Inc.   ]
' [   USE, COPY, MODIFY, OR TRANSFER OF THIS PROGRAM, IN WHOLE OR IN PART,      ]
' [   AND IN ANY FORM OR MEDIA, EXCEPT AS EXPRESSLY PROVIDED FOR BY LICENSE     ]
' [   FROM Mircochip Technology Inc. IS FORBIDDEN.                              ]
' [                                                                             ]
' [================= Revision History ==========================================]
' [ REV.   DATE    OWN  COMMENT                                                 ]
' [ ^^^^   ^^^^    ^^^  ^^^^^^^                                                 ]
' [                                                                             ]
' [ 1.0    15jun98 jpe  - initial release                                       ]
' [=============================================================================]




' Caller can pass in a Voltage to switch to, a wait time in seconds,
' followed by 1 or more pin names to apply that voltage to.  Ex:
'
'   ToggleVdd("2.50", "0.01", "vdd1", "vdd2")
' The syntax  for the call from the program is:
'   function    parameter string
' "ToggleVdd" "2.50,0.01,vdd1,vdd2"
'  NOTE: quote marks not needed, used here to show which box it goes in
'
'  also valid input is a cell name from the spec sheet.  an example of this would be:
'   function    parameter string
' "ToggleVdd" "_toggle_vdd_min,_delay_time,vdd1,vdd2"
'
Public Function ToggleVdd(argc As Long, argv() As String) As Long

    Dim PowerPinList As String                   ' List of pins to toggle
    Dim OriginalVoltages As Variant              ' Keeps track of voltages on entry
    Dim NewVoltages As Variant                   ' Array of voltages to toggle to
    Dim dToggleVoltage As Double                 ' Voltage to toggle to
    Dim lChans() As Long                         ' Array of DPS channels
    Dim lNumChans As Long                        ' Number of DPS channels
    Dim lNumSites As Long                        ' Number of selected or active sites
    Dim errstr As String                         ' Error Message
    Dim ii As Integer                            ' Loop variable

    ' Check that there are at least 3 arguments
    If argc < 3 Then
        MsgBox "Error: ToggleVdd requires at least 3 arguments!", vbCritical
        Exit Function
    End If
    
    ' Recreate pinlist into single comma delimited string (needed for next step)
    PowerPinList = argv(2)
    For ii = 3 To argc - 1
        PowerPinList = "," + PowerPinList + argv(ii)
    Next ii
    
    ' Get list of channels
    Call TheExec.DataManager.GetChanListForSelectedSites(PowerPinList, chDPS, lChans, lNumChans, lNumSites, errstr)
    If errstr <> "" Then
        MsgBox "Error: ToggleVdd error return from GetChanListForSelectedSites:" + _
                Chr$(13) + errstr, vbCritical
        Exit Function
    End If
        
    ' Get original voltages
    OriginalVoltages = TheHdw.DPS.chans(lChans).PrimaryVoltages
    
    ' Set new voltages
    dToggleVoltage = ResolveArgv(argv(0))
    ReDim NewVoltages(lNumChans - 1)
    For ii = 0 To lNumChans - 1
        NewVoltages(ii) = dToggleVoltage
    Next ii
    
    TheHdw.DPS.chans(lChans).PrimaryVoltages = NewVoltages
    TheHdw.DPS.chans(lChans).OutputSource = dpsPrimaryVoltage
    
    ' Wait
    Call TheHdw.wait(ResolveArgv(argv(1)))
    
    ' Restore original voltages
    TheHdw.DPS.chans(lChans).PrimaryVoltages = OriginalVoltages
    TheHdw.DPS.chans(lChans).OutputSource = dpsPrimaryVoltage ' Need this again??
    
End Function

' Function: devPowerDown
' Purpose:  The function is used to ensure that a clean POR is performed
'           when entering into test.
' Params:   pdTime      Double      Power Down Wait Time
' Returns:  n/a
' Revision: Date:       Engineer    Description:
'           28-APR-2006 DePaul      Initial Release
Public Function devPowerDown(argc As Long, argv() As String) As Long
    
    Dim pdTime As Double

    If argc < 1 Then
        MsgBox "Error: devPowerDown() function requires at least one argument!", vbCritical
        Exit Function
    End If
    
    ' Resolve the power down wait time
    pdTime = ResolveArgv(argv(0))
    
    ' Have the pin levels power down all active sites, wait the requested pdTime, and
    ' then reapply the power levels to the pins and wait 100us before proceeding.
    Call TheHdw.PinLevels.PowerDown
    Call TheHdw.wait(pdTime)
    Call TheHdw.PinLevels.ApplyPower
    Call TheHdw.wait(0.0001)

End Function

' Function: failAllActiveSites
' Purpose:  This function will report a fail on all active sites.
'           It is intended for use when an error is detected in VB code
'           execution and the executing function needs to exit abruptly.
'           Calling this function will set the test result of the currently
'           executing test as a FAIL for all active sites and report
'           a FAIL as a functional result on the datalog.
'
' Revision: Date:       Engineer    Description:
'           16-AUG-2007 Aristizabal Initial Release
Public Function failAllActiveSites()
    
    Dim SiteNum As Long
    If TheExec.Sites.SelectFirst <> loopDone Then
        Do
            SiteNum = TheExec.Sites.SelectedSite
            Call TheExec.DataLog.WriteFunctionalResult(SiteNum, TheExec.Sites.Site(SiteNum).testnumber, logTestFail)
            TheExec.Sites.Site(SiteNum).TestResult = siteFail
            
        Loop Until TheExec.Sites.SelectNext(loopTop) = loopDone
    End If

End Function

' Subroutine:   cto_src_ramp
' Purpose:      Ramp CTO source level based on number of steps and delay between steps
' Params:       start       double      value to start ramping CTO
'               finish      double      value to finish ramping CTO
'               cto_ary     long        CTO source channel array
' Returns:      N/A
Public Sub cto_src_ramp(ByVal start As Double, ByVal finish As Double, ByRef cto_ary() As Long)
    Dim indx As Integer
    Dim num_step As Integer
    Dim DLY_TIME As Double
    Dim step_Size As Double
    Dim sign As Integer
    
    step_Size = 0.2
    num_step = Int(Abs(start - finish) / step_Size)
'    num_step = 10

    sign = 1
    If finish < start Then sign = -1

    DLY_TIME = 0.00002
    
    If num_step = 0 Then GoTo exit_cto_src_ramp
    For indx = 0 To num_step Step 1
            TheHdw.CTO.chans(cto_ary).LevelValue(ctoSrc) = start + sign * step_Size * (indx / num_step)
'           TheHdw.CTO.chans(cto_ary).LevelValue(ctoSrc) = start + (finish - start) * indx / num_step
           TheHdw.wait (DLY_TIME)
    Next indx
    
exit_cto_src_ramp:
    TheHdw.CTO.chans(cto_ary).LevelValue(ctoSrc) = finish
    TheHdw.wait (DLY_TIME)
End Sub

Public Function rampVdd(argc As Long, argv() As String) As Long

    Dim PowerPinList As String
    Dim ii As Long
    Dim jj As Long
    Dim errstr As String
    Dim lChans() As Long                         ' Array of DPS channels
    Dim lNumChans As Long                        ' Number of DPS channels
    Dim lNumSites As Long                        ' Number of selected or active sites
    
    Dim OriginalVoltages As Variant
    Dim NewVoltages As Variant                   ' Array of voltages to toggle to
    Dim lNumRampSteps As Long
    
    Dim dRampStepTime As Double
    Dim dRampStepVolt As Double
    Dim dTotalRampTime As Double
    
    ' Check that there are at least 3 arguments
    If argc < 4 Then
        MsgBox "Error: rampVdd requires at least 4 arguments!", vbCritical
        Exit Function
    End If
    
    lNumRampSteps = ResolveArgv(argv(2))
    dTotalRampTime = ResolveArgv(argv(1))
    
    dRampStepTime = dTotalRampTime / lNumRampSteps
    dRampStepVolt = 2.2 / lNumRampSteps       ' Hardcoded to 2.2 V for now

    ' Recreate pinlist into single comma delimited string (needed for next step)
    PowerPinList = argv(3)
    For ii = 4 To argc - 1
        PowerPinList = "," + PowerPinList + argv(ii)
    Next ii
    
    ' Get list of channels
    Call TheExec.DataManager.GetChanListForSelectedSites(PowerPinList, chDPS, lChans, lNumChans, lNumSites, errstr)
    If errstr <> "" Then
        MsgBox "Error: ToggleVdd error return from GetChanListForSelectedSites:" + _
                Chr$(13) + errstr, vbCritical
        Exit Function
    End If
    
    ReDim NewVoltages(lNumChans - 1)
    
    ' Get original voltages
    OriginalVoltages = TheHdw.DPS.chans(lChans).PrimaryVoltages
    
    ' Drop voltage to 0 V
    For ii = 0 To lNumChans - 1
        NewVoltages(ii) = 0
    Next ii
    
    TheHdw.DPS.chans(lChans).PrimaryVoltages = NewVoltages
    TheHdw.DPS.chans(lChans).OutputSource = dpsPrimaryVoltage
    
    ' Wait the specified hold time
    Call TheHdw.wait(ResolveArgv(argv(0)))
    
    ' Begin ramping voltage
    For jj = 1 To lNumRampSteps
        For ii = 0 To lNumChans - 1
            NewVoltages(ii) = NewVoltages(ii) + dRampStepVolt
        Next ii
        
        TheHdw.DPS.chans(lChans).PrimaryVoltages = NewVoltages
        TheHdw.DPS.chans(lChans).OutputSource = dpsPrimaryVoltage
        
        Call TheHdw.wait(dRampStepTime)
    Next jj
        
    ' Restore original voltages - NOTE: This should be voltage we want to run test at,
    ' so we should not stay at the hardcoded value above!!!
    TheHdw.DPS.chans(lChans).PrimaryVoltages = OriginalVoltages
    TheHdw.DPS.chans(lChans).OutputSource = dpsPrimaryVoltage


End Function
'
Public Function rampDPS_andLevels(argc As Long, argv() As String) As Long

'   Function will ramp DPS and pin levels from 1.8V down to a user
'   passed value.

    Dim power_Pin As String                     ' List of pins to toggle
    Dim ramp_Voltage As Variant
    Dim final_Voltage As Double                 ' Voltage to ramp to.
                                                
'   argv(0) - final Vdd
'   argv(1) - time to delay after Vdd change
'   argv(2) - Vdd pin name
'   argv(3) - dut mode (test or normal mode - TM or NM)

'   Check that there are 4 arguments.
    If argc <> 4 Then
        MsgBox "Error: rampDPS_andLevels requires 4 arguments!", vbCritical
        Exit Function
    End If
    
'   Get power pin name.
    power_Pin = Trim(argv(2))
        
    final_Voltage = ResolveArgv(argv(0))
    
    For ramp_Voltage = 1.8 To final_Voltage Step -0.02
        TheHdw.DPS.pins(power_Pin).ForceValue(dpsPrimaryVoltage) = ramp_Voltage
        If UCase(Trim(argv(3))) = "NM" Then Call TheHdw.PinLevels.pins("nmclr").ModifyLevel(chVDriveHi, CDbl(ramp_Voltage))
        TheHdw.wait (0.0001)
    Next ramp_Voltage
    
'   Pin levels not really ramped since they should have been driven to ground during Vdd ramp.  Levels can
'   now be programmed to match final Vdd.
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVDriveHi, final_Voltage)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVCompareLo, (final_Voltage / 2) - 0.15)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVCompareHi, (final_Voltage / 2) + 0.15)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVT, final_Voltage / 2)
    
'   Test mode
    If UCase(Trim(argv(3))) = "TM" Then
        Call TheHdw.PinLevels.pins("nmclr").WriteHighVoltageParams(8, 0.05, 0.000001)
    End If
    
    Call TheHdw.wait(ResolveArgv(argv(1)))
        
End Function
'
Public Function StringToDec(ByVal RawString As String, ByVal NumOfBits As Integer) As Long

    Dim i           As Integer
    Dim j           As Integer
    Dim TempResult  As Long

'   Convert value from HRAM to a decimal value.

        TempResult = 0
        j = NumOfBits

        For i = 1 To NumOfBits
            If Mid(RawString, i, 1) = "H" Then
                TempResult = TempResult + 2 ^ (j - 1)
            End If
            j = j - 1
        Next i

    StringToDec = TempResult
End Function
'
Public Sub disable_one_bit_Mode(ByVal pin_Name As String)

    ' Disable HRAM one bit mode by channel instead of by pin name.  Disabling
    ' by pin name can cause problems for later tests that use HRAM
    ' not in one bit mode.

    Dim ch_Nums() As Long
    Dim ch_Count As Long
    Dim site_count As Long
    Dim err As String
    
    Call TheExec.DataManager.GetChanList(pin_Name, -1, chIO, ch_Nums, ch_Count, site_count, err)
    TheHdw.Digital.HRAM.chans(ch_Nums).OneBitMode = False

End Sub
'
Public Sub adjustLevels(ByVal vddValue As Double)

    TheHdw.DPS.pins("vdd").ForceValue(dpsPrimaryVoltage) = vddValue
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVDriveHi, vddValue)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVCompareLo, (vddValue / 2) - 0.15)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVCompareHi, (vddValue / 2) + 0.15)
    Call TheHdw.PinLevels.pins("iopins").ModifyLevel(chVT, vddValue / 2)

End Sub

