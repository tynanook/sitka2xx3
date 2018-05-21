Attribute VB_Name = "PinPmu_T_Adjusted"
Option Explicit

' IG/XL Pin PMU Test Template
' (c) Teradyne, Inc, 1997, 1998, 1999
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date        Description
' 09/27/99    Release 3.30 Development
'
'
Const cMaxSite = 32                 'ps_t Max number of sites for storing the base Ipd value.
Dim dblBaseIpd(cMaxSite) As Double  'ps_t array to hold the base IPD by site.
Dim strBaseIpd(cMaxSite) As String  'ps_t array to hold the string value of the base IPD by site for dataloging.
Dim dblIdd(cMaxSite) As Double      'ps_t array to hold the measurement of this instance by site.
                                    '     This is used to pass the value to an interpose function
Const blnDebug = True               'ps_t

Dim Arg_DcCategory As String, Arg_DcSelector As String, _
Arg_AcCategory As String, Arg_AcSelector As String, _
Arg_Timing As String, Arg_Edgeset As String, _
Arg_Levels As String, Arg_HspStartLabel As String, _
Arg_StartOfBodyF As String, Arg_PrePatF As String, _
Arg_PreTestF As String, Arg_PostTestF As String, _
Arg_PostPatF As String, Arg_EndOfBodyF As String, _
Arg_PreconditionPat As String, Arg_HoldStatePat As String, _
Arg_PcpStopLabel As String, Arg_WaitFlags As String, _
Arg_DriverLO  As String, Arg_DriverHI  As String, _
Arg_DriverZ As String, Arg_FloatPins As String, _
Arg_Pinlist As String, Arg_MeasureMode As String, _
Arg_Irange As String, Arg_SettlingTime As String
Dim Arg_HiLoLimValid As String, Arg_HiLimit As String, _
Arg_LoLimit As String, Arg_ForceCond1 As String, _
Arg_ForceCond2 As String, Arg_Fload As String, _
Arg_RelayMode As String, Arg_FlagWaitTimeout As String, _
Arg_StartOfBodyFInput As String, Arg_PrePatFInput As String, _
Arg_PreTestFInput As String, Arg_PostTestFInput As String, _
Arg_PostPatFInput As String, Arg_EndOfBodyFInput As String, _
Arg_PcpStartLabel As String, Arg_PcpCheckPatGen As String, _
Arg_HspStopLabel As String, Arg_SamplingTime As String, _
Arg_Samples As String, Arg_HspCheckPatGen As String, _
Arg_HspResumePat As String, Arg_VClampLo As String, _
Arg_VClampHi As String, Arg_Util1 As String, _
Arg_Util0 As String
Dim Arg_StoreBaseIpd As String              'ps_t
Dim Arg_AdjustIpd As String                 'ps_t

Public Const DelayTime = 0.00051
Public byPassDelay As Boolean
Public byPassModeCheck As Boolean


Private Const ARGNUM_HSPSTARTLABEL = 0
Private Const ARGNUM_STARTOFBODYF = 1
Private Const ARGNUM_PREPATF = 2
Private Const ARGNUM_PRETESTF = 3
Private Const ARGNUM_POSTTESTF = 4
Private Const ARGNUM_POSTPATF = 5
Private Const ARGNUM_ENDOFBODYF = 6
Private Const ARGNUM_PRECONDITIONPAT = 7
Private Const ARGNUM_HOLDSTATEPAT = 8
Private Const ARGNUM_PCPSTOPLABEL = 9
Private Const ARGNUM_WAITFLAGS = 10
Private Const ARGNUM_DRIVERLO = 11
Private Const ARGNUM_DRIVERHI = 12
Private Const ARGNUM_DRIVERZ = 13
Private Const ARGNUM_FLOATPINS = 14
Private Const ARGNUM_PINLIST = 15
Private Const ARGNUM_MEASUREMODE = 16
Private Const ARGNUM_IRANGE = 17
Private Const ARGNUM_SETTLINGTIME = 18
Private Const ARGNUM_HILOLIMVALID = 19
Private Const ARGNUM_HILIMIT = 20
Private Const ARGNUM_LOLIMIT = 21
Private Const ARGNUM_FORCECOND1 = 22
Private Const ARGNUM_FORCECOND2 = 23
Private Const ARGNUM_FLOAD = 24
Private Const ARGNUM_RELAYMODE = 25
Private Const ARGNUM_FLAGWAITTIMEOUT = 26
Private Const ARGNUM_STARTOFBODYFINPUT = 27
Private Const ARGNUM_PREPATFINPUT = 28
Private Const ARGNUM_PRETESTFINPUT = 29
Private Const ARGNUM_POSTTESTFINPUT = 30
Private Const ARGNUM_POSTPATFINPUT = 31
Private Const ARGNUM_ENDOFBODYFINPUT = 32
Private Const ARGNUM_PCPSTARTLABEL = 33
Private Const ARGNUM_PCPCHECKPATGEN = 34
Private Const ARGNUM_HSPSTOPLABEL = 35
Private Const ARGNUM_HSPCHECKPATGEN = 36
Private Const ARGNUM_SAMPLINGTIME = 37
Private Const ARGNUM_SAMPLES = 38
Private Const ARGNUM_HSPRESUMEPAT = 39
Private Const ARGNUM_VCLAMPLO = 40
Private Const ARGNUM_VCLAMPHI = 41
Private Const ARGNUM_UTIL1 = 42
Private Const ARGNUM_UTIL0 = 43
'Private Const ARGNUM_MAXARG = ARGNUM_UTIL0
Private Const ARGNUM_StoreBaseIpd = 46                'ps_t dgnuarin
Private Const ARGNUM_AdjustIpd = 47                   'ps_t dgnuarin
Private Const ARGNUM_MAXARG = ARGNUM_AdjustIpd        'ps_t dgnuarin

Private tt_LoopControl As Integer
Private tt_LoopCount As Integer
Private tt_ForceVal As Double
Private tt_Samples As Long
Private tt_WaitFlagsTrue As Long
Private tt_WaitFlagsFalse As Long






' The TestTemplate function simply calls the PreBody, Body, and PostBody
' functions.  The TestTemplate function is called from the tester executive
' code during normal execution rather than calling the PreBody, Body, and
' PostBody individually as a performance optimization.
Function TestTemplate() As Integer
    Dim PreBodyResult As Integer
    
    ' Call PreBody, the code setting up general timing & levels, registering
    '   functions, and initializing hardware sub-systems
    PreBodyResult = PreBody()
    
    If PreBodyResult = TL_SUCCESS Then
        ' Call Body, the code performing DUT testing, and also used during
        '   test debug looping
        Call Body
    
        ' Call PostBody, the code verifying proper test execution, clearing
        '   hw&sw registers as needed
        Call PostBody
            
        TestTemplate = TL_SUCCESS
    Else
        TestTemplate = TL_ERROR
    End If
End Function                                   'End of TestTemplate Function.



Function PreBody() As Integer
    Dim ReturnStatus As Long
    If TheExec.Flow.IsRunning = False Then Exit Function
    'First, acquire the values of the parameters for this instance
    '   from the Data Manager
    Call GetTemplateParameters

    ' Register interpose function names with flow control routines which may
    '   need to invoke them
    Call tl_SetInterpose(TL_C_PREPATF, Arg_PrePatF, Arg_PrePatFInput, _
        TL_C_POSTPATF, Arg_PostPatF, Arg_PostPatFInput, _
        TL_C_PRETESTF, Arg_PreTestF, Arg_PreTestFInput, _
        TL_C_POSTTESTF, Arg_PostTestF, Arg_PostTestFInput)

    ' Optionally power down instruments and power supplies
    If (Arg_RelayMode <> TL_C_RELAYPOWERED) Then Call TheHdw.PinLevels.PowerDown

    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active levels sheet, if needed
    Call TheHdw.PinLevels.ConnectAllPins

    ' Set drive state on specified utility pins.
    If Arg_Util0 <> TL_C_EMPTYSTR Then Call tl_SetUtilState(Arg_Util0, 0)
    If Arg_Util1 <> TL_C_EMPTYSTR Then Call tl_SetUtilState(Arg_Util1, 1)
    
    ' Instruct functional voltages/currents hardware drivers to acquire
    '   drive/receive values from the DataManager and apply them.
    If (Arg_Levels <> TL_C_EMPTYSTR) And (Arg_RelayMode = TL_C_RELAYPOWERED) Then _
        Call TheHdw.PinLevels.ApplyPower

    ' Instruct functional timing hardware drivers to acquire timing values
    '   from the DataManager and apply them.
    If Arg_Timing <> TL_C_EMPTYSTR Then Call TheHdw.Digital.Timing.Load

    ' Remove specified DUT pins, if any, from connection to tester
    '   pin-electronics and other resources
    If Arg_FloatPins <> TL_C_EMPTYSTR Then Call tl_SetFloatState(Arg_FloatPins)

    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern
    '   burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If Arg_DriverLO <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverLO, chStartLo)
    If Arg_DriverHI <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverHI, chStartHi)
    If Arg_DriverZ <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverZ, chStartOff)
    
    ' Set init-state driver conditions on specified pins
    ' Setting init state causes the pin to drive the specified value.  Init
    '   state is set once, during the prebody, before the first pattern burst.
    ' Default is to leave the pin driving whatever value it last drove during
    '   the previous pattern burst.
    If Arg_DriverLO <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverLO, chInitLo)
    If Arg_DriverHI <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverHI, chInitHi)
    If Arg_DriverZ <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverZ, chInitOff)

    ' Initialize the decoded values for the flag condition settings.
    tt_WaitFlagsTrue = 0
    tt_WaitFlagsFalse = 0

    If Arg_HoldStatePat <> TL_C_EMPTYSTR And (InStr(Arg_WaitFlags, "XXXX") = 0) Then
        ' If needed, decode the 'WaitFlags' string to detect which
        '   of the CPU  pattern wait flags are to be set true, and
        '   which are to be set false.
        Call tl_tm_GetFlagsTrueAndFalse(Arg_WaitFlags, _
            tt_WaitFlagsTrue, tt_WaitFlagsFalse)
    End If

    ' Set the LoopCounter for making either 1 or 2 measurements within Body.
    If (Arg_ForceCond2 <> TL_C_EMPTYSTR) Then
        tt_LoopCount = 2
    Else
        tt_LoopCount = 1
    End If

    ' Decide how many samples are to be made & averaged for a test-measurement.
    If Arg_SamplingTime <> TL_C_EMPTYSTR Then
        ' If a duration is specified, during which samples are to be made,
        '   calculate the number of Samples, by dividing sampling time
        '   by the PPMU's conversion time.
        Dim PmuTime As Double
        PmuTime = TheHdw.PPMU.pins(Arg_Pinlist).PredictMeasuringTime(CLng(Arg_Irange))
        If PmuTime <> 0 Then
            tt_Samples = Int(CDbl(Val(Arg_SamplingTime) / PmuTime))
            If tt_Samples = 0 Then tt_Samples = 1
        Else
            tt_Samples = 1
        End If
    Else
        If Arg_Samples = TL_C_EMPTYSTR Then
            ' If no SamplingTime or Samples count is provided, define a default.
            tt_Samples = 1
        Else
            tt_Samples = CLng(Arg_Samples)
        End If
    End If
    
    ' Set the PPMU driver pass/fail limits for the Pinlist,
    '   to support the Smart-range functionality of the instrument.
    With TheHdw.PPMU.pins(Arg_Pinlist)
        If Arg_HiLimit <> TL_C_EMPTYSTR Then
            .TestLimitHigh = CDbl(Val(Arg_HiLimit))
        Else
            If Arg_MeasureMode = TL_C_VOLTAGE Then
                .TestLimitHigh = TL_MAX_PPMU_VOLTAGE
            Else
                .TestLimitHigh = tl_tm_GetMaxRangeVal(TheHdw.PPMU.MeasIRangeList)
            End If
        End If
        If Arg_LoLimit <> TL_C_EMPTYSTR Then
            .TestLimitLow = CDbl(Val(Arg_LoLimit))
        Else
            If Arg_MeasureMode = TL_C_VOLTAGE Then
                .TestLimitLow = TL_MIN_PPMU_VOLTAGE
            Else
                .TestLimitLow = tl_tm_GetMinRangeVal(TheHdw.PPMU.MeasIRangeList)
            End If
        End If
    End With

    PreBody = TL_SUCCESS

End Function



Function Body() As Integer
    Dim RangeVal As Double
    Dim ReturnStatus As Long
    Dim temp As String
    Dim PeRlyClsd As String
    Body = TL_SUCCESS
    If TheExec.Flow.IsRunning = False Then Exit Function
    
    ' Run the 'StartOfBodyF' interpose function, if specified.
    If Arg_StartOfBodyF <> TL_C_EMPTYSTR Then
        Call TheExec.Flow.CallFuncWithArgs(Arg_StartOfBodyF, Arg_StartOfBodyFInput)
        If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
    End If
    
    If (Arg_PreconditionPat <> TL_C_EMPTYSTR) Or (Arg_HoldStatePat <> TL_C_EMPTYSTR) Then
        'If patterns used, enable the timeout counter
        TheHdw.Digital.Patgen.TimeoutEnable = True
    End If

    If (Arg_Fload = TL_C_YES) Or (Arg_VClampLo <> TL_C_EMPTYSTR) Or (Arg_VClampHi <> TL_C_EMPTYSTR) Then
        ' Connect pins to functional loading Init state
        Call tl_SetInitState(Arg_Pinlist, chInitOff)
    End If

    ' Connect pins to the Ppmu, if specified, to save execution time & relay life
    If (Arg_RelayMode <> TL_C_RELAYPOWERED) Then
        'power down the device under test
        Call TheHdw.PinLevels.PowerDown
        'set the Ppmu in a current forcing mode
        With TheHdw.pins(Arg_Pinlist)
            'force 0 Amperes with the Ppmu
            .PPMU.ForceVoltage(ppmu2mA) = 0#
            'add delay for mode switching, if specified
            If Not byPassDelay Then
                TheHdw.wait (DelayTime)
            End If
            If (Arg_Fload = TL_C_YES) Or (Arg_VClampLo <> TL_C_EMPTYSTR) Or (Arg_VClampHi <> TL_C_EMPTYSTR) Then
                .Relays.Connect rlyPPMU_PE      ' connect PPMU and PE relays
            Else
                .PPMU.Connect                   ' connect the Ppmu relay only
            End If
        End With
        'power up the remaining tester resources in the order specified by the levels sheet
        Call TheHdw.PinLevels.ApplyPower
    End If

    ' This For/Next is the major structure of the Template Body.  The loop is
    '   run either 1 or 2 times, based upon conditions found in the parameters
    '   list.
    ' Within the loop, force values are programmed, ranges and pins specified,
    '   DUT conditioning is performed, and the measurements are made.
    For tt_LoopControl = 1 To tt_LoopCount
        ' It is possible that the first time through the loop removed all
        '   active sites; if so, then exit this loop.
        If TheExec.Sites.ActiveCount = 0 Then
            Exit For
        End If
        
        ' Run a PreconditionPat pattern, if any
        If Arg_PreconditionPat <> TL_C_EMPTYSTR Then
            Call tl_RunPatternStartStop(Arg_PreconditionPat, _
                Arg_PcpStartLabel, Arg_PcpStopLabel, Arg_PcpCheckPatGen)
            If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
        End If

        ' Begin/Run a HoldStatePat pattern, if any.
        If Arg_HoldStatePat <> TL_C_EMPTYSTR Then
            Call tl_HoldStatePatBegin(Arg_HoldStatePat, Arg_HspStartLabel, Arg_HspStopLabel, _
                Arg_HspCheckPatGen, tt_WaitFlagsTrue, tt_WaitFlagsFalse, Arg_FlagWaitTimeout)
            If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
        End If

        ' Set Ppmu force value for the Pinlist
        If tt_LoopControl = 1 Then
            tt_ForceVal = CDbl(Val(Arg_ForceCond1))
        ElseIf tt_LoopControl = 2 Then
            tt_ForceVal = CDbl(Val(Arg_ForceCond2))
        End If

        With TheHdw.pins(Arg_Pinlist)
            If Arg_MeasureMode = TL_C_VOLTAGE Then
                If Arg_VClampLo <> TL_C_EMPTYSTR Then
                    Call .PinLevels.ModifyLevel(chVCL, CDbl(Val(Arg_VClampLo)))
                End If
                If Arg_VClampHi <> TL_C_EMPTYSTR Then
                    Call .PinLevels.ModifyLevel(chVCH, CDbl(Val(Arg_VClampHi)))
                End If
                If Arg_Fload = TL_C_NO Then
                    'set the Ppmu in a current forcing mode, and voltage measure mode
                    .PPMU.ForceCurrent(CLng(Arg_Irange)) = tt_ForceVal
                    If (Arg_VClampLo <> TL_C_EMPTYSTR) Or (Arg_VClampHi <> TL_C_EMPTYSTR) Then
                        'if clamps are used, then the PE relay is closed, and this then
                        '   requires that the current loads be set to zero.
                        Call .PinLevels.ModifyLevel(chISource, 0)
                        Call .PinLevels.ModifyLevel(chISink, 0)
                    End If
                End If
                If Arg_Fload = TL_C_YES Then
                    ' Override Functional Load levels on Pin Electronics
                    
                    ' set the PPMU itself to force 0A
                    .PPMU.ForceCurrent(CLng(Arg_Irange)) = 0#
    
                    ' Modify programmed load on Pin Electronics
                    If tt_ForceVal > 0 Then
                        ' set VT to ensure that load will be applied
                        Call .PinLevels.ModifyLevel(chVT, TL_MAX_VT_LEVEL)
                        ' set the load current
                        Call .PinLevels.ModifyLevel(chISource, tt_ForceVal)
                    Else
                        ' set VT to ensure that load will be applied
                        Call .PinLevels.ModifyLevel(chVT, TL_MIN_VT_LEVEL)
                        ' set the load current
                        Call .PinLevels.ModifyLevel(chISink, tt_ForceVal)
                    End If
                End If
            Else
                'set the Ppmu in a voltage forcing mode, and current measure mode
                .PPMU.ForceVoltage(CLng(Arg_Irange)) = tt_ForceVal
            End If
        End With

        ' Execute the Test.
        ' The Ppmu will make the pass/fail decision.
        ' The Ppmu is provided a Pinlist.
        ' The Ppmu driver will return a test status value, based upon an
        '   error condition during the test measurement.
        If Arg_SettlingTime = TL_C_EMPTYSTR Then Arg_SettlingTime = "0"
        If (Arg_Fload = TL_C_YES) Or (Arg_VClampLo <> TL_C_EMPTYSTR) Or (Arg_VClampHi <> TL_C_EMPTYSTR) Then
            PeRlyClsd = TL_C_YES
        Else
            PeRlyClsd = TL_C_NO
        End If
        ReturnStatus = Local_tl_PpmuMeasureValue(Arg_Pinlist, Arg_MeasureMode, _
            tt_Samples, CDbl(Val(Arg_LoLimit)), CDbl(Val(Arg_HiLimit)), _
            CDbl(Val(Arg_SettlingTime)), tt_ForceVal, Arg_HiLoLimValid, _
            Arg_RelayMode, Arg_Irange, PeRlyClsd)
        If ReturnStatus <> TL_SUCCESS Then
            'denote error
            temp = TheExec.DataManager.instanceName
            Call TheExec.ErrorLogMessage("PPMU: tl_PpmuMeasureValue " & TL_C_ERRORSTR & ", Instance: " & temp)
            Call TheExec.ErrorReport
            Body = TL_ERROR
        End If

        If (((Arg_Fload = TL_C_YES) Or (Arg_VClampLo <> TL_C_EMPTYSTR) Or _
            (Arg_VClampHi <> TL_C_EMPTYSTR)) And (tt_LoopControl = 1) And (tt_LoopCount = 2)) _
            Or ((Arg_Fload = TL_C_YES) And (Arg_HspCheckPatGen = TL_C_YES)) Then
            ' If functional loads are being used or voltage clamps are used, and we
            '   are on the first of two passes through the loop, restore the PE
            '   loads to the original values
            Call TheHdw.PinLevels.ApplyPower
        End If

        ' Complete a HoldStatePat pattern, if any.
        If Arg_HoldStatePat <> TL_C_EMPTYSTR Then
            Call tl_HoldStatePatFinish(Arg_HoldStatePat, Arg_HspCheckPatGen, _
                Arg_HspResumePat, Arg_FlagWaitTimeout, Arg_HspStopLabel, _
                tt_WaitFlagsTrue, tt_WaitFlagsFalse)
        End If

    ' This is the end of the For/Next loop.
    Next tt_LoopControl

    If (Arg_RelayMode <> TL_C_RELAYPOWERED) Then
        If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
        'check the force mode, if specified
        If (Not byPassModeCheck) And (Arg_MeasureMode = TL_C_CURRENT) Then
            'FVMI mode, set the ppmu voltages to 0
           TheHdw.PPMU.pins(Arg_Pinlist).ForceVoltage(ppmu200uA) = 0#
        Else
            'FIMV set the ppmu currents to 0
            TheHdw.PPMU.pins(Arg_Pinlist).ForceCurrent(2) = 0#
        End If
                       
     End If

NoSitesActive:

    ' Run the 'EndOfBodyF' interpose function, if specified
    If Arg_EndOfBodyF <> TL_C_EMPTYSTR Then _
        Call TheExec.Flow.CallFuncWithArgs(Arg_EndOfBodyF, Arg_EndOfBodyFInput)

End Function



Function PostBody() As Integer
    Dim DriverOFF As String
    If TheExec.Flow.IsRunning = False Then Exit Function

    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, _
        TL_C_POSTTESTF)

    ' Connect specified DUT pins, if any, to tester pin-electronics & power
    If Arg_FloatPins <> TL_C_EMPTYSTR Then Call tl_ConnectTester(Arg_FloatPins)

    ' Return channels to the default start-state condition, as needed
    DriverOFF = tl_tm_CombineCslStrings(Arg_DriverHI, Arg_DriverLO)
    If DriverOFF <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(DriverOFF, chStartOff)

    ' Clear the pins that are masked for this test instance
    Call tl_ClearChannelsMaskedTestInstance

    PostBody = TL_SUCCESS
End Function

Sub GetTemplateParameters()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)
    
    Arg_DcCategory = ArgStr(TL_C_DCCATCOLNUM)
    Arg_DcSelector = ArgStr(TL_C_DCSELCOLNUM)
    Arg_AcCategory = ArgStr(TL_C_ACCATCOLNUM)
    Arg_AcSelector = ArgStr(TL_C_ACSELCOLNUM)
    Arg_Timing = ArgStr(TL_C_TIMESETCOLNUM)
    Arg_Edgeset = ArgStr(TL_C_EDGESETCOLNUM)
    Arg_Levels = ArgStr(TL_C_LEVELSCOLNUM)
    
    Arg_StartOfBodyF = ArgStr(ARGNUM_STARTOFBODYF)
    Arg_PrePatF = ArgStr(ARGNUM_PREPATF)
    Arg_PreTestF = ArgStr(ARGNUM_PRETESTF)
    Arg_PostTestF = ArgStr(ARGNUM_POSTTESTF)
    Arg_PostPatF = ArgStr(ARGNUM_POSTPATF)
    Arg_EndOfBodyF = ArgStr(ARGNUM_ENDOFBODYF)
    Arg_PreconditionPat = ArgStr(ARGNUM_PRECONDITIONPAT)
    Arg_HoldStatePat = ArgStr(ARGNUM_HOLDSTATEPAT)
    Arg_DriverLO = ArgStr(ARGNUM_DRIVERLO)
    Arg_DriverHI = ArgStr(ARGNUM_DRIVERHI)
    Arg_DriverZ = ArgStr(ARGNUM_DRIVERZ)
    Arg_FloatPins = ArgStr(ARGNUM_FLOATPINS)
    Arg_Pinlist = ArgStr(ARGNUM_PINLIST)
    Arg_MeasureMode = ArgStr(ARGNUM_MEASUREMODE)
    Arg_Irange = ArgStr(ARGNUM_IRANGE)
    Arg_SettlingTime = ArgStr(ARGNUM_SETTLINGTIME)
    Arg_HiLoLimValid = ArgStr(ARGNUM_HILOLIMVALID)
    Arg_HiLimit = ArgStr(ARGNUM_HILIMIT)
    Arg_LoLimit = ArgStr(ARGNUM_LOLIMIT)
    Arg_ForceCond1 = ArgStr(ARGNUM_FORCECOND1)
    Arg_ForceCond2 = ArgStr(ARGNUM_FORCECOND2)
    Arg_Fload = ArgStr(ARGNUM_FLOAD)
    Arg_RelayMode = ArgStr(ARGNUM_RELAYMODE)
    Arg_StartOfBodyFInput = ArgStr(ARGNUM_STARTOFBODYFINPUT)
    Arg_PrePatFInput = ArgStr(ARGNUM_PREPATFINPUT)
    Arg_PreTestFInput = ArgStr(ARGNUM_PRETESTFINPUT)
    Arg_PostTestFInput = ArgStr(ARGNUM_POSTTESTFINPUT)
    Arg_PostPatFInput = ArgStr(ARGNUM_POSTPATFINPUT)
    Arg_EndOfBodyFInput = ArgStr(ARGNUM_ENDOFBODYFINPUT)
    Arg_SamplingTime = ArgStr(ARGNUM_SAMPLINGTIME)
    Arg_Samples = ArgStr(ARGNUM_SAMPLES)
    Arg_VClampLo = ArgStr(ARGNUM_VCLAMPLO)
    Arg_VClampHi = ArgStr(ARGNUM_VCLAMPHI)
    Arg_PcpStartLabel = ArgStr(ARGNUM_PCPSTARTLABEL)
    Arg_PcpStopLabel = ArgStr(ARGNUM_PCPSTOPLABEL)
    Arg_PcpCheckPatGen = ArgStr(ARGNUM_PCPCHECKPATGEN)
    Arg_HspStartLabel = ArgStr(ARGNUM_HSPSTARTLABEL)
    Arg_HspStopLabel = ArgStr(ARGNUM_HSPSTOPLABEL)
    Arg_WaitFlags = ArgStr(ARGNUM_WAITFLAGS)
    Arg_FlagWaitTimeout = ArgStr(ARGNUM_FLAGWAITTIMEOUT)
    Arg_HspCheckPatGen = ArgStr(ARGNUM_HSPCHECKPATGEN)
    Arg_HspResumePat = ArgStr(ARGNUM_HSPRESUMEPAT)
    Arg_Util1 = ArgStr(ARGNUM_UTIL1)
    Arg_Util0 = ArgStr(ARGNUM_UTIL0)
    Arg_StoreBaseIpd = ArgStr(ARGNUM_StoreBaseIpd)                 'ps_t
    Arg_AdjustIpd = ArgStr(ARGNUM_AdjustIpd)                       'ps_t


End Sub


Function DatalogType() As Integer
    DatalogType = logParametric
End Function

' End of Execution Section

Public Function RunIE(Optional FocusArg As Integer) As Boolean
    tl_tm_FocusArg = FocusArg
    Call tl_fs_ResetIECtrl(tl_tm_InstanceEditor)
    With tl_tm_InstanceEditor
        .name = "PinPmu_T"
        .PpmuPages = True
        .CondPatPage = True
        .LevTimPage = True
        .PinPage = True
        .InterposePage = True
        .Caption = TL_C_IEPPMUSTR
        .HelpValue = TL_C_PPMU_HELP
        .UserOptPage1 = True                  'ps_t
        .UserOptName = "Sink Current"         'ps_t
        .UserOptArg1.enabled = True           'ps_t
        .UserOptArg1.CheckBoxVisible = False  'ps_t
        .UserOptArg1.EvalBoxEnabled = False   'ps_t
        .UserOptArg1.ComboNotTextBox = True   'ps_t
        .UserOptArg1.ButtonVisible = False    'ps_t
        .UserOptArg1.EvalBoxEnabled = False   'ps_t
        .UserOptArg2.enabled = True           'ps_t
        .UserOptArg2.CheckBoxVisible = False  'ps_t
        .UserOptArg2.EvalBoxEnabled = False   'ps_t
        .UserOptArg2.ComboNotTextBox = True   'ps_t
        .UserOptArg2.ButtonVisible = False    'ps_t
        .UserOptArg2.EvalBoxEnabled = False   'ps_t

    End With
    'InstanceEditor_IE.Show                   'ps_t
    Call tl_fs_StartIE                        'ps_t
    'the return value will be true if the 'Apply' button was not enabled and if the workbook was valid when the form initialized
    RunIE = (Not (tl_tm_FormCtrl.ButtonEnabled)) And tl_tm_BookIsValid
End Function


Sub AssignTemplateValues()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)
    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            .ParameterValue = ArgStr(.Argnum)
        End With
    Next
    
    'if the value is blank, then apply the default value to the spreadsheet and the Arg
    Call tl_tm_ManageDefault(AllPars, ARGNUM_MAXARG)
    
End Sub
Sub ApplyDefaults()
    Call SetupParameters
    
    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            Call tl_tm_PutDefaultIfNeeded(.Argnum, .DefaultValue)
        End With
    Next
    Call tl_tm_CleanUp

End Sub
Function GetArgNames() As String
    Dim CallSetup As Boolean
    CallSetup = False
    If AllPars.count = 0 Then
        Call SetupParameters    'acquire the Argument information, if needed
        CallSetup = True
    End If
    GetArgNames = tl_tm_ListArgNames(ARGNUM_MAXARG)
    If CallSetup = True Then Call tl_tm_CleanUp
End Function


Sub SetupParameters()
    Call tl_tm_SetupCatSelValidation
    Call tl_tm_SetupTimLevValidation
    Call tl_tm_SetupOverlayValidation
    Call tl_tm_SetupInterposeValidation(ARGNUM_STARTOFBODYF, ARGNUM_PREPATF, _
        ARGNUM_PRETESTF, ARGNUM_POSTTESTF, ARGNUM_POSTPATF, ARGNUM_ENDOFBODYF)
    Call tl_tm_SetupInterposeInputValidation(ARGNUM_STARTOFBODYFINPUT, ARGNUM_PREPATFINPUT, _
        ARGNUM_PRETESTFINPUT, ARGNUM_POSTTESTFINPUT, ARGNUM_POSTPATFINPUT, ARGNUM_ENDOFBODYFINPUT)
    Call tl_tm_SetupConditioningPinlistsValidation(ARGNUM_DRIVERLO, ARGNUM_DRIVERHI, _
        ARGNUM_DRIVERZ, ARGNUM_FLOATPINS, ARGNUM_UTIL1, ARGNUM_UTIL0)
    Call tl_tm_SetupPreCondPatValidation(ARGNUM_PRECONDITIONPAT, ARGNUM_PCPSTOPLABEL, _
        ARGNUM_PCPSTARTLABEL, ARGNUM_PCPCHECKPATGEN)
    Call tl_tm_SetupHoldStatePatValidation(ARGNUM_HOLDSTATEPAT, ARGNUM_HSPSTOPLABEL, _
        ARGNUM_HSPSTARTLABEL, ARGNUM_HSPCHECKPATGEN, ARGNUM_HSPRESUMEPAT, _
        ARGNUM_WAITFLAGS, ARGNUM_FLAGWAITTIMEOUT)
    Call tl_tm_SetupLimitsValidation(ARGNUM_HILOLIMVALID, ARGNUM_HILIMIT, ARGNUM_LOLIMIT, _
        tl_tm_ParPpmuLimits, tl_tm_ParPpmuHiLimSpec, tl_tm_ParPpmuLoLimSpec)
    
    'Pinlist,
    With tl_tm_ParPpmuPinlist
        .AllParsAdd
        .Argnum = ARGNUM_PINLIST
        .ParameterName = TL_C_PinlistStr
        .tl_tm_ParSetParam
        .TestIsPin = True
        .ValueChoices = JobData.AllDigitalPins
        .TestNotBlank = True
    End With
        
    'ps_t  Define user option for saving Base IPD values
    AllPars.Add tl_tm_ParUserOpt1
    With tl_tm_ParUserOpt1
        .Argnum = ARGNUM_StoreBaseIpd
        .ParameterDispName = "Save Sink Current"
        .ParameterName = "SaveSinkCurrent"
        .tl_tm_ParSetParam
        .ValueChoices = "YES,NO"
        .TestIsLegalChoice = True
        .DefaultValue = "NO"
    End With

    'ps_t  Define user option for adjusting measured value by the stored Ipd value
    AllPars.Add tl_tm_ParUserOpt2
    With tl_tm_ParUserOpt2
        .Argnum = ARGNUM_AdjustIpd
        .ParameterDispName = "Add Sink Current"
        .ParameterName = "AddSinkCurrent"
        .tl_tm_ParSetParam
        .ValueChoices = "YES,NO"
        .TestIsLegalChoice = True
        .DefaultValue = "NO"
    End With
        
    'MeasureMode,
    With tl_tm_ParPpmuMeasureMode
        .AllParsAdd
        .Argnum = ARGNUM_MEASUREMODE
        .ParameterName = TL_C_MeasureModeStr
        .tl_tm_ParSetParam
        .ValueChoices = TL_C_MMALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .DefaultValue = tl_tm_GetIndexOf(TL_C_MMVSTR)
    End With
    Set tl_tm_MeasConditions.mode = tl_tm_ParPpmuMeasureMode
    'Irange,
    With tl_tm_ParPpmuIRange
        .AllParsAdd
        .Argnum = ARGNUM_IRANGE
        .ParameterName = TL_C_IRangeStr
        .tl_tm_ParSetParam
        .ValueChoices = TheHdw.PPMU.MeasIRangeList
        .DefaultValue = tl_tm_GetMaxRangeValIndex(.ValueChoices)
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(-tl_tm_GetRangeVal(.DefaultValue, .ValueChoices)) _
            & TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(tl_tm_GetRangeVal(.DefaultValue, .ValueChoices))
        .TestNotBlank = True
        .TestIsLegalChoice = True
    End With
    Set tl_tm_MeasConditions.Irange = tl_tm_ParPpmuIRange
    'HClamp,
    With tl_tm_ParPpmuHClamp
        .AllParsAdd
        .Argnum = ARGNUM_VCLAMPHI
        .ParameterName = TL_C_VClampHiStr
        .tl_tm_ParSetParam
        .TestNumeric = True
        .TestInRange = True
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MIN_VCH_LEVEL) & _
            TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MAX_VCH_LEVEL)
        Call .SetEnabler(tl_tm_ParPpmuHClamp, TL_C_NOTBLANK)
    End With
    'LClamp,
    With tl_tm_ParPpmuLClamp
        .AllParsAdd
        .Argnum = ARGNUM_VCLAMPLO
        .ParameterName = TL_C_VClampLoStr
        .TestNumeric = True
        .tl_tm_ParSetParam
        .TestNumeric = True
        .TestInRange = True
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MIN_VCL_LEVEL) & _
            TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MAX_VCL_LEVEL)
        Call .SetEnabler(tl_tm_ParPpmuLClamp, TL_C_NOTBLANK)
    End With
    'SamplingTime,
    Dim PmuTime As Double
    PmuTime = TheHdw.PPMU.ConversionTime
    With tl_tm_ParPpmuSamplingTime
        .AllParsAdd
        .Argnum = ARGNUM_SAMPLINGTIME
        .ParameterName = TL_C_SamplingTimeStr
        .tl_tm_ParSetParam
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & "0" & TL_C_DELIMITERSTD & _
            TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & CStr(10000 * PmuTime)
        .TestInRange = True
        .TestNumeric = True
        Call .SetEnabler(tl_tm_ParPpmuSamplingTime, TL_C_NOTBLANK)
    End With
    'Samples,
    With tl_tm_ParPpmuSamples
        .AllParsAdd
        .Argnum = ARGNUM_SAMPLES
        .ParameterName = TL_C_SamplesStr
        .tl_tm_ParSetParam
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & "0" & TL_C_DELIMITERSTD & _
            TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & "10000"
        .TestInRange = True
        .TestNumeric = True
        Call .SetEnabler(tl_tm_ParPpmuSamples, TL_C_NOTBLANK)
    End With
    'SettlingTime,
    With tl_tm_ParPpmuSettlingTime
        .AllParsAdd
        .Argnum = ARGNUM_SETTLINGTIME
        .ParameterName = TL_C_SettlingTimeStr
        .tl_tm_ParSetParam
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & "0" & TL_C_DELIMITERSTD & _
            TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & "30"
        .TestInRange = True
        .TestNumeric = True
        .DefaultValue = "0"
        Call .SetEnabler(tl_tm_ParPpmuSettlingTime, TL_C_NOTBLANK)
    End With
    'ForceCond1,
    With tl_tm_ParPpmuForceCond1
        .AllParsAdd
        .Argnum = ARGNUM_FORCECOND1
        .ParameterName = TL_C_ForceCond1Str
        .tl_tm_ParSetParam
        .TestNumeric = True
        .ValueChoices = JobData.AllSpecEntries
        .VCisSpecEntries = True
        .TestNotBlank = True
    End With
    Set tl_tm_MeasConditions.Force1 = tl_tm_ParPpmuForceCond1
    'ForceCond2,
    With tl_tm_ParPpmuForceCond2
        .AllParsAdd
        .Argnum = ARGNUM_FORCECOND2
        .ParameterName = TL_C_ForceCond2Str
        .tl_tm_ParSetParam
        .TestNumeric = True
        .ValueChoices = JobData.AllSpecEntries
        .VCisSpecEntries = True
        Call .SetEnabler(tl_tm_ParPpmuForceCond2, TL_C_NOTBLANK)
    End With
    Set tl_tm_MeasConditions.Force2 = tl_tm_ParPpmuForceCond2
    'FLoad,
    With tl_tm_ParPpmuFLoad
        .AllParsAdd
        .Argnum = ARGNUM_FLOAD
        .ParameterName = TL_C_FLOADStr
        .tl_tm_ParSetParam
        .ValueChoices = TL_C_YNALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .DefaultValue = tl_tm_GetIndexOf(TL_C_YNNEGSTR)
    End With
    Set tl_tm_MeasConditions.Fload = tl_tm_ParPpmuFLoad
    'RelayMode,
    With tl_tm_ParPpmuRelayMode
        .AllParsAdd
        .Argnum = ARGNUM_RELAYMODE
        .ParameterName = TL_C_RelayModeStr
        .tl_tm_ParSetParam
        'the valid choices can change, based upon whether Levels is set non-blank
        .ValueChoices = TL_C_RMALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .DefaultValue = tl_tm_GetIndexOf(TL_C_RMCOLDSTR)
        Call .SetEnabler(tl_tm_ParPpmuRelayMode, TL_C_NOTBLANK)
    End With
    
    'Vrange,
    'no parameter by this name appears on the IE form, this is used solely for validation purposes
    With tl_tm_ParPpmuVrange
        .Argnum = -1
        .ParameterName = TL_C_VrangeStr
        .ParameterStr = TL_C_EMPTYSTR
        .ParameterValue = TL_C_EMPTYSTR
        Set .LabelBox = Nothing
        Set .DataBox = Nothing
        .tl_tm_ParSetParam
        .ValueChoices = TheHdw.PPMU.ForceVRangeList
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(tl_tm_GetRangeVal(TL_C_LRANGEINDEX, .ValueChoices)) _
            & TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(tl_tm_GetRangeVal(TL_C_HRANGEINDEX, .ValueChoices))
    End With
    Set tl_tm_MeasConditions.Vrange = tl_tm_ParPpmuVrange
    
    'BadCombinations of template arguments
    
    'It is a bad combination for WaitFlags to be 'XXXX' and HspResume to be 'Yes'
    '   and HoldStatePat to be specified.
    Call CreateBadCombo     ' create a 'Bad Combination' of template arguments.
    Set tl_tm_ParThisBadCombo = AllBadCombo.Item(AllBadCombo.count)
    With tl_tm_ParThisBadCombo
        'register the template arguments, the data, and the type of comparisans.
        Call .SetInputArg(tl_tm_ParFlags, "XXXX", TL_C_EQUAL)
        Call .SetInputArg(tl_tm_ParHspResume, tl_tm_GetIndexOf(TL_C_YNYESSTR), TL_C_EQUAL)
        Call .SetInputArg(tl_tm_ParHoldStatePatName, TL_C_EMPTYSTR, TL_C_NOTEQUAL)
        'specify the message to be used if the validation testing finds a problem.
        .MsgStr = tl_tm_MsgStr(TL_TM_STR_VALIDATEINCOMPATIBLE, _
                    Array(TL_C_ARGstr & CStr(tl_tm_ParFlags.Argnum), _
                          tl_tm_ParFlags.ParameterName, tl_tm_ParHspResume.ParameterName))
    End With
    
    'It is a bad combination if Floatpins and Pinlist share any elements.
    Call CreateBadCombo     ' create a 'Bad Combination' of template arguments.
    Set tl_tm_ParThisBadCombo = AllBadCombo.Item(AllBadCombo.count)
    With tl_tm_ParThisBadCombo
        'register the template arguments, the data, and the type of comparisans.
        Call .SetInputArg(tl_tm_ParFloatPins, TL_C_EMPTYSTR, TL_C_EVALSHARED)
        Call .SetInputArg(tl_tm_ParPpmuPinlist, TL_C_EMPTYSTR, TL_C_EVALSHARED)
        'when there are two arguments being checked for sharing pins, the
        '   validation routine will generate the error message and set
        '   the failing test type indicator number.
    End With
        
        'ps_t It is a bad combination for both SaveBaseIpd and AdjustIpd to be 'Yes'
    Call CreateBadCombo     ' create a 'Bad Combination' of template arguments.
    Set tl_tm_ParThisBadCombo = AllBadCombo.Item(AllBadCombo.count)
    With tl_tm_ParThisBadCombo
        'register the template arguments, the data, and the type of comparisans.
        Call .SetInputArg(tl_tm_ParUserOpt1, "YES", TL_C_EQUAL)
        Call .SetInputArg(tl_tm_ParUserOpt2, "YES", TL_C_EQUAL)
        'specify the message to be used if the validation testing finds a problem.
        .MsgStr = TL_C_ARGstr & CStr(tl_tm_ParUserOpt1.Argnum) & " " & tl_tm_ParUserOpt1.ParameterName _
                & " and " & tl_tm_ParUserOpt2.ParameterName & " can not both be YES.       "
                    
    End With

    
End Sub


Function ValidateParameters(Optional VDCint As Integer) As Integer
    'This function is used, at validation time, to determine whether the data
    '   to be executed is proper, valid, and copacetic.  It can be called by
    '   an Instance Editor, or by the Job Validation routines.
    Dim TestResult As Integer
    Dim temp As String
    Dim RetVal As Boolean
    Dim IrangeOK As Boolean
    Dim CheckHiLim As Boolean
    Dim CheckLoLim As Boolean
    Dim intX As Integer
    Dim MsgStr As String
    '   This has modes to run in.  If a mode of '0' is specified for
    '   input, it is assumed that the mode is TL_C_VALDATAMODEJOBVAL.
    '   The modes that .ValidateParameters can operate in are:
    '   TL_C_VALDATAMODEJOBVAL  -   Job Validation mode; report errors to sheet.
    '   TL_C_VALDATAMODENORMAL  -   Instance Editor mode; Fix the current parameter being evaluated.
    '   TL_C_VALDATAMODENOSTOP  -   Instance Editor mode; Do not stop to fix any parameters.
    '   It can return with different modes, such as:
    '   TL_C_VALDATAMODENOFIX   -   Instance Editor mode; Error found, that specific one was not fixed.
    '   TL_C_VALDATAMODEFIXNONE -   Instance Editor mode; Error(s) found, none were fixed.

    'Success is first assumed; if a problem is noted, ValidateParameters will be
    '   set to failure by this routine.
    ValidateParameters = TL_SUCCESS
    
    If VDCint = 0 Then VDCint = TL_C_VALDATAMODEJOBVAL
    If (VDCint <> TL_C_VALDATAMODENORMAL) And (VDCint <> TL_C_VALDATAMODENOSTOP) _
        And (VDCint <> TL_C_VALDATAMODEJOBVAL) Then
        'denote an error
        temp = TheExec.DataManager.instanceName
        Call TheExec.ErrorLogMessage("ValidateParameters: Improper mode, instance: " & temp)
        Call TheExec.ErrorReport
        ValidateParameters = TL_ERROR
        Exit Function
    End If

    If VDCint = TL_C_VALDATAMODEJOBVAL Then
        With JobData
            'Get list of pins and pin-groups from datatools.
            Call tl_fs_TemplateJobDataPinlistStrings(JobData, VDCint)
        
            'Get lists of Categories, Selectors, Timesets, Edgesets, and Levels
            Call tl_fs_TemplateCatSelStrings(.AvailDcCat, .AvailDcSel, _
                .AvailAcCat, .AvailAcSel, .AvailTimeSetAll, .AvailTimeSetExtended, _
                .AvailEdgeSet, .AvailLevels)
            'Get list of Overlay
            Call tl_fs_TemplateOverlayString(.AvailOverlay)
        End With
        
        'Define the Parameter types and tests to be performed
        Call SetupParameters
        
        'Now, acquire the values of the parameters for this Template Instance
        '   from the DataManager and assign them to the TemplateArg structures.
        Call AssignTemplateValues
    End If

    ValidateParameters = TL_SUCCESS
    
    ' Choose tests to perform
    Call tl_tm_ChooseTests(AllPars, VDCint)
    
    'setup modifiers to parameter tests
    If tl_tm_ParPpmuMeasureMode.ParameterValue = tl_tm_GetIndexOf(TL_C_MMISTR) Then
        tl_tm_ParPpmuIRange.ValueChoices = TheHdw.PPMU.MeasIRangeList
    Else
        tl_tm_ParPpmuIRange.ValueChoices = TheHdw.PPMU.ForceIRangeList
    End If
    'determine the proper range limit for use with validation
    With tl_tm_ParPpmuIRange
        .RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(-tl_tm_GetRangeVal(.ParameterValue, .ValueChoices)) & _
            TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & _
            CStr(tl_tm_GetRangeVal(.ParameterValue, .ValueChoices))
    End With
    
    'set the Levels sheet as a required argument if any of these conditions exist:
    '   RelayMode is cold
    '   (Arg_Fload = TL_C_YES or Arg_VClampLo <> TL_C_EMPTYSTR or Arg_VClampHi <> TL_C_EMPTYSTR) And (ForceCond2 <> TL_C_EMPTYSTR)
    '   (Arg_Fload = TL_C_YES) And (Arg_HspCheckPatGen = TL_C_YES)
    With tl_tm_ParLevels
        If (((tl_tm_ParPpmuFLoad.ParameterValue = TL_C_YES) Or _
                    (tl_tm_ParPpmuLClamp.ParameterValue <> TL_C_EMPTYSTR) Or _
                    (tl_tm_ParPpmuHClamp.ParameterValue <> TL_C_EMPTYSTR)) And _
                    (tl_tm_ParPpmuForceCond2.ParameterValue <> TL_C_EMPTYSTR)) Or _
            ((tl_tm_ParPpmuFLoad.ParameterValue = TL_C_YES) And _
                (tl_tm_ParHspPatGenCheck.ParameterValue = TL_C_YES)) Then
            .TestNotBlank = True
            .TestingEnabled = True
            If (tl_tm_ParPpmuForceCond2.ParameterValue <> TL_C_EMPTYSTR) Then
                If (tl_tm_ParPpmuFLoad.ParameterValue = TL_C_YES) Then
                    MsgStr = tl_tm_MsgStr(TL_TM_STR_LEVELSNEEDED2ERR, _
                                Array(tl_tm_ParPpmuFLoad.ParameterName, _
                                      tl_GetStringOf(TL_C_YNYESSTR), _
                                      tl_tm_ParPpmuForceCond2.ParameterName, _
                                      .ParameterName))
                    Call .SetErrorMsg(TL_C_REQERR, MsgStr)
                End If
                If (tl_tm_ParPpmuLClamp.ParameterValue <> TL_C_EMPTYSTR) Then
                    MsgStr = tl_tm_MsgStr(TL_TM_STR_LEVELSNEEDED2ERR, _
                                Array(tl_tm_ParPpmuLClamp.ParameterName, _
                                      CStr(tl_tm_ParPpmuLClamp.ParameterValue), _
                                      tl_tm_ParPpmuForceCond2.ParameterName, _
                                      .ParameterName))
                    Call .SetErrorMsg(TL_C_REQERR, MsgStr)
                End If
                If (tl_tm_ParPpmuHClamp.ParameterValue <> TL_C_EMPTYSTR) Then
                    MsgStr = tl_tm_MsgStr(TL_TM_STR_LEVELSNEEDED2ERR, _
                                Array(tl_tm_ParPpmuHClamp.ParameterName, _
                                      CStr(tl_tm_ParPpmuHClamp.ParameterValue), _
                                      tl_tm_ParPpmuForceCond2.ParameterName, _
                                      .ParameterName))
                    Call .SetErrorMsg(TL_C_REQERR, MsgStr)
                End If
            End If
            If (tl_tm_ParPpmuFLoad.ParameterValue = TL_C_YES) And _
                (tl_tm_ParHspPatGenCheck.ParameterValue = TL_C_YES) Then
                MsgStr = tl_tm_MsgStr(TL_TM_STR_LEVELSNEEDED3ERR, _
                            Array(tl_tm_ParPpmuFLoad.ParameterName, _
                                  tl_GetStringOf(TL_C_YNYESSTR), _
                                  tl_tm_ParHspPatGenCheck.ParameterName, _
                                  tl_GetStringOf(TL_C_YNYESSTR), _
                                  .ParameterName))
                Call .SetErrorMsg(TL_C_REQERR, MsgStr)
            End If
        Else
            .TestNotBlank = False
            .ClrErrorMsg
        End If
    End With
    
    ' Now run the tests on each Argument
    Call tl_tm_RunTests(AllPars, VDCint, TestResult)
    If TestResult <> TL_SUCCESS Then ValidateParameters = TL_ERROR
    If (TestResult <> TL_SUCCESS) And (VDCint = TL_C_VALDATAMODENORMAL) Then Exit Function
    
'ensure that hilimit is greater than lolimit
' <cc>   TestResult = tl_tm_MeasConditions.LimitsTest(VDCint)
 ' <cc>      If TestResult <> TL_SUCCESS Then ValidateParameters = TL_ERROR
 ' <cc>  If (TestResult <> TL_SUCCESS) And (VDCint = TL_C_VALDATAMODENORMAL) Then Exit Function


'ensure that the chosen irange and vrange, if used, are adequate for forcing and
'measuring quantities specified by forcecond1, forcecond2, hilim, & lolim
    'first, work must be done here to change the H & L values for the irange, if Floading is used
    If Left(CStr(tl_tm_ParPpmuFLoad.ParameterValue), 1) = TL_C_YES Then
        tl_tm_MeasConditions.FLoadStr = _
            TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MAX_IOL_LEVEL) & "A" & TL_C_DELIMITERSTD & _
            TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & CStr(TL_MIN_IOH_LEVEL) & "A"
    End If
   '<cc> TestResult = tl_tm_MeasConditions.RangesTest(CStr(tl_tm_ParPpmuMeasureMode.ParameterValue), VDCint)
    'now, reset the Fload indicator
    tl_tm_MeasConditions.FLoadStr = TL_C_EMPTYSTR
    If TestResult <> TL_SUCCESS Then ValidateParameters = TL_ERROR
    If (TestResult <> TL_SUCCESS) And (VDCint = TL_C_VALDATAMODENORMAL) Then Exit Function

'    Warning: Be aware that interpose functions are not validated

    If VDCint = TL_C_VALDATAMODEJOBVAL Then Call tl_tm_CleanUp
End Function

Function ValidateDriverParameters() As Integer
    Dim RetVal As Long
    ValidateDriverParameters = TL_SUCCESS
    
    Call SetupParameters
    'Now, acquire the values of the parameters for this Template Instance
    '   from the DataManager and assign them to the TemplateArg structures.
    Call AssignTemplateValues
    
    ' Now validate the patterns used
    RetVal = ValPatAndLabels(tl_tm_ParPrecondPatNames, tl_tm_ParPcpStartLabel, tl_tm_ParPcpStopLabel)
    If RetVal = TL_ERROR Then
        ValidateDriverParameters = TL_ERROR
    End If
    RetVal = ValPatAndLabels(tl_tm_ParHoldStatePatName, tl_tm_ParHspStartLabel, tl_tm_ParHspStopLabel)
    If RetVal = TL_ERROR Then
        ValidateDriverParameters = TL_ERROR
    End If
    Call tl_tm_CleanUp
End Function


Private Function Local_tl_PpmuMeasureValue(pins As String, _
    MeasureMode As String, samples As Long, lowLimit As Double, _
    highLimit As Double, SettlingTime As Double, ForceValue As Double, _
    HiLoLimValid As String, RelayMode As String, Irange As String, PeRlyClosed As String) As Long
    
    Dim channels() As Long
    Dim ChanNdx As Long
    Dim measured() As Double
    Dim nsites As Long
    Dim err As String
    Dim nchannels As Long
    Dim thisChan As Integer 'none
    Dim sampleLoop As Long
    Dim SampleAvg As Double
    Dim parmFlag As Long
    Dim loc As Long
    Dim measunitcode As Long    ' UnitType
    Dim forceunitcode As Long   ' UnitType
    Dim testStatus As Long
    Dim ReturnStatus As Long
    Dim lngX     As Long   'none
    Dim forceVal     As Double 'none
    Dim dblAdjustedIpd As Double   'dgnuarin copied
    
    On Error GoTo errHandler        ' Trap driver errors
        
    If (RelayMode <> TL_C_RELAYPOWERED) Then
        ' Set the Settling Timer for the user specified settling time
        Call TheHdw.SetSettlingTimer(SettlingTime)
    End If
    
    'Prepare to test
    If samples = 0 Then samples = 1 ' Samples should never be 0...
    TheHdw.PPMU.samples = samples
    
    ' Don't know what "loc" is used for in Parametric Datalog
    loc = 0
    
     ' Run the PreTestF interpose function
    Call tl_ExecuteInterpose(TL_C_PRETESTF)
    
    If TheExec.Sites.ActiveCount = 0 Then GoTo NoSitesActive
        
        Call InitIdd  'ps_t zero out the array for measurements of this Test instance 'dgnuarin copied from adjusted PPMU
    'ps_t Initialize to zero the Base Ipd array
    If Arg_StoreBaseIpd = "YES" Then
       Call InitBaseIpd
    End If
    
    ' Initialize units for DataLog
    If MeasureMode = TL_C_VOLTAGE Then
        measunitcode = unitVolt
        forceunitcode = unitAmp
    Else
        measunitcode = unitNone         'dgnuarin edited to use unitNone
        forceunitcode = unitVolt
    End If
    
    'Get the complete list of channels
    '  Call TheExec.DataManager.GetChanListForSelectedSites(pins, chIO, channels, _
    '     nchannels, nsites, err)
    Dim sharedChannels() As Long
    Dim nsharedChannels As Long
    Dim siteList() As Long
    ' Changes for shared resources
    Call TheExec.DataManager.GetSharedChanListForSelectedSites(pins, chIO, channels, _
        sharedChannels, nchannels, nsharedChannels, nsites, siteList, err)
   
    If nchannels = 0 Then
        Call tl_tm_LogError("nchannels = 0")
        Local_tl_PpmuMeasureValue = TL_ERROR
        Exit Function
    End If
    
    ' Connect pins to the Ppmu, if specified
    If (RelayMode = TL_C_RELAYPOWERED) Then
        ' Close PPMU relays "hot"
        TheHdw.PPMU.pins(pins).Connect
        If PeRlyClosed = TL_C_YES Then
            Call TheHdw.Digital.Relays.pins(pins).Connect(rlyPPMU_PE)
        End If
        ' Set the Settling Timer for the user specified settling time
        Call TheHdw.SetSettlingTimer(SettlingTime)
    End If

    ' Wait for the settling timer
    Call TheHdw.SettleWait(30#)
    
    ' Measure all the values at once
    If MeasureMode = TL_C_VOLTAGE Then
        Call TheHdw.PPMU.chans(channels).MeasureVoltages(measured)
    Else
        Call TheHdw.PPMU.chans(channels).MeasureCurrents(measured)
    End If
    
    ' Remove pins from the Ppmu
    If (RelayMode = TL_C_RELAYPOWERED) Then
        ' Open PPMU relays "hot"
        TheHdw.PPMU.pins(pins).Disconnect
    End If
    
    If tl_TestTheMeas("Ppmu", measured, (samples * nchannels - 1)) = TL_ERROR Then
        Local_tl_PpmuMeasureValue = TL_ERROR
        Exit Function
    End If

    ' Compare the value (either current or voltage)
    Dim siteIndex As Long
    For ChanNdx = 0 To nchannels - 1
        
        ' Average the results across all samples
        SampleAvg = 0#
        For sampleLoop = 0 To samples * nchannels - 1 Step nchannels
            SampleAvg = SampleAvg + measured(sampleLoop + ChanNdx)
        Next sampleLoop
        SampleAvg = SampleAvg / samples
       
        siteIndex = ChanNdx Mod nsites   'ps_t  Determine the site 'dgnuarin copied from PPMU adjusted
        If Arg_StoreBaseIpd = "YES" Then
               'pst  Store the Idd/Ipd value as the base value
               Call StoreBaseIpd(siteList(siteIndex), SampleAvg)
        End If
                
                'ps_t  Verify that the user wants to subtract off the base Ipd value   'dgnuarin copy this
        If Arg_AdjustIpd = "YES" Then
               'dblAdjustedIpd = sampleAvg - GetBaseIpd(siteList(siteIndex))
               dblAdjustedIpd = Abs(1 / SampleAvg) + Abs(1 / GetBaseIpd(siteList(siteIndex)))
               TheExec.DataLog.WriteComment (Space(55) & "                    Isnk  = " & GetBaseIpd(siteList(siteIndex)))
               TheExec.DataLog.WriteComment (Space(55) & "                    Isrc  = " & SampleAvg)
               TheExec.DataLog.WriteComment (Space(55) & "Abs(1/Isnk) + Abs(1/Isrc) = " & dblAdjustedIpd)
        Else
               'dblAdjustedIpd = Abs(1 / sampleAvg) 'Do not adjust this measurement
               dblAdjustedIpd = SampleAvg 'Do not adjust this measurement
               'TheExec.Datalog.WriteComment (Space(55) & "Isnk        = " & sampleAvg)
        End If
        Call StoreIddMeas(siteList(siteIndex), dblAdjustedIpd)  'ps_t store measurment

        ' Determine whether test passed or failed
        'Call tl_tm_PassOrFail(sampleAvg, lowLimit, highLimit, HiLoLimValid, parmFlag, testStatus)
        Call tl_tm_PassOrFail(dblAdjustedIpd, lowLimit, highLimit, HiLoLimValid, parmFlag, testStatus) 'dgnuarin - hardcoded the limits
        'Call tl_tm_PassOrFail(dblAdjustedIpd, 0, 480000, "2", parmFlag, testStatus) 'dgnuarin - hardcoded the limits
         
'              forceVal = forceValue(chanNdx)
              'forceVal = forceValue
                
' <cc>       ' Log result and report status
        If (nsharedChannels = 0) Then
            Call Local_tl_tm_LogResultReportStatus(channels(ChanNdx), chIO, _
                dblAdjustedIpd, lowLimit, highLimit, parmFlag, testStatus, _
                measunitcode, ForceValue, forceunitcode, loc, , , HiLoLimValid)
        Else 'There are shared channels
         siteIndex = ChanNdx Mod nsites
        Call Local_tl_tm_LogResultReportStatus(channels(ChanNdx), chIO, _
                dblAdjustedIpd, lowLimit, highLimit, parmFlag, testStatus, _
                measunitcode, ForceValue, forceunitcode, loc, , , HiLoLimValid)
        End If

        ' Log result and report status
    '    If (nsharedChannels = 0) Then
    '        Call Local_tl_tm_LogResultReportStatus(channels(chanNdx), chIO, _
    '            dblAdjustedIpd, 0, 480000, parmFlag, testStatus, _
     '           unitNone, forceVal, unitVolt, loc)
     '   Else 'DPS channels are shared
     '       siteIndex = chanNdx Mod nsites
     '       Call Local_tl_tm_LogResultReportStatus(channels(chanNdx), chIO, _
     '           dblAdjustedIpd, 0, 480000, parmFlag, testStatus, _
     '           unitNone, forceVal, unitVolt, loc, "", siteList(siteIndex))
     '   End If
        
    Next ChanNdx
    
    ' Run the PostTestF interpose function
    If TheExec.Sites.ActiveCount <> 0 Then Call tl_ExecuteInterpose(TL_C_POSTTESTF)
    
NoSitesActive:
    
    Local_tl_PpmuMeasureValue = ReturnStatus
    On Error GoTo 0
    Exit Function
    
errHandler:
    ' count errors
    ReturnStatus = ReturnStatus + 1
    Resume Next
End Function

Private Sub Local_tl_tm_LogResultReportStatus(ByVal ChannelNumber As Long, chtype As Long, _
    SampleAvg As Double, lowLimit As Double, highLimit As Double, _
    parmFlag As Long, testStatus As Long, _
    units As Long, ForceValue As Double, forceUnits As Long, loc As Long, _
    Optional PinNameInput As String, Optional siteNumber As Long = -1, _
    Optional HiLoLimitValid As String = CStr(TL_C_HILIM1LOLIM1))
    Dim ReturnStatus As Long

    Dim PinName As String
    Dim thisSite As Long
    Dim testnumber As Long

    ' Get pin name and the site information
    Call TheHdw.PinSiteFromChan(ChannelNumber, chtype, PinName, thisSite)
    
    If PinNameInput <> TL_C_EMPTYSTR Then
        'this is done to handle a name for a ganged case of bpmu tested pins.
        PinName = PinNameInput
        'since the pins are ganged, do not report a real channel number
        ChannelNumber = -1
    End If

    ' Assign test number
    If (siteNumber = -1) Then
        testnumber = TheExec.Sites.Site(thisSite).testnumber
    Else
        thisSite = siteNumber
        testnumber = TheExec.Sites.Site(thisSite).testnumber
    End If
    
'<cc>    ' Send results to Datalog.
'    ' Decide which WriteParametricResult method to call, depending on the value of
'    ' the HiLoLimitValid flag .
    Select Case HiLoLimitValid
      Case TL_C_HILIM1LOLIM1 ' Hi and Lo limits defined.
        Call TheExec.DataLog.WriteParametricResult(thisSite, testnumber, _
                testStatus, parmFlag, PinName, ChannelNumber, _
                lowLimit, SampleAvg, highLimit, units, ForceValue, forceUnits, loc)
      Case TL_C_HILIM1LOLIM0 ' Only Hi limit defined.
        Call TheExec.DataLog.WriteParametricResultOptLoHi(thisSite, testnumber, _
                testStatus, parmFlag, PinName, ChannelNumber, _
                SampleAvg, units, ForceValue, forceUnits, , highLimit, loc)
      Case TL_C_HILIM0LOLIM1 ' Only Lo limit defined.
        Call TheExec.DataLog.WriteParametricResultOptLoHi(thisSite, testnumber, _
                testStatus, parmFlag, PinName, ChannelNumber, _
                SampleAvg, units, ForceValue, forceUnits, lowLimit, , loc)
      Case TL_C_HILIM0LOLIM0 ' Neither Hi or Lo limit defined.
        Call TheExec.DataLog.WriteParametricResultOptLoHi(thisSite, testnumber, _
                testStatus, parmFlag, PinName, ChannelNumber, _
                SampleAvg, units, ForceValue, forceUnits, , , loc)
    End Select
  
    ' Send results to Datalog
   ' Call TheExec.Datalog.WriteParametricResult(thisSite, testnumber, _
   ''         testStatus, parmFlag, PinName, ChannelNumber, _
   '         lowLimit, sampleAvg, highLimit, units, forceValue, forceUnits, loc)
    If blnDebug Then
        Dim lngResult As Long
        'ps_t print out the mode of the test on the datalog
        If TheExec.DataLog.IsCurrentParaTestLogging(testnumber, lngResult) Then
            If Arg_StoreBaseIpd = "YES" Then
               'TheExec.Datalog.WriteComment (Space(55) & "Stored Sink Value = " & GetBaseIpdString(thisSite))
            End If
            If Arg_AdjustIpd = "YES" Then
               'TheExec.Datalog.WriteComment (Space(55) & "Sink Value" & " 1/" & GetBaseIpdString(thisSite) & " Added")
            End If
        End If  'If TheExec.Datalog.IsCurrentParaTestLogging(testnumber, lngResult) Then
    End If      'If blnDebug Then
  
    ' Report Status
    If testStatus <> logTestPass Then
        TheExec.Sites.Site(thisSite).TestResult = siteFail
    Else
        TheExec.Sites.Site(thisSite).TestResult = sitePass
    End If
    
End Sub

'********************************************************************
'
'ps_t **** custom functions written for the Delta IPD measurements
'
'*******************************************************************
Public Function InitBaseIpd() As Long
'*****************************************************
' Purpose: Zero out the Ipd array which holds the base Ipd
' Inputs:  None
' Returns: TL_SUCCESS or TL_ERROR
'
'*****************************************************

  On Error GoTo errHandler
  
  Dim i As Long
  
  For i = 0 To cMaxSite - 1
     dblBaseIpd(i) = 0      'Initialize the array
     strBaseIpd(i) = ""
  Next i


  InitBaseIpd = TL_SUCCESS
'
Exit Function 'normal exit of function
errHandler:
    InitBaseIpd = TL_ERROR
    Call TheExec.ErrorLogMessage("Function InitBaseIpd had Error" & VBA.vbCrLf & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0
    Call TheExec.ErrorReport
'
End Function  'InitBaseIpd

Public Function InitIdd() As Long
'*****************************************************
' Purpose: Zero out the Ipd array for the measurements of this test instance
' Inputs:  None
' Returns: TL_SUCCESS or TL_ERROR
'
'*****************************************************

  On Error GoTo errHandler
  
  Dim i As Long
  
  For i = 0 To cMaxSite - 1
     dblIdd(i) = 0      'Initialize the array
  Next i


  InitIdd = TL_SUCCESS
'
Exit Function 'normal exit of function
errHandler:
    InitIdd = TL_ERROR
    Call TheExec.ErrorLogMessage("Function InitIdd had Error" & VBA.vbCrLf & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0
    Call TheExec.ErrorReport
'
End Function  'InitIdd



Public Function GetBaseIpd(ByVal lngSite As Long) As Double
'*****************************************************
' Purpose: To return the measurement value which was saved from a previous test instance
' Inputs:  lngSite       Site number
' Returns: the measurement value which was saved from a previous test instance
'
'*****************************************************
  On Error GoTo errHandler
  
  'Test for error condition, return zero since that will no effect on the measurement.
  If (lngSite >= 0) And (lngSite < cMaxSite) Then
     'Since Ipd can not be negative, clip any noise values below zero
     If dblBaseIpd(lngSite) > 0 Then
        GetBaseIpd = dblBaseIpd(lngSite)
     Else           'If dblBaseIpd(lngSite) > 0 Then
        GetBaseIpd = 0
     End If         'If dblBaseIpd(lngSite) > 0 Then
  Else              'If (lngSite >= 0) And (lngSite < cMaxSite) Then
     GetBaseIpd = 0
  End If            'If (lngSite >= 0) And (lngSite < cMaxSite) Then

Exit Function 'normal exit of function
errHandler:
    GetBaseIpd = 0
'    Call TheExec.ErrorLogMessage("Function GetBaseIpd had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    On Error GoTo 0
'    Call TheExec.ErrorReport
'
End Function  'GetBaseIpd


Public Function GetBaseIpdString(ByVal lngSite As Long) As String
'*****************************************************
' Purpose: To return the measurement value as a string which was saved from a previous test instance
' Inputs:  lngSite       Site number
' Returns: the measurement value which was saved from a previous test instance
'
'*****************************************************
  On Error GoTo errHandler
  
  'Test for error condition,
  If (lngSite >= 0) And (lngSite < cMaxSite) Then
     GetBaseIpdString = strBaseIpd(lngSite)
  Else              'If (lngSite >= 0) And (lngSite < cMaxSite) Then
     GetBaseIpdString = " Error in GetBaseIpdString, site out of range "
  End If            'If (lngSite >= 0) And (lngSite < cMaxSite) Then

Exit Function 'normal exit of function
errHandler:
    GetBaseIpdString = " Error in GetBaseIpdString "
'    Call TheExec.ErrorLogMessage("Function GetBaseIpdString had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    On Error GoTo 0
'    Call TheExec.ErrorReport
'
End Function  'GetBaseIpdString


Public Function StoreBaseIpd(lngSite As Long, ByVal dblIpdValue As Double) As Long
'*****************************************************
' Purpose: To save the measurement value to be used in a later test instance
' Inputs:  lngSite       Site number
'          dblIpdValue   Value to be stored for this site
' Returns: TL_SUCCESS or TL_ERROR
'
'*****************************************************

  On Error GoTo errHandler
  
  'dblIpdValue = CDbl(lngSite) * (0.000001) 'Debug code for offline
  
  'Test for out of range condition
  If (lngSite >= 0) And (lngSite < cMaxSite) Then
     dblBaseIpd(lngSite) = dblIpdValue
     strBaseIpd(lngSite) = DblToEngStr(dblIpdValue, "A")
  Else
     StoreBaseIpd = TL_ERROR
     Exit Function
  End If            'If (lngSite >= 0) And (lngSite < cMaxSite) Then

  StoreBaseIpd = TL_SUCCESS
'
Exit Function 'normal exit of function
errHandler:
    StoreBaseIpd = TL_ERROR
    Call TheExec.ErrorLogMessage("Function StoreBaseIpd had Error" & VBA.vbCrLf & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0
    Call TheExec.ErrorReport
'
End Function  'StoreBaseIpd

Public Function DblToEngStr(ByVal dblValue As Double, ByVal strUnit As String) As String
'*****************************************************
' Purpose: To convert a double real number to a string in engineering format
' Inputs:  dblIpdValue   Value to be converted
' Returns: A string in engineering format
'
'*****************************************************

  On Error GoTo errHandler
  Dim blnNegNum As Boolean
  Dim dblLocValue As Double
  Dim dblMult As Double
  Dim lngRange As Long
  Dim strTmp As String
  Dim blnNegExponet As Boolean
  Dim lngTmp As Long
  
  If (dblValue < 0) Then
     blnNegNum = True
     dblLocValue = -1 * dblValue
  Else
     blnNegNum = False
     dblLocValue = dblValue
  End If
  
  If dblLocValue < 1 Then
     blnNegExponet = True
     dblMult = 1000
  Else
     blnNegExponet = False
     dblMult = 0.001
  End If
  
  lngRange = 0
  
  If (dblLocValue < 0.000000000000001) Or (dblLocValue > 1000000#) Then '<1E-15 or >1E6
        strTmp = format(dblLocValue, "+0.0000E+00") & " " & strUnit
        'add sign or space
        If blnNegNum Then
           strTmp = "-" & strTmp
        Else
           strTmp = " " & strTmp
        End If
        
        'Pad spaces to the left so the string is alway the same length
        lngTmp = 14 + VBA.Len(strUnit)
        strTmp = Space(14) & strTmp
        strTmp = Right(strTmp, lngTmp)
  Else
        While ((dblLocValue < 1) Or (dblLocValue >= 1000)) And (lngRange < 20)
          If dblLocValue <> 1 Then
          dblLocValue = dblLocValue * dblMult
          lngRange = lngRange + 1
          End If
        Wend
        strTmp = format(dblLocValue, "###.0000")
        If blnNegExponet Then
          Select Case lngRange
          Case 0
                 strTmp = strTmp & "  " & strUnit
          Case 1
                 strTmp = strTmp & " m" & strUnit
          Case 2
                 strTmp = strTmp & " u" & strUnit
          Case 3
                 strTmp = strTmp & " n" & strUnit
          Case 4
                 strTmp = strTmp & " p" & strUnit
          Case 5
                 strTmp = strTmp & " f" & strUnit
          Case Else   ' Other values.
                 strTmp = strTmp & " ??" & strUnit
          End Select
        Else
          Select Case lngRange
          Case 0
                 strTmp = strTmp & "  " & strUnit
          Case 1
                 strTmp = strTmp & " k" & strUnit
          Case Else   ' Other values.
                 strTmp = strTmp & " ??"
          End Select
        End If
        
        'add sign or space
        If blnNegNum Then
           strTmp = "-" & strTmp
        Else
           strTmp = " " & strTmp
        End If
        
        'Pad spaces to the left so the string is alway the same length
        lngTmp = 12 + VBA.Len(strUnit)
        strTmp = Space(12) & strTmp
        strTmp = Right(strTmp, lngTmp)
  End If  'If (dblLocValue < 0.000000000000001) Or (dblLocValue > 1000#) Then '<1E-15 or >1E6
  
  
  DblToEngStr = strTmp
'
Exit Function 'normal exit of function
errHandler:
    DblToEngStr = "Error in function DblToEngStr, converting to engineering format"
    Call TheExec.ErrorLogMessage("Function DblToEngStr had Error" & VBA.vbCrLf & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0
    Call TheExec.ErrorReport
'
End Function  'DblToEngStr

Public Function GetIddMeas(ByVal lngSite As Long) As Double
'*****************************************************
' Purpose: To save the measurement value to be used in a later test instance
' Inputs:  lngSite       Site number
' Returns: The measurement in this template for the site
'          Will return a large negative number in the case of an error
'
'*****************************************************
  On Error GoTo errHandler
  
  'Test for error condition, return zero since that will no effect on the measurement.
  If (lngSite >= 0) And (lngSite < cMaxSite) Then
     GetIddMeas = dblIdd(lngSite)
  Else
     GetIddMeas = -1E-20   'Set for an error condition so that it can be trapped out
  End If

Exit Function 'normal exit of function
errHandler:
    GetIddMeas = -1E-20   'Set for an error condition so that it can be trapped out
'    Call TheExec.ErrorLogMessage("Function GetIddMeas had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    On Error GoTo 0
'    Call TheExec.ErrorReport
'
End Function  'GetIddMeas

Public Function StoreIddMeas(lngSite As Long, ByVal dblIpdValue As Double) As Long
'*****************************************************
' Purpose: To save the measurement value so that it can be accessed from VB
' Inputs:  lngSite       Site number
'          dblIpdValue   Value to be stored for this site
' Returns: TL_SUCCESS or TL_ERROR
'
'*****************************************************

  On Error GoTo errHandler
  
  'Test for out of range condition
  If (lngSite >= 0) And (lngSite < cMaxSite) Then
     dblIdd(lngSite) = dblIpdValue
  Else
     StoreIddMeas = TL_ERROR
     Exit Function
  End If            'If (lngSite >= 0) And (lngSite < cMaxSite) Then

  StoreIddMeas = TL_SUCCESS
'
Exit Function 'normal exit of function
errHandler:
    StoreIddMeas = TL_ERROR
    Call TheExec.ErrorLogMessage("Function StoreIddMeas had Error" & VBA.vbCrLf & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0
    Call TheExec.ErrorReport
'
End Function  'StoreIddMeas




