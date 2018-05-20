Attribute VB_Name = "VDD_Voltage_Disturb"
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
Dim PowerPinList As String
Dim OriginalVoltages As Variant
Dim NewVoltages As Variant
Dim VdisParam As Double
Dim chans() As Long
Dim nchannels As Long
Dim nsites As Long
Dim errstr As String
Dim i As Integer

    ' Were we called properly?
    If argc < 3 Then
        MsgBox "Error-ToggleVdd expected at least 3 arguments"
        Exit Function
    End If
    
    ' Recreate pinlist into single comma delimited string (needed for next step)
    PowerPinList = argv(2)
    For i = 3 To argc - 1
        PowerPinList = "," + PowerPinList + argv(i)
    Next i
    
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
    
    
    ' Get original voltages
'    Call tl_DpsGetPrimaryVoltages(chans, OriginalVoltages)
    OriginalVoltages = TheHdw.DPS.chans(chans).PrimaryVoltages
    ' Set new voltages
    VdisParam = ResolveArgv(argv(0))
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
'    Call tl_wait(ResolveArgv(argv(1)))
     Call TheHdw.Wait(ResolveArgv(argv(1)))
    
    ' Restore original voltages
'    Call tl_DpsSetPrimaryVoltages(chans, OriginalVoltages)
    TheHdw.DPS.chans(chans).PrimaryVoltages = OriginalVoltages
'     Call tl_DpsSetOutputSourceChannels(chans, TL_DPS_PrimaryVoltage)
    TheHdw.DPS.chans(chans).OutputSource = dpsPrimaryVoltage
End Function


Public Function ResolveArgv(s As String) As Double
    Dim v As Variant
    Dim varname As String
    Dim s_variant As Variant
    Dim ret_val As Double
    Dim cell_text As Variant
    Dim tname As String
    Dim var_val As Double
        
    s_variant = s
    
    If (IsNumeric(s_variant)) Then
        ret_val = CDbl(Val(s))
        'MsgBox "argument is a fixed number = " & ret_val
    Else
        ret_val = TheExec.VariableValue(s)
        'MsgBox "varname is " & s & " value = " & ret_val
    End If
    
    ResolveArgv = ret_val

End Function
 
Public Function IntCycle_power(argc As Long, argv() As String) As Long
    Select Case argc
    Case 2
        Call cycle_power(CDbl(argv(0)), CDbl(argv(1)))
    Case 3
        Call cycle_power(CDbl(argv(0)), CDbl(argv(1)), CDbl(argv(2)))
    Case 4
        Call cycle_power(CDbl(argv(0)), CDbl(argv(1)), CDbl(argv(2)), CDbl(argv(3)))
    Case Else
        
    End Select
End Function

Public Function WaitXSec(argc As Long, argv() As String) As Long
    If argc > 0 Then TheHdw.Wait (CDbl(argv(0)))
End Function
