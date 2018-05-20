Attribute VB_Name = "User_Procs"
' [=============================================================================]
' [                    Visual Basic Module                                      ]
' [                        User_Procs                                           ]
' [=============================================================================]
' [                                                                             ]
' [                   MICROCHIP TECHNOLOGY INC.                                 ]
' [                   2355 WEST CHANDLER BLVD.                                  ]
' [                   CHANDLER AZ 85224-6199                                    ]
' [                       (480) 792-7200                                        ]
' [                                                                             ]
' [================= Copyright Statement =======================================]
' [                                                                             ]
' [   THIS PROGRAM AND ITS VECTORS ARE  PROPERTY OF Microchip Technology Inc.   ]
' [   USE, COPY, MODIFY, OR TRANSFER OF THIS PROGRAM, IN WHOLE OR IN PART,      ]
' [   AND IN ANY FORM OR MEDIA, EXCEPT AS EXPRESSLY PROVIDED FOR BY LICENSE     ]
' [   FROM Microchip Technology Inc. IS FORBIDDEN.                              ]
' [                                                                             ]
' [================= Revision History ==========================================]
' [ REV.   DATE    OWN  COMMENT                                                 ]
' [ ^^^^   ^^^^    ^^^  ^^^^^^^                                                 ]
' [                                                                             ]
' [ 0.5    06jun01 svo  - Created. Came from orig UI_Procs.                     ]
' [                       Added skeleton Exec Interpose Functions               ]
' [ 0.5a   06jun01 svo  - Configured in 9707b_d15 test program                  ]
' [ 1.0    15oct01 svo  - Mchp Beta release V8.0                                ]
' [ 1.0    14aug02 svo  - Mchp Production Release V5 (no changes)               ]
' [ 1.0    14oct02 svo  - Mchp Production Release V6 (no changes)               ]
' [ 1.0    15nov02 svo  - Mchp Beta release V9.0 (no changes)                   ]
' [ 1.0    20nov02 svo  - Mchp Production Release V7 (no changes)               ]
' [ 1.0    03dec02 svo  - Mchp Development Mchp_User_Procs_Dev_V12.03.02.1      ]
' [                       From Production Release V7                            ]
' [ 1.0    13dec02 svo  - Mchp Production Release V8 (no changes)               ]
' [ 1.0    13dec02 svo  - Mchp Development Mchp_User_Procs_Dev_V12.13.02.1      ]
' [ 1.0    19dec02 svo  - Mchp Production Release V9 - no changes               ]
' [ 2.0    13Mar03 jpe  - Mchp Production Release V11                           ]
' [                     - Add error handler routines to the functions           ]
' [                     - Add "Call ml_OnProgramValidated to OnProgramValidated ]
' [=============================================================================]
' Module Instructions:
' 1. Import this module into the test program (in VBA, ALT-F11, slect File=>Import FIle).
' 2. Uncomment out the Exec interpose function that is needed.
'    Select the text and then on the Edit toolbar (view=>toolbars=>edit) select uncomment block
' 3. Add the needed VBA code.


Option Explicit


' ********************************************************************************
'                 EXEC INTERPOSE FUNCTIONS
'                       (Teradyne)
' ********************************************************************************



' Function OnTesterInitialized() As Integer
''   Immediately at the conclusion of the initialization process
'
'   On Error GoTo ErrHandler
'
'   'Any user code goes here
'
'   OnTesterInitialized = TL_SUCCESS
'Exit Function      'normal exit of function
'ErrHandler:
'    OnTesterInitialized = TL_ERROR
'    Call gError.AddError(VBA.err.Number, "TestProgram::OnTesterInitialized", VBA.err.Description, True)
'      'Uncomment out the code below if you want a message box to notify the user of an error.
'      'The gError.AddError method, writes it out to the immediate window and the dataloger.
'      'For this routine, dataloger is usually not running, so need to force an error for it to be seen.
'    Call TheExec.ErrorLogMessage("Function OnTesterInitialized had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    Call TheExec.ErrorReport
'    On Error GoTo 0
' End Function       ' OnTesterInitialized
'




 Function OnProgramLoaded() As Integer
'   Immediately at the conclusion of the load process

   On Error GoTo errHandler

   'Any user code goes here
   'This replaces the Mchp Pre_Validation_Init routine
   'Remove pins from the TDR calibration Here or in OnProgramValidated
   'Example:
   'TheHdw.Digital.ACCalExcludePins ("pin1, pin2")
   'For 9707b
   'TheHdw.Digital.ACCalExcludePins ("mclr,ra4")
   
    TheHdw.Digital.ACCalExcludePins ("VBAT_PMU,XTAL_CONT,MW_DIG_TRIG")   'TW101
    Call AddDatalogButtons

   OnProgramLoaded = TL_SUCCESS
Exit Function      'normal exit of function
errHandler:
    OnProgramLoaded = TL_ERROR
'    Call gerror.AddError(VBA.Err.Number, "TestProgram::OnProgramLoaded", VBA.Err.Description, True)
      'Uncomment out the code below if you want a message box to notify the user of an error.
      'The gError.AddError method, writes it out to the immediate window and the dataloger.
      'For this routine, at load time the dataloger is not running, so need to force an error for it to be seen.
    Call TheExec.ErrorLogMessage("Function OnProgramLoaded had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    Call TheExec.ErrorReport
    On Error GoTo 0
 End Function       ' OnProgramLoaded
'


Function OnProgramValidated() As Integer
    
'    Dim jobN As String
'    Dim devN As String
'    Dim envN As String
     Dim lSite As Long
         
'    Call TheExec.DataManager.GetJobContext(jobN, devN, envN)
    
    'Call SetupCtoArray
    
    ' Populate array of all site numbers
    ReDim all_sites(0 To TheExec.Sites.ExistingCount - 1)
    For lSite = 0 To TheExec.Sites.ExistingCount - 1
        all_sites(lSite) = lSite
    Next lSite
    
    'Call spec_array_onValidate
    'Call globals_onValidate
    'Call OscSpec_OnValidate
    'Call dut_info_onValidate
    'Call AdcLin_onValidate
    ' Call autoDlog_onValidate
    

    
    OnProgramValidated = TL_SUCCESS
    
    'If UCase(Trim(Right(TheExec.CurrentJob, 5))) = "tw101" Then Call RFOnProgramValidated_TW101  'Added for TW101 RF tests (TW101)
    'End If
    
    Call RFOnProgramValidated_LoRa
    
End Function


' Function OnProgramValidated() As Integer
''   Immediately at the conclusion of the validate process
'   On Error GoTo ErrHandler
'
'   'Any user code goes here
'
'    Call ml_OnProgramValidated 'Initial values for MCHP library functions
'
'   OnProgramValidated = TL_SUCCESS
'Exit Function      'normal exit of function
'ErrHandler:
'    OnProgramValidated = TL_ERROR
'    Call gError.AddError(VBA.err.Number, "TestProgram::OnProgramValidated", VBA.err.Description, True)
'      ''Uncomment out the code below if you want a message box to notify the user of an error.
'      ''The gError.AddError method, writes it out to the immediate window and the dataloger.
'    'Call TheExec.ErrorLogMessage("Function OnProgramValidated had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    'Call TheExec.ErrorReport
'    'On Error GoTo 0
' End Function  'OnProgramValidated
'




' Function OnTDRCalibrated() As Integer
''   Immediately at the conclusion of the TDR calibration process
'   On Error GoTo ErrHandler
'
'   'Any user code goes here
'
'   OnTDRCalibrated = TL_SUCCESS
'Exit Function      'normal exit of function
'ErrHandler:
'    OnTDRCalibrated = TL_ERROR
'    Call gError.AddError(VBA.err.Number, "TestProgram::OnTDRCalibrated", VBA.err.Description, True)
'      ''Uncomment out the code below if you want a message box to notify the user of an error.
'      ''The gError.AddError method, writes it out to the immediate window and the dataloger.
'    'Call TheExec.ErrorLogMessage("Function OnTDRCalibrated had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    'Call TheExec.ErrorReport
'    'On Error GoTo 0
' End Function  'OnTDRCalibrated
'




' Function OnProgramStarted() As Integer
''   Immediately after “pre-job reset” when the test program starts
'   On Error GoTo ErrHandler
'
'   'Any user code goes here
'
'   OnProgramStarted = TL_SUCCESS
'Exit Function      'normal exit of function
'ErrHandler:
'    OnProgramStarted = TL_ERROR
'    Call gError.AddError(VBA.err.Number, "TestProgram::OnProgramStarted", VBA.err.Description, True)
'      ''Uncomment out the code below if you want a message box to notify the user of an error.
'      ''The gError.AddError method, writes it out to the immediate window and the dataloger.
'    'Call TheExec.ErrorLogMessage("Function OnProgramStarted had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    'Call TheExec.ErrorReport
'    'On Error GoTo 0
' End Function  'OnProgramStarted
'



' Function OnProgramEnded() As Integer
''   Immediately prior to “post-job reset” when the test program completes
'   On Error GoTo ErrHandler
'
'   'Any user code goes here
'
'   OnProgramEnded = TL_SUCCESS
'Exit Function      'normal exit of function
'ErrHandler:
'    OnProgramEnded = TL_ERROR
'    Call gError.AddError(VBA.err.Number, "TestProgram::OnProgramEnded", VBA.err.Description, True)
'      ''Uncomment out the code below if you want a message box to notify the user of an error.
'      ''The gError.AddError method, writes it out to the immediate window and the dataloger.
'    'Call TheExec.ErrorLogMessage("Function OnProgramEnded had an Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
'    'Call TheExec.ErrorReport
'    'On Error GoTo 0
' End Function  'OnProgramEnded



