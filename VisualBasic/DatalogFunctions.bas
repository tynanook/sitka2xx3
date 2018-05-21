Attribute VB_Name = "DatalogFunctions"
Option Explicit
Public ButtonsAdded As Boolean
Public HramEnabled As Boolean
Public DlogButton As Integer    'Set to 1 if Datalog Off
                                'Set to 2 if Datalog All DC
                                'Set to 3 if Datalog Fail DC


Public Sub AddDatalogButtons()

Dim i As Integer
Dim j As Integer
Dim NewCommand As CommandBarButton
     
On Error GoTo errHandler
   

If ButtonsAdded = True Then Exit Sub

Call CreateSetupFiles   ' Creates the Datalog Setup File
    
    
For i = 1 To Application.CommandBars.count
    If Application.CommandBars(i).Visible = True Then
        'Debug.Print Application.CommandBars(i).Name
        
        
        If Application.CommandBars(i).name = "IG-XL Toolbar" Then
        
            'Datalog Off
            Set NewCommand = Application.CommandBars(i).Controls.Add(msoControlButton, , , , True)
            NewCommand.Caption = "Datalog Off"
            NewCommand.DescriptionText = "Datalog Off"
            NewCommand.TooltipText = "Datalog Off"
            NewCommand.OnAction = "DatalogOff"
            NewCommand.FaceId = 342
            NewCommand.Enabled = True
            NewCommand.Visible = True
            NewCommand.BeginGroup = True
                   
            'Datalog All DC
            Set NewCommand = Application.CommandBars(i).Controls.Add(msoControlButton, , , , True)
            NewCommand.Caption = "Datalog All DC"
            NewCommand.DescriptionText = "Datalog All DC"
            NewCommand.TooltipText = "Datalog All DC"
            NewCommand.OnAction = "DatalogAllDC"
            NewCommand.FaceId = 343
            NewCommand.Enabled = True
            NewCommand.Visible = True
            
            
               'Datalog Fail DC
            Set NewCommand = Application.CommandBars(i).Controls.Add(msoControlButton, , , , True)
            NewCommand.Caption = "Datalog Fail DC"
            NewCommand.DescriptionText = "Datalog Fail DC"
            NewCommand.TooltipText = "Datalog Fail DC"
            NewCommand.OnAction = "DatalogFailDC"
            NewCommand.FaceId = 352
            NewCommand.Enabled = True
            NewCommand.Visible = True
            
                'Capture HRAM Fails
            Set NewCommand = Application.CommandBars(i).Controls.Add(msoControlButton, , , , True)
            NewCommand.Caption = "Capture HRAM Fails"
            NewCommand.DescriptionText = "Capture HRAM Fails"
            NewCommand.TooltipText = "Capture HRAM Fails"
            NewCommand.OnAction = "CaptureHRAMFails"
            NewCommand.FaceId = 214
            NewCommand.Enabled = True
            NewCommand.Visible = True
            NewCommand.BeginGroup = True
      
         
            End If
        End If
    Next

ButtonsAdded = True
HramEnabled = False


Call DatalogOff

Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function AddButtons had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'AddButtons


Public Sub DatalogOff()

Dim NewCommand As CommandBarButton
    
On Error GoTo errHandler
    
TheExec.DataLog.Setup.DatalogSetUp.WindowOutput = False     'Stop Datalog from the window
TheExec.DataLog.Setup.LotSetup.DatalogOn = False             'Turns off the Datalog
TheExec.DataLog.Setup.DatalogSetUp.SelectSetupFile = False   'Un-select the setup file
TheExec.DataLog.Setup.DatalogSetUp.HeaderEveryRun = False    'Off Header every Time

DlogButton = 1

Call ToggleDCButton
    
Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function DatalogOff had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'DatalogOff



Public Sub DatalogAllDC()

Dim NewCommand As CommandBarButton

On Error GoTo errHandler

Call CheckSetupFiles

TheExec.DataLog.Setup.LotSetup.DatalogOn = True             'Turns the Datalog On
TheExec.DataLog.Setup.DatalogSetUp.WindowOutput = True      'Create a windonw to output the Datalog
TheExec.DataLog.Setup.DatalogSetUp.SetupFile _
        = "C:\Temp\DlogAllDC"                               'Point to the save datalog files
TheExec.DataLog.Setup.DatalogSetUp.SelectSetupFile = True   'Select the setup file
TheExec.DataLog.Setup.DatalogSetUp.HeaderEveryRun = True    'On Header every Time
TheExec.DataLog.ApplySetup                                  'Applies the selected setup file
    
DlogButton = 2

Call ToggleDCButton

Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function DatalogAllDC had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'DatalogAllDC



Public Sub DatalogFailDC()

Dim NewCommand As CommandBarButton
    
On Error GoTo errHandler

Call CheckSetupFiles

TheExec.DataLog.Setup.LotSetup.DatalogOn = True             'Turns the Datalog On
TheExec.DataLog.Setup.DatalogSetUp.WindowOutput = True      'Create a windonw to output the Datalog

TheExec.DataLog.Setup.DatalogSetUp.SetupFile _
        = "C:\Temp\DlogFailDC"                              'Point to the save datalog files"
TheExec.DataLog.Setup.DatalogSetUp.SelectSetupFile = True   'Select the setup file
TheExec.DataLog.ApplySetup                                  'Applies the selected setup file

DlogButton = 3

Call ToggleDCButton

Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function DatalogFailDC had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'DatalogFailDC


Public Sub CaptureHRAMFails()

Dim NewCommand As CommandBarButton

On Error GoTo errHandler

'If HramEnabled = True Then
'    Call TheHdw.Digital.HRAM.SetCapture(captAll, True)              'capture all, and compress repeats
'    Call TheHdw.Digital.HRAM.GetTrigger(trigFirst, False, 0, False) 'Trigger on first fail and stop on full disabled
'    Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Capture HRAM Fails")
'    NewCommand.state = msoButtonUp
'    NewCommand.TooltipText = "Capture HRAM Fails:OFF"
'    HramEnabled = False
'    Exit Sub
'End If
    
      
'If HramEnabled = False Then
    Call TheHdw.Digital.HRAM.SetCapture(captFailSTV, True)          'capture Fais + stv, and compress repeats
    Call TheHdw.Digital.HRAM.GetTrigger(trigFirst, False, 0, True)  'Trigger on first fail and stop on full
    Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Capture HRAM Fails")
    NewCommand.State = msoButtonDown
    NewCommand.TooltipText = "Capture HRAM Fails:ON"
    HramEnabled = True
    'Exit Sub
'End If
    
Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function CaptureHRAMFails had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'CaptureHRAMFails

Public Sub ToggleDCButton()

Dim i As Integer
Dim NewCommand As CommandBarButton

On Error GoTo errHandler

Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog Off")
    NewCommand.State = msoButtonUp
Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog All DC")
    NewCommand.State = msoButtonUp
Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog Fail DC")
    NewCommand.State = msoButtonUp


Select Case DlogButton

    Case 1
        Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog Off")
        NewCommand.State = msoButtonDown
  
    Case 2
        Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog All DC")
        NewCommand.State = msoButtonDown
        
    Case 3
        Set NewCommand = Application.CommandBars("IG-XL Toolbar").Controls("Datalog Fail DC")
        NewCommand.State = msoButtonDown
    
    
End Select

Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function CaptureHRAMFails had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0


End Sub 'ToggleDCButton

Public Sub CreateSetupFiles()

Dim filenum As Integer

On Error GoTo errHandler

filenum = FreeFile

Open "C:\Temp\DlogAllDC" For Output As #filenum
Print #filenum, "1.0|0|1|0|0|0|0|0|0|||0|1|0|0|1|0|0|"
Print #filenum, "1|DlogAllDC|1|0|0|1|"
Print #filenum, "0|Default|1|0|0|0|0|0|2|0|0|0|0|0|0|0|0|0|0|0|0|"
Close #filenum

filenum = FreeFile

Open "C:\Temp\DlogFailDC" For Output As #filenum
Print #filenum, "1.0|0|1|0|0|0|0|0|0|||0|1|0|0|1|0|0|"
Print #filenum, "1|DlogFailDC|1|2|0|1|"
Print #filenum, "0|Default|1|0|0|0|0|0|2|0|0|0|0|0|0|0|0|0|0|0|0|"
Close #filenum
 
Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function CreateSetupFiles had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0


End Sub 'CreateSetupFiles

Public Sub CheckSetupFiles()

On Error GoTo errHandler

If Dir("C:\Temp", vbDirectory) = "" Then
    MkDir ("C:\Temp")
End If

If Dir("C:\Temp\DlogAllDC") = "" Or Dir("C:\Temp\DlogFailDC") = "" Then
    Call CreateSetupFiles
End If

Exit Sub 'normal exit of function
errHandler:
    Debug.Print ("Function CheckSetupFiles had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0


End Sub 'CheckSetupFiles
