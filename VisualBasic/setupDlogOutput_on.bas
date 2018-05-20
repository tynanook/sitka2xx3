Attribute VB_Name = "setupDlogOutput_on"
 
' [=============================================================================================]
' [ MCU32 Autodatalog Module                                                                    ]
' [=============================================================================================]
' [=== Revision History ========================================================================]
' [ A0 : 03/29/2012 - ST - Add chanmap on file name
' [ A1 : 06/06/2012 - ST - Add time stamp when not find lotID
' [ A1 : 06/06/2012 - ST - Change save location from \\chip\j750summary\ to \\chip\datalogs\
' [ A3 : 07/19/2012 - ST - Turn key txt and stdf file get it at \\chip\MCU32\MCU32_AtoDlog_Setup\XXXX0\DatalogSetup.txt
' [ A4 : 01/07/2013 - ST - Update setup autodlog location to \\chip\datalogs
' [ A5 : 02/11/2013 - ST - Update turn key checking module
' [ A6 : 02/29/2013 - ST - Support Probe test auto datalog
' [ A7 : 03/13/2013 - ST - Add Scribe, lot namber and mpc from OI data to dlog file name.
' [ A8 : 03/26/2013 - ST - Add Network checking for FT auto datalog
' [ A9 : 04/04/2013 - ST - Fix not show scribe number in probe auto datalog in file name.
' [ B0 : 04/23/2013 - ST - Fix bug when change scribe
' [ B1 : 05/20/2013 - ST - Change path location from network \\chip\datalog to Local drive c:\Prod_STDF
' [ B2 : 05/21/2013 - SC - Add new path to ltore Dlog at \\chip\datalog\MCU32_AtoDlog\TMD to supportSMART Box.
' [ B3 : 05/21/2013 - ST - Separate Atodlog setup file for FT & WS.
' [ B4 : 07/08/2013 - ST - Change save datalog path to c:\Prod_STDF\MCU32_AutoDlog\ support OI 7.7.1
' [    : 07/08/2013 - ST - Strip test still locate at c:\Prod_STDF\qualdata due to ST OI does not support
' [ B5 : 08/05/2013 - ST - Support OI 7.7.1 & Probe Check out & Remove not in use fuction.
' [ B6 : 08/15/2013 - ST - fix para not show
' [ B7 : 08/16/2013 - ST - fix Probe bug
' [ B8 : 11/06/2013 - ST - Support OI 7.7.7
' [=============================================================================================]



Option Explicit

Dim StartLotID As String ' Use for compare start new lot
Dim StartScribe As String ' Use for compare start Scribe
Dim countNumber As Integer
Dim flagSetupAuto As Boolean
Dim dlogFile As String


Public Sub setupDlogOutput()

On Error GoTo ErrHandler

    Dim LotID As String
    
   ' StartLotID = CurrentLotNum
    LotID = TheExec.Datalog.setup.LotSetup.LotID
    StartLotID = LotID

    StartLotID = Replace(StartLotID, ".", "_")
    If StartLotID = "" Then
        StartLotID = "NO_LOT_ID"
    End If
    'StartScribe = CurrentScribe
    ' ------------- Check setup datalog folder in drive c ------------------------
    
    Call CheckSetupFiles

    Call DlogManager

Exit Sub 'normal exit of function
ErrHandler:
    Debug.Print ("Function DatalogAllDC had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub


Public Sub DlogManager()
On Error GoTo ErrHandler

    Dim job As String
    Dim Part As String
    Dim env As String
    Dim chanMap As String
    Dim probeScribe As String
    Dim mchpMPC As String
    Dim testerLotID As String
    Dim mchpDevice As String
    Dim currentTime As String
    Dim LotID As String
   
    job = TheExec.CurrentJob
    Part = TheExec.CurrentPart
    env = TheExec.CurrentEnv
    chanMap = TheExec.CurrentChanMap
    probeScribe = "" 'DEbug for DM920
    mchpMPC = "ZY004" 'DEbug for DM920
   
     LotID = TheExec.Datalog.setup.LotSetup.LotID
     'LotID = Replace(LotID, ".", "_")
     testerLotID = LotID
    'testerLotID = CurrentLotNum

    testerLotID = Replace(StartLotID, ".", "_")
    If testerLotID = "" Then
        testerLotID = "NO_LOT_ID"
    End If
    mchpDevice = UCase(Left(TheExec.ExcelHandle.ActiveWorkbook.Name, 5))
    currentTime = Format(Now, "MM-dd-yy")
    
    If InStr(chanMap, "J") Then
        Call createProbeDlog(mchpDevice, job, mchpMPC, Part, env, testerLotID, currentTime, chanMap, probeScribe)
    Else
        Call createFTDlog(mchpDevice, job, mchpMPC, Part, env, testerLotID, currentTime, chanMap)
    End If
        

Exit Sub
ErrHandler:
    Debug.Print ("Fuction DlogManeger had Error")
    On Error GoTo 0
End Sub

Public Sub createProbeDlog(probeDevice As String, probeJob As String, probeMPC As String, probePart As String, probeEnv As String, probeLotID As String, probeTimeStamp As String, probeChanMap As String, probeScribe As String)
On Error GoTo ErrHandler
    
    Dim dlogDir As String
    dlogDir = "\\chip\datalogs\"
    dlogFile = dlogDir & "\" & probeDevice & "_" & probeJob & "_" & probeMPC & "_" & probePart & "_" & probeChanMap & "_" & probeEnv & "_" & probeLotID & "_" & probeScribe & "_" & probeTimeStamp
    Call ApplyDlogSetup("Probe_Mode", dlogFile)
Exit Sub
ErrHandler:
    Debug.Print ("Fuction createProbeDlog had Error")
    On Error GoTo 0
End Sub

Public Sub createFTDlog(ftDevice As String, ftJob As String, ftMPC As String, ftPart As String, ftEnv As String, ftLotID As String, ftTimeStamp As String, ftChanMap As String)
On Error GoTo ErrHandler

    Dim dlogDir As String

    dlogDir = "\\chip\datalogs\"
    If Not (DirExists(dlogDir)) Then
        MkDir dlogDir
    End If
    
    If (checkStripLocation = True) Then
        dlogDir = dlogDir & "\" & "qualdata"
    Else
        dlogDir = dlogDir & "\" & "WSG_AutoDlog"
    End If
    
    If Not (DirExists(dlogDir)) Then
        MkDir dlogDir
    End If
    dlogDir = dlogDir & "\" & ftDevice
    If Not (DirExists(dlogDir)) Then
        MkDir dlogDir
    End If
        ftTimeStamp = ftTimeStamp & "_" & Format(Now, "hh-nn")
    dlogFile = dlogDir & "\" & ftDevice & "_" & ftJob & "_" & ftMPC & "_" & ftPart & "_" & ftChanMap & "_" & ftEnv & "_" & ftLotID & "_" & ftTimeStamp
    Call ApplyDlogSetup("FT_Mode", dlogFile)
Exit Sub
ErrHandler:
    Debug.Print ("Fuction createFTDlog had Error")
    On Error GoTo 0
End Sub

Public Function checkStripLocation() As Boolean
On Error GoTo ErrHandler
    
    Dim chanMap As String
    chanMap = TheExec.CurrentChanMap
    If InStr(chanMap, "x9") Or InStr(chanMap, "x10") Or InStr(chanMap, "x15") Or InStr(chanMap, "x20") Or InStr(chanMap, "x21") Or InStr(chanMap, "x24") Or InStr(chanMap, "x25") Or InStr(chanMap, "x32") Then
        checkStripLocation = True
    Else
        checkStripLocation = False
    End If
Exit Function
ErrHandler:
    Debug.Print ("Fuction checkStripLocation had Error")
    On Error GoTo 0
End Function

Public Sub ApplyDlogSetup(dlogMode As String, dlogPart As String)
    Dim SetupDlogFile As String
    Dim fso As Object
    Dim CommandList() As String
    Dim fileLine As String
    Dim temp As String
    Dim ctr As Long
    Dim tmp() As String
    Dim fCommandFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If dlogMode = "FT_Mode" Then
        SetupDlogFile = "\\chip\datalogs\MCU32_AutoDlog\MCU32_AutoDlog_Setup\FT\DatalogSetup.txt"  ' ---- FT --
    Else
        SetupDlogFile = "\\chip\datalogs\MCU32_AutoDlog\MCU32_AutoDlog_Setup\Probe\DatalogSetup.txt"  ' ---- Probe --
    End If
    
     If fso.FileExists(SetupDlogFile) Then
            Set fCommandFile = fso.OpenTextFile(SetupDlogFile, 1)
            temp = fCommandFile.ReadAll
            fCommandFile.Close
            CommandList = Split(temp, vbCrLf)
            For ctr = LBound(CommandList) To UBound(CommandList)
                fileLine = Trim(CommandList(ctr))
                If InStr(1, fileLine, "=") And Left(fileLine, 1) = "$" Then
                    tmp = Split(fileLine, "=")
                Select Case UCase(Trim(tmp(0)))
                    Case "$TEXTFILE"    ' ---- Check Text type
                        If UCase(Trim(tmp(1))) = "ON" Then
                            TheExec.Datalog.setup.LotSetup.DatalogOn = True
                            TheExec.Datalog.setup.DatalogSetup.SetupFile = "C:\Temp\DlogAllDC"
                                                        TheExec.Datalog.setup.DatalogSetup.SelectSetupFile = True
                            TheExec.Datalog.setup.DatalogSetup.TextOutputFile = dlogPart
                            TheExec.Datalog.setup.DatalogSetup.TextOutput = True
                        Else
                            TheExec.Datalog.setup.LotSetup.DatalogOn = True
                            TheExec.Datalog.setup.DatalogSetup.TextOutput = False
                        End If
                    Case "$STDFFILE"    ' ---- Check STDF type (!!!! FOR FT ONLY !!!!!)
                        If (dlogMode = "FT_Mode") Then
                            If UCase(Trim(tmp(1))) = "ON" Then
                                TheExec.Datalog.setup.DatalogSetup.SetupFile = "C:\Temp\DlogAllDC"
                                TheExec.Datalog.setup.DatalogSetup.STDFOutputFile = dlogPart
                                TheExec.Datalog.setup.DatalogSetup.STDFOutput = True
                            End If
                        End If
                End Select      ' end case
                End If          ' end line check
            Next ctr            ' end for loop
        Else
            TheExec.Datalog.setup.LotSetup.DatalogOn = True
            TheExec.Datalog.setup.DatalogSetup.SetupFile = "C:\Temp\DlogAllDC"
            TheExec.Datalog.setup.DatalogSetup.SelectSetupFile = True
            TheExec.Datalog.setup.DatalogSetup.TextOutputFile = dlogPart
            TheExec.Datalog.setup.DatalogSetup.TextOutput = True
        End If                  ' end check file exists
    TheExec.Datalog.ApplySetup  ' Applies the selected setup file
End Sub

Public Sub CheckSetupFiles()

On Error GoTo ErrHandler
    
    If Dir("C:\Temp", vbDirectory) = "" Then
        MkDir ("C:\Temp")
    End If
    
    If Dir("C:\Temp\DlogAllDC") = "" Or Dir("C:\Temp\DlogFailDC") = "" Then
        Call CreateSetupFiles
    End If

Exit Sub 'normal exit of function
ErrHandler:
    Debug.Print ("Function CheckSetupFiles had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0


End Sub 'CheckSetupFiles

Public Sub CreateSetupFiles()

    Dim filenum As Integer
    
    On Error GoTo ErrHandler
    
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
ErrHandler:
    Debug.Print ("Function CreateSetupFiles had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0

End Sub 'CreateSetupFiles

Function DirExists(DirName As String) As Boolean
    ' Return True if a directory exists
    ' (the directory name can also include a trailing backslash)

    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Public Sub setupNumber()

    On Error GoTo ErrHandler

    Dim currentLotID As String
    Dim probeScribe As String
    Dim job As String
    Dim fso As Object
    Dim LotID As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    currentLotID = CurrentLotNum
    probeScribe = "" 'DEbug for DM920
    job = LCase(TheExec.CurrentJob)
     
     LotID = TheExec.Datalog.setup.LotSetup.LotID
    ' LotID = Replace(LotID, ".", "_")
     currentLotID = LotID
    currentLotID = Replace(currentLotID, ".", "_")
    If currentLotID = "" Then
        currentLotID = "NO_LOT_ID"
    End If

    ' ----------------------------------------------------
        
    If currentLotID <> StartLotID Then
        TheExec.Datalog.setup.DatalogSetup.TextOutput = False
        TheExec.Datalog.ApplySetup
        Call setupDlogOutput
    End If
                
    If probeScribe <> StartScribe Then
        TheExec.Datalog.setup.DatalogSetup.TextOutput = False
        TheExec.Datalog.ApplySetup
        Call setupDlogOutput
    End If
    
        If fso.FileExists(dlogFile & ".txt") Then
        ' Already have datalog file
    Else
                Call setupDlogOutput
        End If

                  
    Exit Sub 'normal exit of function
ErrHandler:
        Debug.Print ("Function DatalogAllDC had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
        On Error GoTo 0

End Sub 'setupDlogOutput

Public Sub SetupFiles()
    
    Dim filenum As Integer
    
    On Error GoTo ErrHandler
    
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
ErrHandler:
    Debug.Print ("Function CreateSetupFiles had Error" & VBA.vbCrLf & "VBA Error number is " & Format(VBA.err.Number) & VBA.vbCrLf & VBA.err.Description & VBA.vbCrLf)
    On Error GoTo 0


End Sub 'CreateSetupFiles

