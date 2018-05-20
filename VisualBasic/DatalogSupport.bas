Attribute VB_Name = "DatalogSupport"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AutoDatalog Module :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module history :
'
' Rung-aroon P.  Jan, 31 2011    Initial create module auto save datalog by read set up file.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function catch_doall() As Long
    ' This function is intended to run as a test instance.
    ' It checks to see if the "Do All" option is selected.
    ' If it is, it will issue a warning as a message box
    ' disable the "Do All" option, fail the test and stop
    ' execution of the flow.
        
    Dim testNum As Long
    Dim Site As Long
    Dim reportFail As Boolean
    
    
    reportFail = False ' by default
            
    ' Additionally, check to make sure DoAll is not enabled!
    If (TheExec.RunOptions.DoAll = True) Then
        TheExec.RunOptions.DoAll = False            ' disable Do All option.
        Call MsgBox("Do All is not supported in this flow and has been disabled." & Chr(13) & _
                    "Please contact ENGINEERING.", vbCritical, " - W A R N I N G - ")
        TheExec.Datalog.WriteComment ("ERROR: Do All option is not supported in this flow.")
        reportFail = True
   
    End If
    
    If reportFail Then
        If TheExec.Sites.SelectFirst <> loopDone Then
            Do
                Site = TheExec.Sites.SelectedSite
                testNum = TheExec.Sites.Site(Site).testnumber
                TheExec.Sites.Site(Site).TestResult = siteFail
                Call TheExec.Datalog.WriteFunctionalResult(Site, testNum, logTestFail)
            Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
        End If
    Else
        ' report pass
        If TheExec.Sites.SelectFirst <> loopDone Then
            Do
                Site = TheExec.Sites.SelectedSite
                testNum = TheExec.Sites.Site(Site).testnumber
                TheExec.Sites.Site(Site).TestResult = sitePass
                Call TheExec.Datalog.WriteFunctionalResult(Site, testNum, logTestPass)
            Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
        End If
    End If
                 

End Function

Function autoDlog_onValidate() As Long

    Dim fso As New FileSystemObject
    Dim fCommandFile As Object
    Dim CommandList() As String
    Dim fileLine As String
    Dim temp As String
    Dim ctr As Long
    Dim tmp() As String
    Dim SetupDlogPath As String
    Dim DatalogPath As String
    Dim CntSite As String
    
    Dim job As String
    Dim dev As String
    Dim env As String
    Dim LotID As String
    Dim jobContext As String
    Dim dateTimeStamp As String
    Dim filePathNameTxt As String
    Dim filePathNameSTD As String
        
    Const AUTODLOG_SUFFIX = ""
    Const AUTODLOG_EXT_TXT = ".txt"
    Const AUTODLOG_EXT_STDF = ".stdf"
      
    ' Put the location of text file reading setup and location to save datalog.
    SetupDlogPath = "\\chip\datalogs\MTAI_SMTD\Autolog_setup\LEBS0\LEBS0_LF\DatalogSetup.txt" 'Put the locate of setup file reading here.
    DatalogPath = "\\chip\datalogs\MTAI_SMTD\AutoDlogs\LEBS0\LEBS0_LF\" 'Put the locate to save datalog here.
    
    Call TheExec.DataManager.GetJobContext(job, dev, env)
    jobContext = dev & "_" & job & "_" & env
    dateTimeStamp = Format(Now, "yy_mm_dd_hh_mm_ss")
   ' LotID = TheExec.Datalog.setup.LotSetup.LotID
    CntSite = LCase(TheExec.CurrentChanMap)
    CntSite = Mid$(CntSite, 3, 1)
    
    If IsNumeric(CntSite) Then
    
        Call turnOffdatalog
    
    Else
    
            If (Mid$(job, 1, 1) Like "q") Or (Mid$(job, 1, 1) Like "s") Then
    
                Call turnOffdatalog
    
            Else
    
        If LCase(TheHdw.Computer.Name) Like "t*j750*" Or LCase(TheHdw.Computer.Name) Like "*mth*" Then
    
            If fso.FileExists(SetupDlogPath) Then
                Set fCommandFile = fso.OpenTextFile(SetupDlogPath, 1)
                temp = fCommandFile.ReadAll
                fCommandFile.Close
        
                CommandList = Split(temp, vbCrLf)
        
                    For ctr = LBound(CommandList) To UBound(CommandList)
                        fileLine = Trim(CommandList(ctr))
                            
                            If InStr(1, fileLine, "=") And Left(fileLine, 1) = "$" Then
               
                                tmp = Split(fileLine, "=")
                                    Select Case UCase(Trim(tmp(0)))
                    
                                     Case "$DLOGPATH"
                                        
                                        If UCase(Trim(tmp(1))) = "" Then
                                            
                                                Set FSOobj = CreateObject("Scripting.FilesystemObject")
    
                                                    If FSOobj.FolderExists(DatalogPath) = False Then
                                                        FSOobj.CreateFolder DatalogPath
                                                        DatalogPath = DatalogPath
                        
                                                    Else
                    
                                                        DatalogPath = DatalogPath
                        
                                                    End If
                
                                                        Set FSOobj = Nothing
    
                                                    Else
                
                                                        DatalogPath = UCase(Trim(tmp(1)))
                
                                            End If
                                                                                    
                                     Case "$DATALOG"
                                       
                                        If UCase(Trim(tmp(1))) = "ON" Then
                                        
                                            TheExec.Datalog.setup.LotSetup.DatalogOn = True
                                            TheExec.Datalog.setup.DatalogSetup.SetupFile = "DCS"
                                            TheExec.Datalog.setup.DatalogSetup.SelectSetupFile = True
                                        
                                        Else
                                        
                                            Call turnOffdatalog
                                            TheExec.Datalog.ApplySetup
                                            Exit Function
                
                                        End If
                                        
                                    Case "$TEXTFILE"
                                        
                                        If UCase(Trim(tmp(1))) = "ON" Then
                                            filePathNameTxt = DatalogPath & jobContext & "_" & dateTimeStamp & AUTODLOG_SUFFIX & AUTODLOG_EXT_TXT
                                            TheExec.Datalog.setup.DatalogSetup.TextOutputFile = filePathNameTxt
                                            TheExec.Datalog.setup.DatalogSetup.TextOutput = True
                                        
                                        Else
                                        
                                            TheExec.Datalog.setup.DatalogSetup.TextOutput = False
                
                                        End If
                                        
                                    Case "$STDFFILE"
                                      
                                        If UCase(Trim(tmp(1))) = "ON" Then
                                        
                                            filePathNameSTD = DatalogPath & jobContext & "_" & dateTimeStamp & AUTODLOG_SUFFIX & AUTODLOG_EXT_STDF
                                            TheExec.Datalog.setup.DatalogSetup.STDFOutputFile = filePathNameSTD
                                            TheExec.Datalog.setup.DatalogSetup.STDFOutput = True
                                        
                                        Else
                                        
                                            TheExec.Datalog.setup.DatalogSetup.STDFOutput = False
                
                                        End If
                                        
                                    Case "$HEADEREVERYTIME"
                                        
                                        If UCase(Trim(tmp(1))) = "ON" Then
                                        
                                            TheExec.Datalog.setup.DatalogSetup.HeaderEveryRun = True
                                        
                                        Else
                                        
                                            TheExec.Datalog.setup.DatalogSetup.HeaderEveryRun = False
                
                                        End If
                                        
                                     Case "$WINDOWOUTPUT"
                                        
                                        If UCase(Trim(tmp(1))) = "ON" Then
                                        
                                            TheExec.Datalog.setup.DatalogSetup.WindowOutput = True
                                        
                                        Else
                                        
                                            TheExec.Datalog.setup.DatalogSetup.WindowOutput = False
                
                                        End If
                                 
                                        End Select
                    
                                    End If
                    
                                Next ctr
    
                            End If
                            
                            End If
                                            
                        End If
                        
                        TheExec.Datalog.ApplySetup
                        
                    End If
                                       
End Function
    
Function turnOffdatalog()
       
        TheExec.Datalog.setup.LotSetup.DatalogOn = False
        TheExec.Datalog.setup.LotSetup.DeviceNumber = "1"
        TheExec.Datalog.setup.DatalogSetup.HeaderEveryRun = False
        TheExec.Datalog.setup.DatalogSetup.SetupFile = ""
        TheExec.Datalog.setup.DatalogSetup.SelectSetupFile = False
        TheExec.Datalog.setup.DatalogSetup.TextOutput = False
        TheExec.Datalog.setup.DatalogSetup.TextOutputFile = ""
        TheExec.Datalog.setup.DatalogSetup.STDFOutput = False
        TheExec.Datalog.setup.DatalogSetup.STDFOutputFile = ""
        TheExec.Datalog.setup.DatalogSetup.WindowOutput = False
        TheExec.RunOptions.DoAll = False
        TheExec.Datalog.ApplySetup
    
End Function
