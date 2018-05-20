Attribute VB_Name = "Module1"


Public Sub DoTheExport()
Dim FName As Variant
Dim Sep As String
Dim wsSheet As Worksheet
Dim nFileNum As Integer
Dim csvPath As String

''
''Sep = InputBox("Enter a single delimiter character (e.g., comma or semi-colon)", _
''"Export To Text File")
'''csvPath = InputBox("Enter the full path to export CSV files to: ")
''
''csvPath = GetFolderName("Choose the folder to export CSV files to:")
''If csvPath = "" Then
''    MsgBox ("You didn't choose an export directory. Nothing will be exported.")
''    Exit Sub
''End If

Dim directory As String
Dim fso As New FileSystemObject
csvPath = ActiveWorkbook.path & "\WorkSheets"
If Not fso.FolderExists(csvPath) Then
    Call fso.CreateFolder(csvPath)
End If
Set fso = Nothing

Sep = ","


For Each wsSheet In Worksheets
wsSheet.Activate
nFileNum = FreeFile
Open csvPath & "\" & _
  wsSheet.Name & ".csv" For Output As #nFileNum
ExportToTextFile CStr(nFileNum), Sep, False
Close nFileNum
Next wsSheet


Call ExportVisualBasicCode


End Sub



Public Sub ExportToTextFile(nFileNum As Integer, _
Sep As String, SelectionOnly As Boolean)

Dim WholeLine As String
Dim RowNdx As Long
Dim ColNdx As Integer
Dim StartRow As Long
Dim EndRow As Long
Dim StartCol As Integer
Dim EndCol As Integer
Dim CellValue As String

'Application.ScreenUpdating = False
On Error GoTo EndMacro:

If SelectionOnly = True Then
With Selection
StartRow = .Cells(1).row
StartCol = .Cells(1).Column
EndRow = .Cells(.Cells.count).row
EndCol = .Cells(.Cells.count).Column
End With
Else
With ActiveSheet.UsedRange
StartRow = .Cells(1).row
StartCol = .Cells(1).Column
EndRow = .Cells(.Cells.count).row
EndCol = .Cells(.Cells.count).Column
End With
End If

For RowNdx = StartRow To EndRow
WholeLine = ""
For ColNdx = StartCol To EndCol
If Cells(RowNdx, ColNdx).Value = "" Then
CellValue = ""
Else
CellValue = Cells(RowNdx, ColNdx).Value
End If
WholeLine = WholeLine & CellValue & Sep
Next ColNdx
'WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))
Print #nFileNum, WholeLine
Next RowNdx

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True

End Sub

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
    
   ' Stop
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If err.Number <> 0 Then
            Stop
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
           ' Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & path
            Debug.Print "Exported " & VBComponent.Name & ":" & "  " & path
        End If

        On Error GoTo 0
    Next
    
    ' Stop
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub



