Attribute VB_Name = "Bin_numbers"
Option Explicit
' [=============================================================================]
' [ DEVICE :   all                                                              ]
' [ MASK NO:   all                                                              ]
' [ SCOPE  :   routine for looking up bin numbers from a worksheet              ]
' [=============================================================================]
' [                                                                             ]
' [                   MICROCHIP TECHNOLOGY INC.                                 ]
' [                   2355 WEST CHANDLER BLVD.                                  ]
' [                   CHANDLER AZ 85224-6199                                    ]
' [                   (480) 792-7200                                            ]
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
' [ 1.0    17may01 jpe  - initial creation                                      ]
' [ 1.1    24may01 jpe  - add the Tnum through Sort Pass colums to the macro    ]
' [ 1.2     1jun01 jpe  - add the ability to sort the parameter name column     ]
' [                     - do the search for the nop opcode, just ignore         ]
' [                       search  fails                                         ]
' [ 1.3    12jul01 jpe  - add the GetBinNumbersCleanUp macro to remove the      ]
' [                       bin number on the flow sheet (colums AD-AH)           ]
' [=============================================================================]
'
'
'
'*******************************************************************************
'  HOW TO USE THIS MACRO:
' 1. Import this file into the test program
' 2. Import the MasterBinList worksheet into the program
' 3. Copy the data from columns I-N of a flow sheet into the corresponding columns of the
'    MasterBinList (A-F) (NOTE: remove the example bin names and numbers that are there)
' 4. Sort the MasterBinList so that the rows not associated with bin names are removed.
'    This is not a requirement, but will make the list more readable.
'    Also you may need to export and then reimport the MasterBinList in order to sort
'    the sheet by the sort bin number because of formating issues (i.e. this column may be
'    text rather than a number).
' 5. Select the flow sheet that you want to check the bin number on.  The macro will take the
'    active flow sheet as the one to compare.
' 6. Run the macro "GetBinNumbers"
'    (From the workbook menu: tools=>macro=>macros (ALT-F8) and then select the macro from the
'     list and select Run)
' 7. Look at the flow sheet in col lngBinDestCol (AD-AH) for the results.
'    Any cells that are different will be shaded yellow
' 8. If you want to update the bin numbers, copy the cells from columns AD-AH onto the
'    flow sheet columns I-N
' 9. Run the macro GetBinNumberCleanUp while on the flow sheet to remove the data in
'    columns AD-AH that GetBinNumber macro put there.



Const FirstRow = 5
Dim strBinLisName As String
Dim lngTestNameCol As Long
Dim lngBinSrcCol As Long
Dim lngTnumSrcCol As Long
Dim strTNameSrcRng As String
Dim strParameterSrcRng As String
Dim lngTNameSrcCol As Long

Dim strFlowSheet As String
Dim lngOpcodeCol As Long
Dim lngParameterCol As Long
Dim lngTnameCol As Long
Dim lngTnumCol As Long
Dim lngBinDestCol As Long
Dim lngSortFailBinCol As Long

Dim blnSearchParameterCol As Boolean



Private Sub InitVariables()
   
   Application.ScreenUpdating = False

   blnSearchParameterCol = True  'set to false to disable search the parameter column on the master bin list
   

   ' Set information for the flow sheet
   strFlowSheet = "flow_pkg"
   lngOpcodeCol = 7         'column G
   lngParameterCol = 8      'column H
   lngTnameCol = 9          'column I
   lngTnumCol = 10          'colunm j
   lngBinDestCol = 30       'column AD
   lngSortFailBinCol = 14   'column N

   ' set information for the master bin list sheet
   strBinLisName = "MasterBinList"
   lngBinSrcCol = 8            'column H, this column contains the bin number to be copied to the flow sheet into column AD.
   lngTnumSrcCol = 4           'column D, this column contains the TNUM to be copied to the flow sheet into column AD.
   strTNameSrcRng = "c2:c6500" 'column on strBinLisName, to be searched and matched with the name from the flow page
   strParameterSrcRng = "b2:b6500" 'column on strBinLisName, to be searched and matched with the name from the flow page
   lngTNameSrcCol = 3          'column C, this is the TName column on the master bin list
   
End Sub  'InitVariables()

Public Sub GetBinNumbers()
  Dim StrMsg As String                  ' string for error handler
  Dim j As Long                         ' loop counter
  Dim i As Long                         ' loop counter
  Dim lngRow As Long                    ' row number of the found cell on the test instance page
  Dim lngLastRow As Long                ' max row on the flow sheet
  Dim strTestName As String             ' the test name from the flow sheet to search for on the Test instance sheet
  Dim rngFind As Range                  ' the range variable for the result of the search for the test name on the test instance sheet.
  Dim vntTmp As Variant
  Dim strOpcode As String
  Dim strFirstAddress As String
    
  
  On Error GoTo ErrorHandler

  
  Call InitVariables

   strFlowSheet = ActiveWorkbook.ActiveSheet.name
   
  
  
  'determine the last row on the flow sheet
  Worksheets(strFlowSheet).Activate
  Worksheets(strFlowSheet).Select
  ActiveSheet.Outline.ShowLevels RowLevels:=5, ColumnLevels:=5
  Range("A1").Select
  ActiveCell.SpecialCells(xlLastCell).Select
  lngLastRow = ActiveCell.row
  
  'verify that the selected sheet is a flow sheet
  If ActiveSheet.Cells(1, 2).value <> "Flow Table" Then
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "Active worksheet is not a flow table, cell B1 <> Flow Table " & VBA.vbCrLf _
          & "cell B1 is " & ActiveSheet.Cells(1, 2).value & VBA.vbCrLf
    MsgBox StrMsg
    Exit Sub
  End If
  
  'verify that the selected flow sheet has the correct revision
  If ActiveSheet.Cells(1, 1).value <> "DFF 1.1" Then
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "Selected flow table is of the incorrect revision, cell A1 <> DFF 1.1" & VBA.vbCrLf _
          & "cell A1 is " & ActiveSheet.Cells(1, 1).value & VBA.vbCrLf
    MsgBox StrMsg
    Exit Sub
  End If
  
  'simple test of the MasterBinList sheet
  Worksheets(strBinLisName).Activate
  If ActiveSheet.Cells(2, 3).value <> "TName" Then
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "Header of MasterBinList does not appear to be correct, cell C2 <> TName " & VBA.vbCrLf _
          & "cell C2 is " & ActiveSheet.Cells(2, 3).value & VBA.vbCrLf
    MsgBox StrMsg
    Exit Sub
  End If
  

  
  'find the bin for each row with a test opcode and copy the bin number.
  For j = FirstRow To lngLastRow
    Worksheets(strFlowSheet).Activate
    strOpcode = ActiveSheet.Cells(j, lngOpcodeCol).value
    If (strOpcode = "Test") Or (strOpcode = "nop") Then
       strTestName = ActiveSheet.Cells(j, lngTnameCol).value
       If strTestName = "" Then
          strTestName = ActiveSheet.Cells(j, lngParameterCol).value
       End If
       If (strTestName <> "") Then
          Worksheets(strBinLisName).Activate
          With Worksheets(strBinLisName)
              Set rngFind = .Range(strTNameSrcRng).Find(strTestName)
              If blnSearchParameterCol Then  'if enabled will search the parameter column on the master bin list
                 If rngFind Is Nothing Then
                    Set rngFind = .Range(strParameterSrcRng).Find(strTestName)  'look in the Test name column
                    If Not rngFind Is Nothing Then
                       strFirstAddress = rngFind.Address
                       Do
                         lngRow = rngFind.row
                         vntTmp = ActiveSheet.Cells(lngRow, lngTNameSrcCol).value
                         If ActiveSheet.Cells(lngRow, lngTNameSrcCol).value <> Empty Then  'test to see if there is a name in the tname column and then skip it if there is.
                            Set rngFind = .Range(strParameterSrcRng).FindNext(rngFind)
                            'MsgBox ("found address = " & rngFind.address)
                         Else
                            Exit Do  'found correct cell
                         End If   'If ActiveSheet.Cells(lngRow, lngTnameCol).value = "" Then
                       Loop While Not rngFind Is Nothing And rngFind.Address <> strFirstAddress
                       If Not rngFind Is Nothing Then  'recheck that the tname parameter is blank, the above code misses one condition
                          lngRow = rngFind.row
                          If ActiveSheet.Cells(lngRow, lngTNameSrcCol).value <> Empty Then
                             Set rngFind = Nothing  'remove the result becasuse it is wrong
                          End If   'If ActiveSheet.Cells(lngRow, 3).value <> Empty Then
                       End If      'If Not rngFind Is Nothing Then
                    End If         'If Not rngFind Is Nothing Then
                 End If            'If rngFind Is Nothing Then
              End If               'If blnSearchParameterCol Then
              
              If Not rngFind Is Nothing Then
                  lngRow = rngFind.row
                  Worksheets(strBinLisName).Select
                  'vntTmp = .Cells(lngRow, lngBinSrcCol).value
'                  .Cells(lngRow, lngBinSrcCol).Copy
                  .Range(.Cells(lngRow, lngTnumSrcCol), .Cells(lngRow, lngBinSrcCol)).Copy
                  Worksheets(strFlowSheet).Activate
                  Worksheets(strFlowSheet).Select
                  Worksheets(strFlowSheet).Paste Destination:=Worksheets(strFlowSheet).Cells(j, lngBinDestCol)
              Else     'failed to find the string
                  If (strOpcode = "Test") Then
                     StrMsg = "Failed to find string: " & strTestName & VBA.vbCrLf _
                            & "Row number on flow sheet is: " & VBA.format(j) & VBA.vbCrLf _
                            & "Enter OK to continue, Cancel to abort search"
                     vntTmp = MsgBox(StrMsg, vbOKCancel)
                     If vntTmp = vbCancel Then
                        Exit Sub
                     End If  'If vntTmp
                  End If     'If "Test"
              End If         'If Not rngFind Is Nothing Then
          End With           'Worksheets(strBinLisName)
       End If                'If (strTestName <> "") Then
    End If                   'If (strOpcode = "Test") Or (strOpcode = "nop") Then
  Next j                     'For j = FirstRow To lngLastRow
  

  'Compare the TNUM-Sort FAIL cells with the master bin list and then highlight any difference.
   Worksheets(strFlowSheet).Activate
  For j = FirstRow To lngLastRow
    For i = 0 To 4
       If ActiveSheet.Cells(j, lngTnumCol + i).value <> ActiveSheet.Cells(j, lngBinDestCol + i).value Then
          ActiveSheet.Cells(j, lngBinDestCol + i).Interior.ColorIndex = 6
          ActiveSheet.Cells(j, lngBinDestCol + i).Interior.Pattern = xlSolid
       End If       'If ActiveSheet.Cells(j, lngTnumCol + i).value <>
    Next i          'For i = 0 To 4
  Next j            'For j = FirstRow To lngLastRow
    
    
    
  'reset the selected cell to top of flow sheet and hide the cells
  Worksheets(strFlowSheet).Activate
  Worksheets(strFlowSheet).Select
  ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
  Range("A5").Select
  Application.ScreenUpdating = True
    
    
  Exit Sub
    
ErrorHandler:
    
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf _
          & VBA.err.Description & VBA.vbCrLf
    MsgBox StrMsg
    On Error GoTo 0
   Application.ScreenUpdating = True
    

End Sub  'GetBinNumbers


Public Sub GetBinNumbersCleanUp()
  Dim StrMsg As String                  ' string for error handler
  Dim j As Long                         ' loop counter
  Dim i As Long                         ' loop counter
  Dim lngRow As Long                    ' row number of the found cell on the test instance page
  Dim lngLastRow As Long                ' max row on the flow sheet
  Dim strTestName As String             ' the test name from the flow sheet to search for on the Test instance sheet
  Dim rngFind As Range                  ' the range variable for the result of the search for the test name on the test instance sheet.
  Dim vntTmp As Variant
  Dim strOpcode As String
  Dim strFirstAddress As String
    
  
  On Error GoTo ErrorHandler

  
  Call InitVariables

   strFlowSheet = ActiveWorkbook.ActiveSheet.name
   
  
  
  'determine the last row on the flow sheet
  Worksheets(strFlowSheet).Activate
  Worksheets(strFlowSheet).Select
  ActiveSheet.Outline.ShowLevels RowLevels:=5, ColumnLevels:=5
  Range("A1").Select
  ActiveCell.SpecialCells(xlLastCell).Select
  lngLastRow = ActiveCell.row
  
  'verify that the selected sheet is a flow sheet
  If ActiveSheet.Cells(1, 2).value <> "Flow Table" Then
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "Active worksheet is not a flow table, cell B1 <> Flow Table " & VBA.vbCrLf _
          & "cell B1 is " & ActiveSheet.Cells(1, 2).value & VBA.vbCrLf
    MsgBox StrMsg
    Exit Sub
  End If
  
  'verify that the selected flow sheet has the correct revision
  If ActiveSheet.Cells(1, 1).value <> "DFF 1.1" Then
     StrMsg = "Error in maco GetBinNumbers" & VBA.vbCrLf _
          & "Selected flow table is of the incorrect revision, cell A1 <> DFF 1.1" & VBA.vbCrLf _
          & "cell A1 is " & ActiveSheet.Cells(1, 1).value & VBA.vbCrLf
    MsgBox StrMsg
    Exit Sub
  End If
  
 

  
  'Remove the bin numbers from the flow sheet that the GetBinNumbers macro added.
   Worksheets(strFlowSheet).Activate
  For j = FirstRow To lngLastRow
    For i = 0 To 4
          ActiveSheet.Cells(j, lngBinDestCol + i).ClearContents
          ActiveSheet.Cells(j, lngBinDestCol + i).Interior.ColorIndex = xlNone
    Next i          'For i = 0 To 4
  Next j            'For j = FirstRow To lngLastRow
    
    
    
  'reset the selected cell to top of flow sheet and hide the cells
  Worksheets(strFlowSheet).Activate
  Worksheets(strFlowSheet).Select
  ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
  Range("A5").Select
    
    
  Application.ScreenUpdating = True
  Exit Sub
    
ErrorHandler:
    
     StrMsg = "Error in maco GetBinNumbersCleanUp" & VBA.vbCrLf _
          & "VBA Error number is " & format(VBA.err.Number) & VBA.vbCrLf _
          & VBA.err.Description & VBA.vbCrLf
    MsgBox StrMsg
    On Error GoTo 0
    Application.ScreenUpdating = True
    

End Sub  'GetBinNumbersCleanUp




