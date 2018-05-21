VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hram3kDataRdSer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private intBusSize As Integer                   ' bus size, default = 8
Private strPinName As String                    ' Pins that make up the bus
Private lngHramDataOut() As Long                ' Array to store Hram data output
Private intProgrammed As Integer                ' Integer to keep track of programmed status bit 0 = pin names, bit 1 = cycle count
Private intCounter As Integer                   ' Misc. counter used for offline runs
Private intCapSize As Integer                   ' variable to store the size of the capture to be made...
Private intRawCapSize As Integer                ' variable to track the raw data storage size...

' this get property returns an integer showing the currently programmed bus size
Public Property Get BusSize() As Integer
    BusSize = intBusSize
End Property

' this let property sets the bus size to a new value.  You MUST reprogram the pin names after setting this property
Public Property Let BusSize(intNewValue As Integer)
    intBusSize = intNewValue
    intProgrammed = 0                                           'set bits 1 and 3 false
End Property


' this get property returns a 0 if the module is ready to do an Hram Read,
' bit 0 = 1 for pin names not programmed, bit 1 = 1 for cycle count not programmed
Public Property Get Status() As Integer
    Status = intProgrammed
End Property


' this get property returns a long value for the specified cycle and site
Public Property Get HramData(ByVal site As Integer, ByVal cycle As Integer) As Long
    HramData = lngHramDataOut(site, cycle)
End Property


' this get property returns an array long data corresponding to all captured cycles and active sites
Public Property Get hramDataArray() As Variant
    hramDataArray = lngHramDataOut
End Property


' this get property returns the currently programmed pin names for the bus
Public Property Get BusPinNames() As String
    BusPinNames = strPinName
End Property

' this let property sets the pin names for the class object.  You MUST set this property before trying to read the hram
Public Property Let BusPinNames(strNewValue As String)
Dim strSplit() As String

    strSplit = Split(strNewValue, ",")
    If (UBound(strSplit) = intBusSize - 1) Then
        strPinName = strNewValue
        intProgrammed = intProgrammed Or 1                                              'set bit 0 to "true"
    Else
        MsgBox ("wrong number of pins assigned to DataBus during pin name assignment")
    End If
End Property


' class initialization method, initializes the class object at creation...
Private Sub Class_Initialize()
    intBusSize = 8                                                                      'set default bus size
    intProgrammed = 0                                                                   'set bit 0 and 1 to "false"
End Sub

Public Property Let CaptureSize(Cycles2Capture As Integer)
Dim SiteStatus As Long
Dim thisSite As Long

    'check for valid number of cycles to capture
    Select Case Cycles2Capture
        Case Is < 1
            TheExec.DataLog.WriteComment ("ERROR:  Attempt to read less than 1 HRAM cycle")
            SiteStatus = TheExec.Sites.SelectFirst
            While SiteStatus <> loopDone
                thisSite = TheExec.Sites.SelectedSite
                TheExec.Sites.site(thisSite).BinNumber = 0
                TheExec.Sites.site(thisSite).Result = resultFail
                SiteStatus = TheExec.Sites.SelectNext(SiteStatus)
            Wend
            Exit Property
         Case Is > 2800
            TheExec.DataLog.WriteComment ("ERROR:  Attempt to read more than 2800 HRAM cycles")
            SiteStatus = TheExec.Sites.SelectFirst
            While SiteStatus <> loopDone
                thisSite = TheExec.Sites.SelectedSite
                TheExec.Sites.site(thisSite).BinNumber = 0
                TheExec.Sites.site(thisSite).Result = resultFail
                SiteStatus = TheExec.Sites.SelectNext(SiteStatus)
            Wend
            Exit Property
    End Select

    intCapSize = Cycles2Capture
    intRawCapSize = Fix(intCapSize / 11 * 12) + 24
    intProgrammed = intProgrammed Or 2                                                        'set bit 1 to "true"
    ReDim lngHramDataOut(0 To TheExec.Sites.ExistingCount - 1, 0 To intCapSize - 1) As Long   'redim output data array
    
End Property

Public Property Get CaptureSize() As Integer
    CaptureSize = intCapSize
End Property


'setup patgen to capture hram using stv in the pattern
Public Sub PrimeCaptureSTV()

    TheHdw.Digital.Patgen.NoHaltMode = noHaltAlways

    With TheHdw.Digital.HRAM
        .pins(strPinName).OneBitMode = True
        .SetTrigger trigFirst, False, 0, False          'change stop on full to false for 3Kx1 mode...
        .SetCapture captSTV, False
        .Size = intRawCapSize                           'have to use raw cap size because we have to capture 12/11ths samples so we can throw out every 12 sample
    End With

End Sub


' read hram method, this is the guts of the class, run the pattern from the func template or from outside code, this code only processes the Hram
Public Sub ReadHRAM(Optional PrintData As Boolean = False)

Dim bitcounter As Integer
Dim intCounter As Integer

Dim CycleCount As Long
Dim hramCycle As Long
Dim SiteStatus As Long
Dim thisSite As Long

Dim HramValue As Double
    
Dim RawHramData() As Integer
    
    ' You've got to program the pin names before you try to read the hram, if you haven't, you get this error
    ' and the program will bin out bin 0
    If (intProgrammed And 1) = 0 Then
        TheExec.DataLog.WriteComment ("ERROR: Attempt to read HRAM before initializaton of pins")
        BinZeroFail
        Exit Sub
    End If
    If (intProgrammed And 2) = 0 Then
        TheExec.DataLog.WriteComment ("ERROR: Attempt to read HRAM without specifying number of cycles")
        BinZeroFail
        Exit Sub
    End If
      
    'read the data out of HRAM on a site by site basis and stuff into Hram data out array.
    SiteStatus = TheExec.Sites.SelectFirst
    While SiteStatus <> loopDone
        thisSite = TheExec.Sites.SelectedSite
        
        ' Get the data from Hram...
        If TheExec.TesterMode = testModeOnline Then
            Call TheHdw.Digital.HRAM.pins(strPinName).ReadOneBitData(0, CLng(intRawCapSize), RawHramData)
        Else
            ReDim RawHramData(intRawCapSize) As Integer
            For intCounter = 0 To intRawCapSize
                RawHramData(intCounter) = intCounter
            Next intCounter
        End If
        
        CycleCount = 0
        hramCycle = 0
        While CycleCount < intCapSize
            If Not (hramCycle Mod 12 = 0) Then
                lngHramDataOut(thisSite, CycleCount) = RawHramData(hramCycle)
            
                'datalog readback values
                If PrintData Then _
                    Call TheExec.DataLog.WriteParametricResultOptLoHi(thisSite, TheExec.Sites.site(thisSite).testnumber, logTestNoPF, parmTestLim, _
                        "Hram Rd # " & CStr(hramCycle) & " ", 0, CDbl(lngHramDataOut(thisSite, CycleCount)), unitNone, 0, unitNone)
                TheExec.Sites.site(thisSite).IncrementTestNumber                  'increment tnum
                       
                CycleCount = CycleCount + 1
            End If
            hramCycle = hramCycle + 1
        Wend
        
        SiteStatus = TheExec.Sites.SelectNext(SiteStatus)
    Wend

    TheHdw.Digital.HRAM.pins(strPinName).OneBitMode = False
    


End Sub

' Method to clear the data structure
Public Sub ClearData()
Dim thisSite As Long
Dim hramCycle As Long

    For thisSite = 0 To TheExec.Sites.ExistingCount - 1
        For hramCycle = 0 To intCapSize
            lngHramDataOut(thisSite, hramCycle) = 0
        Next hramCycle
    Next thisSite
    
End Sub

Private Sub BinZeroFail()
Dim SiteStatus As Long
Dim thisSite As Long

    SiteStatus = TheExec.Sites.SelectFirst
    While SiteStatus <> loopDone
        thisSite = TheExec.Sites.SelectedSite
        TheExec.Sites.site(thisSite).BinNumber = 0
        TheExec.Sites.site(thisSite).Result = resultFail
        SiteStatus = TheExec.Sites.SelectNext(SiteStatus)
    Wend

End Sub





