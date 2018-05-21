Attribute VB_Name = "LEBSX_GLOBALS"
Option Explicit

' Calibration Variables
Public bgBiasCalVal() As Integer
Public bgVoltageCalVal() As Integer
Public bgCalVoltage() As Double
Public bgCalBiasCurrent() As Double
Public lfIntOscCalVals() As Long
Public HfIntOsc_Values() As Long
Public HfIntOscTempCo_Values() As Long
Public LdoCalVal() As Long                      'added LDO Cal Val global variable RCA

' Array of all site numbers
' This can be used, for example, in block overlay functions where an array of sites to be
' overlayed is required
Public all_sites() As Long
'


Public Function globals_onValidate()

    Dim NOS As Long
    NOS = TheExec.Sites.ExistingCount - 1

    ReDim bgBiasCalVal(0 To NOS)
    ReDim bgVoltageCalVal(0 To NOS)
    ReDim lfIntOscCalVals(0 To NOS)
    ReDim HfIntOsc_Values(0 To NOS)
    ReDim HfIntOscTempCo_Values(0 To NOS)
    ReDim LdoCalVal(0 To NOS)                   'added ldoCalVal to be redimensioned. - MT (09-11-20)

End Function

'------------------------------------------------------------------------------------------
' Function:     DummyFunc
' Purpose:      Does nothing.  This can be used in situations in which two different tests
'               would be able to use the same exact pattern except one of them requires
'               a VB function to be called out from the pattern.  The test in which this is
'               not necessary can make a call to this function instead.
' Params:       None
' Returns:      None
'------------------------------------------------------------------------------------------
Public Function DummyFunc(argc As Long, argv() As String) As Long
    'Call MsgBox("You just called Dummy Func!", vbInformation, "DummyFunc")
End Function

