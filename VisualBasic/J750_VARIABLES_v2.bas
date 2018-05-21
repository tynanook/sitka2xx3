Attribute VB_Name = "J750_VARIABLES_v2"
Option Explicit
'---------------------------------------------------------------------------------------------------------
' Module:       J750_VARIABLES
' Purpose:      Routine to dereference specification values using native IG-XL commands.  To optimize test
'               time, the value dereferenced will be stored in an array so that the native IG-XL command is
'               only necessary on the first resolution.
'
' Comments:     The first version of this module dereferenced every variable using evaluate() command. This
'               removed the native IG-XL commands from the equation.  Version 2 was developed to use native
'               IG-XL commands (i.e. VariableValue) to eliminate any test concerns.
'
'               Following are test times that examplify the reason for this modules creation, which is to
'               minimize specification resolution at run-time.
'
'               Native IG-XL VariableValue() Function Test Times
'               - ~200ms when switching test context and resolving variable
'               - ~50ms when resolving variable, no context switching
'
'               Native IG-XL GetInstanceContext() Function Test Time
'               - ~0.05ms to resolve test instance context
'
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------Revision History---------------------------------------------
'---------------------------------------------------------------------------------------------------------
' Rev   Author      Date        Description
' <0>   DePaul      08/05/2008  - Initial release
'---------------------------------------------------------------------------------------------------------



Private spec_info() As String           ' Specification Information: Name, DC Cat, DC Sel, AC Cat, AC Sel
Private spec_data() As Variant          ' Specification Data
Private spec_array_init As Boolean      ' Specification Array Initialized

' Subr:     spec_array_onValidate
' Purpose:  Redimension specification array and set the zero element
' Params:   None
' Returns:  None
Public Sub spec_array_onValidate()
        spec_array_init = False
End Sub

' Func:     resolve_spec
' Purpose:  Use native IG-XL command to resolve function first time and add value to array.
'           Subsequent resolves will use stored array value.  Code setup to handle characterization
' Params:   spec_name       String      specification to be resolved
' Returns:  Double          specification value
Public Function resolve_spec(ByVal spec_name As String, Optional ByVal test_name As String = "") As Variant
    Dim spec_cnt As Long                    ' Specification Array Index Counter Variable
    Dim spec_info_str As String             ' Specification Information
    
    ' Test Instance Context Variables
    Dim ret_dc_cat As String                ' DC Category
    Dim ret_dc_sel As String                ' DC Selector
    Dim ret_ac_cat As String                ' AC Category
    Dim ret_ac_sel As String                ' AC Selector
    Dim ret_time_set As String              ' Time Set Sheet Name
    Dim ret_edge_set As String              ' Edge Set Sheet Name
    Dim ret_pin_lvls As String              ' Pin Levels Sheet Name
    Dim ret_overlay As String               ' Overlay value
    
    ' If specification array is not initialized then reinitialize the array
    If Not spec_array_init Then
        ReDim spec_info(0 To 0) As String
        ReDim spec_data(0 To 0) As Variant
        spec_info(0) = "spec_name,dc_cat,dc_sel,ac_cat,ac_sel,value"
        spec_data(0) = Empty
        spec_array_init = True
    End If
    
    ' If the test name is provided, then resolve the specification using native IG-XL command
    If Trim(test_name) <> "" Then
        resolve_spec = TheExec.VariableValue(spec_name, test_name)
        Exit Function
    End If
    
    ' If the flow is not running, then resolve the specification using native IG-XL command
    If Not TheExec.Flow.IsRunning Then
        resolve_spec = TheExec.VariableValue(spec_name)
        Exit Function
    End If
    
    ' If the test is characterizing (i.e. shmoo), then use native IG-XL command to return specification value
    If TheExec.Flow.IsCharacterizing Then
        resolve_spec = TheExec.VariableValue(spec_name)
        Exit Function
    End If

    ' Resolve test context information for specification array search
    Call TheExec.DataManager.GetInstanceContext(ret_dc_cat, ret_dc_sel, ret_ac_cat, ret_ac_sel, ret_time_set, ret_edge_set, ret_pin_lvls, ret_overlay)
    spec_info_str = LCase(spec_name) & "," & LCase(ret_dc_cat) & "," & LCase(ret_dc_sel) & "," & LCase(ret_ac_cat) & "," & LCase(ret_ac_sel)
    
    ' Search specification array for variable information
    For spec_cnt = LBound(spec_info) To UBound(spec_info) Step 1
        If spec_info(spec_cnt) = spec_info_str Then Exit For
    Next spec_cnt
    
    ' If not found resolve variable using native IG-XL command with current test context
    ' Otherwise, split the information and return the result
    If spec_cnt > UBound(spec_info) Then
        resolve_spec = TheExec.VariableValue(spec_name)
        
        ' Store specification info and value into specification array
        ReDim Preserve spec_info(LBound(spec_info) To (UBound(spec_info) + 1)) As String
        ReDim Preserve spec_data(LBound(spec_info) To UBound(spec_info)) As Variant
        spec_info(UBound(spec_info)) = spec_info_str
        spec_data(UBound(spec_data)) = resolve_spec
    Else
        ' Return the specification data
        resolve_spec = spec_data(spec_cnt)
    End If
End Function
