Attribute VB_Name = "RunVBT"
' This ALWAYS GENERATED file contains wrappers for VBT tests.
' Do not edit.

Private Sub HandleUntrappedError()
    ' Sanity clause
    If TheExec Is Nothing Then
        MsgBox "IG-XL is not running!  VBT tests cannot execute unless IG-XL is running."
        Exit Sub
    End If
    ' If the last site has failed out, let's ignore the error
    If TheExec.Sites.ActiveCount = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 5.10 aborts gracefully.)
    Do While TheExec.Sites.InSerialLoop
        Call TheExec.Sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    TheExec.ErrorLogMessage "Test " + TheExec.DataManager.instanceName + ": VBT error #" + Trim(str(err.Number)) + " '" + err.Description + "'"
End Sub

Public Function Empty_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    Dim p1 As New InterposeName
    p1.Value = v(0).Value
    Dim p2 As New InterposeName
    p2.Value = v(1).Value
    Dim p3 As New InterposeName
    p3.Value = v(2).Value
    Dim p4 As New InterposeName
    p4.Value = v(3).Value
    Dim p5 As New InterposeName
    p5.Value = v(4).Value
    Dim p6 As New InterposeName
    p6.Value = v(5).Value
    Dim p7 As New PinList
    p7.Value = v(12).Value
    Dim p8 As New PinList
    p8.Value = v(13).Value
    Dim p9 As New PinList
    p9.Value = v(14).Value
    Dim p10 As New PinList
    p10.Value = v(15).Value
    Dim p11 As New PinList
    p11.Value = v(16).Value
    Dim p12 As New PinList
    p12.Value = v(17).Value
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Empty_T__ = VBT_Empty_T.Empty_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function axrfZigbeeBasicTest__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    axrfZigbeeBasicTest__ = VBT_AXRFBasic.axrfZigbeeBasicTest()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function read_cal_factors__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    read_cal_factors__ = VBT_RF.read_cal_factors()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rf_power_meas_t39a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rf_power_meas_t39a__ = VBT_RF.rf_power_meas_t39a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function forced_tx_mode_t39a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' forced_tx_mode_t39a__ = VBT_RF.forced_tx_mode_t39a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rf_power_meas_t48a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rf_power_meas_t48a__ = VBT_RF.rf_power_meas_t48a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function sleep_current_t48a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' sleep_current_t48a__ = VBT_RF.sleep_current_t48a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rf_adv_tx_mode_t48a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rf_adv_tx_mode_t48a__ = VBT_RF.rf_adv_tx_mode_t48a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rf_tx_off_20ms_t48a__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rf_tx_off_20ms_t48a__ = VBT_RF.rf_tx_off_20ms_t48a(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function tx_config_mrf34ta__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' tx_config_mrf34ta__ = VBT_RF.tx_config_mrf34ta(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rn2483_tx868_cw__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rn2483_tx868_cw__ = VBT_RF.rn2483_tx868_cw(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rn2483_fsk_pkt_rcv__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rn2483_fsk_pkt_rcv__ = VBT_RF.rn2483_fsk_pkt_rcv(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rn2483_id__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rn2483_id__ = VBT_RF.rn2483_id(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function rn2483_i_sleep__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' rn2483_i_sleep__ = VBT_RF.rn2483_i_sleep(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function reset_all_leds__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' reset_all_leds__ = VBT_LEDS.reset_all_leds(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function reset_module_static__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' reset_module_static__ = VBT_LEDS.reset_module_static(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function set_pass_fail_leds__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    ' set_pass_fail_leds__ = VBT_LEDS.set_pass_fail_leds(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function selected_site__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    selected_site__ = VBT_LEDS.selected_site()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































