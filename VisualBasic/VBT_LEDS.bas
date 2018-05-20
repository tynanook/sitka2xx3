Attribute VB_Name = "VBT_LEDS"
Option Explicit


Public Const LED_PULSE = 0.15
'
Public Function reset_all_leds(argc As Long, argv() As String) As Long   'Initialize/Reset U11 - ALL LEDS OFF
  
    'This function resets all Pass/Fail LEDs to the OFF state
       
    Dim nSiteIndex As Long
    
    
    On Error GoTo errHandler
    
    Call TheHdw.Digital.Patgen.Halt
    
'Debug.Print "reset_all_leds..."
            
        TheHdw.pins("LEDs_OFF").InitState = chInitHi
        
            TheHdw.Wait (LED_PULSE)
        
        TheHdw.pins("LEDs_OFF").InitState = chInitLo
        
       
    max_site = 0 'Initialize global integer from Sites_control

    
    Passing_Site_Flag = False
    
    Call first_active_test_sites
    
    
    Exit Function
    
errHandler:

    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into production abort routine
    reset_all_leds = TL_ERROR

End Function
Public Function reset_module_static(argc As Long, argv() As String) As Long


    
    'This function resets the LoRa module statically by activating MCLR_nRESET
    
    Dim nSiteIndex As Long
    
    On Error GoTo errHandler
     
    For nSiteIndex = 0 To TheExec.Sites.ExistingCount - 1
    
        If TheExec.Sites.Site(nSiteIndex).Active = True Then
        
            ' Invoke MCLR_nRESET low, then high to effect a static module reset.
            
        TheHdw.pins("MCLR_nRESET").InitState = chInitHi
        
            TheHdw.Wait (0.1)
        
        TheHdw.pins("MCLR_nRESET").InitState = chInitLo
        
            TheHdw.Wait (0.01)
        
        TheHdw.pins("MCLR_nRESET").InitState = chInitHi
        
            TheHdw.Wait (0.1)
        
        End If
        
    Next nSiteIndex
    
    Exit Function
    
errHandler:

    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into prouduction abort routine
    reset_module_static = TL_ERROR

End Function

Public Function set_pass_fail_leds(argc As Long, argv() As String) As Long


    
    'This function sets Red site LED for Fail, and Green site LED for Pass.
    'At least one test passes if the test program makes it to this function.

    'sites_active() as boolean 'global
    
    Dim nSiteIndex As Long
    'Dim PF_State As Long
    
    
    On Error GoTo errHandler

    
    'Debug.Print "set_pass_fails_leds..."
    
   If TheExec.Sites.Site(0).Active = True Then
    
        If (sites_tested(0) = True) Then 'test for pass
             
            If Not (TheExec.Sites.Site(0).LastTestResult = 2) Then      'Site 0 = Module 2
            
                     TheHdw.pins("GRN2_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN2_ON").InitState = chInitLo
                             
             Else
             
                     TheHdw.pins("RED2_ON").InitState = chInitHi
         
                                    TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED2_ON").InitState = chInitLo
                      
             End If
             
        End If
        
   End If
   
   If TheExec.Sites.Site(1).Active = True Then
   
        If (sites_tested(1) = True) Then
        
             If Not (TheExec.Sites.Site(1).LastTestResult = 2) Then      'Site 1 = Module 3
            
                     TheHdw.pins("GRN3_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN3_ON").InitState = chInitLo
                             
             Else
             
                     TheHdw.pins("RED3_ON").InitState = chInitHi
         
                                    TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED3_ON").InitState = chInitLo
                      
             End If
             
         End If
         
   End If
   
   
   If TheExec.Sites.Site(2).Active = True Then
   
        If (sites_tested(2) = True) Then
        
            If Not (TheExec.Sites.Site(2).LastTestResult = 2) Then   'Site 2 = Module 4
            
                     TheHdw.pins("GRN4_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN4_ON").InitState = chInitLo
                             
             Else
             
                     TheHdw.pins("RED4_ON").InitState = chInitHi
         
                                    TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED4_ON").InitState = chInitLo
                      
             End If
         
         End If
         
   End If
   
   If TheExec.Sites.Site(3).Active = True Then
   
        If (sites_tested(3) = True) Then
        
            If Not (TheExec.Sites.Site(3).LastTestResult = 2) Then   'Site 3 = Module 1
            
                     TheHdw.pins("GRN1_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN1_ON").InitState = chInitLo
                             
             Else
             
                     TheHdw.pins("RED1_ON").InitState = chInitHi
         
                                    TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED1_ON").InitState = chInitLo
                      
             End If
        
        End If
           
   End If
    
        Passing_Site_Flag = True      'notify OnProgramEnded that a site passed
    
    Exit Function
    
errHandler:



    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into prouduction abort routine
    set_pass_fail_leds = TL_ERROR

End Function

Public Function selected_site()


End Function
