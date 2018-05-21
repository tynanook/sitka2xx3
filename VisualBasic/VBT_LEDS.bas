Attribute VB_Name = "VBT_LEDS"
Option Explicit

Public Passing_Site0_Flag As Boolean
Public Passing_Site1_Flag As Boolean
Public Passing_Site2_Flag As Boolean
Public Passing_Site3_Flag As Boolean


Public Const LED_PULSE = 0.15

Public Function reset_all_leds(argc As Long, argv() As String) As Long   'Initialize/Reset U11 - ALL LEDS OFF

    'This function resets all Pass/Fail LEDs to the OFF state
       
    Dim nSiteIndex As Long
    
    
    On Error GoTo errHandler
    
    Call TheHdw.Digital.Patgen.Halt
    
'Debug.Print "reset_all_leds..."

If FIRSTRUN = True Then

    Call init_leds
    
    FIRSTRUN = False
    
Else

End If
            
        TheHdw.pins("LEDs_OFF").InitState = chInitHi
        
            TheHdw.Wait (LED_PULSE)
        
        TheHdw.pins("LEDs_OFF").InitState = chInitLo
        
       
    max_site = 0 'Initialize global integer from Sites_control
    
    Passing_Site0_Flag = False  'Individual passing site flags fixed a multi-site issue.
    Passing_Site1_Flag = False
    Passing_Site2_Flag = False
    Passing_Site3_Flag = False

    

    
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
    
    'This function sets Red site LED for test flow Fail, and Green site LED for test flow Pass.
    'At least one test passes if the test program makes it to this function.
    'This function is the only function that will set a passing GREEN LED for a site.

    'sites_active() as boolean 'global
    
    Dim nSiteIndex As Long
    'Dim PF_State As Long
    
    
    On Error GoTo errHandler

    
    'Debug.Print "set_pass_fails_leds..."
    
   If TheExec.Sites.Site(0).Active = True Then
    
        If (sites_tested(0) = True) Then 'test for pass
             
            If TheExec.Sites.Site(0).LastTestResult = 2 Then      'Site 0 = Module 1
            
                     TheHdw.pins("RED1_ON").InitState = chInitHi
         
                        TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED1_ON").InitState = chInitLo
                     
                     Passing_Site0_Flag = False
                     
                                       
             Else
                     TheHdw.pins("GRN1_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN1_ON").InitState = chInitLo
                     
                     Passing_Site0_Flag = True      'notify OnProgramEnded that site 0 passed
                      
             End If
             
        End If
        
   End If
   
   If TheExec.Sites.Site(1).Active = True Then
   
        If (sites_tested(1) = True) Then
        
             If TheExec.Sites.Site(1).LastTestResult = 2 Then      'Site 1 = Module 2
             
                     TheHdw.pins("RED2_ON").InitState = chInitHi
         
                        TheHdw.Wait (LED_PULSE)
                                    
                     TheHdw.pins("RED2_ON").InitState = chInitLo
                     
                     Passing_Site1_Flag = False
                             
             Else
                               
                     TheHdw.pins("GRN2_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN2_ON").InitState = chInitLo
                     
                     Passing_Site1_Flag = True      'notify OnProgramEnded that site 1 passed
                      
             End If
             
         End If
         
   End If
   
   
   If TheExec.Sites.Site(2).Active = True Then
   
        If (sites_tested(2) = True) Then
        
            If TheExec.Sites.Site(2).LastTestResult = 2 Then   'Site 2 = Module 3
            
            
                    TheHdw.pins("RED3_ON").InitState = chInitHi
         
                        TheHdw.Wait (LED_PULSE)
                                    
                    TheHdw.pins("RED3_ON").InitState = chInitLo
                    
                    Passing_Site2_Flag = False
                             
             Else
             
                    TheHdw.pins("GRN3_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                    TheHdw.pins("GRN3_ON").InitState = chInitLo
                     
                    Passing_Site2_Flag = True       'notify OnProgramEnded that site 2 passed
             
             End If
         
         End If
         
   End If
   
   If TheExec.Sites.Site(3).Active = True Then
   
        If (sites_tested(3) = True) Then
        
            If TheExec.Sites.Site(3).LastTestResult = 2 Then   'Site 3 = Module 4
            
                    TheHdw.pins("RED4_ON").InitState = chInitHi
         
                        TheHdw.Wait (LED_PULSE)
                                    
                    TheHdw.pins("RED4_ON").InitState = chInitLo
            
                     
                    Passing_Site3_Flag = False
                             
             Else
             
                     TheHdw.pins("GRN4_ON").InitState = chInitHi
         
                         TheHdw.Wait (LED_PULSE)
                      
                     TheHdw.pins("GRN4_ON").InitState = chInitLo
                     
                     Passing_Site3_Flag = True      'notify OnProgramEnded that site 3 passed
                      
             End If
        
        End If
           
   End If
    

    
    Exit Function
    
errHandler:



    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into prouduction abort routine
    set_pass_fail_leds = TL_ERROR

End Function

Public Function selected_site()


End Function

Public Function init_leds() As Long  'Initialize/Reset U11 Blink LEDs
  
    'This function blinks all Pass/Fail LEDs, then resets them to the OFF state
    
Dim led_level_high As Double
Dim led_level_low As Double

On Error GoTo errHandler

 'Flash LEDs once up and down at program validation

    led_level_high = 4.99
    led_level_low = 0.01
    
    
        'Deactivate LED signals
        TheHdw.PPMU.pins("GRN1_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("GRN2_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        TheHdw.PPMU.pins("GRN3_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("GRN4_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_low
    
        TheHdw.PPMU.pins("LEDS_OFF").ForceVoltage(ppmu2mA) = led_level_low
    
    
        'Connect all LED signals
        TheHdw.PPMU.pins("GRN1_ON").Connect
        TheHdw.PPMU.pins("RED1_ON").Connect
        TheHdw.PPMU.pins("GRN2_ON").Connect
        TheHdw.PPMU.pins("RED2_ON").Connect
        
        TheHdw.PPMU.pins("GRN3_ON").Connect
        TheHdw.PPMU.pins("RED3_ON").Connect
        TheHdw.PPMU.pins("GRN4_ON").Connect
        TheHdw.PPMU.pins("RED4_ON").Connect
        
        TheHdw.PPMU.pins("LEDS_OFF").Connect
        
        'Turn Site 0 Green ON
        TheHdw.PPMU.pins("GRN1_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN1_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 0 Red ON
        TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("RED1_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 1 Green ON
        TheHdw.PPMU.pins("GRN2_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN2_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 1 Red ON
        TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("RED2_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 2 Green ON
        TheHdw.PPMU.pins("GRN3_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN3_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 2 Red ON
        TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("RED3_ON").ForceVoltage(ppmu2mA) = led_level_low
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 3 Green ON
        TheHdw.PPMU.pins("GRN4_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN4_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        TheHdw.Wait (LED_PULSE)
        TheHdw.Wait (LED_PULSE)
        
        
        'Turn Site 3 Red ON
        TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("RED4_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        TheHdw.Wait (LED_PULSE)
        TheHdw.Wait (LED_PULSE)
        
        'Turn Site 3 Green ON
        TheHdw.PPMU.pins("GRN4_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN4_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        'Turn Site 2 Green ON
        TheHdw.PPMU.pins("GRN3_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN3_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        'Turn Site 1 Green ON
        TheHdw.PPMU.pins("GRN2_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN2_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        'Turn Site 0 Green ON
        TheHdw.PPMU.pins("GRN1_ON").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("GRN1_ON").ForceVoltage(ppmu2mA) = led_level_low
        
        
        TheHdw.Wait (LED_PULSE)
        TheHdw.Wait (LED_PULSE)
        
        
        'Turn OFF all LEDs
        TheHdw.PPMU.pins("LEDS_OFF").ForceVoltage(ppmu2mA) = led_level_high
        TheHdw.Wait (LED_PULSE)
        TheHdw.PPMU.pins("LEDS_OFF").ForceVoltage(ppmu2mA) = led_level_low
        
        'Disconnect all LED signals
        TheHdw.PPMU.pins("LEDS_OFF").Disconnect
        TheHdw.PPMU.pins("GRN1_ON").Disconnect
        TheHdw.PPMU.pins("RED1_ON").Disconnect
        TheHdw.PPMU.pins("GRN2_ON").Disconnect
        TheHdw.PPMU.pins("RED2_ON").Disconnect
        TheHdw.PPMU.pins("GRN3_ON").Disconnect
        TheHdw.PPMU.pins("RED3_ON").Disconnect
        TheHdw.PPMU.pins("GRN4_ON").Disconnect
        TheHdw.PPMU.pins("RED4_ON").Disconnect
        

        
        
    
    Exit Function
    
errHandler:

    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    'Call TheHdw.Digital.Patgen.Halt
    
    If AbortTest Then Exit Function Else Resume Next      'Hook into production abort routine
    init_leds = TL_ERROR

End Function
