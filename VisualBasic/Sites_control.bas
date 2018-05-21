Attribute VB_Name = "Sites_control"
Option Explicit
Public sites_failed() As Boolean
Public max_site As Long                'number of actual sites being tested (initialized to zero elsewhere)
Public remaining_sites As Long

'Public sites_tested() As Boolean            'array showing which sites tested in the flow (True = tested)
'Public sites_active() As Boolean    'array showing which sites are active(passed) during last test in flow
'Public sites_inactive() As Boolean
Public site0_failed As Boolean
Public site1_failed As Boolean
Public site2_failed As Boolean
Public site3_failed As Boolean

Public first_tested_flag As Boolean
Public Passing_Site_Flag As Boolean


Dim site_bin() As Integer
Dim site_sort() As Integer
Dim site_bin_sort_control() As Integer



Dim num_site_per_testersite As Integer  'number of sub-sites in one tester site
Dim total_fail As Integer

Dim siteStatus() As Boolean  'added code for IDDWR
 
Public Function init_sites_array() As Long

Dim site_num As Long
Dim i As Long

max_site = 32    'Total actual number of sites being tested. This number needs to be changed depending on the number of sites
num_site_per_testersite = 2
total_fail = 0

ReDim sites_failed(max_site - 1)    'array containing failed sites. value 1 means the site failed. 0 means the site didn't fail
ReDim site_bin(max_site - 1)        'array containing the bin number of each sites
ReDim site_sort(max_site - 1)       'array containing sort number of each sites
ReDim site_bin_sort_control(max_site - 1)

TheExec.Sites.SetAllActive (True)

For site_num = 0 To max_site - 1    'reset the arrays
     sites_failed(site_num) = False
     site_bin(site_num) = 2         'sdriscoll_102600: temp. changed from 1 to 2 since I can't find where passing bin gets assigned
     site_sort(site_num) = 2        'sdriscoll_102600: temp. changed from 1 to 2 since I can't find where passing sort gets assigned
     site_bin_sort_control(site_num) = 0
Next site_num
TheExec.RunOptions.DoAll = True     'set doall flag to begin
End Function

Public Function stuff_bin_sort(argc As Long, argv() As String) As Long
'Because sort and bin number can only be set after the end of body
'so this function is put at every StartBody of interpose function to check
'the status and stuff bin and sort number
  
'pass in: IDDW_test switch(argv(0)), 'site-group' number (argv(1))
  
  Dim site_num As Long
  Dim i As Long

For site_num = 0 To max_site - 1
    If ((site_bin_sort_control(site_num) = 1) And (sites_failed(site_num) = True)) Then
        site_bin(site_num) = TheExec.Sites.Site(site_num \ num_site_per_testersite).BinNumber
        site_sort(site_num) = TheExec.Sites.Site(site_num \ num_site_per_testersite).SortNumber
        site_bin_sort_control(site_num) = 2
    End If
Next site_num

For i = 0 To 15
    If TheExec.Sites.Site(i).Active = False Then
        site_bin(i * 2) = -1
        site_sort(i * 2) = -1
        site_bin(i * 2 + 1) = -1
        site_sort(i * 2 + 1) = -1
    End If
Next

'If (total_fail = max_site) Then
'    Call print_bin_sort
'    Call ComResultsToOI
'    Call SendConfigSTDF
'   theexec.RunOptions.DoAll = False
'End If

'*** added in for IDDW testing 9/24/14
If argc = 1 Then
  Call ActivateSelectedSites(CLng(argv(0)))
End If



End Function

Public Function stuff_bin_sort_func() As Long
Dim site_num As Long
Dim i As Long

For site_num = 0 To max_site - 1
    If ((site_bin_sort_control(site_num) = 1) And (sites_failed(site_num) = True)) Then
        site_bin(site_num) = TheExec.Sites.Site(site_num \ num_site_per_testersite).BinNumber
        site_sort(site_num) = TheExec.Sites.Site(site_num \ num_site_per_testersite).SortNumber
        site_bin_sort_control(site_num) = 2
    End If
Next site_num

For i = 0 To 15
    If TheExec.Sites.Site(i).Active = False Then
        site_bin(i * 2) = -1
        site_sort(i * 2) = -1
        site_bin(i * 2 + 1) = -1
        site_sort(i * 2 + 1) = -1
    End If
Next

'If (total_fail = max_site) Then
'    Call print_bin_sort
'    Call ComResultsToOI
'    Call SendConfigSTDF
'    theexec.RunOptions.DoAll = False
'End If

End Function

'This function gets executed at the end of the flow to print a summary of the bin and sort #

Public Function print_bin_sort() As Long
Dim site_num As Long
For site_num = 0 To max_site - 1
    If ((site_bin_sort_control(site_num) = 1) And (sites_failed(site_num) = True)) Then
        site_bin(site_num) = TheExec.Sites.Site(site_num \ 4).BinNumber
        site_sort(site_num) = TheExec.Sites.Site(site_num \ 4).SortNumber
        site_bin_sort_control(site_num) = 2
    End If
Next site_num

For site_num = 0 To max_site - 1
    TheExec.DataLog.WriteComment ("   " + CStr(site_num) + " Bin Number is: " + CStr(site_bin(site_num)) _
                                + "   " + "Sort Number is: " + CStr(site_sort(site_num)))
Next site_num

End Function

'This function finds the first failure for each test and prints the text "Sub Sites Failure" before the first failure,
'in each test

Public Function find_first_fail(first_fail_flag)

If (first_fail_flag = 0) Then
    TheExec.DataLog.WriteComment ("   Sub Sites Failure")
    first_fail_flag = 1
End If

End Function


'Same function as the previous one, but this can be called from other interpose functions.

Public Function Find_failed_site_call() As Long

Dim site_num As Long
Dim first_fail_flag As Integer
Dim Curr_chanMap As String
Dim loopstatus As Long

first_fail_flag = 0   'to keep track of the first failure

Curr_chanMap = TheExec.CurrentChanMap       'Check if this is the Engineering DIB

  loopstatus = TheExec.Sites.SelectFirst
  ' Loop while there are more sites
  While loopstatus <> loopDone
      ' Do something site specific here
      ' Cycle to the next site
      site_num = TheExec.Sites.SelectedSite


    If (TheHdw.pins("SO_0").FailCount(site_num) > 0) Then
        If sites_failed(site_num * num_site_per_testersite) Then
            Call find_first_fail(first_fail_flag)
            TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                        + "   " + CStr(site_num * num_site_per_testersite) + "    FAIL")
        ElseIf Not sites_failed(site_num * num_site_per_testersite) Then
            sites_failed(site_num * num_site_per_testersite) = True
            site_bin_sort_control(site_num * num_site_per_testersite) = 1
            total_fail = total_fail + 1
            Call find_first_fail(first_fail_flag)
            TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                         + "   " + CStr(site_num * num_site_per_testersite) + "    FAIL")
        End If
    End If

   If Curr_chanMap = "x8man8dip" Or Curr_chanMap = "x8man8dip_FA" Then GoTo SiteCallMap
     If (TheHdw.pins("SO_1").FailCount(site_num) > 0) Then
        If sites_failed(site_num * num_site_per_testersite + 1) Then
            Call find_first_fail(first_fail_flag)
            TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                        + "   " + CStr(site_num * num_site_per_testersite + 1) + "    FAIL")
        ElseIf Not sites_failed(site_num * num_site_per_testersite + 1) Then
            sites_failed(site_num * num_site_per_testersite + 1) = True
            site_bin_sort_control(site_num * num_site_per_testersite + 1) = 1
            total_fail = total_fail + 1
            Call find_first_fail(first_fail_flag)
            TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                         + "   " + CStr(site_num * num_site_per_testersite + 1) + "    FAIL")
        End If
    End If

 '   If (thehdw.pins("CS_2, SCK_2, SI_2, SO_2, HOLD_2, WP_2").FailCount(site_num) > 0) Then
 '       If sites_failed(site_num * num_site_per_testersite + 2) Then
 '           Call find_first_fail(first_fail_flag)
 '           theexec.Datalog.WriteComment ("   " + CStr(theexec.Sites.site(site_num).TestNumber) _
 '                                       + "   " + CStr(site_num * num_site_per_testersite + 2) + "    FAIL")
 '       ElseIf Not sites_failed(site_num * num_site_per_testersite + 2) Then
 '           sites_failed(site_num * num_site_per_testersite + 2) = True
 '           site_bin_sort_control(site_num * num_site_per_testersite + 2) = 1
 '           total_fail = total_fail + 1
 '           Call find_first_fail(first_fail_flag)
 '           theexec.Datalog.WriteComment ("   " + CStr(theexec.Sites.site(site_num).TestNumber) _
 '                                        + "   " + CStr(site_num * num_site_per_testersite + 2) + "    FAIL")
 '       End If
 '   End If
 '
 '   If (thehdw.pins("CS_3, SCK_3, SI_3, SO_3, HOLD_3, WP_3").FailCount(site_num) > 0) Then
 '       If sites_failed(site_num * num_site_per_testersite + 3) Then
 '           Call find_first_fail(first_fail_flag)
 '           theexec.Datalog.WriteComment ("   " + CStr(theexec.Sites.site(site_num).TestNumber) _
 '                                       + "   " + CStr(site_num * num_site_per_testersite + 3) + "    FAIL")
 '       ElseIf Not sites_failed(site_num * num_site_per_testersite + 3) Then
 '           sites_failed(site_num * num_site_per_testersite + 3) = True
 '           site_bin_sort_control(site_num * num_site_per_testersite + 3) = 1
 '           total_fail = total_fail + 1
 '           Call find_first_fail(first_fail_flag)
 '           theexec.Datalog.WriteComment ("   " + CStr(theexec.Sites.site(site_num).TestNumber) _
 '                                        + "   " + CStr(site_num * num_site_per_testersite + 3) + "    FAIL")
 '       End If
 '   End If

SiteCallMap:
    loopstatus = TheExec.Sites.SelectNext(loopstatus)
Wend

TheExec.DataLog.WriteComment ("   Tester Sites")

End Function

Public Function Find_failed_site_param_test(argc As Long, argv() As String) As Long
'pass in: site_offset (argv(0)), IDDW_test switch(argv(1)), 'site-group' number (argv(2))
  
  Dim site_num As Long
  Dim site_offset As Long
  Dim all_sites As Long
  Dim first_fail_flag As Integer

  first_fail_flag = 0

  site_offset = CLng(argv(0))   'site offset number. This number is set in the parametric
                                'test template. - changed to argv(2) to accomodate IDDW test 9-25-14.



 'the prober will set the activecount of the sites.
 If TheExec.Sites.ActiveCount < max_site \ num_site_per_testersite Then
   all_sites = TheExec.Sites.ActiveCount * num_site_per_testersite
   Else: all_sites = max_site
 End If
 
 all_sites = 32

 For site_num = 0 To all_sites \ num_site_per_testersite - 1
   'If this site already failed other tests before
   If sites_failed(site_num * num_site_per_testersite + site_offset) Then
     If TheExec.Sites.Site(site_num).Active = True Then
       If (TheExec.Sites.Site(site_num).LastTestResultRaw = resultFail) Then
         Call find_first_fail(first_fail_flag)
         TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                     + "   " + CStr(site_num * num_site_per_testersite + site_offset) + "    FAIL")
       End If
     End If
   End If

   'If this site hasn't failed yet. This is the first time it fails
   If Not sites_failed(site_num * num_site_per_testersite + site_offset) Then
     If TheExec.Sites.Site(site_num).Active = True Then
       If (TheExec.Sites.Site(site_num).LastTestResultRaw = resultFail) Then
         sites_failed(site_num * num_site_per_testersite + site_offset) = True
         site_bin_sort_control(site_num * num_site_per_testersite + site_offset) = 1
         Call find_first_fail(first_fail_flag)
         TheExec.DataLog.WriteComment ("   " + CStr(TheExec.Sites.Site(site_num).testnumber) _
                                     + "   " + CStr(site_num * num_site_per_testersite + site_offset) + "    FAIL")
         total_fail = total_fail + 1
       End If
     End If
   End If
 Next site_num
 
 '*** added in for IDDW testing 9/24/14
If argc = 2 Then
'  Call reActivateSelectedSites(CLng(argv(1)))          ' No Defined --> Remove PTT 05/07/15
End If

 TheExec.DataLog.WriteComment ("   Tester Sites")

End Function


'Prints out the values inside each of the arrays. For debug only
Public Function print_array() As Long

    Dim site_num As Long

    TheExec.DataLog.WriteComment ("sites_failed")
    For site_num = 0 To max_site - 1
        TheExec.DataLog.WriteComment (CStr(sites_failed(site_num)))
    Next site_num

    TheExec.DataLog.WriteComment ("site_bin")
    For site_num = 0 To max_site - 1
        TheExec.DataLog.WriteComment (CStr(site_bin(site_num)))
    Next site_num

    TheExec.DataLog.WriteComment ("site_sort")
    For site_num = 0 To max_site - 1
        TheExec.DataLog.WriteComment (CStr(site_sort(site_num)))
    Next site_num

    TheExec.DataLog.WriteComment ("site_bin_sort_control")
    For site_num = 0 To max_site - 1
        TheExec.DataLog.WriteComment (CStr(site_bin_sort_control(site_num)))
    Next site_num

End Function

Public Function SendConfigSTDF() As Long

    Dim site_num As Long
    TheExec.DataLog.WriteComment ("<ProberSite>")
    TheExec.DataLog.WriteComment (CStr(GetSetting("MCHPOI", "J750", "PROBERSITE")))
    TheExec.DataLog.WriteComment ("<site_bin_data>")

    For site_num = 0 To max_site - 1
        If sites_failed(site_num) = True Then
            TheExec.DataLog.WriteComment (CStr(Int(site_num / num_site_per_testersite)) + "," _
            + CStr(site_num) + "," + CStr(site_sort(site_num)) + "," + "F")
        Else
            TheExec.DataLog.WriteComment (CStr(Int(site_num / num_site_per_testersite)) + "," _
            + CStr(site_num) + "," + CStr(site_sort(site_num)) + "," + "P") 'sdriscoll_102600: modded pass string to match fail style.
        End If
    Next site_num

End Function




'-----------------------------------------------------------------------
' OI Data Exchange Module
'-----------------------------------------------------------------------
'Communicates 32 sites results and bin # to OI

Public Function ComResultsToOI() As Long

    Dim i As Integer
    Dim TestInProgress As Integer
    ReDim binarray(max_site - 1)
    ReDim sortarray(max_site - 1)
    ReDim failarray(max_site - 1)

    'TestInProgress must be sent to the registry to inform OI on the test program status. TestInProgress must
    'be sent 1 to start bin display. When TestInProgress is sent as -1 OI will stop polling for bin results

    TestInProgress = 1
    Call SaveSetting("MCHPOI", "J750", "STATUS", TestInProgress)

    'Put bin data into registry
    For i = 0 To max_site - 1
        Call SaveSetting("MCHPOI", "J750", "BINSITE" + CStr(i), site_bin(i))

        'Put sort data into registry
        Call SaveSetting("MCHPOI", "J750", "SORTSITE" + CStr(i), site_sort(i))

        'Need to assign to pass/fail result this way because the prober driver expects these macros
        'FFAIL=2, FPASS=1
        If sites_failed(i) = True Then
            Call SaveSetting("MCHPOI", "J750", "P/F_SITE" + CStr(i), 2)
        Else
            Call SaveSetting("MCHPOI", "J750", "P/F_SITE" + CStr(i), 1)
        End If
    Next i

End Function


Public Function Enable_StoreInactiveSites(argc As Long, argv() As String) As Long

'Interpose function to store inactive site numbers and enable any site that has been disabled because of
'test failures.  Should be called as a StartOfBody function in a Test Instance.

  Dim Site As Long
  
  ReDim siteStatus(TheExec.Sites.ExistingCount - 1) As Boolean
  ReDim sites_inactive(TheExec.Sites.ExistingCount - 1) As Boolean
  
    'Debug.Print "Enable_StoreInactiveSites..."

  ' Loop through all sites and store active status
  
  For Site = 0 To TheExec.Sites.ExistingCount - 1
  
    siteStatus(Site) = TheExec.Sites.Site(Site).Active
    
    
        'Debug.Print "Site = "; site
        'Debug.Print "Status = "; siteStatus(site)

    
  Next Site
  


  'make all sites active
  Call TheExec.Sites.SetAllActive(True)
  
  'powerdown, for some reason the inactive site will not power up unless you power down
  TheHdw.PinLevels.PowerDown
  
'possible wait here
  
  'apply power to all site since they should all be active now.
  
  Call TheHdw.PinLevels.ApplyPower

End Function

Public Function DisableInactiveSites(argc As Long, argv() As String) As Long

'Interpose function to take stored SiteStatus and use that information to disable previously disabled
'tester sites.  Should be called as an EndOfBody function in a Test Instance

  Dim Site As Long
  
  'Debug.Print "DisableInactiveSites..."
  


  'we need to make sure all sites are active so we can remove power
  Call TheExec.Sites.SetAllActive(True)
  
  
  'power down all sites
  TheHdw.PinLevels.PowerDown

  ' Loop through all sites and de-activate the original inactive sites...
  
For Site = 0 To TheExec.Sites.ExistingCount - 1


       
  
    If Not siteStatus(Site) Then TheExec.Sites.Site(Site).Active = False
    
        'Debug.Print "Inactive = "; site
        'Debug.Print "Status = "; siteStatus(site)
        
        sites_inactive(Site) = True 'global
           'If (sites_tested(site) = True) Then 'test for pass
        
        If sites_tested(Site) = True Then
       
            If (TheExec.Sites.Site(Site).LastTestResult = 2) Then
                 Select Case Site
                    Case 0
                        site0_failed = True
                        'Debug.Print "Site "; Site; " FAILED"
                    Case 1
                        site1_failed = True
                        'Debug.Print "Site "; Site; " FAILED"
                    Case 2
                        site2_failed = True
                        'Debug.Print "Site "; Site; " FAILED"
                    Case 3
                        site3_failed = True
                        'Debug.Print "Site "; Site; " FAILED"
                    Case Else
                End Select
                
            End If
            
        End If
    
Next Site
   
 

End Function

Public Function first_active_test_sites() As Long
'
'This function determines active sites testing.

  Dim Site As Long

    'Debug.Print "first_active_test_sites..."
          
    ReDim sites_tested(TheExec.Sites.ExistingCount - 1) As Boolean
    ReDim siteStatus(TheExec.Sites.ExistingCount - 1) As Boolean
 

  ' Loop through all sites and store active status
  
  site0_failed = False      'initialize site fail indicators for use later in the flow.
  site1_failed = False
  site2_failed = False
  site3_failed = False
  
  For Site = 0 To TheExec.Sites.ExistingCount - 1
  

        sites_tested(Site) = TheExec.Sites.Site(Site).Active
        siteStatus(Site) = TheExec.Sites.Site(Site).Active

        'Debug.Print "Site "; site;
        'Debug.Print " Testing = "; sites_tested(site)


'        If (sites_tested(site) = True) Then
'            max_site = max_site + 1
'        End If


  Next Site
  
  
    'Debug.Print "max_site = "; max_site


End Function
Public Function disable_inactive_sites() As Long

'duplicate of DisableInactiveSites for use as a called function from within a test function, rather than from a Template

  Dim Site As Long
  
  'Debug.Print "DisableInactiveSites..."
  
'  ReDim sites_inactive(Site)
'  ReDim sites_tested(Site)


  'we need to make sure all sites are active so we can remove power
  Call TheExec.Sites.SetAllActive(True)
  
  
  'power down all sites
  TheHdw.PinLevels.PowerDown

  ' Loop through all sites and de-activate the original inactive sites...
  
For Site = 0 To TheExec.Sites.ExistingCount - 1


       
  
    If Not siteStatus(Site) Then TheExec.Sites.Site(Site).Active = False
    
        'Debug.Print "Inactive = "; site
        'Debug.Print "Status = "; siteStatus(site)
        
        sites_inactive(Site) = True 'global
           'If (sites_tested(site) = True) Then 'test for pass
        
        If sites_tested(Site) = True Then
       
            If (TheExec.Sites.Site(Site).LastTestResult = 2) Then
                 Select Case Site
                    Case 0
                        site0_failed = True
                        'Debug.Print "Site "; site; " FAILED"
                    Case 1
                        site1_failed = True
                        'Debug.Print "Site "; site; " FAILED"
                    Case 2
                        site2_failed = True
                        'Debug.Print "Site "; site; " FAILED"
                    Case 3
                        site3_failed = True
                        'Debug.Print "Site "; site; " FAILED"
                    Case Else
                End Select
                
            End If
            
        End If
    
Next Site
   
 

End Function

Public Function enable_store_inactive_sites() As Long

'duplicate of Enable_StoreInactiveSites for use as a called function from within another function instead of a template

  Dim Site As Long
  
  ReDim siteStatus(TheExec.Sites.ExistingCount - 1) As Boolean
  ReDim sites_inactive(TheExec.Sites.ExistingCount - 1) As Boolean
  
    'Debug.Print "Enable_StoreInactiveSites..."

  ' Loop through all sites and store active status
  
  For Site = 0 To TheExec.Sites.ExistingCount - 1
  
    siteStatus(Site) = TheExec.Sites.Site(Site).Active
    
    
        'Debug.Print "Site = "; site
        'Debug.Print "Status = "; siteStatus(site)

    
  Next Site
  


  'make all sites active
  Call TheExec.Sites.SetAllActive(True)
  
  'powerdown, for some reason the inactive site will not power up unless you power down
  TheHdw.PinLevels.PowerDown
  
'possible wait here
  
  'apply power to all site since they should all be active now.
  
  Call TheHdw.PinLevels.ApplyPower

End Function
