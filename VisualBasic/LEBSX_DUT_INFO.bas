Attribute VB_Name = "LEBSX_DUT_INFO"
'---------------------------------------------------------------------------------------------------------
' Module:       LEBDX_DUT_INFO
' Author:       David DePaul
' Purpose:      Routines to perform DUT information memory content dump to datalog
' Comments:     - Redesigned based on 160k DUT information VB module routines.
'               - Program is modular to any product type based on constant setups
'               - Read patterns must be loaded during validation through onValidate subroutine
'               - DFM testing has been blocked because LEAEX does not have EEPROM
'               - MIDBANDS returned, then MSB+1 [based on word size] will be set, and "M"s will be output
'               - GLITCH returned, then MSB+2 [based on word size] will be set, and "G"s will be output
'
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------Revision History---------------------------------------------
'---------------------------------------------------------------------------------------------------------
' Rev   Author      Date        Description
' <0>   DePaul      01/29/2008  - Initial release
'
' <1>   M. Hudiani  06/15/2009  - Modify to comply to LEAV.
'                                 Modify to use LEAV PFM/DFM size and LEAV mask name.
' <2>   R. Asuncion 10/07/2009  - Modified to comply to LEBD.
'---------------------------------------------------------------------------------------------------------
Option Explicit

' DUT Info Constants
Private Const HRAM_SIZE As Long = 256                               ' Number of History RAM Vectors
Private Const PFM_WORD_SIZE As Long = 14                            ' Number of bits per PFM word
Private Const PFM_ADDR_SIZE As Long = 14                            ' Number of bits per PFM address
Private Const DFM_WORD_SIZE As Long = 8                             ' Number of bits per DFM word
Private Const DFM_ADDR_SIZE As Long = 8                             ' Number of bits per DFM address
Private Const PFM_SET As String = "eprd_pm_dut_info_nrg_pat"        ' Program Memory Read Pattern Set Name
Private Const PFM_PAT As String = "eprd_pm_dut_info_nrg"            ' Program Memory Read Pattern Name
Private Const PAT_PM_LBL_READ As String = "di_pm_read"              ' Program Memory Pattern Read Start Label
Private Const PAT_PM_LBL_EXIT As String = "di_pm_exit"              ' Program Memory Pattern Exit Start Label
Private Const PFM_FA_SET As String = "eprd_pm_fa_dut_info_nrg_pat"  ' Program Memory FA Read Pattern Set Name
Private Const PFM_FA_PAT As String = "eprd_pm_fa_dut_info_nrg"      ' Program Memory FA Read Pattern Name
Private Const PAT_PM_FA_LBL_READ As String = "di_pm_fa_read"        ' Program Memory FA Pattern Read Start Label
Private Const PAT_PM_FA_LBL_EXIT As String = "di_pm_fa_exit"        ' Program Memory FA Pattern Exit Start Label
Private Const DFM_SET As String = "DOES_NOT_EXIST"                  ' Data Memory Read Pattern Set Name
Private Const DFM_PAT As String = "DOES_NOT_EXIST"                  ' Data Memory Read Pattern Name
Private Const PAT_DM_LBL_READ As String = "di_dm_read"              ' Data Memory Pattern Read Start Label
Private Const PAT_DM_LBL_EXIT As String = "di_dm_exit"              ' Data Memory Pattern Exit Start Label
Private Const TFM_SET As String = "eprd_tm_dut_info_nrg_pat"        ' Test Memory Read Pattern Set Name
Private Const TFM_PAT As String = "eprd_tm_dut_info_nrg"            ' Test Memory Read Pattern Name
Private Const PAT_TM_LBL_READ As String = "di_tm_read"              ' Test Memory Pattern Read Start Label
Private Const PAT_TM_LBL_EXIT As String = "di_tm_exit"              ' Test Memory Pattern Exit Start Label
Private Const EPRD_PINS As String = "eprd_bus"                      ' Ext. Para. Read Pin Group Name
Private Const ICSPDAT_PIN As String = "icspdat"                     ' ICSPDAT Pin Group Name

' Product Specific Constants
Private Const PFM_LEBD0 As Long = 2048                              ' Program Memory Size   <2> RCA
Private Const DFM_LEBD0 As Long = 0                                 ' Data Memory Size      <2> RCA
Private Const TFM_LEBD0 As Long = 512                               ' Test Memory Size
Private Const TFM_START_ADDR As Long = 8192                         ' Test Memory Start Address


' Memory Content Arrays
Private pfm() As Long                                               ' Program Memory Array
Private pfm_fa() As Long                                            ' Program Memory FA Array
Private dfm() As Long                                               ' Data Memory Array
Private tfm() As Long                                               ' Test Memory Array

' Memory Array Sizes
Private pfm_size As Long
Private dfm_size As Long
Private tfm_size As Long

' Original Trigger & Capture Setup
Private hram_setup As Boolean
Private RetTrig As TrigType
Private RetWaitForEvent As Boolean
Private RetPreTrigCycleCnt As Long
Private RetStopOnFull As Boolean
Private RetCapt As CaptType
Private RetCompressRepeats As Boolean

' Subroutine:   dut_info_onValidate
' Purpose:      Setup memory arrays and load patterns for DUT information testing
' Params:       None
' Returns:      None
Public Sub dut_info_onValidate()
    Dim RetJobName As String
    Dim RetPartName As String
    Dim RetEnv As String
    
    Call TheExec.DataManager.GetJobContext(RetJobName, RetPartName, RetEnv)
    
    If LCase(RetJobName) Like "*dut*info*" Then
        
        Select Case UCase(Trim(RetPartName))
        ' <2> RCA
        Case "LEBD1", "LEBD2", "LEBD3", "LEBD4", "LEBD5", "LEBD6"
            pfm_size = PFM_LEBD0
            dfm_size = DFM_LEBD0
            tfm_size = TFM_LEBD0
        Case Else
            Call MsgBox("Error: unknown mask name [" & RetPartName & "]... defaulting to LEBD1 PFM, DFM, and TFM Sizes.", vbOKOnly, "Error")
            pfm_size = PFM_LEBD0
            dfm_size = DFM_LEBD0
            tfm_size = TFM_LEBD0
        End Select
        
        ' Rediminsion the Arrays based on existing site coutn and memory sizes
        If pfm_size <> 0 Then ReDim pfm(0 To TheExec.Sites.ExistingCount - 1, 0 To pfm_size - 1) As Long
        If pfm_size <> 0 Then ReDim pfm_fa(0 To TheExec.Sites.ExistingCount - 1, 0 To pfm_size - 1) As Long
        If dfm_size <> 0 Then ReDim dfm(0 To TheExec.Sites.ExistingCount - 1, 0 To dfm_size - 1) As Long
        If tfm_size <> 0 Then ReDim tfm(0 To TheExec.Sites.ExistingCount - 1, 0 To tfm_size - 1) As Long
        
        ' Load the DUT info patterns into LVM
        If pfm_size <> 0 Then Call TheHdw.Digital.Patterns.Pat(PFM_SET).Load
        If pfm_size <> 0 Then Call TheHdw.Digital.Patterns.Pat(PFM_FA_SET).Load
        If dfm_size <> 0 Then Call TheHdw.Digital.Patterns.Pat(DFM_SET).Load
        If tfm_size <> 0 Then Call TheHdw.Digital.Patterns.Pat(TFM_SET).Load
        
    End If
    
End Sub


' Function:     dut_info
' Purpose:      Perform the DUT information flow by filling memory arrays and
'               outputing data to datalog.
' Params:       None
' Returns:      Long        Success/Failure [Not Used]
Public Function dut_info(argc As Long, argv() As String) As Long
    Call dut_info_hram_setup        ' Setup HRAM For Testing
    
    ' Read The DUT Memory Contents
    If pfm_size <> 0 Then Call dut_info_read_hram(pfm, PFM_PAT, PFM_WORD_SIZE, 2, EPRD_PINS, True, PAT_PM_LBL_READ, PAT_PM_LBL_EXIT)
    If pfm_size <> 0 Then Call dut_info_read_hram(pfm_fa, PFM_FA_PAT, PFM_WORD_SIZE, 2, EPRD_PINS, True, PAT_PM_FA_LBL_READ, PAT_PM_FA_LBL_EXIT)
    If dfm_size <> 0 Then Call dut_info_read_hram(dfm, DFM_PAT, DFM_WORD_SIZE, 8, ICSPDAT_PIN, False, PAT_DM_LBL_READ, PAT_DM_LBL_EXIT)
    If tfm_size <> 0 Then Call dut_info_read_hram(tfm, TFM_PAT, PFM_WORD_SIZE, 2, EPRD_PINS, True, PAT_TM_LBL_READ, PAT_TM_LBL_EXIT)
    
    ' Output The DUT Memory Contents
    Call dut_info_output
    
    Call dut_info_hram_reset        ' Reset HRAM To Original Settings
End Function

' Subroutine:   dut_info_read_hram
' Purpose:      Reburst the pattern and read the memory contents into the memory array
' Params:       mem_ary()       Long        Memory Array Reference Variable
'               mem_pat         String      Pattern Name to reburst
'               mem_word_size   Long        Number of bits per memory word
'               mem_word_hram   Long        Number of vectors of HRAM per memory word
'               mem_word_pins   String      Name of HRAM vector pins
'               mem_msb_first   Boolean     Whether, MSB or LSB first vector
'               mem_lbl_read    String      Pattern Label For Read
'               mem_lbl_exit    String      Pattern Label For Exit
Private Sub dut_info_read_hram(ByRef mem_ary() As Long, ByVal mem_pat As String, ByVal mem_word_size As Long, ByVal mem_word_hram As Long, ByVal mem_word_pins As String, ByVal mem_msb_first As Boolean, ByVal mem_lbl_read As String, ByVal mem_lbl_exit As String)
    Dim Site As Long                ' Site Number
    Dim bit_idx As Long             ' Bit Index
    Dim mem_idx As Long             ' Memory Array Index
    Dim hram_idx As Long            ' HRAM Array Index
    Dim hram_data As String         ' HRAM Result Data String
    Dim hram_data_idx As Long       ' HRAM Vector Data Index (references mem_word_hram)
    
    ' Run the Initialization Pattern
    Call TheHdw.Digital.Patterns.Pat(mem_pat).Run("")
    mem_idx = 0
    Do While (mem_idx < UBound(mem_ary, 2))
        Call TheHdw.Digital.Patterns.Pat(mem_pat).Run(mem_lbl_read)
        
        For hram_idx = 0 To HRAM_SIZE - 1 Step mem_word_hram
            If TheExec.Sites.SelectFirst <> loopDone Then
                Do
                    Site = TheExec.Sites.SelectedSite
                    mem_ary(Site, mem_idx) = 0
                    
                    hram_data = ""
                    For hram_data_idx = 0 To mem_word_hram - 1
                        If mem_msb_first Then
                            hram_data = hram_data & TheHdw.Digital.HRAM.pins(mem_word_pins).PinData(hram_idx + hram_data_idx)
                        Else
                            hram_data = TheHdw.Digital.HRAM.pins(mem_word_pins).PinData(hram_idx + hram_data_idx) & hram_data
                        End If
                    Next hram_data_idx
                                
                    ' Make sure HRAM data result is proper string length
                    If mem_msb_first Then hram_data = Left(hram_data, mem_word_size) Else hram_data = Right(hram_data, mem_word_size)
                    
                    For bit_idx = 0 To mem_word_size - 1 Step 1
                        Select Case UCase(Mid(hram_data, mem_word_size - bit_idx, 1))
                        Case "H": mem_ary(Site, mem_idx) = mem_ary(Site, mem_idx) + (2 ^ bit_idx)
                        Case "L": mem_ary(Site, mem_idx) = mem_ary(Site, mem_idx)
                        Case "M": mem_ary(Site, mem_idx) = mem_ary(Site, mem_idx) Or 2 ^ mem_word_size
                        Case "G": mem_ary(Site, mem_idx) = mem_ary(Site, mem_idx) Or 2 ^ mem_word_size + 1
                        End Select
                    Next bit_idx
                Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
            End If
            
            ' Increment memory index and verify memory index does not exceed array size.
            mem_idx = mem_idx + 1
            If mem_idx > UBound(mem_ary, 2) Then Exit For
        Next hram_idx
    Loop
    Call TheHdw.Digital.Patterns.Pat(mem_pat).Run(mem_lbl_exit)
End Sub

' Subroutine:   dut_info_output
' Purpose:      Read and Display via datalog memory contents of the DUT
' Params:       None
' Returns:      Nothing
Private Sub dut_info_output()
    Dim Site As Long            ' Site Number
    Dim mem_idx As Long         ' Memory Address Index
    
    ' Loop through each active device
    If TheExec.Sites.SelectFirst <> loopDone Then
        Do
            Site = TheExec.Sites.SelectedSite
            Call TheExec.DataLog.WriteComment("  ")
            Call TheExec.DataLog.WriteComment("   ---- SITE:" & str(Site) & " ----")
           
DUT_INFO_OUTPUT_PFM:
            If pfm_size = 0 Then GoTo DUT_INFO_OUTPUT_TFM:
            Call TheExec.DataLog.WriteComment("   ---- Program Memory ----")
            Call TheExec.DataLog.WriteComment("   Address | ---0 ---1 ---2 ---3 ---4 ---5 ---6 ---7 ---8 ---9 ---A ---B ---C ---D ---E ---F")
            For mem_idx = 0 To UBound(pfm, 2) Step 16
                Call TheExec.DataLog.WriteComment("     " & dut_info_hex(mem_idx, PFM_ADDR_SIZE) & " |" _
                                                          & dut_info_hex(pfm(Site, mem_idx + 0), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 1), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 2), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 3), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 4), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 5), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 6), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 7), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 8), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 9), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 10), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 11), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 12), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 13), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 14), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm(Site, mem_idx + 15), PFM_WORD_SIZE))
            Next mem_idx

DUT_INFO_OUTPUT_TFM:
            If tfm_size = 0 Then GoTo DUT_INFO_OUTPUT_DFM
            Call TheExec.DataLog.WriteComment(" ")
            Call TheExec.DataLog.WriteComment("   ---- Test Memory ----")
            Call TheExec.DataLog.WriteComment("   Address | ---0 ---1 ---2 ---3 ---4 ---5 ---6 ---7 ---8 ---9 ---A ---B ---C ---D ---E ---F")
            For mem_idx = 0 To UBound(tfm, 2) Step 16
                Call TheExec.DataLog.WriteComment("     " & dut_info_hex(mem_idx + TFM_START_ADDR, PFM_ADDR_SIZE) & " |" _
                                                          & dut_info_hex(tfm(Site, mem_idx + 0), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 1), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 2), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 3), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 4), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 5), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 6), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 7), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 8), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 9), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 10), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 11), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 12), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 13), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 14), PFM_WORD_SIZE) _
                                                          & dut_info_hex(tfm(Site, mem_idx + 15), PFM_WORD_SIZE))
            Next mem_idx
            
DUT_INFO_OUTPUT_DFM:
            If dfm_size = 0 Then GoTo DUT_INFO_OUTPUT_PFM_FA
            Call TheExec.DataLog.WriteComment(" ")
            Call TheExec.DataLog.WriteComment("   ---- Data Memory ----")
            Call TheExec.DataLog.WriteComment("   Address | -0 -1 -2 -3 -4 -5 -6 -7 -8 -9 -A -B -C -D -E -F")
            For mem_idx = 0 To UBound(dfm, 2) Step 16
                Call TheExec.DataLog.WriteComment("     " & dut_info_hex(mem_idx, DFM_ADDR_SIZE) & " |" _
                                                          & dut_info_hex(dfm(Site, mem_idx + 0), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 1), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 2), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 3), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 4), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 5), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 6), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 7), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 8), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 9), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 10), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 11), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 12), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 13), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 14), DFM_WORD_SIZE) _
                                                          & dut_info_hex(dfm(Site, mem_idx + 15), DFM_WORD_SIZE))
            Next mem_idx
            
DUT_INFO_OUTPUT_PFM_FA:
            If pfm_size = 0 Then GoTo DUT_INFO_OUTPUT_FINISHED
            Call TheExec.DataLog.WriteComment(" ")
            Call TheExec.DataLog.WriteComment("   ---- Program Memory FA Verify ----")
            Call TheExec.DataLog.WriteComment("   Address | ---0 ---1 ---2 ---3 ---4 ---5 ---6 ---7 ---8 ---9 ---A ---B ---C ---D ---E ---F")
            For mem_idx = 0 To UBound(pfm_fa, 2) Step 16
                Call TheExec.DataLog.WriteComment("     " & dut_info_hex(mem_idx, PFM_ADDR_SIZE) & " |" _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 0), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 1), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 2), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 3), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 4), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 5), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 6), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 7), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 8), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 9), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 10), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 11), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 12), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 13), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 14), PFM_WORD_SIZE) _
                                                          & dut_info_hex(pfm_fa(Site, mem_idx + 15), PFM_WORD_SIZE))
            Next mem_idx
            
DUT_INFO_OUTPUT_FINISHED:
            Call TheExec.DataLog.WriteComment("   -----------------------------------------------------------------------------------------")
        Loop While TheExec.Sites.SelectNext(loopTop) <> loopDone
    End If
   
End Sub

' Subroutine:   dut_info_hram_setup
' Purpose:      Retrieve original HRAM setup and then force HRAM setup to capture only STV
' Params:       N/A
' Returns:      N/A
Private Sub dut_info_hram_setup()
    ' Get Original HRAM Setup Prior to Updating If HRAM Has Not Previously Been Setup
    If Not hram_setup Then
        Call TheHdw.Digital.HRAM.GetTrigger(RetTrig, RetWaitForEvent, RetPreTrigCycleCnt, RetStopOnFull)
        Call TheHdw.Digital.HRAM.GetCapture(RetCapt, RetCompressRepeats)
    End If
    
    ' Setup HRAM to capture STV
    Call TheHdw.Digital.HRAM.SetTrigger(trigSTV, False, 0, False)
    Call TheHdw.Digital.HRAM.SetCapture(captSTV, False)
    hram_setup = True
End Sub

' Subroutine:   dut_info_hram_reset
' Purpose:      Reset HRAM setup to original settings.
' Params:       N/A
' Returns:      N/A
Private Sub dut_info_hram_reset()
    ' Reset HRAM setup to original state
    Call TheHdw.Digital.HRAM.SetTrigger(RetTrig, RetWaitForEvent, RetPreTrigCycleCnt, RetStopOnFull)
    Call TheHdw.Digital.HRAM.SetCapture(RetCapt, RetCompressRepeats)
    hram_setup = False
End Sub

' Function:     dut_info_hex
' Purpose:      Return a string containing the hex value of memory data. The string
'               is to be sized based on the number of bits per word
' Params:       data            Long            Data to convert to HEX string
'               word_size       Long            Number of bits data represents
Private Function dut_info_hex(ByVal data As Long, ByVal word_size As Long) As String
    Dim hex_str As String       ' HEX string working variable
    Dim hex_len As Long         ' HEX string length based on word_size
    Dim str_idx As Long         ' HEX string index variable
    
    ' Determine HEX string length based on word_size
    hex_len = Int(word_size / 4)
    If (word_size Mod 4) <> 0 Then hex_len = hex_len + 1
    
    ' Determine if data returned as mid-band, glitch, or actual
    If data And 2 ^ (word_size) Then
        For str_idx = 1 To hex_len Step 1
            hex_str = "M" & hex_str
        Next str_idx
    ElseIf data And 2 ^ (word_size + 1) Then
        For str_idx = 1 To hex_len Step 1
            hex_str = "G" & hex_str
        Next str_idx
    Else
        hex_str = Hex(data And (2 ^ word_size - 1))
        str_idx = Len(hex_str)
        Do While (str_idx < hex_len)
            hex_str = "0" & hex_str
            str_idx = str_idx + 1
        Loop
    End If
    
    dut_info_hex = " " & hex_str
End Function

