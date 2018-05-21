Attribute VB_Name = "AdcLin_v2"
Option Explicit

' [=============================================================================]
' [ DEVICE :   All devices with A/D's                                           ]
' [ MASK NO:   any                                                              ]
' [ SCOPE  :   Currently debugged and checked out for 160k PICs with 10 bit     ]
' [            A/D converters, running 10 hits per code.                        ]
' [=============================================================================]
' [                                                                             ]
' [                   MICROCHIP TECHNOLOGY INC.                                 ]
' [                   2355 WEST CHANDLER BLVD.                                  ]
' [                   CHANDLER AZ 85224-6199                                    ]
' [                   (602) 963-7373                                            ]
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
' [ ^^^^   ^^^^    ^^^  ^^^^^^^                                                 ]                                                   ]
' [ v2r0  06jan06  da   - Initial release of AdcLin_v2                          ]
' [ v2r1  09sep07  da   - Fixed bug to support 0.5LSB first transition point    ]
' [                     - Added support to detect 'sparkle' defined here as     ]
' [                       a transition of 2 or mode codes at once (up or down)  ]
' [                     - Updated parametric reporting to report units in LSB   ]
' [                     - Public/Global variables required changed to support   ]
' [                       sparkle implementation                                ]
' [ v2r2  20aug08  da   - Added support for Partial Code Testing                ]
' [ v2r3  06jan09  da   - Added new debug mode: ADC_DB_summary                  ]
' [                       which must be defined in _specific module             ]
' [                                                                             ]
' [                                                                             ]
' [                                                                             ]
' [=============================================================================]


' ** NOTICE: **
' This code was originally taken from test program DEBN0_B6 (which in turn
' appeared to come from Teradyne documentation) and heavily modified
' by David Aristizabal in Jan of 2006 for the purpose of creating a general
' module for computing the A/D DC parameters associated with testing
' A/D's on PIC's or otherwise.






' *****************************************************************************
' FUNCTION:    AdcLin_analysis
'
' Analyses the results collected from the DUT, computes requested parameters
' such as offset, gain, DNL, INL, absolute error and missing codes, determines
' pass/fail conditions and sends the results to the datalog.
'      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' REQUIREMENTS:
' * This function requires the current test instance to be running from the
'   IG-XL ADC Template.  This function collects most of the user parameters
'   from what was filled into the template.
'
' * This module requires the existance of module AdcLin_v2_Specific which
'   performs all the DUT/pattern/implementation-specific actions to collect
'   the A/D results from the DUT.
'
' * This function is is NOT an interpose function and should not be specified
'   directly as an interpose function in the test template.  Instead, the test
'   template must call the end-of body function the AdcLin_v2_specific module,
'   which in turn calls this function passing the required parameters.
'
' * The AdcLin_v2_Specific module must define and initialize the following
'   public/global variables
'
'   Sparkle Options:
'   Note: The template form has no input fields for sparkle options,
'         so the options are defined here and values can be assiged
'         on the startOfBodyIF
'      adc_testSparkle
'      adc_sparkleSensitivity
'      adc_sparkleLimit
'
'   Debugging Variables must be defined in this module:
'      ADC_DB_MODE                   Bitmask that enables debug modes
'      ADC_DB_printResultsPerCode    Debug mode Bit flag
'      ADC_DB_printHits              Debug mode Bit flag
'      ADC_DB_vbCodeTime             Debug mode Bit flag
'      ADC_DB_summary                Debug mode Bit flag
'      adc_db_tmpStr
'      adc_db_tOverall              ' Overall time in analysis function
'      adc_db_tOverallRef
'      adc_db_tGetArgs              ' time taken to get template arguments
'      adc_db_tGetArgsRef
'
'    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' PARAMETERS:
'   ADresults()        Input: 2D Array, whos dimensions are given by
'                      other input parameters ADresults_Rows and ADresults_Cols.
'                      It contains the digital result for every conversion made
'                      during the test.
'
'   ADresults_Rows     Input: Specifies the number of rows in the array.
'                      This is the first dimension.
'                      The number of rows MUST match the number of SITES
'                      in the current channelmap.
'
'   ADresults_Cols     Input: Specifies the number of columns in the array.
'                      This is the second dimension.
'                      This must be larger or equal to the expected total number
'                      hits/conversions made during the test.
'    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' RETURNS:
'  AdcLin_analysis returns a status code as a long.
'  Currently it is set up to always return 0.
' *****************************************************************************
Public Function AdcLin_analysis(ByRef ADresults() As Integer, _
                                ByVal ADresults_Rows As Long, _
                                ByVal ADresults_Cols As Integer) As Long

 Stop
' '20170217 - ty commented out this code due to many variables not existing....
'
'
'    ' These variables and constants are used to access the instance
'    ' arguments when needed by User Review Functions or Interpose
'    ' Functions.  These declarations may be made with Function
'    ' scope, or with Module scope if needed by more than one Function.
'    ' These declarations should be copied from the template module,
'    ' in this case CtoAdc_T in the Template.xla VBA project.
'    ' This technique may be used with any template, noting that each
'    ' template has a unique set of arguments which can be copied from
'    ' the appropriate template module.
'    ' Variables to hold the instance argument values
'    Dim Arg_DcCategory As String, Arg_DcSelector As String, _
'    Arg_AcCategory As String, Arg_AcSelector As String, _
'    Arg_Timing As String, Arg_Edgeset As String, _
'    Arg_Levels As String, Arg_Bits As String, _
'    Arg_Pinlist As String, Arg_AnalogIn As String, _
'    Arg_VrefLVal As String, Arg_VrefHVal As String, _
'    Arg_HistoSave As String, Arg_CalRef As String, _
'    Arg_Pattern As String, Arg_AinData As String, _
'    Arg_DifErr As String, Arg_IntErr As String, _
'    Arg_OffErr As String, Arg_GainErr As String, _
'    Arg_MissErr As String, Arg_Hits As String, _
'    Arg_VEnd As String, Arg_VStart As String, _
'    Arg_SaveResults As String, Arg_CtoRange As String, _
'    Arg_Serial As String, Arg_PreconditionPat As String, _
'    Arg_PcpStartLabel As String, Arg_PcpStopLabel As String, _
'    Arg_DriverLO  As String, Arg_DriverHI  As String, _
'    Arg_DriverZ As String, Arg_FloatPins As String
'    Dim Arg_StartOfBodyF As String, Arg_PrePatF As String, _
'    Arg_PreTestF As String, Arg_PostTestF As String, _
'    Arg_PostPatF As String, Arg_EndOfBodyF As String, _
'    Arg_StartOfBodyFInput As String, Arg_PrePatFInput As String, _
'    Arg_PreTestFInput As String, Arg_PostTestFInput As String, _
'    Arg_PostPatFInput As String, Arg_EndOfBodyFInput As String, _
'    Arg_RelayMode As String, Arg_UserReviewF As String, _
'    Arg_TestType As String, Arg_PcpCheckPatGen As String, _
'    Arg_ReviewType As String, Arg_DataPoints As String, _
'    Arg_LimitType As String, Arg_TransitionPoint As String, _
'    Arg_NormalizationMethod As String, Arg_AbsErr As String, _
'    Arg_Monotonicity As String, Arg_Util1 As String, _
'    Arg_Util0 As String, Arg_VrefLPin As String, _
'    Arg_VrefHPin As String, Arg_CheckBox As String, _
'    Arg_LsbBeyond As String, Arg_RefOrEnd As String
'
'    ' Constants required to retrieve each argument from the argument array
'
'    Const ARGNUM_BITS = 0
'    Const ARGNUM_PINLIST = 1
'    Const ARGNUM_ANALOG = 2
'    Const ARGNUM_VREFLVAL = 3
'    Const ARGNUM_VREFHVAL = 4
'
'    Const ARGNUM_HISTO = 5
'    Const ARGNUM_CALREF = 6
'    Const ARGNUM_PATTERN = 7
'    Const ARGNUM_AINCODE = 8
'    Const ARGNUM_DIFERR = 9
'    Const ARGNUM_INTERR = 10
'    Const ARGNUM_OFFERR = 11
'    Const ARGNUM_GAINERR = 12
'    Const ARGNUM_MISSERR = 13
'    Const ARGNUM_HITS = 14
'    Const ARGNUM_VENDVAL = 15
'    Const ARGNUM_VSTARTVAL = 16
'    Const ARGNUM_DATAPOINTS = 17
'    Const ARGNUM_SAVERESULTS = 18
'
'    Const ARGNUM_CTORANGE = 19
'    Const ARGNUM_SERIAL = 20
'    Const ARGNUM_PRECONDITIONPAT = 21
'    Const ARGNUM_PCPSTARTLABEL = 22
'    Const ARGNUM_PCPSTOPLABEL = 23
'    Const ARGNUM_DRIVERLO = 24
'    Const ARGNUM_DRIVERHI = 25
'    Const ARGNUM_DRIVERZ = 26
'    Const ARGNUM_FLOATPINS = 27
'    Const ARGNUM_STARTOFBODYF = 28
'    Const ARGNUM_PREPATF = 29
'    Const ARGNUM_PRETESTF = 30
'    Const ARGNUM_POSTTESTF = 31
'
'    Const ARGNUM_POSTPATF = 32
'    Const ARGNUM_ENDOFBODYF = 33
'    Const ARGNUM_STARTOFBODYFINPUT = 34
'    Const ARGNUM_PREPATFINPUT = 35
'    Const ARGNUM_PRETESTFINPUT = 36
'    Const ARGNUM_POSTTESTFINPUT = 37
'    Const ARGNUM_POSTPATFINPUT = 38
'    Const ARGNUM_ENDOFBODYFINPUT = 39
'    Const ARGNUM_RELAYMODE = 40
'    Const ARGNUM_USERREVIEWF = 41
'    Const ARGNUM_TESTTYPE = 42
'    Const ARGNUM_PCPCHECKPATGEN = 43
'
'    Const ARGNUM_REVIEWTYPE = 44
'    Const ARGNUM_LIMITTYPE = 45
'    Const ARGNUM_TRANSITIONPOINT = 46
'    Const ARGNUM_NORMMETHOD = 47
'    Const ARGNUM_ABSERR = 48
'    Const ARGNUM_MONOTONICITY = 49
'    Const ARGNUM_UTIL1 = 50
'    Const ARGNUM_UTIL0 = 51
'    Const ARGNUM_VREFLPIN = 52
'    Const ARGNUM_VREFHPIN = 53
'    Const ARGNUM_CHECKBOX = 54
'    Const ARGNUM_LSBBEYOND = 55
'    Const ARGNUM_REFOREND = 56
'
'    Const ARGNUM_MAXARG = ARGNUM_REFOREND
'
'    '' Note: These additional definitions are in Template.xla, CtoSupport module
'    ''    Public Const ADC_BIT_DIFF = 0   'this is used to define the bit position in the
'    ''    Public Const ADC_BIT_INT = 1    '   Arg_CheckBox which denotes the enabling of
'    ''    Public Const ADC_BIT_OFF = 2    '   the corresponding test.
'    ''    Public Const ADC_BIT_GAIN = 3
'    ''    Public Const ADC_BIT_MISS = 4
'    ''    Public Const ADC_BIT_ABS = 5
'
'
'    Dim appliedVoltageArray() As Double   ' the ADC stimuli
'    Dim CodeHistogram() As Long           ' the codes returned from the DUT
'    Dim sparkleCount As Long              ' number of sparkle errors detected
'
'    Dim dnlErr As Double                  ' linearity error for each code
'    Dim numCodes As Long                  ' number of codes for the ADC DUT
'    Dim siteStatus As Long
'
'    Dim DNLFlag As Long
'    Dim INLFlag As Long
'    Dim MCFlag As Long
'    Dim OffFlag As Long
'    Dim GainFlag As Long
'    Dim AbsErrFlag As Long
'    Dim sparkleFlag As Long
'
'    Dim CodeVal As Integer
'    Dim parmFlag As Long
'    Dim thisSite As Long
'    Dim nsites As Long
'    Dim err As String
'    Dim PinName As String
'
'    Dim ReturnStatus As Long
'    Dim CodeReturned As Long
'    Dim prevCodeReturned As Long
'    Dim hitsPerCode As Integer           ' User specified hitsPerCode
'    Dim AverageHitsPerCode As Double
'    Dim ExpectedNumDataItems As Long
'    Dim NumDataItems As Long
'    Dim ExpectedNumSites As Integer
'    Dim RetHedSiz As Integer             ' Returned Header Size
'    Dim RetStpSiz As Integer
'
'    Dim TestNumDNL As Long
'    Dim TestNumINL As Long
'    Dim TestNumMC As Long
'    Dim TestNumOff As Long
'    Dim TestNumGain As Long
'    Dim TestNumAbsErr As Long
'    Dim TestNumSparkle As Long
'
'
'    Dim testStatus As Long               ' overall pass/fail result
'    Dim testStatusINL As Long            ' INL pass/fail result
'    Dim testStatusDNL As Long            ' DNL pass/fail result
'    Dim testStatusMC As Long             ' Missing Codes pass/fail result
'    Dim testStatusOff As Long
'    Dim testStatusGain As Long
'    Dim testStatusAbsErr As Long
'    Dim testStatusSparkle As Long
'
'    Dim DnlLimit As Double
'    Dim InlLimit As Double
'    Dim MissingCodesLimit As Double
'    Dim OffLimit As Double
'    Dim GainErrorLimit As Double
'    Dim AbsErrorLimit As Double
'    Dim sparkleLimit As Double
'    Dim sparkleSensitivity As Integer   ' min size of sparkle to be detected
'    Dim lastSparkle As Long
'
'
'    Dim worstDNL As Double
'    Dim worstDNLcode As Long
'    Dim worstINL As Double
'    Dim worstINLcode As Long
'    Dim TotalItemCountPointer As Long
'    Dim codei As Long                    ' general code index variable
'    Dim hitIndex As Long
'    Dim ArgStr() As String               ' array of template argument strings
'
'    Dim NumLsbBeyond As Integer
'    Dim intErr As Double
'    Dim NumMissingCodes As Double
'    Dim lastMC As Long
'    Dim OffsetError As Double
'    Dim GainError As Double
'    Dim AbsError As Double
'    Dim worstAbsErrCode As Long
'    Dim idealNumOfHits As Long
'    Dim actualNumOfHits As Long
'
'    Dim StartVoltage As Double
'    Dim EndVoltage As Double
'
'    Dim worstAbsErr As Double
'    Dim MaxCodeNotFound As Integer
'    Dim PresentCode As Integer
'    Dim absErr As Double
'    Dim TransitionPoint As Double
'    Dim loopstatus As Long
'
'    ' What tests to actually run:
'    Dim checkboxData As Integer
'    Dim testInl As Boolean
'    Dim testDnl As Boolean
'    Dim testMC As Boolean
'    Dim testOffset As Boolean
'    Dim testGain As Boolean
'    Dim testAbsErr As Boolean
'
'    Dim testSparkle As Boolean
'
'
'    ' 20170216 - ty commented the follow section bc the variables don't exist
'    ' *** DEBUG MODE ? ***
'    'If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        'adc_db_tOverallRef = TheExec.Timer
'        'adc_db_tGetArgsRef = TheExec.Timer
'    'End If
'
'
'    ' Getting the test instance arguments.
'    ' First, get the arguments for the current instance
'    Call TheExec.DataManager.GetArgumentList(ArgStr, ARGNUM_MAXARG)
'
'
'    ' Now pick out the arguments we need,
'    ' and place them in convenient variables
'    Arg_Hits = ArgStr(ARGNUM_HITS)          ' number of hits per code
'    Arg_Bits = ArgStr(ARGNUM_BITS)          ' number of ADC output bits
'    Arg_AnalogIn = ArgStr(ARGNUM_ANALOG)    ' Name of Analog Pin
'    Arg_DifErr = ArgStr(ARGNUM_DIFERR)      ' DNL Limit
'    Arg_IntErr = ArgStr(ARGNUM_INTERR)      ' INL Limit
'    Arg_LsbBeyond = ArgStr(ARGNUM_LSBBEYOND)    ' #LSB's beyond range
'    Arg_OffErr = ArgStr(ARGNUM_OFFERR)      ' Offset Error Limit
'    Arg_MissErr = ArgStr(ARGNUM_MISSERR)    ' Num of Missing Codes Limit
'    Arg_VEnd = ArgStr(ARGNUM_VENDVAL)       ' Programmed Ramp End Voltage
'    Arg_VStart = ArgStr(ARGNUM_VSTARTVAL)   ' Programmed Ramp Start Voltage
'    Arg_GainErr = ArgStr(ARGNUM_GAINERR)
'    Arg_AbsErr = ArgStr(ARGNUM_ABSERR)
'    Arg_TransitionPoint = ArgStr(ARGNUM_TRANSITIONPOINT) ' 0-1 transition in LSB
'    Arg_CheckBox = ArgStr(ARGNUM_CHECKBOX)
'    hitsPerCode = CInt(Arg_Hits)
'    DnlLimit = CDbl(Arg_DifErr)
'    InlLimit = CDbl(Arg_IntErr)
'    NumLsbBeyond = CInt(Arg_LsbBeyond)
'    MissingCodesLimit = CDbl(Arg_MissErr)
'    OffLimit = CDbl(Arg_OffErr)
'    EndVoltage = CDbl(Arg_VEnd)
'    StartVoltage = CDbl(Arg_VStart)
'    GainErrorLimit = CDbl(Arg_GainErr)
'    AbsErrorLimit = CDbl(Arg_AbsErr)
'
'    ' Tansition Point
'    ' Template form has two options:
'    '   "0.5 LSB"  -- Arg_TransitionPoint = 0
'    '   "  1 LSB"  -- Arg_TransitionPoint = 0
'    If Arg_TransitionPoint = 0 Then
'        TransitionPoint = 0.5
'    Else
'        TransitionPoint = 1#
'    End If
'
'    'sparkleLimit = adc_sparkleLimit   ' 20170216 - ty commented out bc var not exist
'    sparkleLimit = -1
'    'sparkleSensitivity = adc_sparkleSensitivity
'    sparkleSensitivity = -1             ' 20170216 - ty commented out bc var not exist
'
'    ' Determine what tests to perform/check:
'    checkboxData = CInt(Val(Arg_CheckBox))
'    testMC = True
'    testGain = True
'    testDnl = True
'    testInl = True
'    testOffset = True
'    testAbsErr = True
'    testSparkle = adc_testSparkle
'    If (checkboxData Or (2 ^ ADC_BIT_MISS)) <> checkboxData Then testMC = False
'    If (checkboxData Or (2 ^ ADC_BIT_GAIN)) <> checkboxData Then testGain = False
'    If (checkboxData Or (2 ^ ADC_BIT_DIFF)) <> checkboxData Then testDnl = False
'    If (checkboxData Or (2 ^ ADC_BIT_INT)) <> checkboxData Then testInl = False
'    If (checkboxData Or (2 ^ ADC_BIT_OFF)) <> checkboxData Then testOffset = False
'    If (checkboxData Or (2 ^ ADC_BIT_ABS)) <> checkboxData Then testAbsErr = False
'
'    ' Getting the applied voltage values array.
'    ' Now get the array of values that the Template used as stimulus
'    ' for the DUT.  These ramps are registered under the name of the
'    ' instance.  (The array is not actually used in this function.)
'    Arg_AinData = TheExec.DataManager.instanceName
'    ReturnStatus = tl_GetArrayDouble(Arg_AinData, appliedVoltageArray)
'
'    If ReturnStatus <> TL_SUCCESS Then
'        Call TheExec.ErrorLogMessage(Arg_AinData + _
'                ": No Applied Voltage Array ")
'        Call TheExec.ErrorReport
'        Exit Function
'    End If
'
'
'    ' *** DEBUG MODE ? ***
'    If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        adc_db_tGetArgs = TheExec.Timer(adc_db_tGetArgsRef)
'    End If
'
'
'    ' Calculate how many codes this ADC has
'    numCodes = 2 ^ CInt(Arg_Bits)
'    ExpectedNumDataItems = (numCodes + (NumLsbBeyond * 2)) * hitsPerCode
'    ExpectedNumSites = TheExec.Sites.ExistingCount
'    ReDim CodeHistogram(numCodes - 1)
'
'
'    ' -----------------------  Loop through sites ----------------------
'    loopstatus = TheExec.Sites.SelectFirst
'    While loopstatus <> loopDone
'    thisSite = TheExec.Sites.SelectedSite
'
'        ' get and set test numbers
'        TestNumOff = TheExec.Sites.Site(thisSite).testnumber
'        TestNumGain = TheExec.Sites.Site(thisSite).testnumber + 1
'        TestNumDNL = TheExec.Sites.Site(thisSite).testnumber + 2
'        TestNumINL = TheExec.Sites.Site(thisSite).testnumber + 3
'        TestNumMC = TheExec.Sites.Site(thisSite).testnumber + 4
'        TestNumAbsErr = TheExec.Sites.Site(thisSite).testnumber + 5
'        TestNumSparkle = TheExec.Sites.Site(thisSite).testnumber + 6
'
'        ' see how many data items this site returned
'        If ADresults_Cols < (ExpectedNumDataItems) Then
'            ' Fail the site
'            Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                ": Not enough data points returned for site " + _
'                    CStr(thisSite))
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'            GoTo NextSite
'        End If
'        NumDataItems = ExpectedNumDataItems 'only want to look at valid data items
'
'
'        ' Make sure number of sites returned makes sense:
'        If ADresults_Rows <> ExpectedNumSites Then
'            ' Fail the site
'            Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                ": Number of sites returned with results array (adresults) is invalid.")
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'            GoTo NextSite
'        End If
'
'
'        ' *********** HISTOGRAM CONSTRUCTION & SPARKLE DETECTION ******************
'
'        ' Clear the histogram data:
'        sparkleCount = 0
'        For codei = 0 To numCodes - 1
'            CodeHistogram(codei) = 0
'        Next codei
'
'        prevCodeReturned = ADresults(thisSite, 0)
'        For hitIndex = 0 To NumDataItems - 1
'
'            CodeReturned = ADresults(thisSite, hitIndex)
'            If CodeReturned > numCodes - 1 Then
'                ' an illegal code was returned!
'                Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                    ": Illegal code returned by site " + _
'                        CStr(thisSite))
'                TheExec.Sites.Site(thisSite).TestResult = siteFail
'                GoTo NextSite
'            End If
'            CodeHistogram(CodeReturned) = CodeHistogram(CodeReturned) + 1
'
'            ' Sparkle detection:
'            If Abs(CodeReturned - prevCodeReturned) >= sparkleSensitivity Then
'                sparkleCount = sparkleCount + 1
'                lastSparkle = CodeReturned
'            End If
'
'            prevCodeReturned = CodeReturned
'
'        Next hitIndex
'
'        ' Determine the average number of hits per bin.
'        ' For an ideal converter, the average number of hits per bin
'        ' would be the same as the number of hits per code in the
'        ' stimulus ramp.  For a real converter, offset and gain make
'        ' the average number of hits per bin different from ideal.
'        ' We normalize for offset and gain by comparing
'        ' the observed number of hits for each code with the observed
'        ' average, instead of with the number in the stimulus ramp.
'        ' We are only interested in "interior codes", i.e., not 0 or
'        ' full-scale.  This is because any input voltage less than the
'        ' first transition point (or greater than the last) can produce
'        ' the first or last code, and potentially make these bins much
'        ' more populated than the interior bins.
'
'        AverageHitsPerCode = (NumDataItems - CodeHistogram(0) - CodeHistogram(numCodes - 1)) _
'                             / (numCodes - 2)
'
'        If (AverageHitsPerCode = 0) Then
'            AverageHitsPerCode = 1
'        End If
'
'
'
'        ' *** DEBUG MODE ? ***
'        If (ADC_DB_MODE And (ADC_DB_printResultsPerCode Or ADC_DB_printHits Or ADC_DB_summary)) <> 0 Then
'            adc_db_tmpStr = "::ADC_DB - Site " & thisSite & "::"
'            TheExec.DataLog.WriteComment (adc_db_tmpStr)
'            adc_db_tmpStr = ""
'            If (ADC_DB_MODE And ADC_DB_printHits) <> 0 Then
'               For hitIndex = 0 To NumDataItems - 1
'                    ' print 20 codes per line, separated by ':'
'                    If hitIndex Mod 20 = 0 Then
'                        TheExec.DataLog.WriteComment (adc_db_tmpStr)
'                        adc_db_tmpStr = "::ADC_DB::"
'                    End If
'                    adc_db_tmpStr = adc_db_tmpStr & ADresults(thisSite, hitIndex) & ":"
'               Next hitIndex
'               TheExec.DataLog.WriteComment (adc_db_tmpStr)
'            End If
'        End If
'
'
'
'
'        ' ***********  OFFSET ERROR CALCULATION ****************
'
'        ' ideal number of hits for code 0:
'        idealNumOfHits = (NumLsbBeyond + TransitionPoint) * hitsPerCode
'
'        OffsetError = (CodeHistogram(0) - idealNumOfHits) / hitsPerCode
'
'
'
'
'
'        ' *************  GAIN ERROR CALCULATION ****************
'
'        ' Compute ideal number of hits of last code:
'        idealNumOfHits = (NumLsbBeyond + 1 + (1 - TransitionPoint)) * hitsPerCode
'
'        GainError = (CodeHistogram(numCodes - 1) - idealNumOfHits) / hitsPerCode
'
'        ' Compensate for offset error:
'        ' Note: offset error is added to gain error, because when
'        '       offset error is negative, it naturally makes the
'        '       last code longer making the uncompensated gain error
'        '       larger.
'        GainError = GainError + OffsetError
'
'
'
'
'
'
'        ' ********** DNL, INL, MISSING & ABSOLUTE ERROR CALCULATIONS PER CODE *********
'        worstDNL = 0
'        worstINL = 0
'        worstDNLcode = 0
'        worstINLcode = 0
'        NumMissingCodes = 0
'        lastMC = 0
'        worstAbsErr = 0
'        worstAbsErrCode = 0
'        actualNumOfHits = 0
'
'        intErr = 0 ' Must reset for each site
'
'        ' *** DEBUG MODE ? ***
'        If (ADC_DB_MODE And ADC_DB_printResultsPerCode) <> 0 Then
'            TheExec.DataLog.WriteComment ("::ADC_DB::DNL[" & 0 & "]=" & Format(0, "0.000000") & ";INL[" & 0 & "]=" & Format(0, "0.000000") & ";AbsErr[" & 0 & "]=" & Format(0, "0.000000") & ";HitsPerCode[" & 0 & "]=" & Format(CodeHistogram(0), "0000") & ";")
'        End If
'
'
'        For codei = 1 To numCodes - 2
'            ' *** DNL ***
'            ' Compute the difference between the code width and the
'            ' expected 1 LSB code width.  For each code, the width is
'            ' the ratio of the observed hits per code and the average
'            ' hits per code.
'            ' NOTE: Comparing against the average hits per code has
'            '       the effect of compensating for gain error.
'            '       Also, there is no need to compensate for offset
'            '       because the histogram method is used.
'
'
'            dnlErr = (CDbl(CodeHistogram(codei)) / AverageHitsPerCode) - 1#
'
'            If Abs(dnlErr) > Abs(worstDNL) Then
'                ' a new largest DNL error was observed
'                worstDNL = dnlErr
'                worstDNLcode = codei
'            End If
'
'            ' *** INL ***
'            ' The Sum of differential errors is the integral error
'            intErr = intErr + dnlErr
'            If Abs(intErr) > Abs(worstINL) Then
'                ' a new maximum INL error was observed
'                worstINL = intErr
'                worstINLcode = codei
'            End If
'
'            ' *** Missing Codes ***
'            If CodeHistogram(codei) < 1 Then
'                NumMissingCodes = NumMissingCodes + 1
'                lastMC = codei
'            End If
'
'
'            ' *** Absolute Error ***
'            ' Computed by comparing the actual number of hits counted up
'            ' to the transition (codei-1  to codei) to the ideal number
'            ' of hits expected up to the transition (codei-1  to codei).
'            idealNumOfHits = (NumLsbBeyond + TransitionPoint + codei - 1) * hitsPerCode
'            actualNumOfHits = actualNumOfHits + CodeHistogram(codei - 1)
'            absErr = (actualNumOfHits - idealNumOfHits) / hitsPerCode
'            If Abs(absErr) > Abs(worstAbsErr) Then
'                worstAbsErr = absErr
'                worstAbsErrCode = codei
'            End If
'
'
'            ' *** DEBUG MODE ? ***
'            If (ADC_DB_MODE And ADC_DB_printResultsPerCode) <> 0 Then
'                TheExec.DataLog.WriteComment ("::ADC_DB::DNL[" & codei & "]=" & Format(dnlErr, "0.000000") & ";INL[" & codei & "]=" & Format(intErr, "0.000000") & ";AbsErr[" & codei & "]=" & Format(absErr, "0.000000") & ";HitsPerCode[" & codei & "]=" & Format(CodeHistogram(codei), "0000") & ";")
'            End If
'
'        Next codei
'
'
'        ' Must include last code for MissingCodes and abs error
'        codei = numCodes - 1
'
'        ' *** Missing Codes: check last code ***
'        If CodeHistogram(codei) < 1 Then
'            NumMissingCodes = NumMissingCodes + 1
'            lastMC = codei
'        End If
'
'        ' *** Absolute Error: Last code transition ***
'        idealNumOfHits = (NumLsbBeyond + TransitionPoint + codei - 1) * hitsPerCode
'        actualNumOfHits = actualNumOfHits + CodeHistogram(codei - 1)
'        absErr = (actualNumOfHits - idealNumOfHits) / hitsPerCode
'        If Abs(absErr) > Abs(worstAbsErr) Then
'            worstAbsErr = absErr
'            worstAbsErrCode = codei
'        End If
'
'
'        ' *** DEBUG MODE ? ***
'        If (ADC_DB_MODE And ADC_DB_printResultsPerCode) <> 0 Then
'            TheExec.DataLog.WriteComment ("::ADC_DB::DNL[" & codei & "]=" & Format(0, "0.000000") & ";INL[" & codei & "]=" & Format(0, "0.000000") & ";AbsErr[" & codei & "]=" & Format(absErr, "0.000000") & ";HitsPerCode[" & codei & "]=" & Format(CodeHistogram(codei), "0000") & ";")
'        End If
'        If (ADC_DB_MODE And (ADC_DB_printResultsPerCode Or ADC_DB_summary)) <> 0 Then
'            TheExec.DataLog.WriteComment ("::ADC_DB::Offset=" & OffsetError & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::GainErr=" & GainError & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::DNLErr=" & worstDNL & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstDNLcode=" & worstDNLcode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::INLErr=" & worstINL & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstINLcode=" & worstINLcode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::MissCodes=" & NumMissingCodes & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::AbsErr=" & worstAbsErr & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstAbsErrCode=" & worstAbsErrCode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::AvgHPC=" & AverageHitsPerCode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::Sparkle=" & sparkleCount & ";")
'        End If
'
'
'
'
'
'        ' ***********  TEST PASS/FAIL CONDITIONS ***************
'
'        testStatus = logTestPass
'        testStatusOff = logTestPass
'        testStatusGain = logTestPass
'        testStatusDNL = logTestPass
'        testStatusINL = logTestPass
'        testStatusMC = logTestPass
'        testStatusAbsErr = logTestPass
'        testStatusSparkle = logTestPass
'
'        ' check results against limits
'
'        If (OffsetError > OffLimit) And testOffset Then
'            ' Offset Error was outside limit (too high)
'            testStatusOff = logTestFail
'            testStatus = logTestFail
'            OffFlag = parmHigh
'        End If
'        If (OffsetError < -OffLimit) And testOffset Then
'            ' Offset Error was outside limit (too low)
'            testStatusOff = logTestFail
'            testStatus = logTestFail
'            OffFlag = parmLow
'        End If
'
'        If (GainError > GainErrorLimit) And testGain Then
'            ' Gain Error was outside limit (too high)
'            testStatusGain = logTestFail
'            testStatus = logTestFail
'            GainFlag = parmHigh
'        End If
'        If (GainError < -GainErrorLimit) And testGain Then
'            ' Gain Error was outside limit (too low)
'            testStatusGain = logTestFail
'            testStatus = logTestFail
'            GainFlag = parmLow
'        End If
'
'        If (worstDNL > DnlLimit) And testDnl Then
'            ' DNL was outside limit (too high)
'            testStatusDNL = logTestFail
'            testStatus = logTestFail
'            DNLFlag = parmHigh
'        End If
'        If (worstDNL < -DnlLimit) And testDnl Then
'            ' DNL was outside limit (too low)
'            testStatusDNL = logTestFail
'            testStatus = logTestFail
'            DNLFlag = parmLow
'        End If
'
'        If (worstINL > InlLimit) And testInl Then
'            ' INL was outside limit (too high)
'            testStatusINL = logTestFail
'            testStatus = logTestFail
'            INLFlag = parmHigh
'        End If
'        If (worstINL < -InlLimit) And testInl Then
'            ' INL was outside limit (too low)
'            testStatusINL = logTestFail
'            testStatus = logTestFail
'            INLFlag = parmLow
'        End If
'
'        If (NumMissingCodes > MissingCodesLimit) And testMC Then
'            ' Number of missing codes exceeded test limit
'            testStatusMC = logTestFail
'            testStatus = logTestFail
'            MCFlag = parmHigh
'        End If
'
'        If (worstAbsErr > AbsErrorLimit) And testAbsErr Then
'            ' Absolute Error was outside limit (too high)
'            testStatusAbsErr = logTestFail
'            testStatus = logTestFail
'            AbsErrFlag = parmHigh
'        End If
'        If (worstAbsErr < -AbsErrorLimit) And testAbsErr Then
'            ' Absolute Error was outside limit (too low)
'            testStatusAbsErr = logTestFail
'            testStatus = logTestFail
'            AbsErrFlag = parmLow
'        End If
'
'        If (sparkleCount > sparkleLimit) And testSparkle Then
'            ' Sparkle error detected
'            testStatusSparkle = logTestFail
'            testStatus = logTestFail
'            sparkleFlag = parmHigh
'        End If
'
'       Dim vdd As Double
'       vdd = TheHdw.DPS.pins("vdd").ForceValue(dpsPrimaryVoltage)
'
'        ' ***************   DATALOG OUTPUT   *****************
'        ' Sending results to the datalog.
'        ' NOTES:
'        ' - Instead of pin name, the test parameter is specified
'        '   eg: instead of "ra0", the datalog will show "DNL"
'        ' - Instead of channel number, the worst-case code
'        '   (where applicable) is specified.
'
'
'        If testOffset Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumOff, testStatusOff, OffFlag, "Offset", 0, _
'                -OffLimit, OffsetError, OffLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testGain Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumGain, testStatusGain, GainFlag, "Gain", numCodes - 1, _
'                -GainErrorLimit, GainError, GainErrorLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testDnl Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumDNL, testStatusDNL, DNLFlag, "DNL", worstDNLcode, _
'                -DnlLimit, worstDNL, DnlLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testInl Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumINL, testStatusINL, INLFlag, "INL", worstINLcode, _
'                -InlLimit, worstINL, InlLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testMC Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumMC, testStatusMC, MCFlag, "MC", lastMC, _
'                MissingCodesLimit, NumMissingCodes, MissingCodesLimit, _
'                unitNone, vdd, unitVolt, 0)
'        End If
'
'        If testAbsErr Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumAbsErr, testStatusAbsErr, AbsErrFlag, "AbsErr", worstAbsErrCode, _
'                -AbsErrorLimit, worstAbsErr, AbsErrorLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testSparkle Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumSparkle, testStatusSparkle, sparkleFlag, "Sprkl", lastSparkle, _
'                sparkleLimit, CDbl(sparkleCount), sparkleLimit, _
'                unitNone, vdd, unitVolt, 0)
'        End If
'
'
'
'        ' Setting device pass/fail status.
'
'        ' Report Status
'        If testStatus <> logTestPass Then
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'        Else
'            TheExec.Sites.Site(thisSite).TestResult = sitePass
'        End If
'
'NextSite:
'      loopstatus = TheExec.Sites.SelectNext(loopstatus)
'    Wend
'    ' ------------------   End Loop through sites ----------------------
'
'    ' *** DEBUG MODE ? ***
'    If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        adc_db_tOverall = TheExec.Timer(adc_db_tOverallRef)
'
'        ' Print out all timer results:
'        TheExec.DataLog.WriteComment ("::ADC_DB::VB_get_args_time=" & Format(adc_db_tGetArgs, "0.000000") & "s;")
'        TheExec.DataLog.WriteComment ("::ADC_DB::VB_overall_exec_time=" & Format(adc_db_tOverall, "0.000000") & "s;")
'    End If
'
'exitFunction:
'
'    AdcLin_analysis = 0
End Function


' Generates and registers a CTO Voltage array for Partial Code ADC Linearity testing
Public Function AdcLin_partialCodeCtoVoltageArray(name As String, codeSegments() As Integer, hpc As Integer, vrefp As Double, vrefn As Double, trnPnt As Double, bits As Integer)

    Dim voltages() As Double
    Dim numCodeSegments As Integer
    ' Calculate the number of elements in the voltage array:
    numCodeSegments = UBound(codeSegments) + 1
    
    Dim seg As Integer
    Dim numCodes As Integer
    
    numCodes = 0
    For seg = 0 To numCodeSegments - 1
        numCodes = numCodes + codeSegments(seg, 1) - codeSegments(seg, 0) + 1
    Next seg
    
    Dim numHits As Long
    numHits = numCodes * hpc
    
    ' Dimension array to hold all voltages
    ReDim voltages(0 To numHits - 1)
    
    
    Dim codeWidth As Double ' in voltage
    Dim vstep As Double
    codeWidth = (vrefp - vrefn) / 2 ^ bits
    vstep = codeWidth / hpc
    
    ' Offset to apply to each voltage, based on vrefn and trnPnt
    Dim off As Double       ' in voltage
    off = vrefn + (trnPnt - 1) * codeWidth
    
    ' Fill in the voltages:
    Dim codes As Integer
    Dim v As Double
    Dim i As Integer
    Dim vindex As Long
    Dim code1 As Double
    For seg = 0 To numCodeSegments - 1
        v = codeSegments(seg, 0) * codeWidth + off
        codes = codeSegments(seg, 1) - codeSegments(seg, 0) + 1
        For i = 0 To codes * hpc - 1
            voltages(vindex) = v
            v = v + vstep
            vindex = vindex + 1
        Next i
    Next seg
        
    ' Register CTO Source Array:
    Call ctosupport.tl_DestroyArray(name)
    Call ctosupport.tl_RegisterArrayDouble(voltages, name)

End Function

' *****************************************************************************
' FUNCTION:    AdcLin_partialCodeAnalysis
'
' Analyses the results collected from the DUT, computes requested parameters
' such as offset, gain, DNL, INL, absolute error and missing codes, determines
' pass/fail conditions and sends the results to the datalog.
'      - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' REQUIREMENTS:
' * This function requires the current test instance to be running from the
'   IG-XL ADC Template.  This function collects most of the user parameters
'   from what was filled into the template.
'   It must use a CTO voltage array (instead of the template ramp) method
'   for specifying the input voltages.
'
' * This module requires the existance of module AdcLin_v2_Specific which
'   performs all the DUT/pattern/implementation-specific actions to collect
'   the A/D results from the DUT.
'
' * This function is NOT an interpose function and should not be specified
'   directly as an interpose function in the test template.  Instead, the test
'   template must call the end-of body function the AdcLin_v2_specific module,
'   which in turn calls this function passing the required parameters.
'
' * The AdcLin_v2_Specific module must define and initialize the following
'   public/global variables
'
'   Sparkle Options:
'   Note: The template form has no input fields for sparkle options,
'         so the options are defined here and values can be assiged
'         on the startOfBodyIF
'      adc_testSparkle
'      adc_sparkleSensitivity
'      adc_sparkleLimit
'
'   Debugging Variables must be defined in this module:
'      ADC_DB_MODE                   Bitmask that enables debug modes
'      ADC_DB_printResultsPerCode    Debug mode Bit flag
'      ADC_DB_printHits              Debug mode Bit flag
'      ADC_DB_vbCodeTime             Debug mode Bit flag
'      adc_db_tmpStr
'      adc_db_tOverall              ' Overall time in analysis function
'      adc_db_tOverallRef
'      adc_db_tGetArgs              ' time taken to get template arguments
'      adc_db_tGetArgsRef
'
'    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' PARAMETERS:
'   ADresults()        Input: 2D Array, whose dimensions are given by
'                      other input parameters ADresults_Rows and ADresults_Cols.
'                      It contains the digital result for every conversion made
'                      during the test.
'
'   ADresults_Rows     Input: Specifies the number of rows in the array.
'                      This is the first dimension.
'                      The number of rows MUST match the number of SITES
'                      in the current channelmap.
'
'   ADresults_Cols     Input: Specifies the number of columns in the array.
'                      This is the second dimension.
'                      This must be larger or equal to the expected total number
'                      hits/conversions made during the test.
'
'   codes()            Input: 1D array with a list of the codes to be analyzed for
'                      this partial-code test.
'    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' *****************************************************************************
Public Function AdcLin_partialCodeAnalysis(ByRef ADresults() As Integer, _
                                ByVal ADresults_Rows As Long, _
                                ByVal ADresults_Cols As Integer, _
                                ByRef codes() As Integer)
Stop
''20170216 - ty commented out due to many vars not existing
'
'    ' These variables and constants are used to access the instance
'    ' arguments when needed by User Review Functions or Interpose
'    ' Functions.  These declarations may be made with Function
'    ' scope, or with Module scope if needed by more than one Function.
'    ' These declarations should be copied from the template module,
'    ' in this case CtoAdc_T in the Template.xla VBA project.
'    ' This technique may be used with any template, noting that each
'    ' template has a unique set of arguments which can be copied from
'    ' the appropriate template module.
'    ' Variables to hold the instance argument values
'    Dim Arg_DcCategory As String, Arg_DcSelector As String, _
'    Arg_AcCategory As String, Arg_AcSelector As String, _
'    Arg_Timing As String, Arg_Edgeset As String, _
'    Arg_Levels As String, Arg_Bits As String, _
'    Arg_Pinlist As String, Arg_AnalogIn As String, _
'    Arg_VrefLVal As String, Arg_VrefHVal As String, _
'    Arg_HistoSave As String, Arg_CalRef As String, _
'    Arg_Pattern As String, Arg_AinData As String, _
'    Arg_DifErr As String, Arg_IntErr As String, _
'    Arg_OffErr As String, Arg_GainErr As String, _
'    Arg_MissErr As String, Arg_Hits As String, _
'    Arg_VEnd As String, Arg_VStart As String, _
'    Arg_SaveResults As String, Arg_CtoRange As String, _
'    Arg_Serial As String, Arg_PreconditionPat As String, _
'    Arg_PcpStartLabel As String, Arg_PcpStopLabel As String, _
'    Arg_DriverLO  As String, Arg_DriverHI  As String, _
'    Arg_DriverZ As String, Arg_FloatPins As String
'    Dim Arg_StartOfBodyF As String, Arg_PrePatF As String, _
'    Arg_PreTestF As String, Arg_PostTestF As String, _
'    Arg_PostPatF As String, Arg_EndOfBodyF As String, _
'    Arg_StartOfBodyFInput As String, Arg_PrePatFInput As String, _
'    Arg_PreTestFInput As String, Arg_PostTestFInput As String, _
'    Arg_PostPatFInput As String, Arg_EndOfBodyFInput As String, _
'    Arg_RelayMode As String, Arg_UserReviewF As String, _
'    Arg_TestType As String, Arg_PcpCheckPatGen As String, _
'    Arg_ReviewType As String, Arg_DataPoints As String, _
'    Arg_LimitType As String, Arg_TransitionPoint As String, _
'    Arg_NormalizationMethod As String, Arg_AbsErr As String, _
'    Arg_Monotonicity As String, Arg_Util1 As String, _
'    Arg_Util0 As String, Arg_VrefLPin As String, _
'    Arg_VrefHPin As String, Arg_CheckBox As String, _
'    Arg_LsbBeyond As String, Arg_RefOrEnd As String
'
'    ' Constants required to retrieve each argument from the argument array
'
'    Const ARGNUM_BITS = 0
'    Const ARGNUM_PINLIST = 1
'    Const ARGNUM_ANALOG = 2
'    Const ARGNUM_VREFLVAL = 3
'    Const ARGNUM_VREFHVAL = 4
'
'    Const ARGNUM_HISTO = 5
'    Const ARGNUM_CALREF = 6
'    Const ARGNUM_PATTERN = 7
'    Const ARGNUM_AINCODE = 8
'    Const ARGNUM_DIFERR = 9
'    Const ARGNUM_INTERR = 10
'    Const ARGNUM_OFFERR = 11
'    Const ARGNUM_GAINERR = 12
'    Const ARGNUM_MISSERR = 13
'    Const ARGNUM_HITS = 14
'    Const ARGNUM_VENDVAL = 15
'    Const ARGNUM_VSTARTVAL = 16
'    Const ARGNUM_DATAPOINTS = 17
'    Const ARGNUM_SAVERESULTS = 18
'
'    Const ARGNUM_CTORANGE = 19
'    Const ARGNUM_SERIAL = 20
'    Const ARGNUM_PRECONDITIONPAT = 21
'    Const ARGNUM_PCPSTARTLABEL = 22
'    Const ARGNUM_PCPSTOPLABEL = 23
'    Const ARGNUM_DRIVERLO = 24
'    Const ARGNUM_DRIVERHI = 25
'    Const ARGNUM_DRIVERZ = 26
'    Const ARGNUM_FLOATPINS = 27
'    Const ARGNUM_STARTOFBODYF = 28
'    Const ARGNUM_PREPATF = 29
'    Const ARGNUM_PRETESTF = 30
'    Const ARGNUM_POSTTESTF = 31
'
'    Const ARGNUM_POSTPATF = 32
'    Const ARGNUM_ENDOFBODYF = 33
'    Const ARGNUM_STARTOFBODYFINPUT = 34
'    Const ARGNUM_PREPATFINPUT = 35
'    Const ARGNUM_PRETESTFINPUT = 36
'    Const ARGNUM_POSTTESTFINPUT = 37
'    Const ARGNUM_POSTPATFINPUT = 38
'    Const ARGNUM_ENDOFBODYFINPUT = 39
'    Const ARGNUM_RELAYMODE = 40
'    Const ARGNUM_USERREVIEWF = 41
'    Const ARGNUM_TESTTYPE = 42
'    Const ARGNUM_PCPCHECKPATGEN = 43
'
'    Const ARGNUM_REVIEWTYPE = 44
'    Const ARGNUM_LIMITTYPE = 45
'    Const ARGNUM_TRANSITIONPOINT = 46
'    Const ARGNUM_NORMMETHOD = 47
'    Const ARGNUM_ABSERR = 48
'    Const ARGNUM_MONOTONICITY = 49
'    Const ARGNUM_UTIL1 = 50
'    Const ARGNUM_UTIL0 = 51
'    Const ARGNUM_VREFLPIN = 52
'    Const ARGNUM_VREFHPIN = 53
'    Const ARGNUM_CHECKBOX = 54
'    Const ARGNUM_LSBBEYOND = 55
'    Const ARGNUM_REFOREND = 56
'
'    Const ARGNUM_MAXARG = ARGNUM_REFOREND
'
'    '' Note: These additional definitions are in Template.xla, CtoSupport module
'    ''    Public Const ADC_BIT_DIFF = 0   'this is used to define the bit position in the
'    ''    Public Const ADC_BIT_INT = 1    '   Arg_CheckBox which denotes the enabling of
'    ''    Public Const ADC_BIT_OFF = 2    '   the corresponding test.
'    ''    Public Const ADC_BIT_GAIN = 3
'    ''    Public Const ADC_BIT_MISS = 4
'    ''    Public Const ADC_BIT_ABS = 5
'
'
'    Dim inputv() As Double                ' the ADC stimuli (applied voltage array)
'    Dim sparkleCount As Long              ' number of sparkle errors detected
'
'    Dim dnlErr As Double                  ' linearity error for each code
'    Dim numCodes As Long                  ' number of codes for the ADC DUT
'
'    Dim DNLFlag As Long
'    Dim INLFlag As Long
'    Dim MCFlag As Long
'    Dim OffFlag As Long
'    Dim GainFlag As Long
'    Dim AbsErrFlag As Long
'    Dim sparkleFlag As Long
'
'    Dim thisSite As Long
'
'    Dim ReturnStatus As Long
'    Dim CodeReturned As Long
'    Dim prevCodeReturned As Long
'    Dim ExpectedNumDataItems As Long
'    Dim NumDataItems As Long
'    Dim ExpectedNumSites As Integer
'
'    Dim TestNumDNL As Long
'    Dim TestNumINL As Long
'    Dim TestNumMC As Long
'    Dim TestNumOff As Long
'    Dim TestNumGain As Long
'    Dim TestNumAbsErr As Long
'    Dim TestNumSparkle As Long
'
'
'    Dim testStatus As Long               ' overall pass/fail result
'    Dim testStatusINL As Long            ' INL pass/fail result
'    Dim testStatusDNL As Long            ' DNL pass/fail result
'    Dim testStatusMC As Long             ' Missing Codes pass/fail result
'    Dim testStatusOff As Long
'    Dim testStatusGain As Long
'    Dim testStatusAbsErr As Long
'    Dim testStatusSparkle As Long
'
'    Dim DnlLimit As Double
'    Dim InlLimit As Double
'    Dim MissingCodesLimit As Double
'    Dim OffLimit As Double
'    Dim GainErrorLimit As Double
'    Dim AbsErrorLimit As Double
'    Dim sparkleLimit As Double
'    Dim sparkleSensitivity As Integer   ' min size of sparkle to be detected
'    Dim lastSparkle As Long
'
'
'    Dim worstDNL As Double
'    Dim worstDNLcode As Long
'    Dim worstINL As Double
'    Dim worstINLcode As Long
'    Dim codei As Long                    ' general code index variable
'    Dim i As Long
'    Dim hitIndex As Long
'    Dim ArgStr() As String               ' array of template argument strings
'
'    Dim intErr As Double
'    Dim NumMissingCodes As Double
'    Dim lastMC As Long
'    Dim OffsetError As Double
'    Dim GainError As Double
'    Dim worstAbsErrCode As Long
'
'    Dim vrefH As Double
'    Dim vrefL As Double
'
'    Dim worstAbsErr As Double
'    Dim absErr As Double
'    Dim TransitionPoint As Double
'    Dim loopstatus As Long
'
'    ' What tests to actually run:
'    Dim checkboxData As Integer
'    Dim testInl As Boolean
'    Dim testDnl As Boolean
'    Dim testMC As Boolean
'    Dim testOffset As Boolean
'    Dim testGain As Boolean
'    Dim testAbsErr As Boolean
'
'
'    Dim ideal As Double
'    Dim firstTrn As Double
'    Dim lastTrn As Double
'    Dim offsetErrorV As Double
'    Dim fullScaleErrorV As Double
'    Dim gainErrorV As Double
'    Dim idealCodeWidth As Double
'    Dim avgCodeWidth As Double
'    Dim trnFr As Double
'    Dim trnTo As Double
'    Dim adjTrnFr As Double
'    Dim codeMissing As Boolean
'
'
'    Dim testSparkle As Boolean
'
'
'    ' *** DEBUG MODE ? ***
'    If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        adc_db_tOverallRef = TheExec.Timer
'        adc_db_tGetArgsRef = TheExec.Timer
'    End If
'
'
'    ' Getting the test instance arguments.
'    ' First, get the arguments for the current instance
'    Call TheExec.DataManager.GetArgumentList(ArgStr, ARGNUM_MAXARG)
'
'
'    ' Now pick out the arguments we need,
'    ' and place them in convenient variables
'    Arg_Hits = ArgStr(ARGNUM_HITS)          ' number of hits per code
'    Arg_Bits = ArgStr(ARGNUM_BITS)          ' number of ADC output bits
'    Arg_AnalogIn = ArgStr(ARGNUM_ANALOG)    ' Name of Analog Pin
'    Arg_DifErr = ArgStr(ARGNUM_DIFERR)      ' DNL Limit
'    Arg_IntErr = ArgStr(ARGNUM_INTERR)      ' INL Limit
'    Arg_LsbBeyond = ArgStr(ARGNUM_LSBBEYOND)    ' #LSB's beyond range
'    Arg_OffErr = ArgStr(ARGNUM_OFFERR)      ' Offset Error Limit
'    Arg_MissErr = ArgStr(ARGNUM_MISSERR)    ' Num of Missing Codes Limit
'    Arg_VEnd = ArgStr(ARGNUM_VENDVAL)       ' Programmed Ramp End Voltage
'    Arg_VStart = ArgStr(ARGNUM_VSTARTVAL)   ' Programmed Ramp Start Voltage
'    Arg_VrefLVal = ArgStr(ARGNUM_VREFLVAL)
'    Arg_VrefHVal = ArgStr(ARGNUM_VREFHVAL)
'    Arg_GainErr = ArgStr(ARGNUM_GAINERR)
'    Arg_AbsErr = ArgStr(ARGNUM_ABSERR)
'    Arg_TransitionPoint = ArgStr(ARGNUM_TRANSITIONPOINT) ' 0-1 transition in LSB
'    Arg_CheckBox = ArgStr(ARGNUM_CHECKBOX)
'    Arg_AinData = ArgStr(ARGNUM_AINCODE)    ' Name of applied CTO voltage array
'    DnlLimit = CDbl(Arg_DifErr)
'    InlLimit = CDbl(Arg_IntErr)
'    MissingCodesLimit = CDbl(Arg_MissErr)
'    OffLimit = CDbl(Arg_OffErr)
'    vrefL = CDbl(Arg_VrefLVal)
'    vrefH = CDbl(Arg_VrefHVal)
'    GainErrorLimit = CDbl(Arg_GainErr)
'    AbsErrorLimit = CDbl(Arg_AbsErr)
'
'    ' Tansition Point
'    ' Template form has two options:
'    '   "0.5 LSB"  -- Arg_TransitionPoint = 0
'    '   "  1 LSB"  -- Arg_TransitionPoint = 0
'    If Arg_TransitionPoint = 0 Then
'        TransitionPoint = 0.5
'    Else
'        TransitionPoint = 1#
'    End If
'
'    sparkleLimit = adc_sparkleLimit
'    sparkleSensitivity = adc_sparkleSensitivity
'
'    ' Determine what tests to perform/check:
'    checkboxData = CInt(Val(Arg_CheckBox))
'    testMC = True
'    testGain = True
'    testDnl = True
'    testInl = True
'    testOffset = True
'    testAbsErr = True
'    testSparkle = adc_testSparkle
'    If (checkboxData Or (2 ^ ADC_BIT_MISS)) <> checkboxData Then testMC = False
'    If (checkboxData Or (2 ^ ADC_BIT_GAIN)) <> checkboxData Then testGain = False
'    If (checkboxData Or (2 ^ ADC_BIT_DIFF)) <> checkboxData Then testDnl = False
'    If (checkboxData Or (2 ^ ADC_BIT_INT)) <> checkboxData Then testInl = False
'    If (checkboxData Or (2 ^ ADC_BIT_OFF)) <> checkboxData Then testOffset = False
'    If (checkboxData Or (2 ^ ADC_BIT_ABS)) <> checkboxData Then testAbsErr = False
'
'    ' Getting the applied voltage values array.
'    ' Now get the array of values that the Template used as stimulus
'    ' for the DUT.
'    ReturnStatus = tl_GetArrayDouble(Arg_AinData, inputv)
'
'    If ReturnStatus <> TL_SUCCESS Then
'        Call TheExec.ErrorLogMessage(Arg_AinData + _
'                ": Error trying to get Applied Voltage Array.")
'        Call TheExec.ErrorReport
'        Exit Function
'    End If
'
'
'    ' *** DEBUG MODE ? ***
'    If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        adc_db_tGetArgs = TheExec.Timer(adc_db_tGetArgsRef)
'    End If
'
'
'    ' Calculate how many codes this ADC has
'    numCodes = 2 ^ CInt(Arg_Bits)
'
'    ' Ideal code width (or LSB size)
'    idealCodeWidth = (vrefH - vrefL) / numCodes
'
'    ' !!!DA TODO: consider using redim preserve with all these.  Would it be faster?
'    Dim codeIsIncluded() As Boolean     ' true of code is in the codes array
'    Dim codeFirstSeen() As Long         ' index into inputv
'    Dim codeLastSeen() As Long          ' index into inputv
'
'    ReDim codeFirstSeen(numCodes - 1)
'    ReDim codeLastSeen(numCodes - 1)
'
'    ' Number of codes included in the analysis:
'    Dim numCodesIncluded As Integer
'    numCodesIncluded = UBound(codes) + 1
'
'    ' Initialize codeIsIncluded array:
'    ReDim codeIsIncluded(numCodes - 1)
'    For codei = 0 To numCodes - 1
'        codeIsIncluded(codei) = False
'    Next codei
'    For i = 0 To numCodesIncluded - 1
'        codei = codes(i)
'        codeIsIncluded(codei) = True
'    Next i
'
'    ' Expected number of data items in the ADResults array per site:
'    ExpectedNumDataItems = UBound(inputv) + 1
'    ExpectedNumSites = TheExec.Sites.ExistingCount
'
'    ' -----------------------  Loop through sites ----------------------
'    loopstatus = TheExec.Sites.SelectFirst
'    While loopstatus <> loopDone
'    thisSite = TheExec.Sites.SelectedSite
'
'        ' ********* Validate sizes of arrays ******************
'
'        ' see how many data items this site returned
'        If ADresults_Cols < (ExpectedNumDataItems) Then
'            ' Fail misserably:
'            Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                ": Not enough data points returned for site " + _
'                    CStr(thisSite))
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'            GoTo NextSite
'        End If
'        NumDataItems = ExpectedNumDataItems 'only want to look at valid data items
'
'        ' Make sure number of sites returned makes sense:
'        If ADresults_Rows <> ExpectedNumSites Then
'            ' Fail misserably:
'            Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                ": Number of sites returned with results array (adresults) is invalid.")
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'            GoTo NextSite
'        End If
'
'
'        ' *********** Clear data structures **************************************
'        '!!!DA TODO: how long does this take? -- maybe even add it to the debug timing stuff
'        sparkleCount = 0
'        For codei = 0 To numCodes - 1
'            codeFirstSeen(codei) = -1
'            codeLastSeen(codei) = -1
'        Next codei
'
'        prevCodeReturned = ADresults(thisSite, 0)
'
'        ' *********** Go through ADResults & SPARKLE DETECTION ******************
'
'        For hitIndex = 0 To NumDataItems - 1
'
'            CodeReturned = ADresults(thisSite, hitIndex)
'            If CodeReturned > numCodes - 1 Then
'                ' an illegal code was returned!
'                Call TheExec.DataLog.WriteComment(Arg_AinData + _
'                    ": Illegal code returned by site " + _
'                        CStr(thisSite))
'                TheExec.Sites.Site(thisSite).TestResult = siteFail
'                GoTo NextSite
'            End If
'
'            ' Is this the first time we see this code?
'            If (codeFirstSeen(CodeReturned) < 0) Then
'                codeFirstSeen(CodeReturned) = hitIndex
'            End If
'
'            ' This is the latest time we have seen this code:
'            codeLastSeen(CodeReturned) = hitIndex
'
'            ' Sparkle detection:
'            If Abs(CodeReturned - prevCodeReturned) >= sparkleSensitivity Then
'                ' We may have a sparkle event, but only if it is from or to
'                ' one of the codes included in the analysis:
'                If (codeIsIncluded(CodeReturned) Or codeIsIncluded(prevCodeReturned)) Then
'                    sparkleCount = sparkleCount + 1
'                    lastSparkle = CodeReturned
'                End If
'            End If
'
'            prevCodeReturned = CodeReturned
'
'        Next hitIndex
'
'
'        ' *** DEBUG MODE ? ***
'        If (ADC_DB_MODE And (ADC_DB_printResultsPerCode Or ADC_DB_printHits)) <> 0 Then
'            adc_db_tmpStr = "::ADC_DB - Site " & thisSite & "::"
'            TheExec.DataLog.WriteComment (adc_db_tmpStr)
'            adc_db_tmpStr = ""
'            If (ADC_DB_MODE And ADC_DB_printHits) <> 0 Then
'               For hitIndex = 0 To NumDataItems - 1
'                    ' print 20 codes per line, separated by ':'
'                    If hitIndex Mod 20 = 0 Then
'                        TheExec.DataLog.WriteComment (adc_db_tmpStr)
'                        adc_db_tmpStr = "::ADC_DB::"
'                    End If
'                    adc_db_tmpStr = adc_db_tmpStr & ADresults(thisSite, hitIndex) & ":"
'               Next hitIndex
'               TheExec.DataLog.WriteComment (adc_db_tmpStr)
'            End If
'        End If
'
'
'        ' **************** OFFSET ERROR Calculation ***********************
'
'        ' Ideal vs actual voltage for the first transition:
'        ideal = idealCodeWidth * TransitionPoint + vrefL
'
'        ' Find the transition point from 0 to 1:
'        ' if Needed codes are missing: Report huge error!
'        If (codeFirstSeen(0) < 0) Or (codeFirstSeen(1) < 1) Then
'            firstTrn = -numCodes * idealCodeWidth
'        Else
'            firstTrn = (inputv(codeLastSeen(0)) + inputv(codeFirstSeen(1))) / 2
'        End If
'
'        offsetErrorV = firstTrn - ideal
'        OffsetError = offsetErrorV / idealCodeWidth
'
'        ' **************** GAIN ERROR Calculation ***********************
'
'        ' Ideal vs actual voltage for the last transition:
'        ideal = vrefH - (1 + 1 - TransitionPoint) * idealCodeWidth
'
'        ' if needed codes are missing, report huge error:
'        If (codeFirstSeen(numCodes - 2) < 0) Or (codeFirstSeen(numCodes - 1) < 0) Then
'            lastTrn = 2 * numCodes * idealCodeWidth
'        Else
'            lastTrn = (inputv(codeLastSeen(numCodes - 2)) + inputv(codeFirstSeen(numCodes - 1))) / 2
'        End If
'
'        fullScaleErrorV = lastTrn - ideal
'
'        ' Gain error is compensated for offset:
'        gainErrorV = -fullScaleErrorV + offsetErrorV
'        GainError = gainErrorV / idealCodeWidth
'        ' Note: a positive gain error means that the slope of the
'        '       Digital output vs Analog input line is greater than ideal.
'
'
'        ' ************** Average Code Width ****************************
'
'        ' Determine the average code width.
'        ' For a real converter, offset and gain make
'        ' the average code width different from ideal.
'
'        avgCodeWidth = (lastTrn - firstTrn) / (numCodes - 2)
'
'
'        ' ********** DNL, INL, MISSING & ABSOLUTE ERROR CALCULATIONS PER CODE *********
'        worstDNL = 0
'        worstINL = 0
'        worstDNLcode = 0
'        worstINLcode = 0
'        NumMissingCodes = 0
'        lastMC = 0
'        worstAbsErr = 0
'        worstAbsErrCode = 0
'        codeMissing = False
'        trnTo = 0
'        trnFr = 0
'
'        For i = 0 To numCodesIncluded - 1
'
'            codei = codes(i)
'
'            ' *** Missing Codes ***
'            If codeFirstSeen(codei) < 0 Then
'                codeMissing = True
'            Else
'                codeMissing = False
'            End If
'
'            If codei = 0 Then
'                dnlErr = 0
'                absErr = OffsetError
'                intErr = 0
'
'            ElseIf codei = numCodes - 1 Then ' last code
'                dnlErr = 0
'                absErr = fullScaleErrorV / idealCodeWidth
'                intErr = 0
'
'            Else ' all other codes
'
'                ' Get transition points for this code:
'                If codeMissing Or (codeFirstSeen(codei - 1) < 0) Or (codeFirstSeen(codei + 1) < 0) Then
'                    ' Needed codes are missing: report huge error!
'                    dnlErr = numCodes + 1
'                    absErr = numCodes + 1
'                    intErr = numCodes + 1
'
'                Else
'                    trnTo = (inputv(codeLastSeen(codei - 1)) + inputv(codeFirstSeen(codei))) / 2
'                    trnFr = (inputv(codeLastSeen(codei)) + inputv(codeFirstSeen(codei + 1))) / 2
'
'
'                    ' *** DNL ***
'                    ' Compute the difference between the code width and the
'                    ' average code width.
'                    ' NOTE: Comparing against the average code width
'                    '       the effect of compensating for gain error.
'                    dnlErr = ((trnFr - trnTo) - avgCodeWidth) / idealCodeWidth
'
'
'                    ' *** Absolute Error ***
'                    ' Find where the ideal transition from this code to the next should be:
'                    ideal = (codei + TransitionPoint) * idealCodeWidth + vrefL
'                    ' Compare against the actual transition:
'                    absErr = (trnFr - ideal) / idealCodeWidth
'
'                    ' *** INL ***
'                    ' Compare acutal transition vs ideal transition, but adjust for
'                    ' offset and gain errors:
'                    adjTrnFr = trnFr - offsetErrorV - codei * (avgCodeWidth - idealCodeWidth)
'                    intErr = (adjTrnFr - ideal) / idealCodeWidth
'
'
'                    ' Note: Dividing by "ideal" code width is done to convert
'                    '       from V to LSB units
'
'                End If
'
'            End If
'
'            ' Track worst-case errors:
'            If Abs(dnlErr) > Abs(worstDNL) Then
'                ' a new largest DNL error was observed
'                worstDNL = dnlErr
'                worstDNLcode = codei
'            End If
'
'            If Abs(intErr) > Abs(worstINL) Then
'                ' a new maximum INL error was observed
'                worstINL = intErr
'                worstINLcode = codei
'            End If
'
'            If Abs(absErr) > Abs(worstAbsErr) Then
'                ' a new maximum ABS error was observed
'                worstAbsErr = absErr
'                worstAbsErrCode = codei
'            End If
'
'            If codeMissing Then
'                NumMissingCodes = NumMissingCodes + 1
'                lastMC = codei
'            End If
'
'            ' *** DEBUG MODE ? ***
'            If (ADC_DB_MODE And ADC_DB_printResultsPerCode) <> 0 Then
'                TheExec.DataLog.WriteComment ("::ADC_DB::DNL[" & codei & "]=" & Format(dnlErr, "0.000000") & ";INL[" & codei & "]=" & Format(intErr, "0.000000") & ";AbsErr[" & codei & "]=" & Format(absErr, "0.000000") & ";codeFirstSeen[" & codei & "]=" & Format(codeFirstSeen(codei), "0000") & ";codeLastSeen[" & codei & "]=" & Format(codeLastSeen(codei), "0000") & ";")
'            End If
'
'        Next i
'
'        ' *** DEBUG MODE ? ***
'        If (ADC_DB_MODE And ADC_DB_printResultsPerCode) <> 0 Then
'            TheExec.DataLog.WriteComment ("::ADC_DB::Offset=" & OffsetError & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::GainErr=" & GainError & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::DNLErr=" & worstDNL & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstDNLcode=" & worstDNLcode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::INLErr=" & worstINL & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstINLcode=" & worstINLcode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::MissCodes=" & NumMissingCodes & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::AbsErr=" & worstAbsErr & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::worstAbsErrCode=" & worstAbsErrCode & ";")
'            TheExec.DataLog.WriteComment ("::ADC_DB::Sparkle=" & sparkleCount & ";")
'        End If
'
'
'        ' ***********  TEST PASS/FAIL CONDITIONS ***************
'
'        testStatus = logTestPass
'        testStatusOff = logTestPass
'        testStatusGain = logTestPass
'        testStatusDNL = logTestPass
'        testStatusINL = logTestPass
'        testStatusMC = logTestPass
'        testStatusAbsErr = logTestPass
'        testStatusSparkle = logTestPass
'
'        ' check results against limits
'
'        If (OffsetError > OffLimit) And testOffset Then
'            ' Offset Error was outside limit (too high)
'            testStatusOff = logTestFail
'            testStatus = logTestFail
'            OffFlag = parmHigh
'        End If
'        If (OffsetError < -OffLimit) And testOffset Then
'            ' Offset Error was outside limit (too low)
'            testStatusOff = logTestFail
'            testStatus = logTestFail
'            OffFlag = parmLow
'        End If
'
'        If (GainError > GainErrorLimit) And testGain Then
'            ' Gain Error was outside limit (too high)
'            testStatusGain = logTestFail
'            testStatus = logTestFail
'            GainFlag = parmHigh
'        End If
'        If (GainError < -GainErrorLimit) And testGain Then
'            ' Gain Error was outside limit (too low)
'            testStatusGain = logTestFail
'            testStatus = logTestFail
'            GainFlag = parmLow
'        End If
'
'        If (worstDNL > DnlLimit) And testDnl Then
'            ' DNL was outside limit (too high)
'            testStatusDNL = logTestFail
'            testStatus = logTestFail
'            DNLFlag = parmHigh
'        End If
'        If (worstDNL < -DnlLimit) And testDnl Then
'            ' DNL was outside limit (too low)
'            testStatusDNL = logTestFail
'            testStatus = logTestFail
'            DNLFlag = parmLow
'        End If
'
'        If (worstINL > InlLimit) And testInl Then
'            ' INL was outside limit (too high)
'            testStatusINL = logTestFail
'            testStatus = logTestFail
'            INLFlag = parmHigh
'        End If
'        If (worstINL < -InlLimit) And testInl Then
'            ' INL was outside limit (too low)
'            testStatusINL = logTestFail
'            testStatus = logTestFail
'            INLFlag = parmLow
'        End If
'
'        If (NumMissingCodes > MissingCodesLimit) And testMC Then
'            ' Number of missing codes exceeded test limit
'            testStatusMC = logTestFail
'            testStatus = logTestFail
'            MCFlag = parmHigh
'        End If
'
'        If (worstAbsErr > AbsErrorLimit) And testAbsErr Then
'            ' Absolute Error was outside limit (too high)
'            testStatusAbsErr = logTestFail
'            testStatus = logTestFail
'            AbsErrFlag = parmHigh
'        End If
'        If (worstAbsErr < -AbsErrorLimit) And testAbsErr Then
'            ' Absolute Error was outside limit (too low)
'            testStatusAbsErr = logTestFail
'            testStatus = logTestFail
'            AbsErrFlag = parmLow
'        End If
'
'        If (sparkleCount > sparkleLimit) And testSparkle Then
'            ' Sparkle error detected
'            testStatusSparkle = logTestFail
'            testStatus = logTestFail
'            sparkleFlag = parmHigh
'        End If
'
'        ' ***************   DATALOG OUTPUT   *****************
'        ' Sending results to the datalog.
'        ' NOTES:
'        ' - Instead of pin name, the test parameter is specified
'        '   eg: instead of "ra0", the datalog will show "DNL"
'        ' - Instead of channel number, the worst-case code
'        '   (where applicable) is specified.
'
'        Dim vdd As Double
'        vdd = TheHdw.DPS.pins("vdd").ForceValue(dpsPrimaryVoltage)
'
'        ' get and set test numbers
'        TestNumOff = TheExec.Sites.Site(thisSite).testnumber
'        TestNumGain = TheExec.Sites.Site(thisSite).testnumber + 1
'        TestNumDNL = TheExec.Sites.Site(thisSite).testnumber + 2
'        TestNumINL = TheExec.Sites.Site(thisSite).testnumber + 3
'        TestNumMC = TheExec.Sites.Site(thisSite).testnumber + 4
'        TestNumAbsErr = TheExec.Sites.Site(thisSite).testnumber + 5
'        TestNumSparkle = TheExec.Sites.Site(thisSite).testnumber + 6
'
'
'        If testOffset Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumOff, testStatusOff, OffFlag, "Offset", 0, _
'                -OffLimit, OffsetError, OffLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testGain Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumGain, testStatusGain, GainFlag, "Gain", numCodes - 1, _
'                -GainErrorLimit, GainError, GainErrorLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testDnl Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumDNL, testStatusDNL, DNLFlag, "DNL", worstDNLcode, _
'                -DnlLimit, worstDNL, DnlLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testInl Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumINL, testStatusINL, INLFlag, "INL", worstINLcode, _
'                -InlLimit, worstINL, InlLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testMC Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumMC, testStatusMC, MCFlag, "MC", lastMC, _
'                MissingCodesLimit, NumMissingCodes, MissingCodesLimit, _
'                unitNone, vdd, unitVolt, 0)
'        End If
'
'        If testAbsErr Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumAbsErr, testStatusAbsErr, AbsErrFlag, "AbsErr", worstAbsErrCode, _
'                -AbsErrorLimit, worstAbsErr, AbsErrorLimit, unitLSB, vdd, _
'                unitVolt, 0)
'        End If
'
'        If testSparkle Then
'            Call TheExec.DataLog.WriteParametricResult(thisSite, _
'                TestNumSparkle, testStatusSparkle, sparkleFlag, "Sprkl", lastSparkle, _
'                sparkleLimit, CDbl(sparkleCount), sparkleLimit, _
'                unitNone, vdd, unitVolt, 0)
'        End If
'
'
'
'        ' Setting device pass/fail status.
'
'        ' Report Status
'        If testStatus <> logTestPass Then
'            TheExec.Sites.Site(thisSite).TestResult = siteFail
'        Else
'            TheExec.Sites.Site(thisSite).TestResult = sitePass
'        End If
'
'NextSite:
'        loopstatus = TheExec.Sites.SelectNext(loopstatus)
'    Wend
'    ' ------------------   End Loop through sites ----------------------
'
'    ' *** DEBUG MODE ? ***
'    If (ADC_DB_MODE And ADC_DB_vbCodeTime) <> 0 Then
'        adc_db_tOverall = TheExec.Timer(adc_db_tOverallRef)
'
'        ' Print out all timer results:
'        TheExec.DataLog.WriteComment ("::ADC_DB::VB_get_args_time=" & Format(adc_db_tGetArgs, "0.000000") & "s;")
'        TheExec.DataLog.WriteComment ("::ADC_DB::VB_overall_exec_time=" & Format(adc_db_tOverall, "0.000000") & "s;")
'    End If
    
                                
                                
End Function


'!!!DA TODO: Resolve minor discrepancies between calculations in full-code and partial-code
'            functions. eg.:
'            Edge used for ABS / INL
'            dividing by ideal / avg codewidth/hpc








