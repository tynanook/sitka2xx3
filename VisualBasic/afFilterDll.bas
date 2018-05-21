Attribute VB_Name = "afFilterDll"
' This module contains the information needed for an application
' to use the Filter Analysis Library DLL.
' NB - 'afMeasLibDefs.bas' must be added to the project with this module.
Option Explicit

' =============================================================================================================
' Dll installed in system32 folder
' =============================================================================================================

'**** Type Definitions ****
' afFilterError - defines the different error codes that can be returned
'                 by the functions defined in this DLL.
Public Enum afFilterError
    afFilterInvalidMeasurementLengthSetting = -16384
    afFilterInvalidMeasurementOffsetSetting = -16385
    afFilterInvalidMeasurementSettings = -16386
    afFilterInvalidFilterType = -16387
    afFilterInvalidObjectID = -16388
    afFilterInvalidOSFactor = -16389
    afFilterParameterOutOfRange = -16390
End Enum

' afFilterTraceData - defines the different traces
Public Enum afFilterTraceData
    afFilterPowerVsTime = 0
    afFilterCoeffs = 1
End Enum

' afFilterTypes - defines the types of filter provided by this DLL.
Public Enum afFilterTypes
    afFilterNone = &H0
    afFilterRRC = &H2
    afFilterGaussian = &H5
    afFilterUserDefined = &H3
End Enum


'**** Exported Functions ****

'** Methods **
Declare Function afFilterDll_CreateObject Lib "afFilterDll.dll" (ByRef ptrID As Long) As Long

Declare Function afFilterDll_DestroyObject Lib "afFilterDll.dll" (ByVal nID As Long) As Long

Declare Function afFilterDll_GetTraceDataLength Lib "afFilterDll.dll" (ByVal nID As Long, _
                                                                             ByVal traceType As Long, _
                                                                             ByRef ptrLength As Long) As Long

Declare Function afFilterDll_GetTraceData Lib "afFilterDll.dll" (ByVal nID As Long, _
                                                                       ByVal traceType As Long, _
                                                                       ByRef ptrX As Double, _
                                                                       ByRef ptrY As Double, _
                                                                       ByVal numPoints As Long) As Long
' ptrX must be an array of size numPoints
' ptrY must be an array of size numPoints

Declare Function afFilterDll_GetErrorMsgLength Lib "afFilterDll.dll" (ByVal nID As Long, _
                                                                            ByVal errorCode As Long, _
                                                                            ByRef ptrLength As Long) As Long

Declare Function afFilterDll_GetErrorMsg Lib "afFilterDll.dll" (ByVal nID As Long, _
                                                                      ByVal errorCode As Long, _
                                                                      ByVal ptrBuffer As String, _
                                                                      ByVal bufferSize As Long) As Long ' NB byval on the string

Declare Function afFilterDll_MaxErrorMsgLength_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrLength As Long) As Long

Declare Function afFilterDll_GetVersion Lib "afFilterDll.dll" (ByRef ptrVersion As Long) As Long
'* NOTE - ptrVersion must be a pointer to an array of 4 elements.*'



'* Configuration Methods and Properties *

Declare Function afFilterDll_Blocksize_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal blocksize As Long) As Long
Declare Function afFilterDll_Blocksize_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrBlocksize As Long) As Long

Declare Function afFilterDll_OSFactor_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal oSFactor As Single) As Long
Declare Function afFilterDll_OSFactor_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrOSFactor As Single) As Long

Declare Function afFilterDll_SymbolRate_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal symbolRate As Single) As Long
Declare Function afFilterDll_SymbolRate_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrSymbolRate As Single) As Long

Declare Function afFilterDll_RollOffFactor_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal rollOffFactor As Single) As Long
Declare Function afFilterDll_RollOffFactor_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrRollOffFactor As Single) As Long

Declare Function afFilterDll_MeasurementOffset_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal measurementOffset As Long) As Long
Declare Function afFilterDll_MeasurementOffset_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrMeasurementOffset As Long) As Long

Declare Function afFilterDll_RfLevelCal_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal rfLevelCal As Single) As Long
Declare Function afFilterDll_RfLevelCal_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrRfLevelCal As Single) As Long

Declare Function afFilterDll_FilterType_Set Lib "afFilterDll.dll" (ByVal nID As Long, ByVal type_ As Long) As Long
Declare Function afFilterDll_FilterType_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrType As Long) As Long
'* Apply specified filter *
Declare Function afFilterDll_ApplyFilter Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrDataInI As Single, ByRef ptrDataInQ As Single, ByRef ptrDataOutI As Single, ByRef ptrDataOutQ As Single, ByVal NumSamples As Long, ByVal filterReset As Long) As Long
' ptrDataInI must be an array of size numSamples
' ptrDataInQ must be an array of size numSamples
' ptrDataOutI must be an array of size numSamples
' ptrDataOutQ must be an array of size numSamples
'* Calculate the number of output samples *
Declare Function afFilterDll_GetNumOutputSamples Lib "afFilterDll.dll" (ByVal nID As Long, ByVal inSamps As Long, ByRef ptrOutSamps As Long) As Long

'* Results Methods and Properties *
Declare Function afFilterDll_Results_AveragePower_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrAveragePower As Single) As Long
Declare Function afFilterDll_Results_PeakPower_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrPeakPower As Single) As Long
Declare Function afFilterDll_Results_NumberOfFilterCoeffs_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrNumCoeffs As Long) As Long
Declare Function afFilterDll_Results_GetFilterCoeffs Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrFiltCoeffs As Single, ByVal numCoeffs As Long) As Long
' ptrFiltCoeffs must be an array of size numCoeffs
'* Calculate the minimum number of samples *
Declare Function afFilterDll_Results_MinNumSamples_Get Lib "afFilterDll.dll" (ByVal nID As Long, ByRef ptrMinSamps As Long) As Long
