Attribute VB_Name = "Globals"
' Global Variables
Public OscValues1(31) As Long   'Osc cal value calculated at hot for +2% limit
Public OscValues2(31) As Long   'Osc cal value calculated at hot for -2% limit
Public OscValues3(31) As Long   'Osc cal value 0%

Public sitePos()    As Long     'X-Y Coordinates of die at probe.
Public scribeLot()  As Long     'Lot & scribe number from probe.

Public passcodeVal() As Long


Global AXRFInitialized As Boolean
Global AXRF_Error_Flag As Boolean

Global Initialize_status As Long

Global Dev1 As Long

Global ReferenceTime As Double
Global ElapsedTime As Double
Global TimeLimit As Double

Global count As Long

Global Site_Stat As Long

