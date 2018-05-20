Attribute VB_Name = "NI_Clock_Support"
Option Explicit


Public Function ArmClock20M(argc As Long, argv() As String) As Long


Call itl.Raw.NI.Sync("Dev1").ConnectClkTerminals(NISYNC_CONSTS.Clkin, NISYNC_CONSTS.Clk10In) 'set DDS reference to CLKIN from AXRF

Call itl.Raw.NI.Sync("Dev1").ConnectClkTerminals(NISYNC_CONSTS.Dds, NISYNC_CONSTS.Clkout) 'route DDS to CLKOUT

Call itl.Raw.NI.Sync("Dev1").SetBoolean(niSyncProperties_ClkoutGainEnable, "", True) 'set high clk output (2.5vpk into 50 ohms) Note: EVM fails unless clk output high (True) is enabled


Call itl.Raw.NI.Sync("Dev1").SetDouble_2(niSyncProperties_DdsFreq, 20000000#)

TheHdw.Wait (0.00001) 'debug trap



End Function


Public Function Ref_Clock_On(ByVal ClockFreq As Double) As Long

Call itl.Raw.NI.Sync("Dev1").ConnectClkTerminals(NISYNC_CONSTS.Clkin, NISYNC_CONSTS.Clk10In) 'set DDS reference to CLKIN from AXRF

Call itl.Raw.NI.Sync("Dev1").ConnectClkTerminals(NISYNC_CONSTS.Dds, NISYNC_CONSTS.Clkout) 'route DDS to CLKOUT

Call itl.Raw.NI.Sync("Dev1").SetBoolean(niSyncProperties_ClkoutGainEnable, "", True) 'set high clk output (2.5vpk into 50 ohms) Note: EVM fails unless clk output high (True) is enabled

Call itl.Raw.NI.Sync("Dev1").SetDouble_2(niSyncProperties_DdsFreq, ClockFreq)

End Function
Public Function Ref_Clock_Off() As Long

Call itl.Raw.NI.Sync("Dev1").DisconnectClkTerminals(NISYNC_CONSTS.Dds, NISYNC_CONSTS.Clkout)

End Function

Public Function SimpleAXRFLoopBack(SrcChan As AXRF_CHANNEL, CapChan As AXRF_CHANNEL) As Long

    Dim power As Double
    Dim data(1023) As Double
    With itl.Raw.AF.AXRF

        .Source SrcChan, -20, 433920000#
        .MeasureSetup CapChan, -20, 433920000#
        'Power = .Measure(CapChan)
        TheExec.Datalog.WriteComment ("Measured power from the LoopBack test--- " & power)
        .MeasureArray CapChan, data, AXRF_ARRAY_TYPE_AXRF_FREQ_DOMAIN
'''        power = PlotDouble(data)         '(Debug)


    End With
End Function






