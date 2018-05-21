Attribute VB_Name = "NI_Clock_Support"
Option Explicit


Public Function ArmClock20M(argc As Long, argv() As String) As Long


Call niSync_ConnectClkTerminals(Dev1, NISYNC_VAL_CLKIN, NISYNC_VAL_CLK10) 'set DDS reference to CLKIN from AXRF

Call niSync_ConnectClkTerminals(Dev1, NISYNC_VAL_DDS, NISYNC_VAL_CLKOUT) 'route DDS to CLKOUT

Call niSync_SetAttributeViBoolean(Dev1, NISYNC_ATTR_CLKOUT_GAIN_ENABLE, "", True) 'set high clk output (Dev1, 2.5vpk into 50 ohms) Note: EVM fails unless clk output high (Dev1, True) is enabled


Call niSync_SetAttributeViReal64(Dev1, NISYNC_ATTR_DDS_FREQ, 20000000#)

TheHdw.wait (0.00001) 'debug trap



End Function


Public Function Ref_Clock_On(ByVal ClockFreq As Double) As Long

Call niSync_ConnectClkTerminals(Dev1, NISYNC_VAL_CLKIN, NISYNC_VAL_CLK10) 'set DDS reference to CLKIN from AXRF

Call niSync_ConnectClkTerminals(Dev1, NISYNC_VAL_DDS, NISYNC_VAL_CLKOUT) 'route DDS to CLKOUT

Call niSync_SetAttributeViBoolean(Dev1, NISYNC_ATTR_CLKOUT_GAIN_ENABLE, "", True) 'set high clk output (Dev1, 2.5vpk into 50 ohms) Note: EVM fails unless clk output high (Dev1, True) is enabled

Call niSync_SetAttributeViReal64(Dev1, NISYNC_ATTR_DDS_FREQ, ClockFreq)

End Function
Public Function Ref_Clock_Off() As Long

Call niSync_DisconnectClkTerminals(Dev1, NISYNC_VAL_DDS, NISYNC_VAL_CLKOUT)

End Function

Public Function SimpleAXRFLoopBack(SrcChan As AXRF_CHANNEL, CapChan As AXRF_CHANNEL) As Long

    Dim power As Double
    Dim data(1023) As Double


        TevAXRF_Source SrcChan, -20, 433920000#
        TevAXRF_MeasureSetup CapChan, -20, 433920000#
        'Power = TevAXRF_Measure(CapChan)
        TheExecTevAXRF_DataLogTevAXRF_WriteComment ("Measured power from the LoopBack test--- " & power)
        TevAXRF_MeasureArray CapChan, data(0)(0), AXRF_FREQ_DOMAIN
'''        power = PlotDouble(data)         '(Debug)



End Function







