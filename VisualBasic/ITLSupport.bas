Attribute VB_Name = "ITLSupport"
' ------------------------------------------------------------------
' © 1997-2012 Teradyne, Inc. All Rights Reserved.
'
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software
' has been published.  This software is the trade secret information
' of Teradyne, Inc.  Use of this software is only in accordance with
' the terms of a license agreement from Teradyne, Inc.
' ------------------------------------------------------------------


' ###################################################################
' ###             WARNING DO NOT MODIFY THIS MODULE               ###
' ###                                                             ###
' ###                Teradyne ITL SOFTWARE                        ###
' ###          AUTOMATICALLY GENERATES THIS FILE                  ###
' ###################################################################
Option Explicit

Public NIDCPWR_CONSTS As New nidcpowerConstants
Public NIDMM_CONSTS As New nidmmConstants
Public NIRFSG_CONSTS As New niRFSGConstants
Public NIRFSA_CONSTS As New niRFSAConstants
Public NIRFSC_CONST As New niRFSCConstants
Public NIFGEN_CONSTS As New niFgenConstants
Public NIHSDIO_CONSTS As New niHSDIOConstants
Public NISCOPE_CONSTS As New niScopeConstants
Public NISWITCH_CONSTS As New niSwitchConstants
Public NISYNC_CONSTS As New niSyncConstants
Public GTDIO_CONSTS As New gtDIOConstants

Public itl As itl

Public Sub ITLOnProgramValidated()
     Set itl = Application.COMAddIns("TerITLAddIn").Object
     itl.Exec.TesterSvc.SetTesterExec TheExec
     itl.Exec.LoadConfiguration
End Sub


Public Sub ITLOnProgramStarted()
     Set itl = Application.COMAddIns("TerITLAddIn").Object
     itl.Exec.Initialize
End Sub


