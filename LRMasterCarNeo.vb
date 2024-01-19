'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Master Car Object
'
' Copyright 2018 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 04/25/2018 09:24 am ]
'     Start coding this object...
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports ggcClient
Imports System.Windows.Forms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class LRMasterCarNeo

#Region "Constant"
    Private Const xsSignature As String = "08220326"
    Private Const pxeMODULENAME As String = "LR_Master"
    Private Const pxeMasterTble As String = "LR_Master"
    Private Const pxeDetailTble As String = "LR_Master_Car"
#End Region

#Region "Protected Members"
    Protected p_oAppDrvr As GRider
    Protected p_oDTMaster As DataTable
    Protected p_oDTDetail As DataTable
    Protected p_nEditMode As xeEditMode
    Protected p_oSC As New MySqlCommand

    Protected p_sBranchCd As String
    Protected p_sSourceNo As String
    Protected p_sSourceCd As String
    Protected p_bCancelled As Boolean

    Protected p_bInitTran As Boolean
    Protected p_nAcctStat As Integer = -1
#End Region

#Region "Public Event"
    Public Event MasterRetrieved(ByVal Index As Integer, _
                              ByVal Value As Object)

    Public Event DetailRetrieved(ByVal Index As Integer, _
                          ByVal Value As Object)
#End Region

#Region "Private"
    Private p_oClient As ggcClient.Client
#End Region

#Region "Properties"
    ReadOnly Property AppDriver As GRider
        Get
            Return p_oAppDrvr
        End Get
    End Property

    ReadOnly Property Cancelled
        Get
            Return p_bCancelled
        End Get
    End Property

    Property Branch() As String
        Get
            Return p_sBranchCd
        End Get
        Set(ByVal Value As String)
            p_sBranchCd = Value
        End Set
    End Property

    Property SourceNo() As String
        Get
            Return p_sSourceNo
        End Get
        Set(ByVal Value As String)
            p_sSourceNo = Value
        End Set
    End Property

    Property SourceCd() As String
        Get
            Return p_sSourceCd
        End Get
        Set(ByVal Value As String)
            p_sSourceCd = Value
        End Set
    End Property

    WriteOnly Property AccountStatus() As Integer
        Set(ByVal Value As Integer)
            p_nAcctStat = Value
        End Set
    End Property

    Property Master(ByVal Index As Object) As Object
        Get
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sacctnmbr"
                    Case "dtransact"
                    Case "sclientnm"
                    Case "xaddressx"
                    Case "sremarksx"
                    Case "nprincipl"
                    Case "nintratex"
                    Case "nsrvcchrg"
                    Case "nacctterm"
                    Case "nmonamort"
                    Case "dduedatex"
                    Case "ninterest"
                    Case "npenltyrt"
                    Case "ninschrge"
                    Case "nsrvcchrg"
                    Case "cactivexx"
                    Case "cloantype"
                    Case "cacctstat"
                    Case "dclosedxx"
                    Case "nrebatesx"
                    Case "nrebtotlx"
                    Case "nabalance"
                    Case "ninttotal"
                    Case "sapplicno"
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                        Return DBNull.Value
                End Select
            End If
            Return p_oDTMaster.Rows(0)(Index)
        End Get

        Set(ByVal Value As Object)
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sacctnmbr"
                    Case "dtransact", "dfirstpay", "dlastpaym", "dclosedxx"
                        If Index = "dfirstpay" Then
                            p_oDTMaster(0).Item("dDueDatex") = DateAdd(DateInterval.Month, p_oDTMaster(0).Item("nAcctTerm") - 1, p_oDTMaster(0).Item("dFirstPay"))
                            RaiseEvent MasterRetrieved(12, p_oDTMaster(0).Item("dDueDatex"))
                        End If

                        If Index = "dtransact" Then
                            p_oDTMaster(0).Item("dLastPaym") = p_oDTMaster(0).Item("dTransact")
                        End If
                    Case "sclientnm"
                        getClient(Value, False, False)
                    Case "xaddresxx"
                    Case "sremarksx"
                    Case "nprincipl", "nintratex"
                        If IsNumeric(Value) Then
                            p_oDTMaster(0).Item(Index) = Convert.ToSingle(Value)
                        End If

                        'If principal/interest rate/term then compute for amortization
                        If p_oDTMaster(0).Item("nAcctTerm") > 0 Then
                            'Set the principal as the initial balance
                            p_oDTMaster(0).Item("nABalance") = p_oDTMaster(0).Item("nPrincipl")

                            p_oDTMaster(0).Item("nInterest") = (p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) * p_oDTMaster(0).Item("nAcctTerm") * p_oDTMaster(0).Item("nIntRatex") / 100
                            RaiseEvent MasterRetrieved(6, p_oDTMaster(0).Item("nInterest"))

                            'p_oDTMaster(0).Item("nMonAmort") = Math.Round(((p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) + p_oDTMaster(0).Item("nInterest")) / p_oDTMaster(0).Item("nAcctTerm"), 2)
                            p_oDTMaster(0).Item("nMonAmort") = (Math.Round((p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) / p_oDTMaster(0).Item("nAcctTerm"), 2)) + p_oDTMaster(0).Item("nRebatesx")
                            RaiseEvent MasterRetrieved(13, p_oDTMaster(0).Item("nMonAmort"))
                        End If

                        RaiseEvent MasterRetrieved(22, p_oDTMaster(0).Item("nIntTotal"))
                    Case "nacctterm", "nrebatesx"
                        If IsNumeric(Value) Then
                            p_oDTMaster(0).Item(Index) = Convert.ToSingle(Value)
                        End If

                        p_oDTMaster(0).Item("nInterest") = (p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) * p_oDTMaster(0).Item("nAcctTerm") * p_oDTMaster(0).Item("nIntRatex") / 100
                        p_oDTMaster(0).Item("nABalance") = p_oDTMaster(0).Item("nPrincipl")
                        p_oDTMaster(0).Item("dDueDatex") = DateAdd(DateInterval.Month, p_oDTMaster(0).Item("nAcctTerm") - 1, p_oDTMaster(0).Item("dFirstPay"))
                        p_oDTMaster(0).Item("nMonAmort") = (Math.Round((p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) / p_oDTMaster(0).Item("nAcctTerm"), 2)) + p_oDTMaster(0).Item("nRebatesx")

                        RaiseEvent MasterRetrieved(13, p_oDTMaster(0).Item("nMonAmort"))
                        RaiseEvent MasterRetrieved(6, p_oDTMaster(0).Item("nInterest"))
                        RaiseEvent MasterRetrieved(12, p_oDTMaster(0).Item("dDueDatex"))
                        RaiseEvent MasterRetrieved(22, p_oDTMaster(0).Item("nIntTotal"))
                    Case "dduedatex"
                    Case "ninterest"
                    Case "cloantype"
                    Case "npenltyrt"
                    Case "ninschrge"
                        If IsNumeric(Value) Then
                            p_oDTMaster(0).Item(Index) = Convert.ToSingle(Value)
                        End If

                        p_oDTMaster(0).Item("nInterest") = (p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) * p_oDTMaster(0).Item("nAcctTerm") * p_oDTMaster(0).Item("nIntRatex") / 100
                        p_oDTMaster(0).Item("nMonAmort") = (Math.Round((p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) / p_oDTMaster(0).Item("nAcctTerm"), 2)) + p_oDTMaster(0).Item("nRebatesx")
                        'p_oDTMaster(0).Item("nMonAmort") = Math.Round(((p_oDTMaster(0).Item("nPrincipl") + p_oDTMaster(0).Item("nInsChrge")) + p_oDTMaster(0).Item("nInterest")) / p_oDTMaster(0).Item("nAcctTerm"), 2)

                        RaiseEvent MasterRetrieved(13, p_oDTMaster(0).Item("nMonAmort"))
                        RaiseEvent MasterRetrieved(6, p_oDTMaster(0).Item("nInterest"))
                        RaiseEvent MasterRetrieved(22, p_oDTMaster(0).Item("nIntTotal"))
                    Case "nsrvcchrg"
                    Case "sapplicno"
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                End Select
            End If
            p_oDTMaster(0)(Index) = Value
        End Set
    End Property

    Property Detail(ByVal Index As Object) As Object
        Get
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sengineno"
                    Case "sframenox"
                    Case "sbrandnme"
                    Case "smodelnme"
                    Case "scolornme"
                    Case "nyearmodl"
                    Case "sfilenoxx"
                    Case "screnoxxx"
                    Case "scrnoxxxx"
                    Case "splatenop"
                    Case "dregister"
                    Case "nfinamntx"
                    Case "nsubsidze"
                    Case "nsubcrdtd"
                    Case "ninctvamt"
                    Case "ninsamtxx"
                    Case "nincntpdx"
                    Case "ninsamtpd"
                    Case "dinslstpd"
                    Case "sserialid"
                    Case "smodelidx"
                    Case "sbrandidx"
                    Case "scoloridx"
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                        Return DBNull.Value
                End Select
            End If
            Return p_oDTDetail.Rows(0)(Index)
        End Get

        Set(ByVal Value As Object)
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sengineno"
                    Case "sframenox"
                    Case "sbrandnme"
                    Case "smodelnme"
                    Case "scolornme"
                    Case "nyearmodl"
                    Case "sfilenoxx"
                    Case "screnoxxx"
                    Case "scrnoxxxx"
                    Case "splatenop"
                    Case "dregister"
                    Case "nfinamntx"
                    Case "nsubsidze"
                    Case "nsubcrdtd"
                    Case "ninctvamt"
                    Case "ninsamtxx"
                    Case "nincntpdx"
                    Case "ninsamtpd"
                    Case "dinslstpd"
                    Case "sserialid"
                    Case "smodelidx"
                    Case "sbrandidx"
                    Case "scoloridx"
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                        Exit Property
                End Select
            End If
            p_oDTDetail.Rows(0)(Index) = Value
        End Set
    End Property
#End Region

#Region "Private Function"
    Private Function isUserActive(ByRef loDT As DataTable) As Boolean
        Dim lnCtr As Integer = 0
        Dim lbMember As Boolean = False

        If loDT.Rows(0).Item("cUserType").Equals(0) Then
            For lnCtr = 0 To loDT.Rows.Count - 1
                If loDT.Rows(0).Item("sProdctID").Equals(p_oAppDrvr.ProductID) Then
                    Exit For
                    lbMember = True
                End If
            Next
        Else
            lbMember = True
        End If

        If Not lbMember Then
            MsgBox("User is not a member of this application!!!" & vbCrLf & _
               "Application used is not allowed!!!", vbCritical, "Warning")
        End If

        ' check user status
        If loDT.Rows(0).Item("cUserStat").Equals(xeUserStatus.SUSPENDED) Then
            MsgBox("User is currently suspended!!!" & vbCrLf & _
                     "Application used is not allowed!!!", vbCritical, "Warning")
            Return False
        End If
        Return True
    End Function

    Private Function getSQ_User() As String
        Return "SELECT sUserIDxx" & _
              ", sLogNamex" & _
              ", sPassword" & _
              ", sUserName" & _
              ", nUserLevl" & _
              ", cUserType" & _
              ", sProdctID" & _
              ", cUserStat" & _
              ", nSysError" & _
              ", cLogStatx" & _
              ", cLockStat" & _
              ", cAllwLock" & _
           " FROM xxxSysUser" & _
           " WHERE sLogNamex = ?sLogNamex" & _
              " AND sPassword = ?sPassword"
    End Function

    Private Function getSQL_Model() As String
        Return "SELECT a.sModelIDx" & _
                    ", a.sModelNme" & _
                    ", b.sBrandNme" & _
                    ", b.sBrandIDx" & _
                " FROM Car_Model a" & _
                    ", Car_Brand b" & _
                " WHERE a.sBrandIDx = b.sBrandIDx" & _
                    " AND a.cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Color() As String
        Return "SELECT sColorIDx" & _
                    ", sColorNme" & _
                " FROM Color" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function isEntryOk() As Boolean
        'Check client
        If p_oDTMaster(0).Item("sClientID") = "" Then
            MsgBox("Client Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        'Check serial
        If p_oDTDetail(0).Item("sEngineNo") = "" Then
            MsgBox("Car serial seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If


        'Check validity of transaction date
        If p_oDTMaster(0).Item("dTransact") <= "2016-01-01" And p_oDTMaster(0).Item("dTransact") > p_oAppDrvr.SysDate Then
            MsgBox("Transaction release date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        'Check how much does he intends to borrow
        If Val(p_oDTMaster(0).Item("nPrincipl")) <= 1000 Then
            MsgBox("Loan Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        'Check when will be the exptected released of this loan
        If p_oDTMaster(0).Item("dFirstPay") < p_oDTMaster(0).Item("dTransact") Then
            MsgBox("Expected first pay date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        Return True
    End Function
#End Region

#Region "Public function"
    Function SaveTransaction() As Boolean

        Dim lsSQL As String
        Dim lnRow As Integer

        If p_bCancelled Then Return False

        If Not isEntryOk() Then Return False

        Try
            With p_oDTMaster
                p_oAppDrvr.BeginTransaction()

                Call SplitPayment(.Rows(0)("nPrincipl") _
                                  , .Rows(0)("nInterest") _
                                  , .Rows(0)("nAcctTerm") _
                                  , .Rows(0)("nPaymTotl") _
                                  , .Rows(0)("nIntTotal") _
                                  , 0 _
                                  , 0 _
                                  , 0)

                lsSQL = "INSERT INTO " & pxeMasterTble & " SET" & _
                            "  sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr")) & _
                            ", dTransact = " & dateParm(.Rows(0)("dTransact")) & _
                            ", sCompnyID = " & strParm(.Rows(0)("sCompnyID")) & _
                            ", sBranchCd = " & strParm(.Rows(0)("sBranchCd")) & _
                            ", sApplicNo = " & strParm(.Rows(0)("sApplicNo")) & _
                            ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                            ", sRemarksx = " & strParm(.Rows(0)("sRemarksx")) & _
                            ", nPrincipl = " & CDbl(.Rows(0)("nPrincipl")) & _
                            ", nInterest = " & CDbl(.Rows(0)("nInterest")) & _
                            ", nSrvcChrg = " & CDbl(.Rows(0)("nSrvcChrg")) & _
                            ", nInsChrge = " & CDbl(.Rows(0)("nInsChrge")) & _
                            ", nIntRatex = " & CDbl(.Rows(0)("nIntRatex")) & _
                            ", nAcctTerm = " & CDbl(.Rows(0)("nAcctTerm")) & _
                            ", dFirstPay = " & dateParm(.Rows(0)("dFirstPay")) & _
                            ", nMonAmort = " & CDbl(.Rows(0)("nMonAmort")) & _
                            ", nRebatesx = " & CDbl(.Rows(0)("nRebatesx")) & _
                            ", nPenltyRt = " & CDbl(.Rows(0)("nPenltyRt")) & _
                            ", nLastPaym = " & CDbl(.Rows(0)("nLastPaym")) & _
                            ", dLastPaym = " & dateParm(.Rows(0)("dLastPaym")) & _
                            ", dDueDatex = " & dateParm(.Rows(0)("dDueDatex")) & _
                            ", nPaymTotl = " & CDbl(.Rows(0)("nPaymTotl")) & _
                            ", nPenTotlx = " & CDbl(.Rows(0)("nPenTotlx")) & _
                            ", nDebtTotl = " & CDbl(.Rows(0)("nDebtTotl")) & _
                            ", nCredTotl = " & CDbl(.Rows(0)("nCredTotl")) & _
                            ", nIntTotal = " & CDbl(.Rows(0)("nIntTotal")) & _
                            ", nRebTotlx = " & CDbl(.Rows(0)("nRebTotlx")) & _
                            ", nAmtDuexx = " & CDbl(.Rows(0)("nAmtDuexx")) & _
                            ", nABalance = " & CDbl(.Rows(0)("nABalance")) & _
                            ", nDelayAvg = " & CDbl(.Rows(0)("nDelayAvg")) & _
                            ", cRatingxx = " & strParm(0) & _
                            ", cActivexx = " & strParm(1) & _
                            ", cLoanType = " & strParm(1) & _
                            ", nLedgerNo = " & strParm(0) & _
                            ", sMCActNox = " & strParm("") & _
                            ", sRouteIDx = " & strParm("") & _
                            ", sExAcctNo = " & strParm("") & _
                            ", sModified = " & strParm(p_oAppDrvr.UserID) & _
                            ", dModified = " & dateParm(p_oAppDrvr.SysDate) & _
                        " ON DUPLICATE KEY UPDATE" & _
                            "  dTransact = " & dateParm(.Rows(0)("dTransact")) & _
                            ", sCompnyID = " & strParm(.Rows(0)("sCompnyID")) & _
                            ", sBranchCd = " & strParm(.Rows(0)("sBranchCd")) & _
                            ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                            ", sRemarksx = " & strParm(.Rows(0)("sRemarksx")) & _
                            ", nPrincipl = " & CDbl(.Rows(0)("nPrincipl")) & _
                            ", nInterest = " & CDbl(.Rows(0)("nInterest")) & _
                            ", nSrvcChrg = " & CDbl(.Rows(0)("nSrvcChrg")) & _
                            ", nInsChrge = " & CDbl(.Rows(0)("nInsChrge")) & _
                            ", nIntRatex = " & CDbl(.Rows(0)("nIntRatex")) & _
                            ", nAcctTerm = " & CDbl(.Rows(0)("nAcctTerm")) & _
                            ", dFirstPay = " & dateParm(.Rows(0)("dFirstPay")) & _
                            ", nMonAmort = " & CDbl(.Rows(0)("nMonAmort")) & _
                            ", nRebatesx = " & CDbl(.Rows(0)("nRebatesx")) & _
                            ", nPenltyRt = " & CDbl(.Rows(0)("nPenltyRt")) & _
                            ", nLastPaym = " & CDbl(.Rows(0)("nLastPaym")) & _
                            ", dLastPaym = " & dateParm(.Rows(0)("dLastPaym")) & _
                            ", dDueDatex = " & dateParm(.Rows(0)("dDueDatex"))

                lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                If lnRow <= 0 Then
                    MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    p_oAppDrvr.RollBackTransaction()
                    Return False
                End If
            End With

            If p_oDTDetail.Rows(0)("sEngineNo") <> "" Then
                Dim loSerial As LRCarSerial
                loSerial = New LRCarSerial(p_oAppDrvr)
                loSerial.NewTransaction()

                With loSerial
                    .HasParent = True

                    If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                        .NewTransaction()
                        p_oDTDetail.Rows(0)("sSerialID") = .Master("sSerialID")
                    Else
                        .Master("sSerialID") = p_oDTDetail.Rows(0)("sSerialID")
                        .UpdateTransaction()
                    End If

                    .Master("sEngineNo") = p_oDTDetail.Rows(0)("sEngineNo")
                    .Master("sFrameNox") = p_oDTDetail.Rows(0)("sFrameNox")
                    .Master("sModelCde") = p_oDTDetail.Rows(0)("sModelIDx")
                    .Master("sBrandCde") = p_oDTDetail.Rows(0)("sBrandIDx")
                    .Master("sColorCde") = p_oDTDetail.Rows(0)("sColorIDx")
                    .Master("sFileNoxx") = p_oDTDetail.Rows(0)("sFileNoxx")
                    .Master("sCRENoxxx") = p_oDTDetail.Rows(0)("sCRENoxxx")
                    .Master("sCRNoxxxx") = p_oDTDetail.Rows(0)("sCRNoxxxx")
                    .Master("sPlateNoP") = p_oDTDetail.Rows(0)("sPlateNoP")
                    .Master("dRegister") = p_oDTDetail.Rows(0)("dRegister")
                    .Master("nYearModl") = p_oDTDetail.Rows(0)("nYearModl")
                    .Master("sClientID") = p_oDTMaster.Rows(0)("sClientID")

                    If Not .SaveTransaction Then
                        MsgBox("Unable to save serial info!", vbOKOnly, pxeMODULENAME)
                        p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End With
            End If

            With p_oDTDetail
                lsSQL = "INSERT INTO " & pxeDetailTble & " SET" & _
                            "  sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr")) & _
                            ", sSerialID = " & strParm(.Rows(0)("sSerialID")) & _
                            ", nFinAmntx = " & CDbl(.Rows(0)("nFinAmntx")) & _
                            ", nSubsidze = " & CDbl(.Rows(0)("nSubsidze")) & _
                            ", nInctvAmt = " & CDbl(.Rows(0)("nInctvAmt")) & _
                            ", nInsAmtxx = " & CDbl(p_oDTMaster.Rows(0)("nInsChrge")) & _
                            ", dInsLstPd = " & dateParm(.Rows(0)("dInsLstPd")) & _
                        " ON DUPLICATE KEY UPDATE" & _
                            "  sSerialID = " & strParm(.Rows(0)("sSerialID")) & _
                            ", nFinAmntx = " & CDbl(.Rows(0)("nFinAmntx")) & _
                            ", nSubsidze = " & CDbl(.Rows(0)("nSubsidze")) & _
                            ", nInctvAmt = " & CDbl(.Rows(0)("nInctvAmt")) & _
                            ", nInsAmtxx = " & CDbl(p_oDTMaster.Rows(0)("nInsChrge")) & _
                            ", dInsLstPd = " & dateParm(.Rows(0)("dInsLstPd"))

                lnRow = p_oAppDrvr.Execute(lsSQL, pxeDetailTble)
                If lnRow <= 0 Then
                    MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    p_oAppDrvr.RollBackTransaction()
                    Return False
                End If

                lsSQL = "UPDATE MC_Credit_Application SET" & _
                        "  cTranStat = '4'" & _
                    "WHERE sTransNox = " & strParm(p_oDTMaster.Rows(0)("sApplicNo"))

                lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                'If lnRow <= 0 Then
                '    MsgBox("Unable to Save Transaction!!!" & vbCrLf &
                '            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                '    p_oAppDrvr.RollBackTransaction()
                '    Return False
                'End If
            End With

            If p_oDTMaster.Rows(0)("sClientID") <> "" Then
                If p_oDTMaster.Rows(0)("sApplicNo") = "" Then
                    If Not p_oClient.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If
            End If

            p_oAppDrvr.CommitTransaction()

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        Catch ex As Exception
            p_oAppDrvr.RollBackTransaction()

            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Function PostTransaction() As Boolean
        Dim lsSQL As String
        Dim lnRow As Integer
        Dim loDT As New DataTable

        Try
            With p_oDTMaster
                p_oAppDrvr.BeginTransaction()

                If Not IsDBNull(.Rows(0)("cAcctStat")) Then
                    MsgBox("Unable to Post Transaction!!!" & vbCrLf & _
                            "Please contact GGC SEG/SSG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    Return False
                End If

                lsSQL = "UPDATE " & pxeMasterTble & " SET" & _
                            "  cAcctStat = " & strParm("0") & _
                        " WHERE sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr"))

                Try
                    lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                    If lnRow <= 0 Then
                        MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                                "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                        p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                Catch ex As Exception
                    Throw ex
                End Try

                lsSQL = "SELECT" & _
                            "  a.sReferNox" & _
                        " FROM MC_Credit_Application a" & _
                            ", LR_Master b" & _
                        " WHERE a.sTransNox = b.sApplicNo" & _
                            " AND b.sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr"))

                loDT = New DataTable
                loDT = p_oAppDrvr.ExecuteQuery(lsSQL)
                If loDT.Rows.Count = 0 Then Return False

                lsSQL = "UPDATE Credit_Online_Application SET" & _
                            "  cTranStat = " & strParm("2") & _
                        " WHERE sTransNox = " & strParm(loDT.Rows(0)("sReferNox"))

                Try
                    lnRow = p_oAppDrvr.Execute(lsSQL, "Credit_Online_Application")
                    If lnRow <= 0 Then
                        MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                                "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                        p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                Catch ex As Exception
                    Throw ex
                End Try

                lsSQL = "UPDATE " & pxeDetailTble & " SET" & _
                            "  nFinAmntx = " & CDbl(p_oDTMaster.Rows(0)("nPrincipl")) & _
                            ", nSubCrdtd = " & CDbl(p_oDTDetail.Rows(0)("nSubsidze")) & _
                            ", nIncntPdx = " & CDbl(p_oDTDetail.Rows(0)("nInctvAmt")) & _
                            ", nInsAmtPd = " & CDbl(p_oDTDetail.Rows(0)("nInsAmtxx")) & _
                        " WHERE sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr")) & _
                            " AND sSerialID = " & strParm(p_oDTDetail.Rows(0)("sSerialID"))
                Try
                    lnRow = p_oAppDrvr.Execute(lsSQL, pxeDetailTble)
                    If lnRow <= 0 Then
                        MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                                "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                        p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                Catch ex As Exception
                    Throw ex
                End Try

                lsSQL = "UPDATE Car_Serial SET" & _
                            "  cSoldStat = " & strParm("1") & _
                        " WHERE sSerialID = " & strParm(p_oDTDetail.Rows(0)("sSerialID"))

                lnRow = p_oAppDrvr.Execute(lsSQL, "Car_Serial")
                If lnRow <= 0 Then
                    MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    p_oAppDrvr.RollBackTransaction()
                    Return False
                End If
            End With

            p_oAppDrvr.CommitTransaction()
            Return True
        Catch ex As Exception
            p_oAppDrvr.RollBackTransaction()
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Function CancelTransaction() As Boolean
        Dim lsSQL As String
        Dim lnRow As Integer

        Try
            With p_oDTMaster
                p_oAppDrvr.BeginTransaction()

                If Not IsDBNull(.Rows(0)("cAcctStat")) Then
                    MsgBox("Unable to Cancel Transaction!!!" & vbCrLf & _
                            "Please contact GGC SEG/SSG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    Return False
                End If

                lsSQL = "DELETE FROM " & pxeMasterTble & _
                        " WHERE sAcctNmbr = " & strParm(.Rows(0)("sAcctNmbr"))

                lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                If lnRow <= 0 Then
                    MsgBox("Unable to Cancel Transaction!!!" & vbCrLf & _
                            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    p_oAppDrvr.RollBackTransaction()
                    Return False
                End If
            End With

            p_oAppDrvr.CommitTransaction()
            Return True
        Catch ex As Exception
            p_oAppDrvr.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Function InitTransaction() As Boolean
        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oAppDrvr.BranchCode

        createMasterTable()
        createTempTable()

        p_bInitTran = True
        InitTransaction = True
    End Function

    Function NewTransaction() As Boolean
        If p_sBranchCd = "" Then
            MsgBox("Branch is empty... Please indicate branch!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        Call initMaster()
        Call initDetail()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Function SearchTransaction( _
                            ByVal fsValue As String _
                          , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String
        Dim lsCondition As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMaster(0).Item("sAcctNmbr") Then Return True
            Else
                If fsValue = p_oDTMaster(0).Item("sClientNm") Then Return True
            End If
        End If

        lsSQL = getSQL_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sAcctNmbr = " & strParm(fsValue)
        Else
            lsFilter = "CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName) LIKE " & strParm("%" & fsValue & "%")
        End If

        If p_nAcctStat <> -1 Then
            If p_nAcctStat > -1 Then
                lsCondition = "("
                For pnCtr = 1 To Len(Trim(Str(p_nAcctStat)))
                    lsCondition = lsCondition & " a.cAcctStat = " & _
                                      strParm(Mid(Trim(Str(p_nAcctStat)), pnCtr, 1)) & " OR "
                Next
                lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
            Else
                lsCondition = "a.cAcctStat = " & strParm(p_nAcctStat)
            End If
        End If

        lsSQL = AddCondition(lsSQL, lsCondition)

        Debug.Print(lsSQL)

        Dim loDta As DataRow = KwikSearch(p_oAppDrvr _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sAcctNmbr»sClientNm»sEngineNo»sPlateNoP" _
                                        , "AcctNmbr»Client»EngineNo»Plate", _
                                        , "a.sAcctNmbr»CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName)»d.sEngineNo»d.sPlateNoP" _
                                        , IIf(fbByCode, 0, 1))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sAcctNmbr"))
        End If
    End Function

    Public Function SearchApplication() As Boolean
        Dim lsSQL As String

        lsSQL = "SELECT" & _
                    "  a.sTransNox" & _
                    ", a.dAppliedx" & _
                    ", a.cTranStat" & _
                    ", b.sCompnyNm" & _
                    ", CONCAT(b.sAddressx, ', ', c.sTownName, ' ', d.sProvName) xAddressx" & _
                    ", a.sClientID" & _
                " FROM MC_Credit_Application a" & _
                    ", Client_Master b" & _
                    ", TownCity c" & _
                    ", Province d" & _
                " WHERE a.sClientID = b.sClientID" & _
                    " AND a.cTranStat = " & strParm(xeTranStat.TRANS_POSTED) & _
                    " AND DATEDIFF(SYSDATE(), a.dAppliedx) <= 60" & _
                    " AND b.sTownIDxx = c.sTownIDxx" & _
                    " AND c.sProvIDxx = d.sProvIDxx" & _
                    " AND a.cUnitAppl = '4'" & _
                " ORDER BY a.dAppliedx"

        Dim loDta As DataRow = KwikSearch(p_oAppDrvr _
                                        , lsSQL _
                                        , False _
                                        , "" _
                                        , "sTransNox»sCompnyNm»dAppliedx" _
                                        , "Trans No»Client»Date", _
                                        , "sTransNox»sCompnyNm»dAppliedx" _
                                        , 2)
        If IsNothing(loDta) Then
            Return False
        End If

        With p_oDTMaster
            .Rows(0).Item("sClientID") = loDta.Item("sClientID")
            .Rows(0).Item("sClientNm") = loDta.Item("sCompnyNm")
            .Rows(0).Item("sApplicNo") = loDta.Item("sTransNox")
            .Rows(0).Item("xAddressx") = loDta.Item("xAddressx")
        End With

        Return True
    End Function


    Function OpenTransaction(ByVal fsAcctNmbr As String) As Boolean
        Dim loDT As New DataTable
        Dim lsSQL As String

        p_nEditMode = xeEditMode.MODE_READY

        lsSQL = AddCondition("SELECT * FROM " & pxeMasterTble, "sAcctNmbr = " & strParm(fsAcctNmbr))

        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Return False

        If loDT.Rows.Count > 0 Then
            With p_oDTMaster
                .Rows.Add()
                For nCtr As Integer = 0 To .Columns.Count - 1
                    Select Case .Columns.Item(nCtr).ColumnName
                        Case "sClientNm", "xAddressx"
                        Case "sClientID"
                            Call getClient(loDT.Rows(0)(.Columns.Item(nCtr).ColumnName), True, True)
                        Case Else
                            .Rows(0)(.Columns.Item(nCtr).ColumnName) = loDT.Rows(0)(.Columns.Item(nCtr).ColumnName)
                    End Select
                Next nCtr
            End With

            lsSQL = AddCondition("SELECT * FROM " & pxeDetailTble, "sAcctNmbr = " & strParm(fsAcctNmbr))

            loDT = New DataTable
            loDT = p_oAppDrvr.ExecuteQuery(lsSQL)
            With p_oDTDetail
                .Rows.Add()
                For nCtr As Integer = 0 To loDT.Columns.Count - 1

                    .Rows(0)(loDT.Columns.Item(nCtr).ColumnName) = loDT.Rows(0)(loDT.Columns.Item(nCtr).ColumnName)
                Next nCtr

                Call getSerial(p_oDTDetail.Rows(0)("sSerialID"), True, True)
            End With
        Else
            Call initMaster()
            Call initDetail()
        End If

        Return True
    End Function
#End Region

#Region "Private Procedures"
    Private Sub createMasterTable()
        p_oDTMaster = New DataTable
        With p_oDTMaster
            .Columns.Add("sAcctNmbr", GetType(String)).MaxLength = 10
            .Columns.Add("dTransact", GetType(Date))
            .Columns.Add("sClientNm", GetType(String)).MaxLength = 128
            .Columns.Add("xAddressx", GetType(String)).MaxLength = 128
            .Columns.Add("sRemarksx", GetType(String)).MaxLength = 128
            .Columns.Add("nPrincipl", GetType(Decimal))
            .Columns.Add("nInterest", GetType(Decimal))
            .Columns.Add("nSrvcChrg", GetType(Decimal))
            .Columns.Add("nInsChrge", GetType(Decimal))
            .Columns.Add("nIntRatex", GetType(Decimal))
            .Columns.Add("dFirstPay", GetType(Date))
            .Columns.Add("nAcctTerm", GetType(Integer))
            .Columns.Add("dDueDatex", GetType(Date))
            .Columns.Add("nMonAmort", GetType(Decimal))
            .Columns.Add("nPenltyRt", GetType(Decimal))
            .Columns.Add("nLastPaym", GetType(Decimal))
            .Columns.Add("dLastPaym", GetType(Date))
            .Columns.Add("nLedgerNo", GetType(Decimal))
            .Columns.Add("nPaymTotl", GetType(Decimal))
            .Columns.Add("nPenTotlx", GetType(Decimal))
            .Columns.Add("nDebtTotl", GetType(Decimal))
            .Columns.Add("nCredTotl", GetType(Decimal))
            .Columns.Add("nIntTotal", GetType(Decimal))
            .Columns.Add("nAmtDuexx", GetType(Decimal))
            .Columns.Add("nABalance", GetType(Decimal))
            .Columns.Add("nDelayAvg", GetType(Decimal))
            .Columns.Add("cRatingxx", GetType(Char))
            .Columns.Add("cAcctstat", GetType(Char))
            .Columns.Add("dClosedxx", GetType(Date))
            .Columns.Add("nRebatesx", GetType(Decimal))
            .Columns.Add("nRebTotlx", GetType(Decimal))
            .Columns.Add("cActivexx", GetType(Char))
            .Columns.Add("cLoanType", GetType(Char))
            .Columns.Add("sCompnyID", GetType(String)).MaxLength = 4
            .Columns.Add("sBranchCd", GetType(String)).MaxLength = 4
            .Columns.Add("sClientID", GetType(String)).MaxLength = 12
            .Columns.Add("sApplicNo", GetType(String)).MaxLength = 12
        End With
    End Sub

    Private Sub initMaster()
        With p_oDTMaster
            .Clear()
            .Rows.Add()
            .Rows(0)("sAcctNmbr") = GetNextCode(pxeMasterTble, "sAcctNmbr", True, p_oAppDrvr.Connection, True, p_sBranchCd)
            .Rows(0)("dTransact") = p_oAppDrvr.SysDate
            .Rows(0)("sCompnyID") = "CT"
            .Rows(0)("sBranchCd") = p_sBranchCd
            .Rows(0)("sClientID") = ""
            .Rows(0)("sRemarksx") = ""
            .Rows(0)("sClientNm") = ""
            .Rows(0)("nSrvcChrg") = 0.0
            .Rows(0)("nInsChrge") = 0.0
            .Rows(0)("nIntRatex") = 0.0
            .Rows(0)("nInterest") = 0.0
            .Rows(0)("nSrvcChrg") = 0.0
            .Rows(0)("nPrincipl") = 0.0
            .Rows(0)("nMonAmort") = 0.0
            .Rows(0)("nRebatesx") = 0.0
            .Rows(0)("dFirstPay") = DateAdd(DateInterval.Month, 1, p_oAppDrvr.SysDate)
            .Rows(0)("dDueDatex") = p_oAppDrvr.SysDate
            .Rows(0)("nAcctTerm") = 0
            .Rows(0)("nPenltyRt") = 0.0

            .Rows(0)("cLoanType") = "1"
            .Rows(0)("nLastPaym") = 0.0
            .Rows(0)("dLastPaym") = p_oAppDrvr.SysDate
            .Rows(0)("nPaymTotl") = 0.0
            .Rows(0)("nPenTotlx") = 0.0
            .Rows(0)("nDebtTotl") = 0.0
            .Rows(0)("nCredTotl") = 0.0
            .Rows(0)("nIntTotal") = 0.0
            .Rows(0)("nRebTotlx") = 0.0
            .Rows(0)("nAmtDuexx") = 0.0
            .Rows(0)("nABalance") = 0.0
            .Rows(0)("nDelayAvg") = 0.0
            .Rows(0)("cActivexx") = "1"
            .Rows(0)("nLedgerNo") = 0
            .Rows(0)("sApplicNo") = ""
        End With
    End Sub

    Private Sub initDetail()
        With p_oDTDetail
            .Clear()
            .Rows.Add()
            .Rows(0)("sAcctNmbr") = p_oDTMaster.Rows(0)("sAcctNmbr")
            .Rows(0)("sSerialID") = ""
            .Rows(0)("nFinAmntx") = 0.0
            .Rows(0)("nSubsidze") = 0.0
            .Rows(0)("nSubCrdtd") = 0.0
            .Rows(0)("nInctvAmt") = 0.0
            .Rows(0)("nIncntPdx") = 0.0
            .Rows(0)("nInsAmtxx") = 0.0
            .Rows(0)("nInsAmtPd") = 0.0
            .Rows(0)("dInsLstPd") = p_oDTMaster.Rows(0)("dTransact")
        End With
    End Sub

    Private Sub createTempTable()
        p_oDTDetail = New DataTable
        With p_oDTDetail
            .Columns.Add("sAcctNmbr", GetType(String)).MaxLength = 10
            .Columns.Add("sEngineNo", GetType(String)).MaxLength = 30
            .Columns.Add("sFrameNox", GetType(String)).MaxLength = 30
            .Columns.Add("sModelNme", GetType(String)).MaxLength = 30
            .Columns.Add("sBrandNme", GetType(String)).MaxLength = 30
            .Columns.Add("sColorNme", GetType(String)).MaxLength = 30
            .Columns.Add("sFileNoxx", GetType(String)).MaxLength = 30
            .Columns.Add("sCRENoxxx", GetType(String)).MaxLength = 30
            .Columns.Add("sCRNoxxxx", GetType(String)).MaxLength = 30
            .Columns.Add("sPlateNoP", GetType(String)).MaxLength = 30
            .Columns.Add("dRegister", GetType(Date))
            .Columns.Add("nYearModl", GetType(Integer))
            .Columns.Add("nFinAmntx", GetType(Decimal))
            .Columns.Add("nSubsidze", GetType(Decimal))
            .Columns.Add("nSubCrdtd", GetType(Decimal))
            .Columns.Add("nInctvAmt", GetType(Decimal))
            .Columns.Add("nIncntPdx", GetType(Decimal))
            .Columns.Add("nInsAmtxx", GetType(Decimal))
            .Columns.Add("nInsAmtPd", GetType(Decimal))
            .Columns.Add("dInsLstPd", GetType(Date))
            .Columns.Add("sSerialID", GetType(String)).MaxLength = 12
            .Columns.Add("sModelIDx", GetType(String)).MaxLength = 9
            .Columns.Add("sBrandIDx", GetType(String)).MaxLength = 9
            .Columns.Add("sColorIDx", GetType(String)).MaxLength = 7
        End With
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMaster(0).Item("sClientID") And fsValue <> "" And p_oDTMaster(0).Item("sClientNm") <> "" Then Exit Sub
        Else
            'Do not allow searching of value if fsValue is empty
            If (fsValue = p_oDTMaster(0).Item("sClientNm") And fsValue <> "") Or fsValue = "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oAppDrvr)
        loClient.Parent = "LRMasterCar"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oClient = loClient
                p_oDTMaster(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oDTMaster(0).Item("sClientNm") = p_oClient.Master("sLastName") & ", " & _
                                                   p_oClient.Master("sFrstName") & _
                                                   IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                                   p_oClient.Master("sMiddName")

                p_oDTMaster(0).Item("xAddressx") = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                                    p_oClient.Master("sAddressx") & ", " & _
                                                    p_oClient.Master("sTownName")
            Else
                p_oDTMaster(0).Item("sClientID") = ""
                p_oDTMaster(0).Item("sClientNm") = ""
                p_oDTMaster(0).Item("xAddressx") = ""
            End If

            RaiseEvent MasterRetrieved(2, p_oDTMaster(0).Item("sClientNm"))
            RaiseEvent MasterRetrieved(3, p_oDTMaster(0).Item("xAddressx"))
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                p_oClient = loClient
                p_oDTMaster(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oDTMaster(0).Item("sClientNm") = p_oClient.Master("sLastName") & ", " & _
                                                   p_oClient.Master("sFrstName") & _
                                                   IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                                   p_oClient.Master("sMiddName")
                p_oDTMaster(0).Item("xAddressx") = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                                   p_oClient.Master("sAddressx") & ", " & _
                                                   p_oClient.Master("sTownName")
            End If
        End If

        RaiseEvent MasterRetrieved(2, p_oDTMaster.Rows(0)("sClientNm"))
        RaiseEvent MasterRetrieved(3, p_oDTMaster.Rows(0)("xAddressx"))
    End Sub

    Private Sub SplitPayment( _
            ByVal fnPrincipl As Decimal _
          , ByVal fnInterest As Decimal _
          , ByVal fnAcctTerm As Integer _
          , ByRef fnPaymTotl As Decimal _
          , ByRef fnIntTotal As Decimal _
          , ByRef fnTranAmtx As Decimal _
          , ByRef fnPaidAmtx As Decimal _
          , ByRef fnIntAmtxx As Decimal)

        'Compute for the monthly amortization for the principal and interest
        Dim lnPayAmort As Decimal = fnPrincipl / fnAcctTerm
        Dim lnIntAmort As Decimal = fnInterest / fnAcctTerm

        'Compute for the number of terms paid for the principal and interest
        Dim lnPayTermx As Single = fnPaymTotl / lnPayAmort
        Dim lnIntTermx As Single = fnIntTotal / lnIntAmort

        If fnTranAmtx <= 0 Then Exit Sub

        If lnPayTermx = lnIntTermx Then
            'Distribute payment to interest payment
            If fnTranAmtx < lnIntAmort Then
                'Get the actual interest deducted
                lnIntAmort = fnTranAmtx

                fnIntAmtxx = fnIntAmtxx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnIntAmtxx = fnIntAmtxx + lnIntAmort
                fnTranAmtx = fnTranAmtx - lnIntAmort
            End If

            'Distribute payment to monthly payment
            If fnTranAmtx < lnPayAmort Then
                'Get the actual monthly amortization deducted
                lnPayAmort = fnTranAmtx

                fnPaidAmtx = fnPaidAmtx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnPaidAmtx = fnPaidAmtx + lnPayAmort
                fnTranAmtx = fnTranAmtx - lnPayAmort
            End If
            fnPaymTotl = fnPaymTotl + lnPayAmort
            fnIntTotal = fnIntTotal + lnIntAmort
        ElseIf lnPayTermx < lnIntTermx Then
            'Compute for the amount to be distributed for monthly payment
            'Dim lnDiff As Decimal = (lnPayTermx - lnIntTermx) * lnPayAmort
            Dim lnDiff As Decimal = (lnIntTermx - lnPayTermx) * lnPayAmort
            lnPayAmort = lnDiff
            If fnTranAmtx < lnDiff Then
                'Get the actual monthly amortization
                lnPayAmort = fnTranAmtx

                fnPaidAmtx = fnPaidAmtx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnPaidAmtx = fnPaidAmtx + lnDiff
                fnTranAmtx = fnTranAmtx - lnDiff
            End If
            fnPaymTotl = fnPaymTotl + lnPayAmort
        Else
            'Compute for the amount to be distributed for interest payment
            Dim lnDiff As Decimal = (lnPayTermx - lnIntTermx) * lnIntAmort
            lnIntAmort = lnDiff
            If fnTranAmtx < lnDiff Then
                lnIntAmort = fnTranAmtx
                fnIntAmtxx = fnIntAmtxx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnIntAmtxx = fnIntAmtxx + lnDiff
                fnTranAmtx = fnTranAmtx - lnDiff
            End If
            fnIntTotal = fnIntTotal + lnIntAmort
        End If

        'Execute a recursive function if fnTranAmtx is not yet 0
        If fnTranAmtx > 0 Then
            SplitPayment(fnPrincipl, fnInterest, fnAcctTerm, fnPaymTotl, fnIntTotal, fnTranAmtx, fnPaidAmtx, fnIntAmtxx)
        End If
    End Sub

    Private Function getSerial(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getSerial"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sEngineNo LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sEngineNo = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sSerialID = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            lsCondition = lsCondition & " a.cSoldStat = " & strParm("0")
        End If
        lsSQL = AddCondition(getSQL_Serial, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTDetail
                .Rows(0)("sSerialID") = loDT(0)("sSerialID")
                .Rows(0)("sEngineNo") = loDT(0)("sEngineNo")
                .Rows(0)("sFrameNox") = loDT(0)("sFrameNox")
                .Rows(0)("sBrandNme") = loDT(0)("sBrandNme")
                .Rows(0)("sModelNme") = loDT(0)("sModelNme")
                .Rows(0)("sColorNme") = loDT(0)("sColorNme")
                .Rows(0)("nYearModl") = loDT(0)("nYearModl")
                .Rows(0)("sFileNoxx") = loDT(0)("sFileNoxx")
                .Rows(0)("sCRENoxxx") = loDT(0)("sCRENoxxx")
                .Rows(0)("sCRNoxxxx") = loDT(0)("sCRNoxxxx")
                .Rows(0)("sPlateNoP") = loDT(0)("sPlateNoP")
                .Rows(0)("dRegister") = loDT(0)("dRegister")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sSerialID»sEngineNo»sFrameNox»sPlateNoP", _
                                "SerialID»Engine No»Frame No»Plate", _
                                "", _
                                "", _
                                3)

            If Not IsNothing(loDataRow) Then
                With p_oDTDetail
                    .Rows(0)("sSerialID") = loDataRow("sSerialID")
                    .Rows(0)("sEngineNo") = loDataRow("sEngineNo")
                    .Rows(0)("sFrameNox") = loDataRow("sFrameNox")
                    .Rows(0)("sBrandNme") = loDataRow("sBrandNme")
                    .Rows(0)("sModelNme") = loDataRow("sModelNme")
                    .Rows(0)("sColorNme") = loDataRow("sColorNme")
                    .Rows(0)("nYearModl") = loDataRow("nYearModl")
                    .Rows(0)("sFileNoxx") = loDataRow("sFileNoxx")
                    .Rows(0)("sCRENoxxx") = loDataRow("sCRENoxxx")
                    .Rows(0)("sCRNoxxxx") = loDataRow("sCRNoxxxx")
                    .Rows(0)("sPlateNoP") = loDataRow("sPlateNoP")
                    .Rows(0)("dRegister") = loDataRow("dRegister")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTDetail
            RaiseEvent DetailRetrieved(1, .Rows(0)("sEngineNo"))
            RaiseEvent DetailRetrieved(2, .Rows(0)("sFrameNox"))
            RaiseEvent DetailRetrieved(3, .Rows(0)("sBrandNme"))
            RaiseEvent DetailRetrieved(4, .Rows(0)("sModelNme"))
            RaiseEvent DetailRetrieved(5, .Rows(0)("sColorNme"))
            RaiseEvent DetailRetrieved(6, .Rows(0)("nYearModl"))
            RaiseEvent DetailRetrieved(7, .Rows(0)("sFileNoxx"))
            RaiseEvent DetailRetrieved(8, .Rows(0)("sCRENoxxx"))
            RaiseEvent DetailRetrieved(9, .Rows(0)("sCRNoxxxx"))
            RaiseEvent DetailRetrieved(10, .Rows(0)("sPlateNoP"))
            RaiseEvent DetailRetrieved(11, .Rows(0)("dRegister"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTDetail
            .Rows(0)("sSerialID") = ""
            .Rows(0)("sEngineNo") = ""
            .Rows(0)("sFrameNox") = ""
            .Rows(0)("sBrandNme") = ""
            .Rows(0)("sModelNme") = ""
            .Rows(0)("sColorNme") = ""
            .Rows(0)("nYearModl") = 0
            .Rows(0)("sFileNoxx") = ""
            .Rows(0)("sCRENoxxx") = ""
            .Rows(0)("sCRNoxxxx") = ""
            .Rows(0)("sPlateNoP") = ""
            .Rows(0)("dRegister") = p_oAppDrvr.SysDate
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getModel(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Model, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTDetail
                .Rows(0)("sModelIDx") = loDT(0)("sModelIDx")
                .Rows(0)("sModelNme") = loDT(0)("sModelNme")
                .Rows(0)("sBrandIDx") = loDT(0)("sBrandIDx")
                .Rows(0)("sBrandNme") = loDT(0)("sBrandNme")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sModelIDx»sModelNme»sBrandNme", _
                                "Model ID»Model»Brand", _
                                "", _
                                "a.sModelIDx»a.sModelNme»b.sBrandNme", _
                                2)

            If Not IsNothing(loDataRow) Then
                With p_oDTDetail
                    .Rows(0)("sModelIDx") = loDataRow("sModelIDx")
                    .Rows(0)("sModelNme") = loDataRow("sModelNme")
                    .Rows(0)("sBrandIDx") = loDataRow("sBrandIDx")
                    .Rows(0)("sBrandNme") = loDataRow("sBrandNme")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTDetail
            RaiseEvent DetailRetrieved(3, .Rows(0)("sModelNme"))
            RaiseEvent DetailRetrieved(4, .Rows(0)("sBrandNme"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTDetail
            .Rows(0)("sBrandIDx") = ""
            .Rows(0)("sBrandNme") = ""
            .Rows(0)("sModelIDx") = ""
            .Rows(0)("sModelNme") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getColor(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getColor"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sColorNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sColorNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sColorIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Color, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMaster
                .Rows(0)("sColorCde") = loDT(0)("sColorIDx")
                .Rows(0)("sColorNme") = loDT(0)("sColorNme")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sColorIDx»sColorNme", _
                                "Color ID»Color", _
                                "", _
                                "", _
                                5)

            If Not IsNothing(loDataRow) Then
                With p_oDTDetail
                    .Rows(0)("sColorIDx") = loDataRow("sColorIDx")
                    .Rows(0)("sColorNme") = loDataRow("sColorNme")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTDetail
            RaiseEvent DetailRetrieved(5, .Rows(0)("sColorNme"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTDetail
            .Rows(0)("sColorIDx") = ""
            .Rows(0)("sColorNme") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getSQL_Serial() As String
        Return "SELECT" & _
                    "  a.sSerialID" & _
                    ", a.sEngineNo" & _
                    ", a.sFrameNox" & _
                    ", b.sBrandNme" & _
                    ", c.sModelNme" & _
                    ", d.sColorNme" & _
                    ", a.nYearModl" & _
                    ", a.sFileNoxx" & _
                    ", a.sCRENoxxx" & _
                    ", a.sCRNoxxxx" & _
                    ", a.sPlateNoP" & _
                    ", a.dRegister" & _
                " FROM Car_Serial a" & _
                    " LEFT JOIN Car_Brand b" & _
                        " ON a.sBrandCde = b.sBrandIDx" & _
                    " LEFT JOIN Car_Model c" & _
                        " ON a.sModelCde = c.sModelIDx" & _
                    " LEFT JOIN Color d" & _
                        " ON a.sColorCde = d.sColorIDx"

    End Function

    Private Function getSQL_Browse() As String
        Return "SELECT a.sAcctNmbr" & _
                    ", CONCAT(c.sFrstName, ', ', c.sLastName, ' ', c.sMiddName) sClientNm" & _
                    ", d.sEngineNo" & _
                    ", d.sPlateNoP" & _
            " FROM LR_Master a" & _
                ", LR_Master_Car b" & _
                ", Client_Master c" & _
                ", Car_Serial d" & _
           " WHERE a.sAcctNmbr = b.sAcctNmbr" & _
                " AND a.sClientID = c.sClientID" & _
                " AND b.sSerialID = d.sSerialID"
    End Function
#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub New(ByVal foRider As GRider)
        p_oAppDrvr = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "LRApplication"

        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 2 ' sClientNm
                getClient(fsValue, False, True)
        End Select
    End Sub

    Public Sub SearchDetail(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 3 ' sModelNme
                getModel(fsValue, True, False)
            Case 5 ' sColorNme
                getColor(fsValue, True, False)
        End Select
    End Sub
End Class