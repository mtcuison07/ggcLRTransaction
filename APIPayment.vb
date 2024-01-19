'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     API Payment Object
'
' Copyright 2012 and Beyond
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
'  Mac [ 06/23/2020 05:30 pm ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit On

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient
Imports System.Drawing
Imports Newtonsoft.Json.Linq
Imports System.IO

Public Class APIPayment
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetx As DataTable

    Private p_cTranType As String
    Private p_sParent As String

    Private Const pxeModuleName As String = "APIPayment"
    Private Const pxeSourceCode As String = "xAPI"

    Private Const p_sMasTable As String = "XAPITrans"
    Private Const p_sMsgHeadr As String = "API Payment"

    Private Const p_sDefDatex As String = "1900-01-01"
    Private Const p_sExportDIR As String = "D:\APIPayment.xlsx"

    Private p_oExcel As Object = Nothing
    Private p_oBook As Object = Nothing
    Private p_oSheet As Object = Nothing

    Private p_dTransact As Date
    Private p_sAcctNmbr As String
    Private p_sClientNm As String
    Private p_sPartnerx As String

    Private p_sDateFrom As String
    Private p_sDateThru As String

    Public Event PaymentRetreive(ByVal fnIndex As Integer)
    Public Event MaxRecord(ByVal fnRecord As Integer)
    Public Event FirstRecord()
    Public Event NextRecord()

    Public Property AppDriver() As GRider
        Get
            Return p_oApp
        End Get
        Set(ByVal value As GRider)
            p_oApp = value
        End Set
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    Public ReadOnly Property ItemCount()
        Get
            Return p_oDTDetx.Rows.Count
        End Get
    End Property

    Public WriteOnly Property DatePaid() As Date
        Set(ByVal value As Date)
            p_dTransact = value
        End Set
    End Property

    Public WriteOnly Property AccountNo() As String
        Set(ByVal value As String)
            p_sAcctNmbr = value
        End Set
    End Property

    Public WriteOnly Property Partner() As String
        Set(ByVal value As String)
            p_sPartnerx = value
        End Set
    End Property

    Public WriteOnly Property ClientName() As String
        Set(ByVal value As String)
            p_sClientNm = value
        End Set
    End Property

    Public Property Master(ByVal fnRow As Integer, ByVal foIndex As Object) As Object
        Get
            If Not IsNumeric(foIndex) Then
                Select Case LCase(foIndex)
                    Case "stransnox"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "slognoxxx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sxlientid"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sxapicode"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "srefernox"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sacctnmbr"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "spayloadx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "ctranstat"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dreceived"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dcaptured"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dcilledxx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dpaidxxxx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dcancelld"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sbranchnm"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "spartnerx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sacctnmbr"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sclientnm"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "namountxx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "dtransact"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sornoxxxx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "saddressx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "namtpaidx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "nrebatesx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "sremarksx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case "npenaltyx"
                        Return p_oDTDetx(fnRow)(foIndex)
                    Case Else
                        Return DBNull.Value
                End Select
            Else
                Return p_oDTDetx(fnRow)(foIndex)
            End If
        End Get

        Set(ByVal foValue As Object)
            If p_oDTDetx(fnRow)("cTranStat") <> "0" Then Exit Property

            If Not IsNumeric(foIndex) Then
                Select Case LCase(foIndex)
                    Case "sornoxxxx"
                        p_oDTDetx(fnRow)(foIndex) = foValue
                        RaiseEvent PaymentRetreive(3)
                    Case "sremarksx"
                        p_oDTDetx(fnRow)(foIndex) = foValue
                        RaiseEvent PaymentRetreive(7)
                    Case "namtpaidx"
                        p_oDTDetx(fnRow)(foIndex) = CDbl(foValue)
                        RaiseEvent PaymentRetreive(8)
                    Case "nrebatesx"
                        p_oDTDetx(fnRow)(foIndex) = CDbl(foValue)
                        RaiseEvent PaymentRetreive(10)
                    Case "npenaltyx"
                        p_oDTDetx(fnRow)(foIndex) = CDbl(foValue)
                        RaiseEvent PaymentRetreive(11)
                End Select
            End If
        End Set
    End Property

    Public Function SearchAccount(ByVal fnRow As Integer, ByVal fsValue As String, ByVal fbByCode As Boolean) As Boolean
        Dim lsSQL As String = getSQ_Account()
        Dim loRow As DataRow

        If fnRow <= 0 Then Return False

        If fsValue = "" Then Return False

        loRow = KwikSearch(p_oApp _
                                , lsSQL _
                                , True _
                                , fsValue _
                                , "sAcctNmbr»sClientNm»xAddressx" _
                                , "Account No»Client»Address", _
                                , "a.sAcctNmbr»b.sCompnyNm»CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode)" _
                                , IIf(fbByCode, 0, 1))

        If IsNothing(loRow) Then
            Master(fnRow, "sAcctNmbr") = ""
            Return False
        End If

        p_oDTDetx(fnRow)("sAcctNmbr") = loRow.Item("sAcctNmbr")
        p_oDTDetx(fnRow)("sClientXX") = loRow.Item("sClientID")

        Return True
    End Function

    Public Function SearchPartner(ByVal fsValue As String) As String
        Return getBranch(fsValue, False)
    End Function

    Public Sub New()
        p_dTransact = CDate(p_sDefDatex)
        p_sAcctNmbr = ""
        p_sClientNm = ""
        p_sPartnerx = ""

        p_oDTMstr = New DataTable
        p_oDTDetx = New DataTable
    End Sub


    Public Function LoadTransaction() As Boolean
        If p_sDateFrom = "" Or p_sDateThru = "" Then
            MsgBox("Either DATE FROM or DATE THRU is empty.")
            Return False
        End If

        Return LoadTransaction(p_sDateFrom, p_sDateThru)
    End Function

    Public Function LoadTransaction(ByVal fsDateFrom As String, ByVal fsDateThru As String) As Boolean
        Dim lsSQL As String = getSQ_Master()

        lsSQL = AddCondition(lsSQL, "dReceived BETWEEN " & strParm(fsDateFrom & " 00:00:00") & " AND " & strParm(fsDateThru & " 23:59:30"))

        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        p_sDateFrom = fsDateFrom
        p_sDateThru = fsDateThru

        If p_oDTMstr.Rows.Count = 0 Then
            MsgBox("No record found.", MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        p_oDTDetx = New DataTable

        Dim lnCtr As Integer
        Dim loJSON As JObject
        Dim lsValue As String

        Dim lsID As String
        Dim lsName As String
        Dim lsAddress As String

        RaiseEvent MaxRecord(p_oDTMstr.Rows.Count)
        RaiseEvent FirstRecord()

        p_oDTMstr.Columns.Add("sBranchNm", GetType(String))
        p_oDTMstr.Columns.Add("sPartnerx", GetType(String))
        p_oDTMstr.Columns.Add("nAmountxx", GetType(Double))
        p_oDTMstr.Columns.Add("dTransact", GetType(Date))
        p_oDTMstr.Columns.Add("sClientXX", GetType(String))
        p_oDTMstr.Columns.Add("sClientNm", GetType(String))
        p_oDTMstr.Columns.Add("sAddressx", GetType(String))

        p_oDTMstr.Columns.Add("sORNoxxxx", GetType(String))
        p_oDTMstr.Columns.Add("nAmtPaidx", GetType(Double))
        p_oDTMstr.Columns.Add("nRebatesx", GetType(Double))
        p_oDTMstr.Columns.Add("sRemarksx", GetType(String))

        'mac 2020.10.19
        p_oDTMstr.Columns.Add("nPenaltyx", GetType(Double))

        For lnCtr = 0 To p_oDTMstr.Rows.Count - 1
            p_oDTMstr(lnCtr)("sBranchNm") = getBranch(IFNull(p_oDTMstr(lnCtr)("sClientID")), True)
            p_oDTMstr(lnCtr)("sPartnerx") = getBranch(IFNull(p_oDTMstr(lnCtr)("sClientID")), True)

            'extract sPayLoadx JSON
            lsValue = p_oDTMstr(lnCtr)("sPayloadx")
            loJSON = JObject.Parse(lsValue)

            If p_oDTMstr(lnCtr)("sAcctNmbr") = "" Then
                p_oDTMstr(lnCtr)("sAcctNmbr") = CStr(loJSON.GetValue("account"))
            End If
            p_oDTMstr(lnCtr)("nAmountxx") = CDbl(CStr(loJSON.GetValue("amount")))
            p_oDTMstr(lnCtr)("dTransact") = Format(CDate(loJSON.GetValue("datetime")), "yyyy-MM-dd")

            lsID = ""
            lsName = ""
            lsAddress = ""

            If getClient(p_oDTMstr(lnCtr)("sAcctNmbr"), lsName, lsAddress, lsID) Then
                p_oDTMstr(lnCtr)("sClientXX") = lsID
                p_oDTMstr(lnCtr)("sClientNm") = lsName
                p_oDTMstr(lnCtr)("sAddressx") = lsAddress
            Else
                p_oDTMstr(lnCtr)("sClientNm") = CStr(loJSON.GetValue("name"))
            End If

            getReceiptInfo(lnCtr, p_oDTMstr(lnCtr)("sAcctNmbr"), p_oDTMstr(lnCtr)("sTransNox"))

            RaiseEvent NextRecord()
        Next

        p_oDTMstr.DefaultView.Sort = "dTransact, sAcctNmbr, sClientNm, sReferNox"

        Return True
    End Function

    Private Function getReceiptInfo(ByVal fnRow As Integer, ByVal fsAcctNmbr As String, ByVal fsReferNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = "SELECT * FROM LR_Payment_Master" & _
                " WHERE sSourceCd = " & strParm(pxeSourceCode) & _
                    " AND sSourceNo = " & strParm(fsReferNox) & _
                    " AND sAcctNmbr = " & strParm(fsAcctNmbr) & _
                    " AND cPostedxx <> '3'"

        Dim loDT As DataTable = p_oApp.ExecuteQuery(lsSQL)

        p_oDTMstr(fnRow)("sORNoxxxx") = ""
        p_oDTMstr(fnRow)("nAmtPaidx") = 0.0#
        p_oDTMstr(fnRow)("nRebatesx") = 0.0#
        p_oDTMstr(fnRow)("nPenaltyx") = 0.0#
        p_oDTMstr(fnRow)("sRemarksx") = ""

        If loDT.Rows.Count = 1 Then
            p_oDTMstr(fnRow)("sORNoxxxx") = loDT(0)("sReferNox")
            p_oDTMstr(fnRow)("nAmtPaidx") = loDT(0)("nAmountxx")
            p_oDTMstr(fnRow)("nRebatesx") = loDT(0)("nRebatesx")
            p_oDTMstr(fnRow)("nPenaltyx") = loDT(0)("nPenaltyx")
            p_oDTMstr(fnRow)("sRemarksx") = loDT(0)("sRemarksx")
        End If

        loDT = Nothing

        Return True
    End Function

    Public Function UnFilter() As Boolean
        p_oDTDetx = New DataTable
        p_oDTDetx = p_oDTMstr.Clone

        Return True
    End Function

    Public Function Filter() As Boolean
        If TypeName(p_oDTMstr) = "Nothing" Then Return False

        Dim lsCondition = ""

        If p_sAcctNmbr <> "" Then
            lsCondition = "sAcctNmbr LIKE " & strParm(p_sAcctNmbr & "%")
        End If

        If CDate(p_dTransact) <> CDate(p_sDefDatex) Then
            lsCondition = IIf(lsCondition = "", "", lsCondition & " AND ") & "dTransact = " & dateParm(Format(p_dTransact, "yyyy-MM-dd"))
        End If

        If p_sClientNm <> "" Then
            lsCondition = IIf(lsCondition = "", "", lsCondition & " AND ") & "sClientNm LIKE " & strParm(p_sClientNm & "%")
        End If

        If p_sPartnerx <> "" Then
            lsCondition = IIf(lsCondition = "", "", lsCondition & " AND ") & "sBranchNm LIKE " & strParm(p_sPartnerx & "%")
        End If

        Dim loFilter() As DataRow = p_oDTMstr.Select(lsCondition)

        p_oDTDetx = New DataTable

        If loFilter.Length <> 0 Then
            p_oDTDetx = loFilter.CopyToDataTable
        End If

        Return True
    End Function

    Public Function ReleaseOR() As Boolean
        Dim loMCPayment As ARPayment
        Dim loMPPayment As ARPayment_MP

        p_oApp.BeginTransaction()

        Dim lnCtr As Integer
        Dim lnRow As Integer = 0
        Dim lsSQL As String

        For lnCtr = 0 To ItemCount - 1
            If p_oDTDetx(lnCtr)("cTranStat") = "0" And Trim(p_oDTDetx(lnCtr)("sORNoxxxx")) <> "" Then
                If Strings.Left(p_oDTDetx(lnCtr)("sAcctNmbr"), 1).ToLower = "m" Then
                    loMCPayment = New ARPayment(p_oApp, "2")
                    loMCPayment.Parent = "APIPayment"

                    If loMCPayment.NewTransaction() Then
                        loMCPayment.Master("sAcctNmbr") = p_oDTDetx(lnCtr)("sAcctNmbr")
                        loMCPayment.Master("sClientID") = p_oDTDetx(lnCtr)("sClientXX")
                        loMCPayment.Master("dTransact") = p_oDTDetx(lnCtr)("dTransact")
                        loMCPayment.Master("sReferNox") = p_oDTDetx(lnCtr)("sORNoxxxx")

                        loMCPayment.Master("nAmountxx") = p_oDTDetx(lnCtr)("nAmtPaidx")
                        loMCPayment.Master("nRebatesx") = p_oDTDetx(lnCtr)("nRebatesx")
                        loMCPayment.Master("nPenaltyx") = p_oDTDetx(lnCtr)("nPenaltyx")
                        loMCPayment.Master("sSourceNo") = p_oDTDetx(lnCtr)("sTransNox")
                        loMCPayment.Master("sSourceCd") = pxeSourceCode

                        loMCPayment.Master("sRemarksx") = p_oDTDetx(lnCtr)("sRemarksx")

                        If Not loMCPayment.SaveTransaction Then
                            MsgBox("Unable to save MC Payment Info.", vbInformation, "Warning")
                            GoTo endwithRoll
                        End If
                    End If
                Else
                    loMPPayment = New ARPayment_MP(p_oApp, "2")
                    loMPPayment.Parent = "APIPayment"

                    If loMPPayment.NewTransaction() Then
                        loMPPayment.Master("sAcctNmbr") = p_oDTDetx(lnCtr)("sAcctNmbr")
                        loMPPayment.Master("sClientID") = p_oDTDetx(lnCtr)("sClientXX")
                        loMPPayment.Master("dTransact") = p_oDTDetx(lnCtr)("dTransact")
                        loMPPayment.Master("sReferNox") = p_oDTDetx(lnCtr)("sORNoxxxx")

                        loMPPayment.Master("nAmountxx") = p_oDTDetx(lnCtr)("nAmtPaidx")
                        loMPPayment.Master("nRebatesx") = p_oDTDetx(lnCtr)("nRebatesx")
                        loMPPayment.Master("nPenaltyx") = p_oDTDetx(lnCtr)("nPenaltyx")
                        loMPPayment.Master("sSourceNo") = p_oDTDetx(lnCtr)("sTransNox")
                        loMPPayment.Master("sSourceCd") = pxeSourceCode

                        loMPPayment.Master("sRemarksx") = p_oDTDetx(lnCtr)("sRemarksx")

                        If Not loMPPayment.SaveTransaction Then
                            MsgBox("Unable to save MC Payment Info.", vbInformation, "Warning")
                            GoTo endwithRoll
                        End If
                    End If
                End If

                lsSQL = "UPDATE XAPITrans SET" & _
                            "  sAcctNmbr = " & strParm(p_oDTDetx(lnCtr)("sAcctNmbr")) & _
                            ", cTranStat = '1'" & _
                            ", dCaptured = " & datetimeParm(p_oApp.SysDate) & _
                        " WHERE sTransNox = " & strParm(p_oDTDetx(lnCtr)("sTransNox"))

                If p_oApp.Execute(lsSQL, "XAPITrans") <= 0 Then
                    MsgBox("Unable to save API Payment Info.", vbInformation, "Warning")
                    GoTo endwithRoll
                End If

                lnRow += 1
            End If
        Next

        p_oApp.CommitTransaction()

        If lnRow = 0 Then
            MsgBox("No rows affected.", vbInformation, "Notice")
        Else
            MsgBox("OR Released Successfuly.", vbInformation, "Notice")
        End If

        Return True
endwithRoll:
        p_oApp.RollBackTransaction()
        MsgBox("Exception detected.", vbExclamation, "Warning")
        Return False
    End Function

    Private Function getSQ_Account() As String
        Return "SELECT" & _
                       "  a.sAcctNmbr" & _
                       ", b.sCompnyNm sClientNm" & _
                       ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                       ", a.nPNValuex" & _
                       ", a.nDownPaym" & _
                       ", a.nGrossPrc" & _
                       ", a.nMonAmort" & _
                       ", a.nCashBalx" & _
                       ", a.nAcctTerm" & _
                       ", a.nABalance" & _
                       ", a.nAmtDuexx" & _
                       ", a.nRebatesx" & _
                       ", a.sClientID" & _
                       ", IFNULL(e.sEngineNo, '') sEngineNo" & _
                       ", IFNULL(e.sFrameNox, '') sFrameNox" & _
                       ", IFNULL(f.sModelNme, '') sModelNme" & _
                       ", IFNULL(g.sColorNme, '') sColorNme" & _
                       ", IFNULL(h.sGCardNox, '') sGCardNox" & _
                       ", IFNULL(i.cDigitalx, '') cDigitalx" & _
               " FROM MC_AR_Master a" & _
                    " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" & _
                    " LEFT JOIN TownCity c ON b.sTownIDxx = c.sTownIDxx" & _
                    " LEFT JOIN Province d ON c.sProvIDxx = d.sProvIDxx" & _
                    " LEFT JOIN MC_Serial e ON a.sSerialID = e.sSerialID" & _
                    " LEFT JOIN MC_Model f ON e.sModelIDx = f.sModelIDx" & _
                    " LEFT JOIN Color g ON e.sColorIDx = g.sColorIDx" & _
                    " LEFT JOIN MC_Serial_Service h ON h.sSerialID = a.sSerialID" & _
                    " LEFT JOIN G_Card_Master i ON h.sGCardNox = i.sGCardNox AND i.cCardStat = '4'" & _
                " WHERE nAcctTerm > 0"
    End Function

    Private Function getSQ_Master() As String
        Dim lsSQL As String = "SELECT" & _
                                    "  sTransNox" & _
                                    ", sLogNoxxx" & _
                                    ", sClientID" & _
                                    ", sXAPICode" & _
                                    ", sReferNox" & _
                                    ", sAcctNmbr" & _
                                    ", sPayloadx" & _
                                    ", cTranStat" & _
                                    ", dReceived" & _
                                    ", dCaptured" & _
                                    ", dBilledxx" & _
                                    ", dPaidxxxx" & _
                                    ", dCancelld" & _
                                " FROM " & p_sMasTable & _
                                " WHERE (sAcctNmbr NOT LIKE '____GK1946' OR IFNULL(sAcctNmbr, '') = '')"

        If p_cTranType = "" Then
            Return lsSQL
        Else
            Return ""
        End If
    End Function

    Private Function getBranch(ByVal fsValue As String, ByVal fbByCode As Boolean) As String
        If fsValue = "" Then Return ""

        Dim lsSQL As String = "SELECT" & _
                                    "  sClientID" & _
                                    ", sClientNm" & _
                                    ", sBranchCd" & _
                                " FROM xxxSysClient"

        Dim lsCondition As String = ""

        If fbByCode Then
            lsCondition = "sClientID = " & strParm(fsValue)
        Else
            lsCondition = "sClientNm LIKE " & strParm(fsValue & "%")
        End If

        Debug.Print(AddCondition(lsSQL, lsCondition))
        Dim loRS As DataTable = p_oApp.ExecuteQuery(AddCondition(lsSQL, lsCondition))

        If loRS.Rows.Count = 0 Then
            Return ""
        ElseIf loRS.Rows.Count = 1 Then
            Return loRS(0)("sClientNm")
        Else
            Dim loDR As DataRow = KwikSearch(p_oApp, lsSQL, fsValue, "sClientID»sClientNm»sBranchCd", "Client ID»Branch Name»Branch Code", , "sClientID»sClientNm»sBranchCd", IIf(fbByCode, 0, 1))

            Try
                Return loDR("sClientNm")
            Catch ex As Exception
                Return ""
            End Try
        End If
    End Function

    Private Function getClient(ByVal fsAcctNmbr As String, ByRef fsClientNm As String, ByRef fsAddressx As String, ByRef fsClientID As String) As Boolean
        If fsAcctNmbr = "" Then Return False

        Dim lsSQL As String = "SELECT" & _
                                    "  b.sClientID" & _
                                    ", b.sCompnyNm" & _
                                    ", TRIM(CONCAT(IFNULL(b.sHouseNox, ''), ' ', b.sAddressx, ', ', IF(IFNULL(c.sBrgyName, '') <> '', CONCAT(c.sBrgyName, ', ', d.sTownName, ' ', e.sProvName), CONCAT(f.sTownName, ' ', g.sProvName)))) xAddressx" & _
                                " FROM MC_AR_Master a" & _
                                    ", Client_Master b" & _
                                        " LEFT JOIN Barangay c ON b.sBrgyIDxx = c.sBrgyIDxx" & _
                                        " LEFT JOIN TownCity d ON c.sTownIDxx = d.sTownIDxx" & _
                                        " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
                                        " LEFT JOIN TownCity f ON b.sTownIDxx = f.sTownIDxx" & _
                                        " LEFT JOIN Province g ON f.sProvIDxx = g.sProvIDxx" & _
                                " WHERE a.sClientID = b.sClientID" & _
                                    " AND a.sAcctNmbr = " & strParm(fsAcctNmbr)

        Dim loRS As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If loRS.Rows.Count = 0 Then
            Return False
        End If

        fsClientID = loRS(0)("sClientID")
        fsClientNm = loRS(0)("sCompnyNm")
        fsAddressx = IFNull(loRS(0)("xAddressx"))

        Return True
    End Function

    Public Function Export() As Boolean
        Dim lnCtr As Integer
        Dim lbSuccess As Boolean = False

        Try
            If Not OpenWorkBook(p_oApp.AppPath & "/Reports/APIPayment.xlsx") Then Return False

            'export data
            RaiseEvent MaxRecord(ItemCount)
            RaiseEvent FirstRecord()

            For lnCtr = 0 To ItemCount - 1
                RaiseEvent NextRecord()
                p_oSheet.Range("A" & lnCtr + 2).value = Format(CDate(Master(lnCtr, "dTransact")), "yyyy-MM-dd")
                p_oSheet.Range("B" & lnCtr + 2).value = Master(lnCtr, "sAcctNmbr")
                p_oSheet.Range("C" & lnCtr + 2).value = Master(lnCtr, "sClientNm")
                p_oSheet.Range("D" & lnCtr + 2).value = Master(lnCtr, "sAddressx")
                p_oSheet.Range("E" & lnCtr + 2).value = Master(lnCtr, "sPartnerx")
                p_oSheet.Range("F" & lnCtr + 2).value = Master(lnCtr, "sReferNox")
                p_oSheet.Range("G" & lnCtr + 2).value = Format(CDbl(Master(lnCtr, "nAmountxx")), "#,##0.00")
                p_oSheet.Range("H" & lnCtr + 2).value = Master(lnCtr, "sORNoxxxx")


            Next

            p_oBook.SaveAs("D:\APIPayment.xlsx")
            lbSuccess = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Warning")
        Finally
            closeWorkBook()
        End Try

        Return lbSuccess
    End Function

    Private Function OpenWorkBook(ByVal lsFile As String) As Boolean
        If lsFile = "" Then
            MsgBox("Invalid filename detected.", vbCritical, "Warning")
            Return False
        End If

        If Not File.Exists(lsFile) Then
            MsgBox("File does not exist.", vbCritical, "Warning")
            Return False
        End If

        Try
            p_oExcel = CreateObject("Excel.Application")
            p_oBook = p_oExcel.Workbooks.Open(lsFile)
            p_oSheet = p_oBook.Worksheets(1)
            Debug.Print("Excel File Opened.")
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, pxeModuleName)

            releaseObject(p_oExcel)
            releaseObject(p_oBook)
            releaseObject(p_oExcel)

            Return False
        End Try

        Return True
    End Function

    Public Sub closeWorkBook()
        If Not IsNothing(p_oBook) Then p_oBook.Close()
        If Not IsNothing(p_oExcel) Then p_oExcel.Quit()

        releaseObject(p_oExcel)
        releaseObject(p_oBook)
        releaseObject(p_oExcel)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
