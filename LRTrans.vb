'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Trans Object
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
'  Kalyptus [ 06/07/2016 03:45 pm ]
'      Started creating this object.
'  Questions: 
'       1. What is the difference between sSourceCD and cTranType if all transactions are coming from LR_Payment_Master
'       2. Split the transaction amount and interest amount
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class LRTrans
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private Const p_sMasTable As String = "LR_Master"
    Private Const p_sDtlTable As String = "LR_Ledger"

    Private Const p_sMsgHeadr As String = "LR Transaction"

    Private p_dTransact As Date
    Private p_bIsOffice As Boolean
    Private p_sCollctID As String
    Private p_sRemarksx As String
    Private p_nTranAmtx As Decimal
    Private p_sSourceNo As String
    Private p_sSourceCD As String
    Private p_sReferNox As String
    Private p_sAcctNmbr As String

    Private p_nPaidAmtx As Decimal
    Private p_nIntAmtxx As Decimal
    Private p_nRebatesx As Decimal
    Private p_nPenaltyx As Decimal
    Private p_nInsChrge As Decimal

    Private p_nDebitAmt As Decimal
    Private p_nCredtAmt As Decimal
    Private p_nAmtDuexx As Decimal
    Private p_nMonDelay As Integer
    Private p_cTrantype As String

    Public Property AccountNo() As String
        Get
            Return p_sAcctNmbr
        End Get
        Set(ByVal value As String)
            p_sAcctNmbr = value
        End Set
    End Property

    Public Property Transact_Date() As Date
        Get
            Return p_dTransact
        End Get
        Set(ByVal value As Date)
            p_dTransact = value
        End Set
    End Property

    Public Property isOffice() As Boolean
        Get
            Return p_bIsOffice
        End Get
        Set(ByVal value As Boolean)
            p_bIsOffice = value
        End Set
    End Property

    Public Property Collector() As String
        Get
            Return p_sCollctID
        End Get
        Set(ByVal value As String)
            p_sCollctID = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return p_sRemarksx
        End Get
        Set(ByVal value As String)
            p_sRemarksx = value
        End Set
    End Property

    Public Property Amount() As Decimal
        Get
            Return p_nTranAmtx
        End Get
        Set(ByVal value As Decimal)
            p_nTranAmtx = value
        End Set
    End Property

    Public Property Interest() As Decimal
        Get
            Return p_nIntAmtxx
        End Get
        Set(ByVal value As Decimal)
            p_nIntAmtxx = value
        End Set
    End Property

    Public Property Penalty() As Decimal
        Get
            Return p_nPenaltyx
        End Get
        Set(ByVal value As Decimal)
            p_nPenaltyx = value
        End Set
    End Property

    Public Property Rebates() As Decimal
        Get
            Return p_nRebatesx
        End Get
        Set(ByVal value As Decimal)
            p_nRebatesx = value
        End Set
    End Property

    Public Property Insurance() As Decimal
        Get
            Return p_nInsChrge
        End Get
        Set(ByVal value As Decimal)
            p_nInsChrge = value
        End Set
    End Property

    Public Property SourceNo() As String
        Get
            Return p_sSourceNo
        End Get
        Set(ByVal value As String)
            p_sSourceNo = value
        End Set
    End Property

    Public Property ReferNo() As String
        Get
            Return p_sReferNox
        End Get
        Set(ByVal value As String)
            p_sReferNox = value
        End Set
    End Property

    Public Function Debit() As Boolean
        p_sSourceCD = "DEBT"

        Return SaveTransaction()
    End Function

    Public Function Credit() As Boolean
        p_sSourceCD = "CRDT"

        Return SaveTransaction()
    End Function

    Public Function Payment() As Boolean
        p_sSourceCD = "PYMT"

        Return SaveTransaction()
    End Function

    Private Function SaveTransaction() As Boolean
        Dim ldClosedxx As String

        Try
            p_oDTMstr = GetMaster()

            'Check for the dLastPaym since we will not allow transactions below the last payment date
            If p_oDTMstr(0).Item("dLastPaym") > p_dTransact Then
                MsgBox("Transaction date is prior to the last transaction date!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            ldClosedxx = ""

            Select Case p_sSourceCD
                Case "PYMT"
                    p_nPaidAmtx = p_nTranAmtx
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nPaidAmtx
                    p_cTrantype = "0"
                    p_nDebitAmt = 0
                    p_nCredtAmt = 0
                    p_nMonDelay = getDelay(p_oDTMstr, p_dTransact)
                    p_nAmtDuexx = p_nMonDelay * p_oDTMstr(0).Item("nMonAmort")
                    p_nInsChrge = p_oDTMstr(0).Item("nInsChrge")

                    If p_oDTMstr(0).Item("nABalance") <= 0 And p_oDTMstr(0).Item("nABalance") + p_nTranAmtx > 0 Then
                        ldClosedxx = dateParm(p_dTransact)
                    End If

                Case "CRDT"
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nTranAmtx
                    p_nIntAmtxx = 0
                    p_nPaidAmtx = 0
                    p_cTrantype = "1"
                    p_nPenaltyx = 0
                    p_nDebitAmt = 0
                    p_nCredtAmt = p_nTranAmtx
                    p_nMonDelay = getDelay(p_oDTMstr, p_dTransact)
                    p_nAmtDuexx = p_nMonDelay * p_oDTMstr(0).Item("nMonAmort")

                    If p_oDTMstr(0).Item("nABalance") <= 0 And p_oDTMstr(0).Item("nABalance") + p_nTranAmtx > 0 Then
                        ldClosedxx = dateParm(p_dTransact)
                    End If
                Case "DEBT"
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") + p_nTranAmtx
                    p_nIntAmtxx = 0
                    p_nPaidAmtx = 0
                    p_cTrantype = "2"
                    p_nPenaltyx = 0
                    p_nDebitAmt = p_nTranAmtx
                    p_nCredtAmt = 0
                    p_nMonDelay = getDelay(p_oDTMstr, p_dTransact)
                    p_nAmtDuexx = p_nMonDelay * p_oDTMstr(0).Item("nMonAmort")
                Case "PLTY"
                    p_nIntAmtxx = 0
                    p_nPaidAmtx = 0
                    p_cTrantype = "3"
                    p_nPenaltyx = p_nTranAmtx
                    p_nDebitAmt = 0
                    p_nCredtAmt = 0
                    p_nMonDelay = getDelay(p_oDTMstr, p_dTransact)
                    p_nAmtDuexx = p_nMonDelay * p_oDTMstr(0).Item("nMonAmort")
            End Select

            'Create the SQL Statement to insert the transaction to the LR_Ledger
            Dim lsSQLLdgr As String

            lsSQLLdgr = "INSERT INTO LR_Ledger" & _
                       " SET sAcctNmbr = " & strParm(p_sAcctNmbr) & _
                          ", sBranchCD = " & strParm(Left(p_sSourceNo, 4)) & _
                          ", nEntryNox = " & p_oDTMstr(0).Item("nLedgerNo") + 1 & _
                          ", dTransact = " & dateParm(p_dTransact) & _
                          ", cOffPaymx = " & strParm(IIf(p_bIsOffice, "1", "0")) & _
                          ", sCollIDxx = " & strParm(p_sCollctID) & _
                          ", sReferNox = " & strParm(p_sReferNox) & _
                          ", sRemarksx = " & strParm(p_sRemarksx) & _
                          ", cTrantype = " & strParm(p_cTrantype) & _
                          ", nPaidAmtx = " & p_nPaidAmtx & _
                          ", nIntAmtxx = " & p_nIntAmtxx & _
                          ", nPenaltyx = " & p_nPenaltyx & _
                          ", nDebitAmt = " & p_nDebitAmt & _
                          ", nCredtAmt = " & p_nCredtAmt & _
                          ", nAmtDuexx = " & p_nAmtDuexx & _
                          ", nRebatesx = " & p_nRebatesx & _
                          ", nABalance = " & p_oDTMstr(0).Item("nABalance") & _
                          ", nMonDelay = " & Math.Round(p_nMonDelay, 2) & _
                          ", sSourceCd = " & strParm(p_sSourceCD) & _
                          ", sSourceNo = " & strParm(p_sSourceNo) & _
                          ", dModified = " & dateParm(p_oApp.getSysDate) & _
                          ", cPostedxx = '0'"

            'Create SQL Statement that will update the LR_Master
            Dim lsSQLMstr As String
            lsSQLMstr = "UPDATE LR_Master" & _
                       " SET nIntTotal = nIntTotal + " & p_nIntAmtxx & _
                          ", nPaymTotl = nPaymTotl + " & p_nPaidAmtx & _
                          ", nPenTotlx = nPenTotlx + " & p_nPenaltyx & _
                          ", nDebtTotl = nDebtTotl + " & p_nDebitAmt & _
                          ", nCredTotl = nCredTotl + " & p_nCredtAmt & _
                          ", nRebTotlx = nRebTotlx + " & p_nRebatesx & _
                          ", nABalance = nABalance - " & ((p_nPaidAmtx + p_nCredtAmt) - p_nDebitAmt) & _
                          ", nAmtDuexx = " & p_nAmtDuexx & _
                          ", nLastPaym = " & ((p_nIntAmtxx + p_nPaidAmtx + p_nCredtAmt + p_nRebatesx) - p_nDebitAmt) & _
                          ", dLastPaym = " & dateParm(p_dTransact) & _
                          ", nLedgerNo = " & (p_oDTMstr(0).Item("nLedgerNo") + 1) & _
                       " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
            Call p_oApp.Execute(lsSQLLdgr, "LR_Ledger")
            Call p_oApp.Execute(lsSQLMstr, "LR_Master")

            If p_oDTMstr(0).Item("nABalance") <= 0 Then
                Dim lnDelayAvg As Single

                lnDelayAvg = getAveDelay(p_oDTMstr, p_dTransact)

                If ldClosedxx = "" Then
                    lsSQLMstr = "UPDATE LR_Master" & _
                               " SET dClosedxx = " & p_dTransact & _
                                  ", cAcctstat = " & strParm("1") & _
                                  ", nDelayAvg = " & lnDelayAvg & _
                                  ", cRatingxx = " & strParm(getRating(lnDelayAvg, "")) & _
                                  ", cActivexx = " & strParm("0") & _
                               " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
                Else
                    lsSQLMstr = "UPDATE LR_Master" & _
                               " SET cAcctstat = " & strParm("1") & _
                                  ", nDelayAvg = " & lnDelayAvg & _
                                  ", cRatingxx = " & strParm(getRating(lnDelayAvg, "")) & _
                                  ", cActivexx = " & strParm("0") & _
                               " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
                End If

                Call p_oApp.Execute(lsSQLMstr, "LR_Master")
            ElseIf p_oDTMstr(0).Item("nABalance") <= (p_oDTMstr(0).Item("nMonAmort") + 10) Then
                Dim lnDelayAvg As Single

                lnDelayAvg = getAveDelay(p_oDTMstr, p_dTransact)

                lsSQLMstr = "UPDATE LR_Master" & _
                           " SET nDelayAvg = " & lnDelayAvg & _
                              ", cRatingxx = " & strParm(getRating(lnDelayAvg, "")) & _
                           " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
                Call p_oApp.Execute(lsSQLMstr, "LR_Master")
            End If
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        Return True
    End Function

    Public Function GetMaster() As DataTable
        Return GetMaster(p_sAcctNmbr)
    End Function

    Private Function GetMaster(ByVal fsAcctNmbr As String) As DataTable
        Dim lsSQL As String

        Try
            lsSQL = "SELECT" & _
                           "  sAcctNmbr" & _
                           ", dTransact" & _
                           ", nPrincipl" & _
                           ", nInterest" & _
                           ", dFirstPay" & _
                           ", nAcctTerm" & _
                           ", nMonAmort" & _
                           ", nPenltyRt" & _
                           ", nLastPaym" & _
                           ", dLastPaym" & _
                           ", nPaymTotl" & _
                           ", nPenTotlx" & _
                           ", nDebtTotl" & _
                           ", nCredTotl" & _
                           ", nRebTotlx" & _
                           ", nAmtDuexx" & _
                           ", nABalance" & _
                           ", nDelayAvg" & _
                           ", cRatingxx" & _
                           ", cAcctstat" & _
                           ", dClosedxx" & _
                           ", cActivexx" & _
                           ", nLedgerNo" & _
                           ", sCollatID" & _
                           ", nIntTotal" & _
                           ", dDueDatex" & _
                           ", nIntRatex" & _
                           ", nInsChrge" & _
                           ", nRebatesx" & _
                     " FROM " & p_sMasTable & _
                     " WHERE sAcctNmbr = " & strParm(fsAcctNmbr)
            Return p_oApp.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
            Return Nothing
        End Try
    End Function

    Public Function getDelay(foMaster As DataTable, fdTransact As Date) As Single
        Dim ldDueDate As Date
        Dim lnAmtDuex As Double
        Dim lnActTerm As Single

        With foMaster(0)
            ldDueDate = fdTransact
            If fdTransact > .Item("dDueDatex") Then ldDueDate = .Item("dDueDatex")

            lnActTerm = getMonthTerm(.Item("dFirstPay"), ldDueDate)

            lnAmtDuex = (.Item("nMonAmort") * lnActTerm) + _
                        .Item("nDebtTotl")

            lnAmtDuex = lnAmtDuex - _
                        .Item("nPaymTotl") - _
                        .Item("nCredTotl")

            lnAmtDuex = IIf(lnAmtDuex < 0, 0, lnAmtDuex)

            getDelay = Math.Round(lnAmtDuex / .Item("nMonAmort"), 2)
        End With
    End Function

    Function getMonthTerm( _
            ByVal dFirstPay As Date, _
            ByVal dTransact As Date
            ) As Integer
        getMonthTerm = DateDiff(DateInterval.Month, dFirstPay, dTransact) + 1
        getMonthTerm = IIf(Day(dFirstPay) > Day(dTransact), getMonthTerm - 1, getMonthTerm)
    End Function

    Function getAveDelay(
            foMaster As DataTable, _
            dTransact As Date) As Double
        Dim lsSQL As String
        Dim loData As DataTable

        Dim lnDelayxxx As Double, lnTotDelay As Double
        Dim lnTranAmtx As Double

        Dim lanDayMon(11) As Integer, lnDayOfMon As Integer
        Dim lnCtr1 As Integer, lnCtr2 As Integer
        Dim ldTranDate As Date
        Dim lnDebtTotl As Double, lnCrdtTotl As Double


        With foMaster(0)
            ' Pastdue account's delay is based on the past due month
            If dTransact > .Item("dDueDatex") Then
                Return DateDiff(DateInterval.Month, .Item("dDueDatex"), dTransact)
            End If

            lsSQL = "SELECT" & _
                        "  dTransact" & _
                        ", nABalance" & _
                     " FROM LR_Ledger" & _
                     " WHERE sAcctNmbr = " & strParm(.Item("sAcctNmbr")) & _
                     " ORDER BY dTransact"

            loData = p_oApp.ExecuteQuery(lsSQL)

            If loData.Rows.Count = 0 Then
                Return getMonthTerm(.Item("dFirstPay"), dTransact)
            End If

            lnTotDelay = 0.0
            lnCrdtTotl = 0.0
            lnDebtTotl = 0.0

            lnTranAmtx = 0.0
            ldTranDate = .Item("dFirstPay")
            lnDayOfMon = Day(.Item("dFirstPay"))

            lnCtr2 = 0
            For lnCtr1 = 1 To .Item("nAcctTerm")
                lnTranAmtx = 0
                If lnCtr2 <= loData.Rows.Count - 1 Then

                    Do While DateDiff(DateInterval.Day, loData(lnCtr2).Item("dTransact"), ldTranDate) >= 1
                        Debug.Print(loData(lnCtr2).Item("dTransact"))
                        lnTranAmtx = (.Item("nMonAmort") * .Item("nAcctTerm")) - .Item("nABalance")
                        lnCtr2 = lnCtr2 + 1
                        'kalyptus - 2016.10.22 11:40am
                        'Change lnCtr2 >= loData.Rows.Count - 1 from lnCtr2 <= loData.Rows.Count - 1
                        If lnCtr2 >= loData.Rows.Count - 1 Then Exit Do
                    Loop
                End If

                lnDelayxxx = (lnCtr1 * .Item("nMonAmort") - lnTranAmtx) / .Item("nMonAmort")
                lnTotDelay = lnTotDelay + lnDelayxxx
                ldTranDate = DateAdd("m", lnCtr1, ldTranDate)
            Next
        End With

        lnTotDelay = Math.Round(lnTotDelay / lnCtr1, 2)
        If lnTotDelay < 0.0# Then
            Return 0.0
        Else
            Return lnTotDelay * (-1)
        End If
    End Function

    Function getRating(ByVal nDelay As Double, _
          ByVal cRating As String) As String

        If cRating = "l" Then
            ' Blacklisted account are account that are impounded...
            getRating = cRating
        ElseIf nDelay <= 0 Then
            nDelay = nDelay * (-1)
            getRating = "x"
            If nDelay > 0.5 Then
                getRating = "g"
            End If
        Else
            If nDelay < 6 Then
                getRating = "f"
            ElseIf nDelay < 12 Then
                getRating = "p"
            Else
                getRating = "b"
            End If
        End If
    End Function

    'kalyptus - 2017.05.20 01:16pm
    'Recalculates a particular LR Account
    Public Function Recalculate(ByVal fsAcctNmbr As String) As Boolean
        Dim loDtaMstr As DataTable = GetMaster(fsAcctNmbr)

        'If no record was found, exit immediately 
        If Not loDtaMstr.Rows.Count = 1 Then
            MsgBox("Unable to recalculate!" & vbCrLf & _
                   "No record found for " & p_sAcctNmbr & " ...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Dim lnPaymTotl As Decimal = 0
        Dim lnIntTotal As Decimal = 0
        Dim lnPenTotlx As Decimal = 0
        Dim lnDebtTotl As Decimal = 0
        Dim lnCredTotl As Decimal = 0
        Dim lnRebTotlx As Decimal = 0
        Dim lnABalance As Decimal = loDtaMstr(0).Item("nPrincipl")

        Dim lnLastPaym As Decimal = 0
        Dim ldLastPaym As Date = loDtaMstr(0).Item("dTransact")

        Dim lnMonDelay As Decimal = 0
        Dim lnAmtDuexx As Decimal = 0

        'kalyptus - 2022.10.14 01:49pm
        'Add ldClosedxx in recording what might be the actual closing date of the account
        Dim ldClosedxx As Object

        Try
            Dim lsSQL As String = "SELECT *" & _
                                 " FROM LR_Ledger" & _
                                 " WHERE sAcctnmbr = " & strParm(fsAcctNmbr) & _
                                 " ORDER BY sAcctNmbr, dTransact, nEntryNox"
            Dim loDtaLdgr As DataTable = p_oApp.ExecuteQuery(lsSQL)

            'If account has no ledger then exit immediately
            If loDtaLdgr.Rows.Count = 0 Then
                MsgBox("Unable to recalculate!" & vbCrLf & _
                       "No record found for the ledger of " & p_sAcctNmbr & " ...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            ldClosedxx = Nothing

            Dim lnRow As Integer
            For lnRow = 0 To loDtaLdgr.Rows.Count - 1
                lnPaymTotl = lnPaymTotl + loDtaLdgr(lnRow).Item("nPaidAmtx")
                lnIntTotal = lnIntTotal + loDtaLdgr(lnRow).Item("nIntAmtxx")
                lnPenTotlx = lnPenTotlx + loDtaLdgr(lnRow).Item("nPenaltyx")
                lnDebtTotl = lnDebtTotl + loDtaLdgr(lnRow).Item("nDebitAmt")
                lnCredTotl = lnCredTotl + loDtaLdgr(lnRow).Item("nCredtAmt")
                lnRebTotlx = lnRebTotlx + IFNull(loDtaLdgr(lnRow).Item("nRebatesx"), 0)
                lnABalance = lnABalance - loDtaLdgr(lnRow).Item("nPaidAmtx") - loDtaLdgr(lnRow).Item("nCredtAmt") + loDtaLdgr(lnRow).Item("nDebitAmt")

                'kalyptus - 2022.10.14 01:49pm
                'Set the closing date based on the transaction date of this transaction
                'if this transaction makes our balance equal or less than zero(0)
                If lnABalance <= 0 And Not IsDate(ldClosedxx) Then
                    ldClosedxx = loDtaLdgr(lnRow).Item("dTransact")
                End If

                loDtaLdgr(lnRow).Item("nABalance") = lnABalance
                loDtaLdgr(lnRow).Item("nEntryNox") = lnRow + 1
                lnMonDelay = getDelay(loDtaMstr, loDtaLdgr(lnRow).Item("dTransact"))
                lnAmtDuexx = lnMonDelay * loDtaMstr(0).Item("nMonAmort")
                loDtaLdgr(lnRow).Item("nMonDelay") = Math.Round(lnMonDelay, 2)
                loDtaLdgr(lnRow).Item("nAmtDuexx") = lnAmtDuexx

                lsSQL = ADO2SQL(loDtaLdgr _
                              , lnRow _
                              , "LR_Ledger" _
                              , "sAcctNmbr = " & strParm(fsAcctNmbr) & " AND sSourceCd = " & strParm(loDtaLdgr(lnRow).Item("sSourceCd")) & " AND sSourceNo = " & strParm(loDtaLdgr(lnRow).Item("sSourceNo")))
                'Save the detail of there are updates in the ledger...
                If lsSQL <> "" Then
                    p_oApp.Execute(lsSQL, "LR_Ledger")
                End If

                'Last transaction should be from payment transactor penalty only
                If loDtaLdgr(lnRow).Item("cTrantype") = 0 Or loDtaLdgr(lnRow).Item("cTrantype") = "3" Then
                    ldLastPaym = loDtaLdgr(lnRow).Item("dTransact")
                    lnLastPaym = loDtaLdgr(lnRow).Item("nPaidAmtx") + loDtaLdgr(lnRow).Item("nIntAmtxx") + loDtaLdgr(lnRow).Item("nPenaltyx")
                End If
            Next

            loDtaMstr(0).Item("nABalance") = lnABalance
            loDtaMstr(0).Item("nLedgerNo") = loDtaLdgr.Rows.Count
            loDtaMstr(0).Item("nAmtDuexx") = lnAmtDuexx

            loDtaMstr(0).Item("nPaymTotl") = lnPaymTotl
            loDtaMstr(0).Item("nIntTotal") = lnIntTotal
            loDtaMstr(0).Item("nPenTotlx") = lnPenTotlx
            loDtaMstr(0).Item("nDebtTotl") = lnDebtTotl
            loDtaMstr(0).Item("nCredTotl") = lnCredTotl
            loDtaMstr(0).Item("nCredTotl") = lnCredTotl
            loDtaMstr(0).Item("nRebTotlx") = lnRebTotlx

            If lnLastPaym > 0 Then
                loDtaMstr(0).Item("dLastPaym") = ldLastPaym
                loDtaMstr(0).Item("nLastPaym") = lnLastPaym
            End If

            If lnABalance <= loDtaMstr(0).Item("nMonAmort") + 10 Then
                Dim lnDelayAvg As Single
                lnDelayAvg = getAveDelay(loDtaMstr, p_dTransact)
                loDtaMstr(0).Item("nDelayAvg") = lnDelayAvg
                loDtaMstr(0).Item("cRatingxx") = getRating(lnDelayAvg, "")
            End If

            If lnABalance <= 0 Then
                'loDtaMstr(0).Item("dClosedxx") = loDtaLdgr(loDtaLdgr.Rows.Count - 1).Item("dTransact")
                'kalyptus - 2022.10.14 01:49pm
                'Set the closing date based on the save last transaction that cause the balance to zero(0)
                loDtaMstr(0).Item("dClosedxx") = ldClosedxx
                loDtaMstr(0).Item("cAcctstat") = "1"
                loDtaMstr(0).Item("cActivexx") = "0"
            Else
                'loDtaMstr(0).Item("dClosedxx") = vbNull
                loDtaMstr(0).Item("cAcctstat") = "0"
                loDtaMstr(0).Item("cActivexx") = "1"
            End If

            lsSQL = ADO2SQL(loDtaMstr _
                          , "LR_Master" _
                          , "sAcctNmbr = " & strParm(fsAcctNmbr))
            'Save the master if there are updates in the master...
            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, "LR_Master")
            End If
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        Return True
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
    End Sub
End Class
