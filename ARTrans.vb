'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     AR Trans Object
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
'  Kalyptus [ 09/08/2016 10:50 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcMCInventory

Public Class ARTrans
    Public Const xeActStatActive As Integer = 0
    Public Const xeActStatClosed As Integer = 1
    Public Const xeActStatDead As Integer = 2
    Public Const xeActStatImpounded As Integer = 3
    Public Const xeActStatDiscarded As Integer = 4

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private Const p_sMasTable As String = "MC_AR_Master"
    Private Const p_sDtlTable As String = "MC_AR_Ledger"

    Private Const p_sMsgHeadr As String = "LR Transaction"

    Private p_sAcctNmbr As String
    Private p_dTransact As Date
    Private p_sReferNox As String
    Private p_sRemarksx As String
    Private p_nTranAmtx As Decimal
    Private p_nRebatesx As Decimal
    Private p_nPenaltyx As Decimal
    Private p_sCollctID As String

    Private p_bIsOffice As Boolean

    Private p_sSourceNo As String
    Private p_cTrantype As String

    Private p_nPaymTotl As Decimal
    Private p_nCashTotl As Decimal
    Private p_nDownTotl As Decimal
    Private p_nCredTotl As Decimal

    Private p_sBranchCD As String

    Public Property Master(ByVal Index As String) As Object
        Get
            Select Case LCase(Index)
                Case "sacctnmbr"
                    Master = p_sAcctNmbr
                Case "dtransact"
                    Master = p_dTransact
                Case "srefernox"
                    Master = p_sReferNox
                Case "sremarksx"
                    'mac 2021.09.15
                    '   trim the remarks to 64 chars only
                    If Len(p_sRemarksx) > 64 Then
                        Master = Left(p_sRemarksx, 64)
                    Else
                        Master = p_sRemarksx
                    End If
                Case "ntranamtx"
                    Master = p_nTranAmtx
                Case "nrebatesx"
                    Master = p_nRebatesx
                Case "npenaltyx"
                    Master = p_nPenaltyx
                Case "scollidxx"
                    Master = p_sCollctID
                Case Else
                    MsgBox("Invalid field name retrieved detected: " & Index & "!")
                    Master = ""
            End Select
        End Get
        Set(ByVal value As Object)
            Select Case LCase(Index)
                Case "sacctnmbr"
                    p_sAcctNmbr = value
                Case "dtransact"
                    p_dTransact = value
                Case "srefernox"
                    p_sReferNox = value
                Case "sremarksx"
                    p_sRemarksx = value
                Case "ntranamtx"
                    p_nTranAmtx = value
                Case "nrebatesx"
                    p_nRebatesx = value
                Case "npenaltyx"
                    p_nPenaltyx = value
                Case "scollidxx"
                    p_sCollctID = value
                Case Else
                    MsgBox("Invalid field assigned detected: " & Index & "!")
            End Select
        End Set
    End Property

    Public Function MonthlyPayment( _
         SourceNo As String, _
         OfficeTrans As Boolean) As Boolean

        p_bIsOffice = OfficeTrans
        p_sSourceNo = SourceNo
        p_cTrantype = "p"

        Return SaveTransaction()
    End Function

    Public Function CashBalance( _
         SourceNo As String, _
         OfficeTrans As Boolean) As Boolean

        p_bIsOffice = OfficeTrans
        p_sSourceNo = SourceNo
        p_cTrantype = "b"

        Return SaveTransaction()
    End Function

    Public Function DownPayment( _
         SourceNo As String, _
         OfficeTrans As Boolean) As Boolean

        p_bIsOffice = OfficeTrans
        p_sSourceNo = SourceNo
        p_cTrantype = "d"

        Return SaveTransaction()
    End Function

    Public Function CreditMemo( _
         ByVal SourceNo As String, _
         ByVal OfficeTrans As Boolean) As Boolean

        p_bIsOffice = OfficeTrans
        p_sSourceNo = SourceNo
        p_cTrantype = "c"

        Return SaveTransaction()
    End Function

    Private Function SaveTransaction() As Boolean

        'Check validity of transaction amount
        If p_nTranAmtx + p_nPenaltyx = 0 Then
            MsgBox("Invalid Transaction Amount Detected!" & vbCrLf & _
                     "Verify your Entry then Try Again!", vbCritical, "Warning")
            Return False
        End If

        'Check validity of collector
        If Not p_bIsOffice Then
            If p_sCollctID = "" Then
                MsgBox("Field Collection has No Valid Collector" & vbCrLf & _
                      "Verify your Entry then Try Again!", vbCritical, "Warning")
                Return False
            End If
        Else
            p_sCollctID = ""
        End If

        'Get the Source Branch of this transaction
        p_sBranchCD = Left(p_sSourceNo, 4)

        Try
            'Load the record of the account
            p_oDTMstr = GetMaster()

            'Check for the dLastPaym since we will not allow transactions below the last payment date
            'she 2017-02-23
            'if null unf dLastPayment then dPurchase ung kukunin nya. my mga ar master kasi na null ung dLastpayment pag new entry 
            'If p_oDTMstr(0).Item("dLastPaym") > p_dTransact
            If IFNull(p_oDTMstr(0).Item("dLastPaym"), p_oDTMstr(0).Item("dPurchase")) > p_dTransact Then
                MsgBox("Transaction date is prior to the last transaction date!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            If Not checkTranType() Then Return False

            'Create the SQL Statement to insert the transaction to the LR_Ledger
            Dim lsSQLLdgr As String

            lsSQLLdgr = "INSERT INTO " & p_sDtlTable & " SET" & _
                                    "  sAcctNmbr = " & strParm(p_sAcctNmbr) & _
                                    ", sBranchCd = " & strParm(p_sBranchCD) & _
                                    ", nEntryNox = " & p_oDTMstr(0).Item("nLedgerNo") + 1 & _
                                    ", dTransact = " & dateParm(p_dTransact) & _
                                    ", cOffPaymx = " & strParm(IIf(p_bIsOffice, "3", "2")) & _
                                    ", sCollIDxx = " & strParm(p_sCollctID) & _
                                    ", sORNoxxxx = " & strParm(p_sReferNox) & _
                                    ", cTrantype = " & strParm(p_cTrantype) & _
                                    ", sRemarksx = " & strParm(Left(p_sRemarksx, 64)) & _
                                    ", nTranAmtx = " & p_nTranAmtx & _
                                    ", nDebitAmt = " & 0 & _
                                    ", nOthersxx = " & p_nPenaltyx & _
                                    ", nRebatesx = " & p_nRebatesx & _
                                    ", nABalance = " & p_oDTMstr(0).Item("nABalance") & _
                                    ", nMonDelay = " & getDelay(p_oDTMstr, p_dTransact) & _
                                    ", dModified = " & dateParm(p_oApp.getSysDate)


            'Create SQL Statement that will update the LR_Master
            Dim lsSQLMstr As String
            lsSQLMstr = "UPDATE " & p_sMasTable & _
                       " SET nPaymTotl = nPaymTotl + " & p_nPaymTotl & _
                          ", nPenTotlx = nPenTotlx + " & p_nPenaltyx & _
                          ", nDownTotl = nDownTotl + " & p_nDownTotl & _
                          ", nCashTotl = nCashTotl + " & p_nCashTotl & _
                          ", nRebTotlx = nRebTotlx + " & p_nRebatesx & _
                          ", nCredTotl = nCredTotl + " & p_nCredTotl & _
                          ", nABalance = nABalance - " & (p_nTranAmtx + p_nRebatesx) & _
                          ", nLastPaym = " & (p_nTranAmtx + p_nRebatesx) & _
                          ", dLastPaym = " & dateParm(p_dTransact) & _
                          ", nLedgerNo = " & (p_oDTMstr(0).Item("nLedgerNo") + 1) & _
                       " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)


            'mac 2020.11.19
            '   added validation, rollback changes if rows affected is <= 0
            If p_oApp.Execute(lsSQLLdgr, p_sDtlTable) <= 0 Then Return False
            If p_oApp.Execute(lsSQLMstr, p_sMasTable) <= 0 Then Return False

            If p_oDTMstr(0).Item("cAcctStat") = xeActStatImpounded Then
                If MsgBox("Unit is Currently Impounded!!!" & vbCrLf & _
                         "Released Motorcycle Now???", vbCritical + vbYesNo, "Confirm") = vbYes Then
                    If ReleaseImpound() = False Then Return False
                End If
            End If

            If p_oDTMstr(0).Item("nABalance") <= 0 Then
                Dim lnDelayAvg As Single

                lnDelayAvg = getAveDelay(p_oDTMstr, p_dTransact)

                'kalyptus-2022.10.22 09:46am
                'Update the dClosedxx field if this is the first time the balance becomes less or equals to zero
                If p_nTranAmtx + p_oDTMstr(0).Item("nABalance") > 0 Or
                    p_oDTMstr(0).Item("cAcctstat") = xeActStatClosed Then

                    lsSQLMstr = "UPDATE " & p_sMasTable &
                               " SET dClosedxx = " & dateParm(p_dTransact) &
                                  ", cAcctstat = " & strParm(xeActStatClosed) &
                                  ", nDelayAvg = " & lnDelayAvg &
                                  ", cRatingxx = " & strParm(getRating(lnDelayAvg, p_oDTMstr(0).Item("cRatingxx"), p_oDTMstr(0).Item("nAcctTerm"))) &
                               " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
                Else
                    lsSQLMstr = "UPDATE " & p_sMasTable & _
                               " SET cAcctstat = " & strParm(xeActStatClosed) & _
                                  ", nDelayAvg = " & lnDelayAvg & _
                                  ", cRatingxx = " & strParm(getRating(lnDelayAvg, p_oDTMstr(0).Item("cRatingxx"), p_oDTMstr(0).Item("nAcctTerm"))) & _
                               " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
                End If

                Call p_oApp.Execute(lsSQLMstr, p_sMasTable)
            End If

            Call UpdateCollectionUnit(p_sAcctNmbr, p_dTransact)
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        Return True
    End Function

    Public Function GetMaster() As DataTable
        Dim lsSQL As String

        Try
            lsSQL = "SELECT" & _
                           "  sAcctNmbr" & _
                           ", sBranchCd" & _
                           ", sClientID" & _
                           ", sCoCltID1" & _
                           ", sCoCltID2" & _
                           ", sApplicNo" & _
                           ", sRouteIDx" & _
                           ", sRemarksx" & _
                           ", sExAcctNo" & _
                           ", sSerialID" & _
                           ", cLoanType" & _
                           ", dPurchase" & _
                           ", nGrossPrc" & _
                           ", nPNValuex" & _
                           ", nDownPaym" & _
                           ", nCashBalx" & _
                           ", dFirstPay" & _
                           ", nAcctTerm" & _
                           ", dDueDatex" & _
                           ", nMonAmort" & _
                           ", nPenaltyx" & _
                           ", nRebatesx" & _
                           ", nLastPaym" & _
                           ", dLastPaym" & _
                           ", nPaymTotl" & _
                           ", nPenTotlx" & _
                           ", nRebTotlx" & _
                           ", nDebtTotl" & _
                           ", nCredTotl" & _
                           ", nAmtDuexx" & _
                           ", nABalance" & _
                           ", nDownTotl" & _
                           ", nCashTotl" & _
                           ", nDelayAvg" & _
                           ", cRatingxx" & _
                           ", cAcctstat" & _
                           ", cMotorNew" & _
                           ", dClosedxx" & _
                           ", cActivexx" & _
                           ", nPActTerm" & _
                           ", nPPNValue" & _
                           ", dTermChng" & _
                           ", nLedgerNo" & _
                           ", nLgrLinex" & _
                           ", nPassLine" & _
                           ", cMCStatxx" & _
                           ", cActTypex" & _
                           ", cPostedxx" & _
                           ", sModified" & _
                           ", dModified" & _
                     " FROM " & p_sMasTable & _
                     " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
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

            If .Item("nABalance") <= 0 Then
                Return 0
            End If

            If fdTransact > .Item("dDueDatex") Then ldDueDate = .Item("dDueDatex")

            lnActTerm = getMonthTerm(.Item("dFirstPay"), ldDueDate)

            'kalyptus - 2020.06.06 03:49pm
            'Freeze the term for 2 months for sales prior to the lockdown period and payments from the lockdown period...
            lnActTerm = lnActTerm - getFreezeMonth(.Item("sAcctNmbr"), ldDueDate)

            lnAmtDuex = (lnActTerm * .Item("nMonAmort")) + _
                        .Item("nDownPaym") + _
                        .Item("nCashBalx") + _
                        .Item("nDebtTotl")
            lnAmtDuex = lnAmtDuex - _
                        .Item("nPaymTotl") - _
                        .Item("nDownTotl") - _
                        .Item("nCashTotl") - _
                        .Item("nRebTotlx") - _
                        .Item("nCredTotl")

            lnAmtDuex = IIf(lnAmtDuex < 0, 0, lnAmtDuex)

            If .Item("nMonAmort") > 0.0# Then
                getDelay = Math.Round(lnAmtDuex / .Item("nMonAmort"), 2)
            Else
                If .Item("dDueDatex") < p_dTransact Then
                    getDelay = 1
                Else
                    getDelay = 0
                End If
            End If
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
                        lnTranAmtx = (.Item("nMonAmort") * .Item("nAcctTerm")) - .Item("nABalance")
                        lnCtr2 = lnCtr2 + 1
                        If lnCtr2 <= loData.Rows.Count - 1 Then Exit Do
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
                       ByVal cRating As String, _
                       ByVal nActTerm As Integer) As String

        If nActTerm = 0 Then
            getRating = "n"
        ElseIf cRating = "l" Then
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

    Private Function checkTranType() As Boolean
        Try
            Select Case p_cTrantype
                Case "p"    'Monthly Payment
                    If p_oDTMstr(0).Item("nDownPaym") > p_oDTMstr(0).Item("nDownTotl") Then
                        MsgBox("Down Balance was not yet Paid!!!" & vbCrLf & vbCrLf & _
                                 "Verify Your Entry then Try Again!!!", vbCritical, "Warning")
                        Return False
                    End If

                    If p_oDTMstr(0).Item("nAcctTerm") = 0 Then
                        MsgBox("Account Has No Monthly Payment!" & vbCrLf & vbCrLf & _
                                 "Verify Your Entry then Try Again!!!", vbCritical, "Warning")
                        Return False
                    End If

                    p_nPaymTotl = p_nTranAmtx
                    p_oDTMstr(0).Item("nPaymTotl") = p_oDTMstr(0).Item("nPaymTotl") + p_nTranAmtx
                    p_oDTMstr(0).Item("nRebTotlx") = p_oDTMstr(0).Item("nRebTotlx") + p_nRebatesx
                    p_oDTMstr(0).Item("nLastPaym") = p_nTranAmtx + p_nRebatesx
                    p_oDTMstr(0).Item("dLastPaym") = p_dTransact
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - (p_nTranAmtx + p_nRebatesx)

                Case "b"    'Cash Balance
                    If p_oDTMstr(0).Item("nAcctTerm") > 0 Then
                        MsgBox("Account Has No Cash Balance!" & vbCrLf & vbCrLf & _
                                 "Verify Your Entry then Try Again!!!", vbCritical, "Warning")
                        Return False
                    End If

                    p_nCashTotl = p_nTranAmtx
                    p_oDTMstr(0).Item("nCashTotl") = p_oDTMstr(0).Item("nCashTotl") + p_nTranAmtx
                    p_oDTMstr(0).Item("nLastPaym") = p_nTranAmtx
                    p_oDTMstr(0).Item("dLastPaym") = p_dTransact
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nTranAmtx
                    p_nRebatesx = 0

                Case "d"    'Down Balance
                    If p_oDTMstr(0).Item("nDownPaym") = p_oDTMstr(0).Item("nDownTotl") Then
                        MsgBox("Account Has No Remaining Down Balance!!!" & vbCrLf & vbCrLf & _
                                 "Verify Your Entry then Try Again!!!", vbCritical, "Warning")
                        Return False
                    End If

                    p_nDownTotl = p_nTranAmtx
                    p_oDTMstr(0).Item("nDownTotl") = p_oDTMstr(0).Item("nDownTotl") + p_nTranAmtx
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nTranAmtx
                    p_nRebatesx = 0

                Case "c"    'Credit Memo
                    'kalyptus - 2018.11.13 05:15pm
                    'Added credit memo
                    p_nCredTotl = p_nTranAmtx
                    p_oDTMstr(0).Item("nDownTotl") = p_oDTMstr(0).Item("nCredTotl") + p_nCredTotl
                    p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nTranAmtx
                    p_nRebatesx = 0
            End Select
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        Return True
    End Function

    Private Function ReleaseImpound() As Boolean
        Dim lsSQL As String
        Dim loData As DataTable

        Try
            lsSQL = " SELECT sTransNox" & _
                    " FROM Impound" & _
                    " WHERE sAcctNmbr = " & strParm(p_oDTMstr(0).Item("sAcctNmbr")) & _
                      " AND dImpoundx = " & dateParm(p_oDTMstr(0).Item("dClosedxx"))

            loData = p_oApp.ExecuteQuery(lsSQL)
            If loData.Rows.Count = 0 Then
                MsgBox("Unable to load Impounded Info!", vbCritical, "Warning")
                Return False
            End If

            lsSQL = "UPDATE Impound SET" & _
                        "  cTranStat = " & strParm(1) & _
                        ", dRedeemxx = " & dateParm(p_dTransact) & _
                        ", dModified = " & dateParm(p_oApp.getSysDate()) & _
                    " WHERE sTransNox = " & strParm(loData(0).Item("sTransNox"))

            If p_oApp.Execute(lsSQL, "Impound") = 0 Then
                MsgBox("Unable to Update Impounded Info " & loData(0).Item("sTransNox") & "!", vbCritical, "Warning")
                Return False
            End If

            lsSQL = " SELECT" & _
                        "  sEngineNo" & _
                        ", sMCInvIDx" & _
                        ", nLedgerNo" & _
                    " FROM MC_Serial" & _
                    " WHERE sSerialID = " & strParm(p_oDTMstr(0).Item("sSerialID"))

            loData = p_oApp.ExecuteQuery(lsSQL)
            If loData.Rows.Count = 0 Then
                MsgBox("Unable to load Motorcycle Info!", vbCritical, "Warning")
                Return False
            End If

            Dim loMCTrans As MCSerialTrans
            loMCTrans = New MCSerialTrans(p_oApp)

            With loMCTrans
                .AppDriver = p_oApp
                .Branch = p_sBranchCD
                If .InitTransaction() = False Then Return False

                .Detail(0, "sSerialID") = p_oDTMstr(0).Item("sSerialID")
                .Detail(0, "sMCInvIDx") = loData(0).Item("sMCInvIDx")
                .Detail(0, "nLedgerNo") = loData(0).Item("nLedgerNo")
                .Detail(0, "cSoldStat") = 1
                If .Release(p_sSourceNo, p_dTransact, xeEditMode.MODE_ADDNEW) = False Then Return False
            End With

            lsSQL = "UPDATE " & p_sMasTable & _
                    " SET cAcctstat = " & strParm(xeActStatActive) & _
                       ", dClosedxx = NULL" & _
                    " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
            Call p_oApp.Execute(lsSQL, p_sMasTable)
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        ReleaseImpound = True
    End Function

    Private Sub UpdateCollectionUnit(ByVal fsAcctNmbr As String, ByVal fdTransact As Date)
        Dim lsSQL As String
        Dim loData As DataTable

        Try
            lsSQL = "SELECT" & _
                           "  sAcctNmbr" & _
                           ", cCollStat" & _
                   " FROM LR_Collection_Unit" & _
                   " WHERE dTransact >= " & dateParm(fdTransact)
            loData = p_oApp.ExecuteQuery(lsSQL)

            If loData.Rows.Count > 0 Then
                'Tagged the account as paid if not yet tagged as paid
                If loData(0).Item("cCollStat") <> "3" Then
                    lsSQL = "UPDATE LR_Collection_Unit SET" & _
                                 "  cCollStat = '3'" & _
                           " WHERE sAcctNmbr = " & strParm(fsAcctNmbr)
                    p_oApp.Execute(lsSQL, "LR_Collection_Unit")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getFreezeMonth(ByVal fsAcctNmbr As String, ByVal fdTransact As Date) As Integer
        Dim lsSQL As String
        Dim loDta As DataTable

        getFreezeMonth = 0

        lsSQL = "SELECT sAcctNmbr, b.*" & _
               " FROM MC_AR_Master a" & _
                   " LEFT JOIN Branch_Lockdown_History b ON a.sBranchCd = b.sBranchCD" & _
               " WHERE sAcctNmbr = " & strParm(fsAcctNmbr) & _
                 " AND b.dDateFrom > a.dFirstPay" & _
                 " AND b.dDateThru < " & dateParm(fdTransact)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        Dim lnCtr As Integer
        For lnCtr = 0 To loDta.Rows.Count - 1
            getFreezeMonth = getFreezeMonth + loDta(lnCtr).Item("nMonthxxx")
        Next

        'kalyptus - 2020.08.06 09:18am
        'Remove Unfreezed month if client grab the promo
        lsSQL = "SELECT DISTINCT a.sOthrInfo nUnfreezd, b.sMainInfo, b.sOthrInfo" & _
               " FROM CCS_Promo_Master a" & _
                    " LEFT JOIN CCS_Promo_Detail b ON a.sTransNox = b.sTransNox" & _
               " WHERE a.sProgrmCD = '0001'" & _
                 " AND b.sOthrInfo BETWEEN a.dDateFrom AND a.dDateThru" & _
                 " AND b.sMainInfo = " & strParm(fsAcctNmbr) & _
                 " AND b.sOthrInfo < " & dateParm(fdTransact)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        For lnCtr = 0 To loDta.Rows.Count - 1
            getFreezeMonth = getFreezeMonth - Val(loDta(lnCtr).Item("nUnfreezd"))
        Next
    End Function


    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
