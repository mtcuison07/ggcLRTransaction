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

Public Class LRTransOld
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
    Private p_nPenaltyx As Decimal

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

    Public Property Others() As Decimal
        Get
            Return p_nPenaltyx
        End Get
        Set(ByVal value As Decimal)
            p_nPenaltyx = value
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

    Public Function Penalty() As Boolean
        p_sSourceCD = "PLTY"

        Return SaveTransaction()
    End Function

    Private Function SaveTransaction() As Boolean

        p_oDTMstr = GetMaster()

        'Check for the dLastPaym since we will not allow transactions below the last payment date
        If p_oDTMstr(0).Item("dLastPaym") > p_dTransact Then
            MsgBox("Transaction date is prior to the last transaction date!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Note: what is the difference between sSourceCD and cTranType if all transactions are coming from LR_Payment_Master
        '      just a hunch...
        Select Case p_sSourceCD
            Case "PYMT"
                p_nIntAmtxx = 0
                'Kung may collateral di nabawas ang interest during release, 
                'kayat d2 sa Monthly Payment natin kolektahin ang interest
                If p_oDTMstr(0).Item("sCollatID") <> "" Then
                    'Kalyptus - 2016.08.04 09:44am
                    '   Replace the equal splitting of amount payment and replace with interest first logic...
                    '   Please do not delete just in case this will be adopted later on...
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    'Dim lnIntRate As Single
                    ''nMonAmort's value is the monthly amortization for the principal only...
                    'lnIntRate = (p_oDTMstr(0).Item("nPrincipl") * p_oDTMstr(0).Item("nIntRatex") / 100)
                    'lnIntRate = lnIntRate / (p_oDTMstr(0).Item("nMonAmort") + lnIntRate)
                    '' get the INTEREST PAYMENT from this Monthly Amortization Payment 
                    'p_nIntAmtxx = Math.Round(lnIntRate * p_nTranAmtx)
                    '' Add the INTEREST PAYMENT to  the INTEREST TOTAL. 
                    'p_oDTMstr(0).Item("nIntTotal") = p_oDTMstr(0).Item("nIntTotal") + p_nIntAmtxx
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Dim lnPrincipl As Decimal = p_oDTMstr(0).Item("nPrincipl")
                    Dim lnInterest As Decimal = p_oDTMstr(0).Item("nInterest")
                    Dim lnAcctTerm As Integer = p_oDTMstr(0).Item("nAcctTerm")
                    Dim lnPaymTotl As Decimal = p_oDTMstr(0).Item("nPaymTotl")
                    Dim lnIntTotal As Decimal = p_oDTMstr(0).Item("nIntTotal")
                    Dim lnTranAmtx As Decimal = p_nTranAmtx
                    Dim lnPaidAmtx As Decimal = 0
                    Dim lnIntAmtxx As Decimal = 0

                    Call SplitPayment(lnPrincipl, lnInterest, lnAcctTerm, lnPaymTotl, lnIntTotal, lnTranAmtx, lnPaidAmtx, lnIntAmtxx)

                    p_nIntAmtxx = lnIntAmtxx
                End If

                p_nPaidAmtx = p_nTranAmtx - p_nIntAmtxx

                p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nABalance") - p_nPaidAmtx

                p_cTrantype = "0"
                p_nPenaltyx = 0
                p_nDebitAmt = 0
                p_nCredtAmt = 0
                p_nMonDelay = getDelay(p_oDTMstr, p_dTransact)
                p_nAmtDuexx = p_nMonDelay * p_oDTMstr(0).Item("nMonAmort")
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
                      ", nABalance = nABalance - " & ((p_nPaidAmtx + p_nCredtAmt) - p_nDebitAmt) & _
                      ", nAmtDuexx = " & p_nAmtDuexx & _
                      ", nLastPaym = " & ((p_nIntAmtxx + p_nPaidAmtx + p_nCredtAmt) - p_nDebitAmt) & _
                      ", dLastPaym = " & dateParm(p_dTransact) & _
                      ", nLedgerNo = " & (p_oDTMstr(0).Item("nLedgerNo") + 1) & _
                   " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)

        Call p_oApp.Execute(lsSQLLdgr, "LR_Ledger")
        Call p_oApp.Execute(lsSQLMstr, "LR_Master")

        If p_oDTMstr(0).Item("nABalance") <= 0 Then
            Dim lnDelayAvg As Single

            lnDelayAvg = getAveDelay(p_oDTMstr, p_dTransact)

            lsSQLMstr = "UPDATE LR_Master" & _
                       " SET dClosedxx = " & dateParm(p_dTransact) & _
                          ", cAcctstat = " & strParm("1") & _
                          ", nDelayAvg = " & lnDelayAvg & _
                          ", cRatingxx = " & strParm(getRating(lnDelayAvg, "")) & _
                          ", cActivexx = " & strParm("0") & _
                       " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
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

        Return True
    End Function

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
        ElseIf lnPayTermx < lnIntTermx Then
            'Compute for the amount to be distributed for monthly payment
            Dim lnDiff As Decimal = (lnPayTermx - lnIntTermx) * lnPayAmort
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
        Else
            'Compute for the amount to be distributed for interest payment
            Dim lnDiff As Decimal = (lnIntTermx - lnPayTermx) * lnIntAmort
            lnIntAmort = lnDiff
            If fnTranAmtx < lnDiff Then
                lnIntAmort = fnTranAmtx
                fnIntAmtxx = fnIntAmtxx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnIntAmtxx = fnIntAmtxx + lnDiff
                fnTranAmtx = fnTranAmtx - lnDiff
            End If
        End If

        fnPaymTotl = fnPaymTotl + lnPayAmort
        fnIntTotal = fnIntTotal + lnIntAmort

        'Execute a recursive function if fnTranAmtx is not yet 0
        If fnTranAmtx > 0 Then
            SplitPayment(fnPrincipl, fnInterest, fnAcctTerm, fnPaymTotl, fnIntTotal, fnTranAmtx, fnPaidAmtx, fnIntAmtxx)
        End If
    End Sub

    Public Function GetMaster() As DataTable
        Dim lsSQL As String

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
                 " FROM " & p_sMasTable & _
                 " WHERE sAcctNmbr = " & strParm(p_sAcctNmbr)
        Return p_oApp.ExecuteQuery(lsSQL)
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

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
    End Sub
End Class
