'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Car Trade Billing Object
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
'  Jheff [ 04/30/2018 2:43 Pm ]
'      Started creating this object.
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ggcAppDriver
Imports System.Globalization

Public Class CTBilling
    Const pxeMODULENAME As String = "CTBilling"
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetl As DataTable
    Private p_nEditMode As xeEditMode
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sAPClientNm As String

    Private Const p_sMasTable As String = "CT_Billing_Master"
    Private Const p_sDetTable As String = "CT_Billing_Detail"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                 ByVal Value As Object)
    Public Event DetailRetrieved(ByVal Row As Integer, _
                                 ByVal Index As Integer, _
                                 ByVal Value As Object)

    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    WriteOnly Property TranStatus() As Integer
        Set(ByVal Value As Integer)
            p_nTranStat = Value
        End Set
    End Property

    Public Property Master(ByVal Index As Object) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                'If Index = 9 Then
                '    Return p_sAPClientNm
                'Else
                Return p_oDTMstr(0)(Index)
                'End If
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal Value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTMstr(0)(Index) = Value
                RaiseEvent MasterRetrieved(Index, p_oDTMstr(0)(Index))
            End If
        End Set
    End Property

    Property Detail(ByVal Row As Integer, ByVal Index As Object) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTDetl(Row)(Index)
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal Value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTDetl(Row)(Index) = Value
            End If
        End Set
    End Property

    Public ReadOnly Property ItemCount() As Integer
        Get
            If Not IsNothing(p_oDTDetl) Then
                Return p_oDTDetl.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Function Initialize() As Boolean
        Dim lsProcName As String = "Initialize"

        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Public Function NewTransaction() As Boolean
        Dim lsProcName As String = "NewTransaction"

        Try
            p_oDTMstr = New DataTable
            Debug.Print(AddCondition(getSQL_Master, "0=1"))
            p_oDTMstr = p_oApp.ExecuteQuery(AddCondition(getSQL_Master, "0=1"))
            Call initMaster()

            p_oDTDetl.Clear()
            Call initDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, pxeMODULENAME & "-" & lsProcName)
            GoTo errProc
        Finally
            RaiseEvent MasterRetrieved(0, p_oDTMstr.Rows(0)("sTransNox"))
            p_nEditMode = xeEditMode.MODE_ADDNEW
        End Try

endProc:
        Return True
        Exit Function
errProc:
        Return False
    End Function

    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String
        Dim loDetail As DataTable

        lsSQL = AddCondition(getSQL_Master, "a.sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        If p_oDTMstr.Rows.Count = 0 Then Return False
        p_sAPClientNm = p_oDTMstr.Rows(0)("xAPClient")

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        lsSQL = AddCondition(getSQL_Detail, "a.sTransNox = " & strParm(fsTransNox))
        loDetail = p_oApp.ExecuteQuery(lsSQL)

        p_oDTDetl.Clear()
        With loDetail
            If loDetail.Rows.Count > 0 Then
                For nCtr As Integer = 0 To .Rows.Count - 1
                    p_oDTDetl.Rows.Add()
                    p_oDTDetl.Rows(nCtr)("sTransNox") = .Rows(nCtr)("sTransNox")
                    p_oDTDetl.Rows(nCtr)("nEntryNox") = .Rows(nCtr)("nEntryNox")
                    p_oDTDetl.Rows(nCtr)("dTransact") = .Rows(nCtr)("dTransact")
                    p_oDTDetl.Rows(nCtr)("sAcctNmbr") = .Rows(nCtr)("sReferNox")
                    p_oDTDetl.Rows(nCtr)("sClientNm") = .Rows(nCtr)("sClientNm")
                    p_oDTDetl.Rows(nCtr)("sEngineNo") = .Rows(nCtr)("sEngineNo")
                    p_oDTDetl.Rows(nCtr)("nPrincipl") = .Rows(nCtr)("nPrincipl")
                    p_oDTDetl.Rows(nCtr)("nInterest") = .Rows(nCtr)("nInterest")
                    p_oDTDetl.Rows(nCtr)("nSubsidze") = .Rows(nCtr)("nSubsidze")
                    p_oDTDetl.Rows(nCtr)("nInctvAmt") = .Rows(nCtr)("nInctvAmt")
                    p_oDTDetl.Rows(nCtr)("cBillType") = .Rows(nCtr)("cBillType")
                    p_oDTDetl.Rows(nCtr)("sDescript") = .Rows(nCtr)("sDescript")
                    p_oDTDetl.Rows(nCtr)("sRemarks1") = .Rows(nCtr)("sRemarks1")
                    p_oDTDetl.Rows(nCtr)("sRemarks2") = .Rows(nCtr)("sRemarks2")
                    p_oDTDetl.Rows(nCtr)("nAmountxx") = .Rows(nCtr)("nAmountxx")
                    p_oDTDetl.Rows(nCtr)("nApproved") = .Rows(nCtr)("nApproved")
                    p_oDTDetl.Rows(nCtr)("sClientID") = .Rows(nCtr)("sClientID")
                    p_oDTDetl.Rows(nCtr)("sSerialID") = .Rows(nCtr)("sSerialID")
                Next nCtr
            End If
        End With

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    Public Function SearchTransaction( _
                              ByVal sValue As String _
                            , Optional ByVal bByCode As Boolean = False) As Boolean

        Dim lsSQL As String
        Dim lsCondition As String

        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If bByCode Then
                If sValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If sValue = p_oDTMstr(0).Item("sCompnyNm") Then Return True
            End If
        End If

        lsSQL = getSQL_Browse()

        Dim lsFilter As String
        If bByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & sValue)
        Else
            lsFilter = "b.sCompnyNm LIKE " & strParm(sValue & "%")
        End If

        If p_nTranStat <> -1 Then
            If p_nTranStat > -1 Then
                lsCondition = "("
                For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
                    lsCondition = lsCondition & " a.cTranStat = " & _
                                      strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
                Next
                lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
            Else
                lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
            End If
        End If

        lsSQL = AddCondition(lsSQL, lsCondition)

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sCompnyNm»dTransact" _
                                        , "TransNox»Company»Date", _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact" _
                                        , IIf(bByCode, 0, 1))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function

    Public Function SearchPayment( _
                              ByVal sValue As String _
                            , Optional ByVal bByCode As Boolean = False) As Boolean

        Dim lsSQL As String
        Dim lsCondition As String

        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If bByCode Then
                If sValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If sValue = p_oDTMstr(0).Item("sCompnyNm") Then Return True
            End If
        End If

        lsSQL = AddCondition(getSQL_Browse(), "LEFT(a.sTransNox, 4) <> " & strParm(p_oApp.BranchCode))

        Dim lsFilter As String
        If bByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & sValue)
        Else
            lsFilter = "b.sCompnyNm LIKE " & strParm(sValue & "%")
        End If

        If p_nTranStat <> -1 Then
            If p_nTranStat > -1 Then
                lsCondition = "("
                For pnCtr = 1 To Len(Trim(Str(p_nTranStat)))
                    lsCondition = lsCondition & " a.cTranStat = " & _
                                      strParm(Mid(Trim(Str(p_nTranStat)), pnCtr, 1)) & " OR "
                Next
                lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
            Else
                lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
            End If
        End If

        lsSQL = AddCondition(lsSQL, lsCondition)
        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sCompnyNm»dTransact" _
                                        , "TransNox»Company»Date", _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact" _
                                        , IIf(bByCode, 0, 1))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function

    Public Function SaveTransaction() As Boolean
        Dim lnRow As Integer

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        Dim lsSQL As String


        Try
            With p_oApp
                .BeginTransaction()
                Dim lnCtr As Integer = 0
                For Each dRow As DataRow In p_oDTDetl.Rows
                    If dRow.Item("sAcctNmbr") = "" Then Exit For
                    If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                        lsSQL = "INSERT INTO " & p_sDetTable & " SET" & _
                                    "  sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox")) & _
                                    ", nEntryNox = " & CDbl(lnCtr + 1) & _
                                    ", sReferNox = " & strParm(dRow.Item("sAcctNmbr")) & _
                                    ", cBillType = " & strParm(IFNull(dRow.Item("cBillType"), "0")) & _
                                    ", sDescript = " & strParm(IFNull(dRow.Item("sDEscript"), "")) & _
                                    ", sRemarks1 = " & strParm(IFNull(dRow.Item("sRemarks1"), "")) & _
                                    ", sRemarks2 = " & strParm(IFNull(dRow.Item("sRemarks2"), "")) & _
                                    ", nAmountxx = " & CDbl(dRow.Item("nAmountxx")) & _
                                    ", nApproved = " & CDbl(dRow.Item("nApproved")) & _
                                    ", dModified = " & dateParm(p_oApp.SysDate)

                        lnRow = .Execute(lsSQL, p_sDetTable)
                        If lnRow = 0 Then GoTo endWithroll
                    Else
                        If p_oDTMstr.Rows(0)("nEntryNox") < lnCtr + 1 Then
                            lsSQL = "INSERT INTO " & p_sDetTable & " SET" & _
                                   "  sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox")) & _
                                   ", nEntryNox = " & CDbl(lnCtr + 1) & _
                                   ", sReferNox = " & strParm(dRow.Item("sAcctNmbr")) & _
                                   ", cBillType = " & strParm(IFNull(dRow.Item("cBillType"), "0")) & _
                                   ", sDescript = " & strParm(IFNull(dRow.Item("sDEscript"), "")) & _
                                   ", sRemarks1 = " & strParm(IFNull(dRow.Item("sRemarks1"), "")) & _
                                   ", sRemarks2 = " & strParm(IFNull(dRow.Item("sRemarks2"), "")) & _
                                   ", nAmountxx = " & CDbl(dRow.Item("nAmountxx")) & _
                                   ", nApproved = " & CDbl(dRow.Item("nApproved")) & _
                                   ", dModified = " & dateParm(p_oApp.SysDate)

                        Else
                            lsSQL = "UPDATE " & p_sDetTable & " SET" & _
                                        "  sReferNox = " & strParm(dRow.Item("sAcctNmbr")) & _
                                        ", cBillType = " & strParm(IFNull(dRow.Item("cBillType"), "0")) & _
                                        ", sDescript = " & strParm(IFNull(dRow.Item("sDEscript"), "")) & _
                                        ", sRemarks1 = " & strParm(IFNull(dRow.Item("sRemarks1"), "")) & _
                                        ", sRemarks2 = " & strParm(IFNull(dRow.Item("sRemarks2"), "")) & _
                                        ", nAmountxx = " & CDbl(dRow.Item("nAmountxx")) & _
                                        ", nApproved = " & CDbl(dRow.Item("nApproved")) & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox")) & _
                                        " AND nEntryNox = " & CDbl(lnCtr + 1)

                        End If
                        lnRow = .Execute(lsSQL, p_sDetTable)
                    End If

                    lnCtr = lnCtr + 1
                Next dRow

                If p_oDTMstr.Rows(0)("nEntryNox") <> p_oDTDetl.Rows.Count Then
                    lsSQL = "DELETE FROM " & p_sDetTable & _
                                " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox")) & _
                                    " AND nEntryNox > " & CDbl(p_oDTDetl.Rows.Count)

                    lnRow = .Execute(lsSQL, p_sDetTable)
                End If

                p_oDTMstr.Rows(0)("nEntryNox") = p_oDTDetl.Rows.Count
                If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                    lsSQL = ADO2SQL(p_oDTMstr, _
                                    p_sMasTable, , , , _
                                    "sCompnyNm»sBranchNm»xAPClient")

                    lnRow = .Execute(lsSQL, p_sMasTable)
                    If lnRow = 0 Then GoTo endWithroll
                Else
                    lsSQL = ADO2SQL(p_oDTMstr, _
                                    p_sMasTable, _
                                    "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , , _
                                    "sCompnyNm»sBranchNm»xAPClient")

                    lnRow = .Execute(lsSQL, p_sMasTable)
                End If

                .CommitTransaction()
            End With

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

endwithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Public Function PayTransaction(ByVal bCreateAP As Boolean) As Boolean
        Dim lnRow As Integer
        Dim lsSQL As String

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        Try
            With p_oApp
                .BeginTransaction()

                lsSQL = "UPDATE " & p_sMasTable & " SET" & _
                                    "  cTranStat = " & strParm(xeTranStat.TRANS_POSTED) & _
                                    ", sRemarks2 = " & strParm(p_oDTMstr.Rows(0).Item("sRemarks2")) & _
                                    ", nApprTotl = " & CDbl(p_oDTMstr.Rows(0).Item("nApprTotl")) & _
                                    ", nApprTotl = " & CDbl(p_oDTMstr.Rows(0).Item("nApprTotl")) & _
                                    ", sApprovBy = " & strParm(p_oApp.UserID) & _
                               " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox"))

                lnRow = .Execute(lsSQL, p_sDetTable)

                Dim lnCtr As Integer = 0
                For Each dRow As DataRow In p_oDTDetl.Rows
                    If p_nEditMode = xeEditMode.MODE_UPDATE Then
                        lsSQL = "UPDATE " & p_sDetTable & " SET" & _
                                    "  sRemarks2 = " & strParm(IFNull(dRow.Item("sRemarks2"), "")) & _
                                    ", nApproved = " & CDbl(dRow.Item("nApproved")) & _
                                " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox")) & _
                                    " AND nEntryNox = " & CDbl(lnCtr + 1)

                    End If
                    lnRow = .Execute(lsSQL, p_sDetTable)

                    lnCtr = lnCtr + 1
                Next dRow

                If bCreateAP Then
                    If Not saveAPTransaction() Then
                        MsgBox("Unable to save client ledger!!!" & vbCrLf & _
                                "Please contact GGC SEG/SSG for asssistance!!!", MsgBoxStyle.Critical, "WARNING")
                        .RollBackTransaction()
                    End If
                Else
                    If Not saveARTransaction() Then
                        MsgBox("Unable to save client ledger!!!" & vbCrLf & _
                                "Please contact GGC SEG/SSG for asssistance!!!", MsgBoxStyle.Critical, "WARNING")
                        .RollBackTransaction()
                    End If
                End If
                .CommitTransaction()
            End With
            Return True
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

endwithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Public Function ReceivedPayment() As Boolean
        Dim lnRow As Integer
        Dim lsSQL As String

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        Try
            With p_oApp
                .BeginTransaction()

                lsSQL = "UPDATE " & p_sMasTable & " SET" & _
                                    "  cTranStat = " & strParm(xeTranStat.TRANS_UNKNOWN) & _
                                    ", sRecvByxx = " & strParm(p_oApp.UserID) & _
                                    ", dRecvDate = " & dateParm(p_oApp.SysDate) & _
                               " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox"))

                lnRow = .Execute(lsSQL, p_sDetTable)
                .CommitTransaction()
            End With
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

endwithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Private Function saveAPTransaction() As Boolean
        Dim loDT As DataTable
        Dim lsSQL As String

        'fixed client id for cartrade waiting for sir marlon revision 
        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", a.nCredLimt" & _
                    ", a.nLedgerNo" & _
                    ", a.nABalance" & _
                    ", a.cAutoHold" & _
                    ", b.sSourceCd" & _
                    ", b.sSourceNo" & _
                    ", b.nAmountOt" & _
                    ", b.nAmountIn" & _
                    ", b.nABalance xABalance" & _
                    ", b.dTransact" & _
                 " FROM AP_Client_Master a" & _
                       " LEFT JOIN AP_Client_Ledger b" & _
                          " ON a.sClientID = b.sClientID" & _
                             " AND b.sSourceCd = 'VMCT'" & _
                             " AND b.sSourceNo = " & strParm(p_oDTMstr.Rows(0)("sTransNox")) & _
                 " WHERE a.sClientID = " & strParm(p_oDTMstr.Rows(0)("sClientID"))

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            MsgBox("Invalid Client ID Detected!!!", vbCritical, "Warning")
            Return False
        End If

        If Not addAPClientTrans(IFNull(loDT.Rows(0)("nLedgerNo"), 0),
                                p_oDTMstr.Rows(0)("sClientID"),
                                p_oDTMstr.Rows(0)("dTransact"),
                                p_oDTMstr.Rows(0)("sTransNox"),
                                0,
                                p_oDTMstr.Rows(0)("nApprTotl"),
                                loDT.Rows(0)("nABalance")) Then Return False

        Return True
    End Function

    Private Function addAPClientTrans(ByVal nLedgerNO As Integer, _
                                    ByVal sClientID As String,
                                    ByVal dTransact As Date,
                                    ByVal sSourceNo As String,
                                    ByVal nAmountOt As Long,
                                    ByVal nAmountIn As Long,
                                    ByVal nABalance As Long) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        With p_oApp
            lsSQL = "UPDATE AP_Client_Master SET" & _
                        "  nLedgerNo = " & Format(nLedgerNO + 1, "0000") & _
                        ", dModified = " & dateParm(.SysDate) & _
                     " WHERE sClientID = " & strParm(sClientID)

            lnRow = .Execute(lsSQL, "AP_Client_Master")
            If lnRow <= 0 Then
                MsgBox("Unable to Update Client Master Info!", vbCritical, "Warning")
                Return False
            End If

            lsSQL = "INSERT INTO AP_Client_Ledger SET" & _
                        "  sClientID = " & strParm(sClientID) & _
                        ", nLedgerNo = " & Format(nLedgerNO + 1, "0000") & _
                        ", dTransact = " & dateParm(dTransact) & _
                        ", sSourceCd = " & strParm("VMCT") & _
                        ", sSourceNo = " & strParm(sSourceNo) & _
                        ", nAmountOt = " & nAmountOt & _
                        ", nAmountIn = " & nAmountIn & _
                        ", nABalance = " & nABalance + (nAmountIn - nAmountOt) & _
                        ", dPostedxx = NULL" & _
                        ", dModified = " & dateParm(.SysDate)

            lnRow = .Execute(lsSQL, "AP_Client_Ledger")
            If lnRow <= 0 Then
                MsgBox("Unable to Update Client Ledger Info!", vbCritical, "Warning")
                Return False
            End If

            Return True
        End With
    End Function

    Private Function saveARTransaction() As Boolean
        Dim loDT As DataTable
        Dim lsSQL As String

        'fixed client id for cartrade wainting for sir marlon revision 
        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", a.nCredLimt" & _
                    ", a.nLedgerNo" & _
                    ", a.nABalance" & _
                    ", a.cAutoHold" & _
                    ", b.sSourceCd" & _
                    ", b.sSourceNo" & _
                    ", b.nAmountOt" & _
                    ", b.nAmountIn" & _
                    ", b.nABalance xABalance" & _
                    ", b.dTransact" & _
                 " FROM AR_Client_Master a" & _
                       " LEFT JOIN AR_Client_Ledger b" & _
                          " ON a.sClientID = b.sClientID" & _
                             " AND b.sSourceCd = 'VMNP'" & _
                             " AND b.sSourceNo = " & strParm(p_oDTMstr.Rows(0)("sTransNox")) & _
                 " WHERE a.sClientID = " & strParm("M00118001429")

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            MsgBox("Invalid Client ID Detected!!!", vbCritical, "Warning")
            Return False
        End If

        If Not addARClientTrans(IFNull(loDT.Rows(0)("nLedgerNo"), 0),
                                "M00118001429",
                                p_oDTMstr.Rows(0)("dTransact"),
                                p_oDTMstr.Rows(0)("sTransNox"),
                                p_oDTMstr.Rows(0)("nApprTotl"),
                                0,
                                loDT.Rows(0)("nABalance")) Then Return False

        Return True
    End Function

    Private Function addARClientTrans(ByVal nLedgerNO As Integer, _
                                    ByVal sClientID As String,
                                    ByVal dTransact As Date,
                                    ByVal sSourceNo As String,
                                    ByVal nAmountOt As Long,
                                    ByVal nAmountIn As Long,
                                    ByVal nABalance As Long) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        With p_oApp
            lsSQL = "UPDATE AR_Client_Master SET" & _
                        " nLedgerNo = " & Format(nLedgerNO + 1, "0000") & _
                        ", dModified = " & dateParm(.SysDate) & _
                     " WHERE sClientID = " & strParm(sClientID)

            lnRow = .Execute(lsSQL, "AR_Client_Master")
            If lnRow <= 0 Then
                MsgBox("Unable to Update Client Master Info!", vbCritical, "Warning")
                Return False
            End If

            lsSQL = "INSERT INTO AR_Client_Ledger SET" & _
                        "  sClientID = " & strParm(sClientID) & _
                        ", nLedgerNo = " & Format(nLedgerNO + 1, "0000") & _
                        ", dTransact = " & dateParm(dTransact) & _
                        ", sSourceCd = " & strParm("VMNP") & _
                        ", sSourceNo = " & strParm(sSourceNo) & _
                        ", nAmountOt = " & nAmountOt & _
                        ", nAmountIn = " & nAmountIn & _
                        ", nABalance = " & nABalance + (nAmountIn - nAmountOt) & _
                        ", dPostedxx = NULL" & _
                        ", dModified = " & dateParm(.SysDate)

            lnRow = .Execute(lsSQL, "AR_Client_Ledger")
            If lnRow <= 0 Then
                MsgBox("Unable to Update Client Ledger Info!", vbCritical, "Warning")
                Return False
            End If

            Return True
        End With
    End Function

    Public Function CloseTransaction() As Boolean
        Dim lnRow As Integer
        Dim lsSQL As String

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        lsSQL = "UPDATE " & p_sMasTable & " SET" & _
                    "  cTranStat = " & strParm(xeTranStat.TRANS_CLOSED) & _
                " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox"))

        lnRow = p_oApp.Execute(lsSQL, p_sMasTable)

        Return True

endwithRoll:
        Return False
    End Function

    Public Function CancelTransaction() As Boolean
        Dim lnRow As Integer
        Dim lsSQL As String

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        If p_oDTMstr.Rows(0)("cTranStat") > xeTranStat.TRANS_OPEN Then
            MsgBox("Unable to Cancel Transaction...", vbCrLf & _
                    "Please verify your entry then try again...")
            Return False
        End If

        lsSQL = "UPDATE " & p_sMasTable & " SET" & _
                    "  cTranStat = " & strParm(xeTranStat.TRANS_CANCELLED) & _
                " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox"))

        lnRow = p_oApp.Execute(lsSQL, p_sMasTable)

        Return True

endwithRoll:
        Return False
    End Function

    Public Function PostTransaction() As Boolean
        Dim lnRow As Integer
        Dim lsSQL As String

        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "WARNING")
            Return False
        End If

        If p_oDTMstr.Rows(0)("cTranStat") > xeTranStat.TRANS_OPEN Then
            MsgBox("Unable to Post Transaction...")
            Return False
        End If

        lsSQL = "UPDATE " & p_sMasTable & " SET" & _
                    "  cTranStat = " & strParm(xeTranStat.TRANS_POSTED) & _
                " WHERE sTransNox = " & strParm(p_oDTMstr.Rows(0).Item("sTransNox"))

        lnRow = p_oApp.Execute(lsSQL, p_sMasTable)

        Return True

endwithRoll:
        Return False
    End Function

    Public Function UpdateTransaction() As Boolean
        If p_nEditMode = xeEditMode.MODE_READY Then p_nEditMode = xeEditMode.MODE_UPDATE
        Return True
    End Function

    Public Function SearchMaster(ByVal nIndex As Integer, Optional ByVal sValue As Object = "") As Boolean
        Select Case nIndex
            Case 1 'company
                Return getCompany(sValue, False)
            Case 2 'branch
                Return getBranch(sValue, False)
            Case 8 'apclient
                Return getAPClient(sValue, False)
        End Select

        Return False
    End Function

    Public Sub SearchDetail(ByVal nDtlRow As Integer, ByVal nIndex As Integer, ByVal sValue As String)
        If nDtlRow > p_oDTDetl.Rows.Count - 1 Then Exit Sub

        Select Case nIndex
            Case 2
                getAccount(nDtlRow, sValue, True, True)
            Case 3
                getAccount(nDtlRow, sValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer

        With p_oDTMstr
            .Rows.Add()
            For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
                Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                    Case "stransnox"
                        p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)
                    Case "ntrantotl"
                        p_oDTMstr(0).Item(lnCtr) = 0.0
                    Case "napprtotl"
                        p_oDTMstr(0).Item(lnCtr) = 0.0
                    Case "dtransact"
                        p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                    Case "drecvdate"
                        p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                    Case "dapprovdt"
                        p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                    Case "nentrynox"
                        p_oDTMstr(0).Item(lnCtr) = 0
                    Case "ctranstat"
                        p_oDTMstr(0).Item(lnCtr) = "0"
                    Case Else
                        p_oDTMstr(0).Item(lnCtr) = ""
                End Select
            Next
        End With

        p_sAPClientNm = ""
    End Sub

    Private Sub initDetail()
        Dim lnCtr As Integer

        With p_oDTDetl
            .Rows.Add()
            For lnCtr = 0 To p_oDTDetl.Columns.Count - 1
                Select Case LCase(p_oDTDetl.Columns(lnCtr).ColumnName)
                    Case "stransnox"
                        .Rows(0)(lnCtr) = p_oDTMstr(0).Item("sTransNox")
                    Case "dtransact"
                        .Rows(0)(lnCtr) = p_oApp.SysDate
                    Case "cbilltype"
                        .Rows(0).Item(lnCtr) = "0"
                    Case "namountxx", "napproved", "nprincipl", "ninterest", "nsubsidze", "ninctvamt"
                        .Rows(0).Item(lnCtr) = 0.0
                    Case "nentrynox"
                        .Rows(0).Item(lnCtr) = 0
                    Case Else
                        .Rows(0).Item(lnCtr) = ""
                End Select
            Next
        End With

    End Sub

    Private Sub createDetailTable()
        p_oDTDetl = New DataTable
        With p_oDTDetl
            .Columns.Add("nEntryNox", GetType(Integer))
            .Columns.Add("dTransact", GetType(Date))
            .Columns.Add("sAcctNmbr", GetType(String)).MaxLength = 10
            .Columns.Add("sClientNm", GetType(String)).MaxLength = 128
            .Columns.Add("sEngineNo", GetType(String)).MaxLength = 30
            .Columns.Add("nPrincipl", GetType(Decimal))
            .Columns.Add("nInterest", GetType(Decimal))
            .Columns.Add("nSubsidze", GetType(Decimal))
            .Columns.Add("nInctvAmt", GetType(Decimal))
            .Columns.Add("cBillType", GetType(Char))
            .Columns.Add("sDescript", GetType(String)).MaxLength = 128
            .Columns.Add("sRemarks1", GetType(String)).MaxLength = 128
            .Columns.Add("sRemarks2", GetType(String)).MaxLength = 128
            .Columns.Add("nAmountxx", GetType(Decimal))
            .Columns.Add("nApproved", GetType(Decimal))
            .Columns.Add("sTransNox", GetType(String)).MaxLength = 12
            .Columns.Add("sClientID", GetType(String)).MaxLength = 12
            .Columns.Add("sSerialID", GetType(String)).MaxLength = 12
        End With
    End Sub

    Public Function AddDetail() As Boolean
        Dim lnRow As Integer

        If IFNull(p_oDTDetl(p_oDTDetl.Rows.Count - 1)("sAcctNmbr")) = "" Then
            Return False
        End If

        p_oDTDetl.Rows.Add()
        lnRow = p_oDTDetl.Rows.Count - 1
        p_oDTDetl(lnRow)("sAcctNmbr") = ""
        p_oDTDetl(lnRow)("nEntryNox") = p_oDTDetl.Rows.Count
        p_oDTDetl(lnRow)("cBillType") = "0"
        p_oDTDetl(lnRow)("nAmountxx") = 0.0
        p_oDTDetl(lnRow)("nApproved") = 0.0
        p_oDTDetl(lnRow)("nprincipl") = 0.0
        p_oDTDetl(lnRow)("ninterest") = 0.0
        p_oDTDetl(lnRow)("nSubsidze") = 0.0
        p_oDTDetl(lnRow)("nInctvAmt") = 0.0
        Return True
    End Function

    Public Sub DeleteDetail(ByVal lnRow As Integer)
        p_oDTDetl.Rows(lnRow).Delete()

        p_oDTDetl.AcceptChanges()

        If p_oDTDetl.Rows.Count = 0 Then AddDetail()
    End Sub

    Private Function getCompany(ByVal sValue As String, ByVal bSearch As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getCompany"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If Not bSearch Then
                lsCondition = "sCompnyNm LIKE " & strParm(sValue & "%")
            Else
                lsCondition = "sCompnyNm = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Company, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMstr
                .Rows(0)("sCompnyID") = loDT(0)("sCompnyID")
                .Rows(0)("sCompnyNm") = loDT(0)("sCompnyNm")
            End With
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sCompnyID»sCompnyNm", _
                                "CompanyID»Company", _
                                "", _
                                "", _
                                2)

            If Not IsNothing(loDataRow) Then
                With p_oDTMstr
                    .Rows(0)("sCompnyID") = loDataRow("sCompnyID")
                    .Rows(0)("sCompnyNm") = loDataRow("sCompnyNm")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMstr
            RaiseEvent MasterRetrieved(2, .Rows(0)("sCompnyNm"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMstr
            .Rows(0)("sCompnyID") = ""
            .Rows(0)("sCompnyNm") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getAPClient(ByVal sValue As String, ByVal bSearch As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getAPClient"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If Not bSearch Then
                lsCondition = "a.sCompnyNm LIKE " & strParm(sValue & "%")
            Else
                lsCondition = "a.sCompnyNm = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Client, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMstr
                .Rows(0)("sClientID") = loDT(0)("sClientID")
                p_sAPClientNm = loDT(0)("sCompnyNm")
            End With
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sClientID»sCompnyNm", _
                                "Client ID»Company", _
                                "", _
                                "a.sClientID»a.sCompnyNm", _
                                2)

            If Not IsNothing(loDataRow) Then
                With p_oDTMstr
                    .Rows(0)("sClientID") = loDataRow("sClientID")
                    p_sAPClientNm = loDataRow("sCompnyNm")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMstr
            RaiseEvent MasterRetrieved(8, p_sAPClientNm)
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMstr
            .Rows(0)("sClientID") = ""
            p_sAPClientNm = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getBranch(ByVal sValue As String, ByVal bSearch As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getBranch"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If Not bSearch Then
                lsCondition = "sBranchNm LIKE " & strParm(sValue & "%")
            Else
                lsCondition = "sBranchNm = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Branch, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMstr
                .Rows(0)("sBranchCd") = loDT(0)("sBranchCd")
                .Rows(0)("sBranchNm") = loDT(0)("sBranchNm")
            End With
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sBranchCd»sBranchNm", _
                                "BranchID»Branch", _
                                "", _
                                "", _
                                2)

            If Not IsNothing(loDataRow) Then
                With p_oDTMstr
                    .Rows(0)("sBranchCd") = loDataRow("sBranchCd")
                    .Rows(0)("sBranchNm") = loDataRow("sBranchNm")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMstr
            RaiseEvent MasterRetrieved(3, .Rows(0)("sBranchNm"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMstr
            .Rows(0)("sBranchCd") = ""
            .Rows(0)("sBranchNm") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Sub getAccount(ByVal fnDtlRow As Integer _
                          , ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)

        If fbIsCode Then
            If fsValue = p_oDTDetl(fnDtlRow).Item("sAcctNmbr") And fsValue <> "" Then Exit Sub
        Else
            If fsValue = IFNull(p_oDTDetl(fnDtlRow).Item("sClientNm"), "") And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sAcctNmbr" & _
                       ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) sClientNm" & _
                       ", d.sEngineNo" & _
                       ", a.nPrincipl" & _
                       ", a.nInterest" & _
                       ", c.nSubsidze" & _
                       ", c.nInctvAmt" & _
                       ", a.dTransact" & _
                " FROM LR_Master a" & _
                    ", Client_Master b" & _
                    ", LR_Master_Car c" & _
                    ", Car_Serial d" & _
               " WHERE a.sClientID = b.sClientID" & _
                 " AND a.sAcctNmbr = c.sAcctNmbr" & _
                 " AND c.sSerialID = d.sSerialID"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                            , lsSQL _
                                            , True _
                                            , fsValue _
                                            , "sAcctNmbr»sClientNm»sEngineNo" _
                                            , "AcctNo»Client»Engine", _
                                            , "a.sAcctNmbr»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)»d.sEngineNo" _
                                            , IIf(fbIsCode, 0, 1))

            If IsNothing(loRow) Then
                p_oDTDetl(fnDtlRow).Item("sAcctNmbr") = ""
                p_oDTDetl(fnDtlRow).Item("sClientNm") = ""
                p_oDTDetl(fnDtlRow).Item("sEngineNo") = ""
                p_oDTDetl(fnDtlRow).Item("nPrincipl") = 0.0
                p_oDTDetl(fnDtlRow).Item("nInterest") = 0.0
                p_oDTDetl(fnDtlRow).Item("nSubsidze") = 0.0
                p_oDTDetl(fnDtlRow).Item("nInctvAmt") = 0.0
                p_oDTDetl(fnDtlRow).Item("dTransact") = Format((p_oApp.getSysDate), xsDATE_MEDIUM)
            Else
                p_oDTDetl(fnDtlRow).Item("sAcctNmbr") = loRow.Item("sAcctNmbr")
                p_oDTDetl(fnDtlRow).Item("sClientNm") = loRow.Item("sClientNm")
                p_oDTDetl(fnDtlRow).Item("sEngineNo") = loRow.Item("sEngineNo")
                p_oDTDetl(fnDtlRow).Item("nPrincipl") = loRow.Item("nPrincipl")
                p_oDTDetl(fnDtlRow).Item("nInterest") = loRow.Item("nInterest")
                p_oDTDetl(fnDtlRow).Item("nSubsidze") = loRow.Item("nSubsidze")
                p_oDTDetl(fnDtlRow).Item("nInctvAmt") = loRow.Item("nInctvAmt")
                p_oDTDetl(fnDtlRow).Item("dTransact") = loRow.Item("dTransact")
            End If

            RaiseEvent DetailRetrieved(fnDtlRow, 1, p_oDTDetl(fnDtlRow).Item("dTransact"))
            RaiseEvent DetailRetrieved(fnDtlRow, 2, p_oDTDetl(fnDtlRow).Item("sAcctNmbr"))
            RaiseEvent DetailRetrieved(fnDtlRow, 3, p_oDTDetl(fnDtlRow).Item("sClientNm"))
            RaiseEvent DetailRetrieved(fnDtlRow, 4, p_oDTDetl(fnDtlRow).Item("sEngineNo"))
            RaiseEvent DetailRetrieved(fnDtlRow, 5, p_oDTDetl(fnDtlRow).Item("nPrincipl"))
            RaiseEvent DetailRetrieved(fnDtlRow, 6, p_oDTDetl(fnDtlRow).Item("nInterest"))
            RaiseEvent DetailRetrieved(fnDtlRow, 7, p_oDTDetl(fnDtlRow).Item("nSubsidze"))
            RaiseEvent DetailRetrieved(fnDtlRow, 8, p_oDTDetl(fnDtlRow).Item("nInctvAmt"))
            Exit Sub
        End If

        If fbIsCode Then
            lsSQL = AddCondition(lsSQL, "a.sAcctNmbr = " & strParm(fsValue))
        Else
            lsSQL = AddCondition(lsSQL, "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) = " & strParm(fsValue))
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTDetl(fnDtlRow).Item("sAcctNmbr") = ""
            p_oDTDetl(fnDtlRow).Item("sClientNm") = ""
            p_oDTDetl(fnDtlRow).Item("sEngineNo") = ""
            p_oDTDetl(fnDtlRow).Item("nPrincipl") = 0
            p_oDTDetl(fnDtlRow).Item("nInterest") = 0
            p_oDTDetl(fnDtlRow).Item("nInctvAmt") = 0
            p_oDTDetl(fnDtlRow).Item("nSubsidze") = 0
            p_oDTDetl(fnDtlRow).Item("dTransact") = p_oApp.getSysDate
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTDetl(fnDtlRow).Item("sAcctNmbr") = loDta(0).Item("sAcctNmbr")
            p_oDTDetl(fnDtlRow).Item("sClientNm") = loDta(0).Item("sClientNm")
            p_oDTDetl(fnDtlRow).Item("sEngineNo") = loDta(0).Item("sEngineNo")
            p_oDTDetl(fnDtlRow).Item("nPrincipl") = loDta(0).Item("nPrincipl")
            p_oDTDetl(fnDtlRow).Item("nInterest") = loDta(0).Item("nInterest")
            p_oDTDetl(fnDtlRow).Item("nSubsidze") = loDta(0).Item("nSubsidze")
            p_oDTDetl(fnDtlRow).Item("nInctvAmt") = loDta(0).Item("nInctvAmt")
            p_oDTDetl(fnDtlRow).Item("dTransact") = loDta(0).Item("dTransact")
        End If

        RaiseEvent DetailRetrieved(fnDtlRow, 1, p_oDTDetl(fnDtlRow).Item("dTransact"))
        RaiseEvent DetailRetrieved(fnDtlRow, 2, p_oDTDetl(fnDtlRow).Item("sAcctNmbr"))
        RaiseEvent DetailRetrieved(fnDtlRow, 3, p_oDTDetl(fnDtlRow).Item("sClientNm"))
        RaiseEvent DetailRetrieved(fnDtlRow, 4, p_oDTDetl(fnDtlRow).Item("sEngineNo"))
        RaiseEvent DetailRetrieved(fnDtlRow, 5, p_oDTDetl(fnDtlRow).Item("nPrincipl"))
        RaiseEvent DetailRetrieved(fnDtlRow, 6, p_oDTDetl(fnDtlRow).Item("nInterest"))
        RaiseEvent DetailRetrieved(fnDtlRow, 7, p_oDTDetl(fnDtlRow).Item("nSubsidze"))
        RaiseEvent DetailRetrieved(fnDtlRow, 8, p_oDTDetl(fnDtlRow).Item("nInctvAmt"))
    End Sub

    Private Function getSQL_Browse() As String
        Return "SELECT" & _
                    "  a.sTransNox" & _
                    ", b.sCompnyNm" & _
                    ", a.dTransact" & _
                " FROM " & p_sMasTable & " a" & _
                    ", Company b" & _
                " WHERE a.sCompnyID = b.sCompnyID"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
        createDetailTable()
    End Sub

    Private Function getSQL_Master() As String
        Return "SELECT" & _
                    "  a.sTransNox" & _
                    ", a.dTransact" & _
                    ", b.sCompnyNm" & _
                    ", c.sBranchNm" & _
                    ", a.nTranTotl" & _
                    ", a.nApprTotl" & _
                    ", a.sRemarks1" & _
                    ", a.sRemarks2" & _
                    ", d.sCompnyNm xAPClient" & _
                    ", a.cTranStat" & _
                    ", a.sRecvByxx" & _
                    ", a.dRecvDate" & _
                    ", a.sApprovBy" & _
                    ", a.dApprovDt" & _
                    ", a.sClientID" & _
                    ", a.sCompnyID" & _
                    ", a.sBranchCD" & _
                    ", a.nEntryNox" & _
                " FROM " & p_sMasTable & " a" & _
                    ", Company b" & _
                    ", Branch c" & _
                    ", Client_Master d" & _
                " WHERE a.sCompnyID = b.sCompnyID" & _
                    " AND a.sBranchCd = c.sBranchCd" & _
                    " AND a.sClientID = d.sClientID"

    End Function

    Private Function getSQL_Detail() As String
        Return "SELECT" & _
                    "  a.nEntryNox" & _
                    ", a.sReferNox" & _
                    ", a.cBillType" & _
                    ", a.sDescript" & _
                    ", a.sRemarks1" & _
                    ", a.sRemarks2" & _
                    ", a.nAmountxx" & _
                    ", a.nApproved" & _
                    ", a.sTransNox" & _
                    ", b.nPrincipl" & _
                    ", b.nInterest" & _
                    ", e.sEngineNo" & _
                    ", CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName) sClientNm" & _
                    ", b.dTransact" & _
                    ", c.sClientID" & _
                    ", e.sSerialID" & _
                    ", d.nSubsidze" & _
                    ", d.nInctvAmt" & _
                " FROM " & p_sDetTable & " a" & _
                    ", LR_Master b" & _
                    ", Client_Master c" & _
                    ", LR_Master_Car d" & _
                    ", Car_Serial e" & _
                " WHERE a.sReferNox = b.sAcctNmbr" & _
                    " AND b.sClientID = c.sClientID" & _
                    " AND b.sAcctNmbr = d.sAcctNmbr" & _
                    " AND d.sSerialID = e.sSerialID" & _
                " ORDER BY a.nEntryNox ASC"
    End Function

    Private Function getSQL_Company() As String
        Return "SELECT" & _
                    "  sCompnyID" & _
                    ", sCompnyNm" & _
                " FROM Company" & _
                " WHERE cRecdStat = " & strParm(xeRecordStat.RECORD_NEW)
    End Function

    Private Function getSQL_Branch() As String
        Return "SELECT" & _
                    "  sBranchCd" & _
                    ", sBranchNm" & _
                " FROM Branch" & _
                " WHERE cRecdStat = " & strParm(xeRecordStat.RECORD_NEW)
    End Function

    Private Function getSQL_Client() As String
        Return "SELECT" & _
                    "  a.sClientID" & _
                    ", a.sCompnyNm" & _
                " FROM Client_Master a" & _
                    ", AP_Client_Master b" & _
                " WHERE a.sClientID = b.sClientID" & _
                    " AND b.cRecdStat = " & strParm(xeLogical.YES)
    End Function
End Class
