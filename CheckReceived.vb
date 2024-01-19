'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Check Received Object
'
' Copyright 2017 and Beyond
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
'  Kalyptus [ 07/08/2017 02:33 pm ]
'      Started creating this object.
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports System.Drawing

Public Class CheckReceived
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetl As DataTable
    Private p_oOthersx As New Others
    Private p_nEditMode As xeEditMode
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String
    Private p_nPurpose As Integer

    Private Const p_sMasTable As String = "Checks_Received"
    Private Const p_sMsgHeadr As String = "Checks Received"

    Public Const xeCheckStatOpen As String = "0"
    Public Const xeCheckStatCleared As String = "1"
    Public Const xeCheckStatBounce As String = "2"
    Public Const xeCheckStatCancelled As String = "3"
    Public Const xeCheckStatHold As String = "4"

    Public Const xePurposeCreate = 0
    Public Const xePurposeChange = 1
    Public Const xePurposeView = 2

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Branch As String
        Get
            Return p_sBranchCd
        End Get
        Set(value As String)
            'If Product ID is LR then do allow changing of Branch
            If p_oApp.ProductID = "LRTrackr" Then
                p_sBranchCd = value
            End If
        End Set
    End Property

    Public ReadOnly Property ItemNo As Integer
        Get
            If p_nEditMode = xeEditMode.MODE_READY And Not p_oDTDetl Is Nothing Then
                Return p_oDTDetl.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 'sBankName
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sBankName) = "" Then
                            getBank(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sBankName
                    Case 5
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.nAmountxx
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case 11
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.sReferNox
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case 12
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.sSourceCD
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                If p_nPurpose = xePurposeCreate Then
                    Select Case Index
                        Case 80 'sBankName
                            getBank(2, 80, value, False, False)
                        Case 1, 2, 3, 13 'sCheckNox, sBankIDxx, sAcctNoxx, sRecTrans
                            p_oDTMstr(0).Item(Index) = value
                        Case 4  'dCheckDte
                            If IsDate(value) Then
                                p_oDTMstr(0).Item(Index) = Convert.ToDateTime(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                        Case 5  'nAmountxx
                            If IsNumeric(value) Then
                                p_oOthersx.nAmountxx = Convert.ToDecimal(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oOthersx.nAmountxx)
                        Case 6  'nClearing
                            If IsNumeric(value) Then
                                p_oDTMstr(0).Item(Index) = Convert.ToInt32(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                        Case 11 'sReferNox
                            p_oOthersx.sReferNox = value
                        Case 12 'sSourceCD  
                            p_oOthersx.sSourceCD = value
                    End Select
                End If
            End If
        End Set
    End Property

    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "sbankname" '80 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sBankName) = "" Then
                            getBank(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sBankName
                    Case "namountxx" '5  
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.nAmountxx
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case "srefernox" ' 11 
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.sReferNox
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case "ssourcecd" '12
                        If p_nPurpose = xePurposeCreate Then
                            Return p_oOthersx.sSourceCD
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                If p_nPurpose = xePurposeCreate Then
                    Select Case LCase(Index)
                        Case "sbankname" '80
                            getBank(2, 80, value, False, False)
                        Case "schecknox", "sbankidxx", "sacctnoxx", "srectrans" '1, 2, 3, 13 
                            p_oDTMstr(0).Item(Index) = value
                        Case "dcheckdte" '4  
                            If IsDate(value) Then
                                p_oDTMstr(0).Item(Index) = Convert.ToDateTime(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                        Case "namountxx" '5  
                            If IsNumeric(value) Then
                                p_oOthersx.nAmountxx = Convert.ToDecimal(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oOthersx.nAmountxx)
                        Case "nclearing" '6  
                            If IsNumeric(value) Then
                                p_oDTMstr(0).Item(Index) = Convert.ToInt32(value)
                            End If
                            RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                        Case "srefernox" ' 11 
                            p_oOthersx.sReferNox = value
                        Case "ssourcecd" '12
                            p_oOthersx.sSourceCD = value
                    End Select
                End If
            End If
        End Set
    End Property

    Public Property Detail(ByVal Row As Integer, ByVal Index As Integer) As Object
        Get
            If p_nEditMode = xeEditMode.MODE_READY Then
                Return p_oDTDetl(Row).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode = xeEditMode.MODE_READY Then
                If p_nPurpose = xePurposeChange Then
                    If Index = 5 Then
                        p_oDTDetl(Row).Item(Index) = value
                    End If
                End If
            End If
        End Set
    End Property

    Public Property Detail(ByVal Row As Integer, ByVal Index As String) As Object
        Get
            If p_nEditMode = xeEditMode.MODE_READY Then
                Return p_oDTDetl(Row).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode = xeEditMode.MODE_READY Then
                If p_nPurpose = xePurposeChange Then
                    If LCase(Index) = "sornoxxxx" Then
                        p_oDTDetl(Row).Item(Index) = value
                    End If
                End If
            End If
        End Set
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    'Property ()
    Public ReadOnly Property BranchCode() As String
        Get
            Return p_sBranchCd
        End Get
    End Property

    Public ReadOnly Property BranchName() As String
        Get
            Return p_sBranchNm
        End Get
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    'Public Function NewTransaction()
    Private Function NewTransaction() As Boolean
        Dim lsSQL As String

        If p_sBranchCd = "" Then
            MsgBox("Branch is empty... Please indicate branch!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        lsSQL = AddCondition(getSQ_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())
        Call initMaster()
        Call InitOthers()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    'Public Function OpenTransaction(String)
    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If



        Call InitOthers()

        If p_nPurpose = xePurposeCreate Then
            p_oDTDetl = Nothing
        Else
            LoadDetail()
        End If

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    Private Sub LoadDetail()
        Dim lsSQL As String

        lsSQL = "SELECT sTransNox" & _
                     ", sReferNox" & _
                     ", sSourceCd" & _
                     ", nAmountxx" & _
                     ", NOW() dTransact" & _
                     ", '' sORNoxxxx" & _
                     ", '' sPRNoxxxx" & _
                     ", '' sAcctNmbr" & _
                     ", '' sClientID" & _
                     ", '' xFullName" & _
                     ", '' xAddressx" & _
                     ", '' sEngineNo" & _
                     ", '' cTranType" & _
                     ", '' sCollIDxx" & _
                     ", 0.00 nTranAmtx" & _
                     ", 0.00 nIntAmtxx" & _
                     ", 0.00 nRebatesx" & _
                     ", 0.00 nPenaltyx" & _
                     ", 'Checks_Received' sChckTble" & _
                     ", '' sTranTble" & _
               " FROM Checks_Received" & _
               " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
               " UNION" & _
               " SELECT sTransNox" & _
                     ", sReferNox" & _
                     ", sSourceCd" & _
                     ", nAmountxx" & _
                     ", NOW() dTransact" & _
                     ", '' sORNoxxxx" & _
                     ", '' sPRNoxxxx" & _
                     ", '' sAcctNmbr" & _
                     ", '' sClientID" & _
                     ", '' xFullName" & _
                     ", '' xAddressx" & _
                     ", '' sEngineNo" & _
                     ", '' cTranType" & _
                     ", '' sCollIDxx" & _
                     ", 0.00 nTranAmtx" & _
                     ", 0.00 nIntAmtxx" & _
                     ", 0.00 nRebatesx" & _
                     ", 0.00 nPenaltyx" & _
                     ", 'Checks_Received_Others' sChckTble" & _
                     ", '' sTranTble" & _
               " FROM Checks_Received_Others" & _
               " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
        p_oDTDetl = p_oApp.ExecuteQuery(lsSQL)

        Dim lnRow As Integer
        'Make sure that check should be from AR/LR Payments
        For lnRow = 0 To p_oDTDetl.Rows.Count - 1
            If p_oDTDetl(lnRow).Item("sTranTble") = "" Then
                Select Case LCase(p_oDTDetl(lnRow).Item("sSourceCD"))
                    Case "prec"
                        getPRecTrans(lnRow)
                    Case "lrpy"
                        getLRPyTrans(lnRow)
                    Case "arpy"
                        getARPyTrans(lnRow)
                End Select
            End If
        Next
    End Sub

    Private Sub getPRecTrans(ByVal fnRow As Integer)
        Dim loDta As DataTable
        Dim lsSQL As String

        lsSQL = "SELECT a.dTransact" & _
                     ", a.sPRNoxxxx" & _
                     ", a.sAcctNmbr" & _
                     ", a.sClientID" & _
                     ", c.sCompnyNm AS xFullName" & _
                     ", CONCAT(IF(IFNULL(c.sHouseNox, '') = '', '', CONCAT(c.sHouseNox, ' ')), c.sAddressx, ', ', d.sTownName, ', ', e.sProvName) AS xAddressx" & _
                     ", f.sEngineNo" & _
                     ", a.cTranType" & _
                     ", a.nTranAmtx" & _
                     ", a.nDiscount nRebatesx" & _
                     ", a.nOthersxx nPenaltyx" & _
                     ", a.sCollctId sCollIDxx" & _
                     ", 'Provisionary_Receipt_Master' sTranTble" & _
               " FROM Provisionary_Receipt_Master a" & _
                    " LEFT JOIN MC_AR_Master b ON a.sAcctNmbr = b.sAcctNmbr" & _
                    " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
                    " LEFT JOIN TownCity d ON c.sTownIDxx = d.sTownIDxx" & _
                    " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
                    " LEFT JOIN MC_Serial f ON b.sSerialID = f.sSerialID" & _
               " WHERE a.sTransNox = " & strParm(p_oDTDetl(fnRow).Item("sReferNox")) & _
               " UNION" & _
               " SELECT a.dTransact" & _
                     ", a.sPRNoxxxx" & _
                     ", a.sAcctNmbr" & _
                     ", a.sClientID" & _
                     ", c.sCompnyNm AS xFullName" & _
                     ", CONCAT(IF(IFNULL(c.sHouseNox, '') = '', '', CONCAT(c.sHouseNox, ' ')), c.sAddressx, ', ', d.sTownName, ', ', e.sProvName) AS xAddressx" & _
                     ", f.sEngineNo" & _
                     ", a.cTranType" & _
                     ", a.nTranAmtx" & _
                     ", a.nDiscount nRebatesx" & _
                     ", a.nOthersxx nPenaltyx" & _
                     ", a.sCollctId sCollIDxx" & _
                     ", 'Provisionary_Receipt_Others' sTranTble" & _
               " FROM Provisionary_Receipt_Others a" & _
                    " LEFT JOIN MC_AR_Master b ON a.sAcctNmbr = b.sAcctNmbr" & _
                    " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
                    " LEFT JOIN TownCity d ON c.sTownIDxx = d.sTownIDxx" & _
                    " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
                    " LEFT JOIN MC_Serial f ON b.sSerialID = f.sSerialID" & _
               " WHERE a.sTransNox = " & strParm(p_oDTDetl(fnRow).Item("sReferNox"))
        '" AND a.cTranStat IN ('0', '1')"
        loDta = p_oApp.ExecuteQuery(lsSQL)

        p_oDTDetl(fnRow).Item("dTransact") = loDta(0).Item("dTransact")
        p_oDTDetl(fnRow).Item("sPRNoxxxx") = loDta(0).Item("sPRNoxxxx")
        p_oDTDetl(fnRow).Item("sAcctNmbr") = loDta(0).Item("sAcctNmbr")
        p_oDTDetl(fnRow).Item("sClientID") = loDta(0).Item("sClientID")
        p_oDTDetl(fnRow).Item("xFullName") = loDta(0).Item("xFullName")
        p_oDTDetl(fnRow).Item("xAddressx") = loDta(0).Item("xAddressx")
        p_oDTDetl(fnRow).Item("sEngineNo") = loDta(0).Item("sEngineNo")
        p_oDTDetl(fnRow).Item("cTranType") = Convert.ToInt16(loDta(0).Item("cTranType")) + 1
        p_oDTDetl(fnRow).Item("nTranAmtx") = loDta(0).Item("nTranAmtx")
        p_oDTDetl(fnRow).Item("nRebatesx") = loDta(0).Item("nRebatesx")
        p_oDTDetl(fnRow).Item("nPenaltyx") = loDta(0).Item("nPenaltyx")
        p_oDTDetl(fnRow).Item("sCollIDxx") = loDta(0).Item("sCollIDxx")
        p_oDTDetl(fnRow).Item("sTranTble") = loDta(0).Item("sTranTble")

        If loDta.Rows.Count > 1 Then
            Dim lnRow As Integer
            For lnRow = 1 To loDta.Rows.Count - 1
                p_oDTDetl.Rows.Add(p_oDTDetl.NewRow)
                p_oDTDetl.Rows(p_oDTDetl.Rows.Count - 1).Item("sTransNox") = p_oDTDetl(0).Item("sTransNox")
                p_oDTDetl.Rows(p_oDTDetl.Rows.Count - 1).Item("sReferNox") = p_oDTDetl(0).Item("sReferNox")
                p_oDTDetl.Rows(p_oDTDetl.Rows.Count - 1).Item("sSourceCd") = p_oDTDetl(0).Item("sSourceCd")
                p_oDTDetl.Rows(p_oDTDetl.Rows.Count - 1).Item("nAmountxx") = p_oDTDetl(0).Item("nAmountxx")
                p_oDTDetl.Rows(p_oDTDetl.Rows.Count - 1).Item("sChckTble") = p_oDTDetl(0).Item("sChckTble")

                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("dTransact") = loDta(0).Item("dTransact")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sPRNoxxxx") = loDta(0).Item("sPRNoxxxx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sAcctNmbr") = loDta(0).Item("sAcctNmbr")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sClientID") = loDta(0).Item("sClientID")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("xFullName") = loDta(0).Item("xFullName")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("xAddressx") = loDta(0).Item("xAddressx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sEngineNo") = loDta(0).Item("sEngineNo")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("cTranType") = Convert.ToInt16(loDta(0).Item("cTranType")) + 1
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("nTranAmtx") = loDta(0).Item("nTranAmtx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("nRebatesx") = loDta(0).Item("nRebatesx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("nPenaltyx") = loDta(0).Item("nPenaltyx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sCollIDxx") = loDta(0).Item("sCollIDxx")
                p_oDTDetl(p_oDTDetl.Rows.Count - 1).Item("sTranTble") = loDta(0).Item("sTranTble")
            Next
        End If
    End Sub

    Private Sub getLRPyTrans(ByVal fnRow As Integer)
        Dim loDta As DataTable
        Dim lsSQL As String

        lsSQL = "SELECT a.dTransact" & _
                     ", a.sReferNox sPRNoxxxx" & _
                     ", a.sAcctNmbr" & _
                     ", a.sClientID" & _
                     ", c.sCompnyNm AS xFullName" & _
                     ", CONCAT(IF(IFNULL(c.sHouseNox, '') = '', '', CONCAT(c.sHouseNox, ' ')), c.sAddressx, ', ', d.sTownName, ', ', e.sProvName) AS xAddressx" & _
                     ", a.cTranType" & _
                     ", a.nAmountxx nTranAmtx" & _
                     ", a.nRebatesx" & _
                     ", a.nIntAmtxx" & _
                     ", a.nPenaltyx" & _
                     ", a.sCollIDxx" & _
                     ", 'LR_Payment_Master_PR' sTranTble" & _
               " FROM LR_Payment_Master_PR a" & _
                    " LEFT JOIN LR_Master b ON a.sAcctNmbr = b.sAcctNmbr" & _
                    " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
                    " LEFT JOIN TownCity d ON c.sTownIDxx = d.sTownIDxx" & _
                    " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
               " WHERE a.sTransNox = " & strParm(p_oDTDetl(fnRow).Item("sReferNox"))

        loDta = p_oApp.ExecuteQuery(lsSQL)

        p_oDTDetl(fnRow).Item("dTransact") = loDta(0).Item("dTransact")
        p_oDTDetl(fnRow).Item("sPRNoxxxx") = loDta(0).Item("sPRNoxxxx")
        p_oDTDetl(fnRow).Item("sAcctNmbr") = loDta(0).Item("sAcctNmbr")
        p_oDTDetl(fnRow).Item("sClientID") = loDta(0).Item("sClientID")
        p_oDTDetl(fnRow).Item("xFullName") = loDta(0).Item("xFullName")
        p_oDTDetl(fnRow).Item("xAddressx") = loDta(0).Item("xAddressx")
        p_oDTDetl(fnRow).Item("cTranType") = loDta(0).Item("cTranType")
        p_oDTDetl(fnRow).Item("nTranAmtx") = loDta(0).Item("nTranAmtx")
        p_oDTDetl(fnRow).Item("nRebatesx") = loDta(0).Item("nRebatesx")
        p_oDTDetl(fnRow).Item("nPenaltyx") = loDta(0).Item("nPenaltyx")
        p_oDTDetl(fnRow).Item("nIntAmtxx") = loDta(0).Item("nIntAmtxx")
        p_oDTDetl(fnRow).Item("sCollIDxx") = loDta(0).Item("sCollIDxx")
        p_oDTDetl(fnRow).Item("sTranTble") = loDta(0).Item("sTranTble")

        'If p_oDTDetl(fnRow).Item("sReferNox") = p_oDTMstr(0).Item("sReferNox") Then
        '    p_oDTDetl(fnRow).Item("nAmountxx") = p_oDTDetl(fnRow).Item("nTranAmtx") + p_oDTDetl(fnRow).Item("nPenaltyx") + p_oDTDetl(fnRow).Item("nIntAmtxx")
        'End If
    End Sub

    Private Sub getARPyTrans(ByVal fnRow As Integer)
        Dim loDta As DataTable
        Dim lsSQL As String

        lsSQL = "SELECT a.dTransact" & _
                     ", a.sReferNox sPRNoxxxx" & _
                     ", a.sAcctNmbr" & _
                     ", a.sClientID" & _
                     ", c.sCompnyNm AS xFullName" & _
                     ", CONCAT(IF(IFNULL(c.sHouseNox, '') = '', '', CONCAT(c.sHouseNox, ' ')), c.sAddressx, ', ', d.sTownName, ', ', e.sProvName) AS xAddressx" & _
                     ", Ifnull(f.sEngineNo,'') sEngineNo" & _
                     ", a.cTranType" & _
                     ", a.nAmountxx nTranAmtx" & _
                     ", a.nRebatesx" & _
                     ", a.nPenaltyx" & _
                     ", a.sCollIDxx" & _
                     ", 'LR_Payment_Master_PR' sTranTble" & _
               " FROM LR_Payment_Master_PR a" & _
                    " LEFT JOIN MC_AR_Master b ON a.sAcctNmbr = b.sAcctNmbr" & _
                    " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
                    " LEFT JOIN TownCity d ON c.sTownIDxx = d.sTownIDxx" & _
                    " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
                    " LEFT JOIN MC_Serial f ON b.sSerialID = f.sSerialID" & _
               " WHERE a.sTransNox = " & strParm(p_oDTDetl(fnRow).Item("sReferNox"))

        '                 " AND a.cPostedxx IN ('0', '1')"

        loDta = p_oApp.ExecuteQuery(lsSQL)

        p_oDTDetl(fnRow).Item("dTransact") = loDta(0).Item("dTransact")
        p_oDTDetl(fnRow).Item("sPRNoxxxx") = loDta(0).Item("sPRNoxxxx")
        p_oDTDetl(fnRow).Item("sAcctNmbr") = loDta(0).Item("sAcctNmbr")
        p_oDTDetl(fnRow).Item("sClientID") = loDta(0).Item("sClientID")
        p_oDTDetl(fnRow).Item("xFullName") = loDta(0).Item("xFullName")
        p_oDTDetl(fnRow).Item("xAddressx") = loDta(0).Item("xAddressx")
        p_oDTDetl(fnRow).Item("sEngineNo") = loDta(0).Item("sEngineNo")
        p_oDTDetl(fnRow).Item("cTranType") = loDta(0).Item("cTranType")
        p_oDTDetl(fnRow).Item("nTranAmtx") = loDta(0).Item("nTranAmtx")
        p_oDTDetl(fnRow).Item("nRebatesx") = loDta(0).Item("nRebatesx")
        p_oDTDetl(fnRow).Item("nPenaltyx") = loDta(0).Item("nPenaltyx")
        p_oDTDetl(fnRow).Item("sCollIDxx") = loDta(0).Item("sCollIDxx")
        p_oDTDetl(fnRow).Item("sTranTble") = loDta(0).Item("sTranTble")

        'If p_oDTDetl(fnRow).Item("sReferNox") = p_oDTMstr(0).Item("sReferNox") Then
        '    p_oDTDetl(fnRow).Item("nAmountxx") = p_oDTDetl(fnRow).Item("nTranAmtx") + p_oDTDetl(fnRow).Item("nPenaltyx")
        'End If

    End Sub

    'Public Function SearchTransaction(String, Boolean, Boolean=False)
    Public Function SearchTransaction( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sCheckNox") Then Return True
            Else
                If fsValue = p_oDTDetl(0).Item("xFullName") Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = Replace(getSQ_Browse(), "<cChckStat>", strDissect(p_nTranStat))
        Else
            lsSQL = Replace(getSQ_Browse(), "<cChckStat>", "'0', '1', '2', '3', '4'")
        End If

        If p_sBranchCd <> "" Then
            lsSQL = Replace(lsSQL, "<sBranchCD>", p_sBranchCd)
        Else
            lsSQL = Replace(lsSQL, "<sBranchCD>", "")
        End If

        'create Kwiksearch filter
        Dim lsFilter As String = ""

        If fbByCode Then
            lsSQL = Replace(lsSQL, "<sAddCondx>", " AND a.sCheckNox LIKE " & strParm(fsValue))
        Else
            lsSQL = Replace(lsSQL, "<sAddCondx>", " AND c.sCompnyNm LIKE " & strParm(fsValue & "%"))
        End If

        'If fbByCode Then
        '    lsFilter = "a.sCheckNox LIKE " & strParm(fsValue)
        'Else
        '    lsFilter = "b.sBankName like " & strParm(fsValue & "%")
        'End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sCompnyNm»sCheckNox»sAcctNoxx»dCheckDte»nAmountxx»sTransNox" _
                                        , "Client»Check No»Account No»Date»Amount»Trans No", _
                                        , "sCompnyNm»sCheckNox»a.sAcctNoxx»a.dCheckDte»a.nAmountxx»a.sTransNox" _
                                        , IIf(fbByCode, 1, 0))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function

    Public Function LoadByCheckInfo( _
            ByVal fsAcctNoxx As String, _
            ByVal fsCheckNox As String, _
            ByVal fdCheckDte As Date) As Boolean
        Dim lsSQL As String

        lsSQL = "SELECT sTransNox" & _
               " FROM Checks_Received" & _
               " WHERE sAcctNoxx = " & strParm(fsAcctNoxx) & _
                 " AND sCheckNox = " & strParm(fsCheckNox) & _
                 " AND dCheckDte = " & dateParm(fdCheckDte) & _
                 IIf(p_nPurpose = xePurposeView, "", " AND IFNULL(cChckStat, '0') = '0'")
        Dim loDta As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            LoadByCheckInfo = NewTransaction()
        Else
            LoadByCheckInfo = OpenTransaction(loDta(0).Item("sTransNox"))
        End If
    End Function

    Public Function SaveTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String = ""

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0).Item("sReferNox") = p_oOthersx.sReferNox
                p_oDTMstr(0).Item("sSourceCD") = p_oOthersx.sSourceCD
                p_oDTMstr(0).Item("nAmountxx") = p_oOthersx.nAmountxx
                p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)

                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.getSysDate)
            Else
                lsSQL = "INSERT INTO Checks_Received_Others" & _
                       " SET sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                          ", sReferNox = " & strParm(p_oOthersx.sReferNox) & _
                          ", sSourceCD = " & strParm(p_oOthersx.sSourceCD) & _
                          ", nAmountxx = " & p_oOthersx.nAmountxx & _
                          ", sModified = " & strParm(p_oApp.UserID) & _
                          ", dModified = " & dateParm(p_oApp.getSysDate)
                p_oApp.Execute(lsSQL, "Checks_Received_Others")

                lsSQL = "UPDATE Checks_Received" & _
                       " SET nAmountxx = nAmountxx + " & p_oOthersx.nAmountxx & _
                       " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            End If

            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, p_sMasTable)
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 2  ' sClientNm
                getBank(2, 80, fsValue, False, True)
        End Select
    End Sub

    Public Function CancelTransaction(ByVal fdORIssued As Date) As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid transaction mode detected...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatCancelled Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
                MsgBox("Check transaction was already tagged as Bounced...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            Else
                MsgBox("Check transaction was already tagged as Cancelled...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "2" Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatCleared Then
                MsgBox("Check transaction was already tagged as Cleared...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        End If

        If MsgBox("Setting the check status to Cancelled will cancell the PARENT transaction!" & vbCrLf & _
                  "Do you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, p_sMsgHeadr) = vbNo Then
            MsgBox("Tagging the check as CANCELLED was terminated...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        Try
            p_oApp.BeginTransaction()
            Dim lnRow As Integer
            Dim lsSQL As String
            For lnRow = 0 To p_oDTDetl.Rows.Count - 1
                'UPDATE Transaction Table Provisionary_Receipt_Master/Others/LR_Payment_Master_PR
                lsSQL = "UPDATE " & p_oDTDetl(lnRow).Item("sTranTble") & _
                        " SET cTranStat = '3'" & _
                        " WHERE sTransNox = " & strParm(p_oDTDetl(lnRow).Item("sReferNox"))
                p_oApp.Execute(lsSQL, p_oDTDetl(lnRow).Item("sTranTble"))
            Next

            lsSQL = "UPDATE Checks_Received" & _
                   " SET cChckStat =  " & strParm(xeCheckStatCancelled) & _
                      ", dStatChng = " & dateParm(fdORIssued) & _
                      ", cTranStat = '3'" & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "Checks_Received")
            p_oDTMstr(0).Item("cChckStat") = xeCheckStatCancelled
            p_oDTMstr(0).Item("dStatChng") = fdORIssued
            p_oDTMstr(0).Item("cTranStat") = "3"

            p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Function ClearCheck(fdORIssued As Date) As Boolean
        If Not (p_nPurpose = xePurposeChange) Then
            MsgBox("Module should not be use for clearing checks...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid transaction mode detected...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatCancelled Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
                MsgBox("Check transaction was already tagged as Bounced...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            Else
                MsgBox("Check transaction was already tagged as Cancelled...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "2" Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatCleared Then
                MsgBox("Check transaction was already tagged as Cleared...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        End If

        Try
            p_oApp.BeginTransaction()

            Dim lnRow As Integer
            Dim lsSQL As String
            For lnRow = 0 To p_oDTDetl.Rows.Count - 1
                Select Case LCase(p_oDTDetl(lnRow).Item("sSourceCD"))
                    Case "prec", "arpy"
                        lsSQL = p_oDTDetl(lnRow).Item("cTranType")
                        Dim loTrans As ARPayment
                        loTrans = New ARPayment(p_oApp, lsSQL)
                        loTrans.Parent = "CheckReceived"
                        If Not loTrans.NewTransaction Then
                            MsgBox("Cannot create new AR Payment transaction...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                            p_oApp.RollBackTransaction()
                            Return False
                        End If

                        If LCase(p_oDTDetl(lnRow).Item("sSourceCD")) = "prec" Then
                            If p_oDTDetl(lnRow).Item("cTranType") <> "2" Then
                                MsgBox("Cannot clear receipt with transaction other than the MONTHLY PAYMENT...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                                p_oApp.RollBackTransaction()
                                Return False
                            End If
                        End If

                        loTrans.Master("cTranType") = p_oDTDetl(lnRow).Item("cTranType")

                        loTrans.Master("dTransact") = fdORIssued
                        loTrans.Master("cPaymForm") = "0"
                        loTrans.Master("sReferNox") = p_oDTDetl(lnRow).Item("sORNoxxxx")
                        loTrans.Master("sAcctNmbr") = p_oDTDetl(lnRow).Item("sAcctNmbr")
                        loTrans.Master("sClientID") = p_oDTDetl(lnRow).Item("sClientID")

                        loTrans.Master("sPaidByID") = ""

                        loTrans.Master("nAmountxx") = p_oDTDetl(lnRow).Item("nTranAmtx")
                        loTrans.Master("nRebatesx") = p_oDTDetl(lnRow).Item("nRebatesx")

                        loTrans.Master("sSourceCD") = "CChk"
                        loTrans.Master("sSourceNo") = p_oDTDetl(lnRow).Item("sTransNox")
                        loTrans.Master("sCollIDxx") = p_oDTDetl(lnRow).Item("sCollIDxx")

                        If loTrans.SaveTransaction Then
                            MsgBox("AR Payment transaction was save successfully...", MsgBoxStyle.OkOnly + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                        Else
                            MsgBox("Cannot save AR Payment transaction...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                            p_oApp.RollBackTransaction()
                            Return False
                        End If
                    Case "lrpy"
                        Dim loTrans As LRPayment
                        loTrans = New LRPayment(p_oApp)
                        loTrans.Parent = "CheckReceived"

                        If Not loTrans.NewTransaction Then
                            MsgBox("Cannot create new LR Payment transaction...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                            p_oApp.RollBackTransaction()
                            Return False
                        End If

                        loTrans.Master("cTranType") = p_oDTDetl(lnRow).Item("cTranType")
                        loTrans.Master("dTransact") = fdORIssued
                        loTrans.Master("cPaymForm") = "0"
                        loTrans.Master("sReferNox") = p_oDTDetl(lnRow).Item("sORNoxxxx")
                        loTrans.Master("sAcctNmbr") = p_oDTDetl(lnRow).Item("sAcctNmbr")
                        loTrans.Master("sClientID") = p_oDTDetl(lnRow).Item("sClientID")
                        'loTrans.Master("sPaidByID") = ""
                        loTrans.Master(91) = p_oDTDetl(lnRow).Item("nTranAmtx") + p_oDTDetl(lnRow).Item("nIntAmtxx")
                        loTrans.Master("nPenaltyx") = p_oDTDetl(lnRow).Item("nPenaltyx")
                        loTrans.Master("sSourceCD") = "CChk"
                        loTrans.Master("sSourceNo") = p_oDTDetl(lnRow).Item("sTransNox")
                        loTrans.Master("sCollIDxx") = p_oDTDetl(lnRow).Item("sCollIDxx")

                        If loTrans.SaveTransaction Then
                            MsgBox("LR Payment transaction was save successfully...", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                        Else
                            MsgBox("Cannot save LR Payment transaction...", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                            p_oApp.RollBackTransaction()
                            Return False
                        End If
                End Select

                'UPDATE Transaction Table Provisionary_Receipt_Master/Others/LR_Payment_Master_PR
                If p_oDTDetl(lnRow).Item("sTranTble") = "Provisionary_Receipt_Master" Then
                    lsSQL = "UPDATE " & p_oDTDetl(lnRow).Item("sTranTble") & _
                            " SET cTranStat = '2'" & _
                            " WHERE sTransNox = " & strParm(p_oDTDetl(lnRow).Item("sReferNox"))
                Else
                    lsSQL = "UPDATE " & p_oDTDetl(lnRow).Item("sTranTble") & _
                            " SET cPostedxx = '2'" & _
                               ", dPostedxx = " & dateParm(p_oApp.getSysDate) & _
                            " WHERE sTransNox = " & strParm(p_oDTDetl(lnRow).Item("sReferNox"))
                End If
                p_oApp.Execute(lsSQL, p_oDTDetl(lnRow).Item("sTranTble"))
            Next

            If p_oDTMstr(0).Item("cDepositd") = "0" Then
                lsSQL = "UPDATE Checks_Received" & _
                        " SET cDepositd = '1'" & _
                           ", dDepositd = " & dateParm(fdORIssued) & _
                        " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                p_oApp.Execute(lsSQL, "Checks_Received")

                p_oDTMstr(0).Item("cDepositd") = "1"
                p_oDTMstr(0).Item("dDepositd") = fdORIssued
            End If

            lsSQL = "UPDATE Checks_Received" & _
                   " SET cChckStat =  " & strParm(xeCheckStatCleared) & _
                      ", dStatChng = " & dateParm(fdORIssued) & _
                      ", cTranStat = '2'" & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "Checks_Received")

            p_oDTMstr(0).Item("cChckStat") = xeCheckStatCleared
            p_oDTMstr(0).Item("dStatChng") = fdORIssued
            p_oDTMstr(0).Item("cTranStat") = "2"

            p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Function BounceCheck(fdORIssued As Date) As Boolean
        If Not (p_nPurpose = xePurposeChange) Then
            MsgBox("Module should not be use for clearing checks...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid transaction mode detected...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatCancelled Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
                MsgBox("Check transaction was already tagged as Bounced...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            Else
                MsgBox("Check transaction was already tagged as Cancelled...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "2" Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatCleared Then
                MsgBox("Check transaction was already tagged as Cleared...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        End If

        If MsgBox("Setting the check status to Bounce will cancell the PARENT transaction!" & vbCrLf & _
                  "Do you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, p_sMsgHeadr) = vbNo Then
            MsgBox("Tagging the check as bounce was cancelled...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        Try
            p_oApp.BeginTransaction()
            Dim lnRow As Integer
            Dim lsSQL As String
            For lnRow = 0 To p_oDTDetl.Rows.Count - 1
                'UPDATE Transaction Table Provisionary_Receipt_Master/Others/LR_Payment_Master_PR
                lsSQL = "UPDATE " & p_oDTDetl(lnRow).Item("sTranTble") & _
                        " SET cTranStat = '3'" & _
                        " WHERE sTransNox = " & strParm(p_oDTDetl(lnRow).Item("sReferNox"))
                p_oApp.Execute(lsSQL, p_oDTDetl(lnRow).Item("sTranTble"))
            Next

            If p_oDTMstr(0).Item("cDepositd") = "0" Then
                lsSQL = "UPDATE Checks_Received" & _
                        " SET cDepositd = '1'" & _
                           ", dDepositd = " & dateParm(fdORIssued) & _
                        " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
                p_oApp.Execute(lsSQL, "Checks_Received")

                p_oDTMstr(0).Item("cDepositd") = "1"
                p_oDTMstr(0).Item("dDepositd") = fdORIssued
            End If

            lsSQL = "UPDATE Checks_Received" & _
                   " SET cChckStat =  " & strParm(xeCheckStatBounce) & _
                      ", dStatChng = " & dateParm(fdORIssued) & _
                      ", cTranStat = '3'" & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "Checks_Received")
            p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce
            p_oDTMstr(0).Item("dStatChng") = fdORIssued
            p_oDTMstr(0).Item("cTranStat") = "3"

            p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Function HoldCheck(fdORIssued As Date) As Boolean
        If Not (p_nPurpose = xePurposeChange) Then
            MsgBox("Module should not be use for clearing checks...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If Not (p_nEditMode = xeEditMode.MODE_READY) Then
            MsgBox("Invalid transaction mode detected...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatCancelled Or p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatBounce Then
                MsgBox("Check transaction was already tagged as Bounced...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            Else
                MsgBox("Check transaction was already tagged as Cancelled...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "2" Then
            If p_oDTMstr(0).Item("cChckStat") = xeCheckStatCleared Then
                MsgBox("Check transaction was already tagged as Cleared...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        End If

        Try
            p_oApp.BeginTransaction()

            Dim lsSQL As String
            lsSQL = "UPDATE Checks_Received" & _
                   " SET cChckStat =  " & strParm(xeCheckStatHold) & _
                      ", dStatChng = " & dateParm(fdORIssued) & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "Checks_Received")
            p_oDTMstr(0).Item("cChckStat") = xeCheckStatHold
            p_oDTMstr(0).Item("dStatChng") = fdORIssued

            p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Public Function CancelCheck(ByVal fsReferNox As String, ByVal fsSourceCD As String) As Boolean
        Dim lsSQL As String
        'Make sure to use CHECK CLEARING in posting a PR transaction using a check...
        '+++++++++++++++
        'Check if the PR is a Check transaction
        lsSQL = "SELECT 'Checks_Received' sTableNme, sTransNox, sReferNox, sSourceCd, nAmountxx" & _
               " FROM Checks_Received" & _
               " WHERE sReferNox = " & strParm(fsReferNox) & _
                 " AND sSourceCd = " & strParm(fsSourceCD) & _
               " UNION" & _
               " SELECT 'Checks_Received_Others' sTableNme, sTransNox, sReferNox, sSourceCd, nAmountxx" & _
               " FROM Checks_Received_Others" & _
               " WHERE sReferNox = " & strParm(fsReferNox) & _
                 " AND sSourceCd = " & strParm(fsSourceCD)
        Dim loDta As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then Return False

        'Was it saved in Check_Payments_Others
        If loDta(0).Item("sTableNme") = "Checks_Received_Others" Then
            'Delete Record from Check_Payments_Others
            lsSQL = "DELETE FROM Checks_Received_Others" & _
                   " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox")) & _
                     " AND sReferNox = " & strParm(loDta(0).Item("sReferNox")) & _
                     " AND sSourceCD = " & strParm(loDta(0).Item("sSourceCD"))
            p_oApp.Execute(lsSQL, "Checks_Received_Others")

            'Deduct amount from Check_Payments
            lsSQL = "UPDATE Checks_Received " & _
                   " SET nAmountxx = nAmountxx - " & loDta(0).Item("nAmountxx") & _
                   " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "Checks_Received")
        Else
            'Is the transaction amount the same with that of Check_Payments
            If p_oDTMstr(0).Item("nAmountxx") = loDta(0).Item("nAmountxx") Then
                'Cancel Check payments - assume 1 check = 1 PR
                lsSQL = "UPDATE Checks_Received" & _
                       " SET cTranStat = '3'" & _
                          ", cChckStat = " & strParm(xeCheckStatCancelled) & _
                       " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox"))
                p_oApp.Execute(lsSQL, "Checks_Received")
            Else
                'Search another check record using the same Check No
                lsSQL = "SELECT sReferNox, sSourceCD" & _
                       " FROM Checks_Received_Others" & _
                       " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox"))
                Dim loDtx As DataTable = p_oApp.ExecuteQuery(lsSQL)

                'Delete the check record found
                lsSQL = "DELETE FROM Checks_Received_Others" & _
                       " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox")) & _
                         " AND sReferNox = " & strParm(loDtx(0).Item("sReferNox")) & _
                         " AND sSourceCD = " & strParm(loDtx(0).Item("sSourceCD"))
                p_oApp.Execute(lsSQL, "Checks_Received_Others")

                'Transfer the reference of the deleted record to the main check record 
                lsSQL = "UPDATE Checks_Received " & _
                       " SET nAmountxx = nAmountxx - " & loDta(0).Item("nAmountxx") & _
                          ", sReferNox = " & strParm(loDtx(0).Item("sReferNox")) & _
                          ", sSourceCD = " & strParm(loDtx(0).Item("sSourceCD")) & _
                       " WHERE sTransNox = " & strParm(loDta(0).Item("sTransNox"))
                p_oApp.Execute(lsSQL, "Checks_Received")
            End If
        End If

        Return True
    End Function

    Private Function isEntryOk() As Boolean
        If Trim(p_oDTMstr(0).Item("sBankIDxx")) = "" Then
            MsgBox("Bank Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If Trim(p_oDTMstr(0).Item("sAcctNoxx")) = "" Then
            MsgBox("Account No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If Trim(p_oDTMstr(0).Item("sCheckNox")) = "" Then
            MsgBox("Check No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If Not IsDate(p_oDTMstr(0).Item("dCheckDte")) Then
            MsgBox("Check Date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oOthersx.nAmountxx <= 0 Then
            MsgBox("Check Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                Case "dcheckdte"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dmodified", "smodified", "ddepositd", "dstatchng"
                Case "cchckstat", "cdepositd", "ctranstat"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "namountxx"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case "nclearing"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sBankName = ""
        p_oOthersx.sReferNox = ""
        p_oOthersx.sSourceCD = ""
        p_oOthersx.nAmountxx = 0.0
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getBank(ByVal fnColIdx As Integer _
                           , ByVal fnColDsc As Integer _
                           , ByVal fsValue As String _
                           , ByVal fbIsCode As Boolean _
                           , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sBankName <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sBankName And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBankIDxx" & _
                       ", a.sBankName" & _
              " FROM Banks a"
        IIf(p_nEditMode = xeEditMode.MODE_ADDNEW, " AND a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBankIDxx»sBankName" _
                                             , "ID»Bank", _
                                             , "a.sBankIDxx»a.sBankName" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBankName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sBankIDxx")
                p_oOthersx.sBankName = loRow.Item("sBankName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBankName)
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sBankIDxx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sBankName = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sBankName = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sBankIDxx")
            p_oOthersx.sBankName = loDta(0).Item("sBankName")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBankName)
    End Sub

    Public Sub SearchBranch(ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)

        If Not p_oApp.ProductID = "LRTrackr" Then Exit Sub

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_sBranchCd And fsValue <> "" Then Exit Sub
        Else
            If fsValue = p_sBranchNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBranchCD" & _
                       ", a.sBranchNm" & _
               " FROM Branch a" & _
               IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBranchCD»sBranchNm" _
                                             , "ID»Company", _
                                             , "a.sBranchCD»a.sBranchNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_sBranchCd = ""
                p_sBranchNm = ""
            Else
                p_sBranchCd = loRow.Item("sBranchCD")
                p_sBranchNm = loRow.Item("sBranchNm")
            End If
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sBranchCD = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sBranchNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_sBranchCd = ""
            p_sBranchNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_sBranchCd = loDta(0).Item("sBranchCD")
            p_sBranchNm = loDta(0).Item("sBranchNm")
        End If
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sCheckNox" & _
                    ", a.sBankIDxx" & _
                    ", a.sAcctNoxx" & _
                    ", a.dCheckDte" & _
                    ", a.nAmountxx" & _
                    ", a.nClearing" & _
                    ", IFNULL(a.cChckStat, '0') cChckStat" & _
                    ", IFNULL(a.cDepositd, '0') cDepositd" & _
                    ", a.dDepositd" & _
                    ", a.dStatChng" & _
                    ", a.sReferNox" & _
                    ", a.sSourceCD" & _
                    ", a.sRecTrans" & _
                    ", IFNULL(a.cTranStat, '0') cTranStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
              " FROM Checks_Received a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT c.sCompnyNm, a.sCheckNox, a.sAcctNoxx, a.dCheckDte, a.nAmountxx, a.sTransNox" & _
              " FROM Checks_Received a" & _
                  ", LR_Payment_Master_PR b" & _
                        " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
              " WHERE a.sTransNox LIKE '<sBranchCD>%'" & _
                " AND a.sReferNox = b.sTransNox" & _
                " AND a.sSourceCd IN ('ARPy', 'LRPy')" & _
                " AND IFNULL(a.cChckStat, '0') IN (<cChckStat>)" & _
                " <sAddCondx>" & _
              " UNION" & _
              " SELECT c.sCompnyNm, a.sCheckNox, a.sAcctNoxx, a.dCheckDte, a.nAmountxx, a.sTransNox" & _
              " FROM Checks_Received a" & _
                  ", Provisionary_Receipt_Master b" & _
                        " LEFT JOIN Client_Master c ON b.sClientID = c.sClientID" & _
              " WHERE a.sTransNox LIKE '<sBranchCD>%'" & _
                " AND a.sReferNox = b.sTransNox" & _
                " AND a.sSourceCd IN ('PRec')" & _
                " AND a.dCheckDte >= '2017-07-01'" & _
                " AND b.cTranType IN ('1')" & _
                " AND IFNULL(a.cChckStat, '0') IN (<cChckStat>)" & _
                " <sAddCondx>"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCd = p_oApp.BranchCode
        p_sBranchNm = p_oApp.BranchName

        p_nPurpose = xePurposeCreate
        p_nTranStat = -1
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnPurpose As Int32, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nPurpose = fnPurpose
        p_nTranStat = fnStatus
    End Sub

    Private Class Others
        Public sBankName As String
        Public sReferNox As String
        Public sSourceCD As String
        Public nAmountxx As Decimal
    End Class
End Class

