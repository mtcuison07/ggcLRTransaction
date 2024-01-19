'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Adjustment Object
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
'  Kalyptus [ 07/09/2016 01:15 pm ]
'      Started creating this object.
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class LRAdjustment
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String

    Private Const p_sMasTable As String = "LR_Adjustment_Master"
    Private Const p_sMsgHeadr As String = "LR Adjustment"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sClientNm
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case 82 ' nABalance 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case 83 ' nInterest 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nInterest
                    Case 84 ' nAcctTerm
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case 85 ' nMonAmort 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case 86 ' nAmtDuexx
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case 87 ' sCompnyNm 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sClientNm
                        getAccount(3, 80, value, False, False)
                    Case 81 To 87
                    Case 1
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 6, 7
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set

    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "sclientnm" ' 80 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" ' 81 
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case "nabalance" ' 82
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case "ninterest" ' 83
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nInterest
                    Case "nacctterm" ' 84
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case "nmonamort" ' 85
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case "namtduexx" ' 86
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case "scompnynm" '87
                        If Trim(IFNull(p_oDTMstr(0).Item(3))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(3, 80, p_oDTMstr(0).Item(3), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "sclientnm"
                        getAccount(3, 80, value, False, False)
                    Case "saddressx", "nabalance", "ninterest", "nacctterm", "nmonamort", "namtduexx", "scompnynm"
                    Case "dtransact", "ndebitamt", "ncredtamt"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
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
    Public Function NewTransaction() As Boolean
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

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    'Public Function SearchWithCondition(String)
    Public Function SearchWithCondition(ByVal fsFilter As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Browse, fsFilter)
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        ElseIf p_oDTMstr.Rows.Count = 1 Then
            Return OpenTransaction(p_oDTMstr(0).Item("sTransNox"))
        Else
            'KwikBrowse here!
            Return True
        End If
    End Function

    'Public Function SearchTransaction(String, Boolean, Boolean=False)
    Public Function SearchTransaction( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oOthersx.sClientNm Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQ_Browse, "a.cPostedxx IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = getSQ_Browse()
        End If

        If p_sBranchCd <> "" Then
            lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
        End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "b.sCompnyNm like " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , "Trans No»Client»Date", _
                                        , "a.sTransNox»b.sCompnyNm»a.dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function

    'Public Function SaveTransaction
    'This object does not implement Update
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

            'Save master table 
            'Note: Update is not allowed!!!
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)
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

    'Public Function CancelTransaction
    Public Function CancelTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = "2" Then
            MsgBox("Request was already posted!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cPostedxx") = "3"
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))
            p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    'Public Function PostTransaction()
    Public Function PostTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = "2" Then
            MsgBox("Application was already posted!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = "3" Then
            MsgBox("Application was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim loTrans As LRTrans
            loTrans = New LRTrans(p_oApp)
            loTrans.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")
            loTrans.Transact_Date = p_oDTMstr(0).Item("dTransact")
            loTrans.Amount = p_oDTMstr(0).Item("nDebitAmt") + p_oDTMstr(0).Item("nCredtAmt")
            loTrans.SourceNo = p_oDTMstr(0).Item("sTransNox")
            loTrans.Remarks = p_oDTMstr(0).Item("sRemarksx")
            loTrans.isOffice = True

            'Check Transaction Type
            If p_oDTMstr(0).Item("nCredtAmt") > 0 Then
                'Credit?
                If Not loTrans.Credit Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            Else
                'Debit?
                If Not loTrans.Debit Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If

            p_oDTMstr(0).Item("cPostedxx") = "2"
            p_oDTMstr(0).Item("dPostedxx") = p_oApp.getSysDate

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
            p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

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
            Case 80 ' sClientNm
                getAccount(3, 80, fsValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dmodified", "smodified", "dpostedxx"
                Case "cpostedxx"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "ndebitamt", "ncredtamt"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
        p_oOthersx.nABalance = 0.0
        p_oOthersx.nAcctTerm = 0
        p_oOthersx.nAmtDuexx = 0.0
        p_oOthersx.nInterest = 0.0
        p_oOthersx.nMonAmort = 0.0
        p_oOthersx.sCompnyNm = ""
    End Sub

    Private Function isEntryOk() As Boolean
        'Check validity of transaction date
        If p_oDTMstr(0).Item("dTransact") <= "2016-01-01" And p_oDTMstr(0).Item("dTransact") > p_oApp.SysDate Then
            MsgBox("Transaction date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check if application has client
        If p_oDTMstr(0).Item("sAcctNmbr") = "" Then
            MsgBox("Account Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Val(p_oDTMstr(0).Item("nDebitAmt")) + Val(p_oDTMstr(0).Item("nCredtAmt")) <= 0 Then
            MsgBox("Transaction Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Trim(p_oDTMstr(0).Item("sReferNox")) = "" Then
            MsgBox("Document/Reference No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = "2" Then
            MsgBox("Application was posted! Posted application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = "3" Then
            MsgBox("Application was cancelled! Cancelled application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getAccount(ByVal fnColIdx As Integer _
                         , ByVal fnColDsc As Integer _
                         , ByVal fsValue As String _
                         , ByVal fbIsCode As Boolean _
                         , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sClientNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sClientNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sAcctNmbr" & _
                       ", b.sCompnyNm sClientNm" & _
                       ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                       ", a.nABalance" & _
                       ", IF(a.nIntTotal = a.nInterest, 0, a.nInterest/nAcctTerm) nInterest" & _
                       ", a.nAcctTerm" & _
                       ", a.nMonAmort" & _
                       ", a.nAmtDuexx" & _
                       ", e.sCompnyNm" & _
                       ", a.sCompnyID" & _
                       ", a.sClientID" & _
               " FROM LR_Master a" & _
                " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" & _
                " LEFT JOIN TownCity c ON b.sTownIDxx = c.sTownIDxx" & _
                " LEFT JOIN Province d ON c.sProvIDxx = d.sProvIDxx" & _
                " LEFT JOIN Company e ON a.sCompnyID = e.sCompnyID"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sAcctNmbr»sClientNm»nABalance»sCompnyNm" _
                                             , "Account No»Client»Balance»Company", _
                                             , "a.sAcctNmbr»b.sCompnyNm»a.nABalance»e.sCompnyNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oDTMstr(0).Item("sClientID") = ""
                Call InitOthers()
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sAcctNmbr")
                p_oDTMstr(0).Item("sClientID") = loRow.Item("sClientID")
                p_oOthersx.sClientNm = loRow.Item("sClientNm")
                p_oOthersx.sAddressx = loRow.Item("xAddressx")
                p_oOthersx.nABalance = loRow.Item("nABalance")
                p_oOthersx.nAcctTerm = loRow.Item("nAcctTerm")
                p_oOthersx.nAmtDuexx = loRow.Item("nAmtDuexx")
                p_oOthersx.nInterest = loRow.Item("nInterest")
                p_oOthersx.nMonAmort = loRow.Item("nMonAmort")
                p_oOthersx.sCompnyNm = loRow.Item("sCompnyNm")

                Dim loLR As New LRTrans(p_oApp)
                loLR.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")
                Dim loLRMstr = loLR.GetMaster()
                p_oOthersx.nAmtDuexx = loLR.getDelay(loLRMstr, p_oDTMstr(0).Item("dTransact")) * p_oOthersx.nMonAmort

            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
            Exit Sub

        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sAcctNmbr = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sClientNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oDTMstr(0).Item("sClientID") = ""
            Call InitOthers()
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sAcctNmbr")
            p_oDTMstr(0).Item("sClientID") = loDta(0).Item("sClientID")
            p_oOthersx.sClientNm = loDta(0).Item("sClientNm")
            p_oOthersx.sAddressx = loDta(0).Item("xAddressx")
            p_oOthersx.nABalance = loDta(0).Item("nABalance")
            p_oOthersx.nAcctTerm = loDta(0).Item("nAcctTerm")
            p_oOthersx.nAmtDuexx = loDta(0).Item("nAmtDuexx")
            p_oOthersx.nInterest = loDta(0).Item("nInterest")
            p_oOthersx.nMonAmort = loDta(0).Item("nMonAmort")
            p_oOthersx.sCompnyNm = loDta(0).Item("sCompnyNm")

            Dim loLR As New LRTrans(p_oApp)
            loLR.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")

            Dim loLRMstr = loLR.GetMaster()
            p_oOthersx.nAmtDuexx = loLR.getDelay(loLRMstr, p_oDTMstr(0).Item("dTransact")) * p_oOthersx.nMonAmort

        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
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
                    ", a.dTransact" & _
                    ", a.sReferNox" & _
                    ", a.sAcctNmbr" & _
                    ", a.sClientID" & _
                    ", a.sRemarksx" & _
                    ", a.nDebitAmt" & _
                    ", a.nCredtAmt" & _
                    ", a.sApproved" & _
                    ", a.sAPprCode" & _
                    ", a.cPostedxx" & _
                    ", a.dPostedxx" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sReferNox" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
              " WHERE a.sClientID = b.sClientID"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCd = p_oApp.BranchCode
        p_sBranchNm = p_oApp.BranchName

        p_nTranStat = -1

    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sAddressx As String
        Public nABalance As Decimal
        Public nInterest As Decimal
        Public nAcctTerm As Integer
        Public nMonAmort As Decimal
        Public nAmtDuexx As Decimal
        Public sCompnyNm As String
    End Class
End Class
