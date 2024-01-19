'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Application Object
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
'  Kalyptus [ 06/02/2016 09:32 am ]
'      Started creating this object.
'  
'  Note: How will I present a void transaction?
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient

Public Class LRApplication
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String
    Private p_bValidCode As Boolean

    Private p_oClient As ggcClient.Client

    Private Const p_sMasTable As String = "LR_Application"
    Private Const p_sMsgHeadr As String = "LR Application"

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
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sAddressx) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
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
                        getClient(2, 80, value, False, False)
                    Case 81 ' sAddressx
                    Case 5, 6 ' nLoanAmtx, nIntRatex
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 1, 7 ' dTransact, dExpRelse
                        If IsDate(value) Then
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
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" ' 81 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
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
                        getClient(2, 80, value, False, False)
                    Case "saddressx"
                    Case "nloanamtx", "nintratex" ' 5, 6 
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "dtransact", "dexprelse" ' 1, 7
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

        p_bValidCode = False
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
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat IN (" & strDissect(p_nTranStat) & ")")
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

        If Trim(p_oDTMstr(0).Item("sApprovCd")) <> "" Then p_bValidCode = isValidCode()

        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String = ""

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            If Not p_oClient.SaveClient Then
                MsgBox("Unable to save client info!", vbOKOnly, p_sMsgHeadr)
                If p_sParent = "" Then p_oApp.RollBackTransaction()
                Return False
            End If

            p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")

            'Save master table 
            'Note: Update is not allowed!!!
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                If Trim(p_oDTMstr(0).Item("sApprovCd")) <> "" Then
                    If p_bValidCode Then
                        lsSQL = "SELECT sTransNox" & _
                               " FROM LR_Pre_Approve" & _
                               " WHERE sClientID = " & p_oDTMstr(0).Item("sClientID") & _
                               " ORDER BY dTransact DESC LIMIT 1"
                        Dim loDt As DataTable
                        loDt = p_oApp.ExecuteQuery(lsSQL)

                        If loDt.Rows.Count = 0 Then
                            MsgBox("Unable to locate the client's PRE Approved Record!", vbOKOnly, p_sMsgHeadr)
                            If p_sParent = "" Then p_oApp.RollBackTransaction()
                            Return False
                        End If

                        p_oDTMstr(0).Item("sSourceNo") = loDt(0).Item("sTransNox")
                        p_oDTMstr(0).Item("sSourceCD") = "PALR"
                        p_oDTMstr(0).Item("cTranStat") = "1"
                    Else
                        p_oDTMstr(0).Item("sApprovCd") = ""
                        p_oDTMstr(0).Item("sSourceNo") = ""
                        p_oDTMstr(0).Item("sSourceCD") = ""
                        p_oDTMstr(0).Item("cTranStat") = "0"

                        RaiseEvent MasterRetrieved(9, p_oDTMstr(0).Item("sApprovCd"))

                        If MsgBox("Approval code is not valid for this transaction..." & vbCrLf & _
                                  "Do you want to continue?", vbOKCancel, p_sMsgHeadr) = vbCancel Then
                            If p_sParent = "" Then p_oApp.RollBackTransaction()
                            Return False
                        End If
                    End If
                End If

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

        '1 = pre-approved
        '2 = approved
        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            If MsgBox("Request was already approved! Do you continue?", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr) = MsgBoxResult.Cancel Then
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cTranStat") = "3"
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))
            p_oApp.Execute(lsSQL, p_sMasTable)

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

        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Application was already APPROVED!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "1" Then
            MsgBox("Application was PRE-APPROVED!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Application was already CANCELLED!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'Please check release date here...
        If p_oDTMstr(0).Item("dExpRelse") < p_oDTMstr(0).Item("dTransact") Then
            MsgBox("Expected release date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cTranStat") = "2"

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
            p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

            'kalyptus - 2020.01.20 03:00pm
            'Make sure that a payee is created for this approved loan...
            Call createPayee()

            Return True
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()

            MsgBox(ex.Message)

            Return False
        End Try

    End Function

    'kalyptus - 2020.01.20 02:56pm
    'Create a payee record for approved AR/Cash applications...
    Private Sub createPayee()
        Dim lsSQL As String
        Dim loDTPayee As DataTable

        lsSQL = "SELECT * FROM Payee WHERE sClientID = " & strParm(p_oDTMstr(0).Item("sClientID"))

        loDTPayee = p_oApp.ExecuteQuery(lsSQL)

        If loDTPayee.Rows.Count > 0 Then
            loDTPayee.Rows.Add(loDTPayee.NewRow())


            loDTPayee(0).Item("sPayeeIDx") = GetNextCode("Payee", "sPayeeIDx", True, p_oApp.Connection, True, p_sBranchCd)
            loDTPayee(0).Item("sPayeeNme") = p_oClient.Master("sFrstName") & _
                                          IIf(Len(p_oClient.Master("sMiddName")) <= 1, "", " " & Left(p_oClient.Master("sMiddName"), 1) & ".") & " " & _
                                          p_oClient.Master("sLastName") & _
                                          IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm"))
            loDTPayee(0).Item("sPrtclrID") = "M001001060"
            loDTPayee(0).Item("sClientID") = p_oClient.Master("sClientID")

            lsSQL = "INSERT INTO Payee" & _
                   " SET sPayeeIDx = " & strParm(loDTPayee(0).Item("sPayeeIDx")) & _
                      ", sPayeeNme = " & strParm(loDTPayee(0).Item("sPayeeNme")) & _
                      ", sPrtclrID = " & strParm(loDTPayee(0).Item("sPrtclrID")) & _
                      ", sClientID = " & strParm(loDTPayee(0).Item("sClientID")) & _
                      ", cRecdStat = '1'" & _
                      ", sModified = " & strParm(p_oApp.UserID) & _
                      ", dModified = " & dateParm(p_oApp.getSysDate)

            p_oApp.Execute(lsSQL, "Payee", Left(p_oDTMstr(0).Item("sTransNox"), 4))

        End If
    End Sub

    'Public Function PostTransaction()
    Public Function ReleaseLoan() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "0" Then
            MsgBox("Application was not yet approved!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Loan was already released!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Application was cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        'Create the LR Master here...
        Dim loLRMstr As LRMaster
        loLRMstr = New LRMaster(p_oApp)
        loLRMstr.Parent = "LRApplication"
        loLRMstr.Branch = p_sBranchCd
        If Not loLRMstr.NewTransaction() Then
            MsgBox("Can not create a Loan Master record for this Application!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If
        loLRMstr.Master("sClientID") = p_oDTMstr(0).Item("sClientID")
        'Use expected released date as possible release date
        loLRMstr.Master("dTransact") = p_oDTMstr(0).Item("dExpRelse")
        loLRMstr.Master("nAcctTerm") = 1
        loLRMstr.Master("sApplicNo") = p_oDTMstr(0).Item("sTransNox")
        loLRMstr.Master("nIntRatex") = p_oDTMstr(0).Item("nIntRatex")
        loLRMstr.Master("nPrincipl") = p_oDTMstr(0).Item("nLoanAmtx")

        'Get the MC Account of this client if transaction is pre-approved from list of existing accounts
        If p_oDTMstr(0).Item("sSourceCd") = "PALR" Then
            lsSQL = "SELECT sAcctNmbr" & _
                   " FROM LR_Pre_Approve" & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sSourceNo"))
            Dim loDT As DataTable = p_oApp.ExecuteQuery(lsSQL)

            If loDT.Rows.Count = 1 Then
                loLRMstr.Master("sMCActNox") = loDT(0).Item("sAcctNmbr")
            Else
                loLRMstr.Master("sMCActNox") = ""
            End If
        End If

        'Show form to get the LR Loan information
        Dim loFrm As frmLRMaster
        loFrm = New frmLRMaster
        loFrm.LoanObject = loLRMstr
        loFrm.ShowDialog()
        If loFrm.Cancelled Then Return False

        Try
            'Begin the actual saving of the ReleaseLoan Module
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cTranStat") = "4"

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
            p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

            'Save the LR Loan Record here...
            If Not loLRMstr.SaveTransaction Then
                MsgBox("Can not save the LR Loan record!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
                If p_sParent = "" Then p_oApp.RollBackTransaction()
                Return False
            End If

            lsSQL = "INSERT INTO LR_Master_Release" & _
                   " SET sAcctNmbr = " & strParm(loLRMstr.Master("sAcctNmbr")) & _
                      ", sReleasBy = " & strParm(p_oApp.UserID) & _
                      ", sReferNox = " & strParm(loFrm.ReferNo) & _
                      ", nAmountxx = " & loLRMstr.Master("nTakeHome") & _
                      ", sBnkActID = " & strParm(loFrm.BankID) & _
                      ", sCheckNox = " & strParm(loFrm.CheckNo)

            If IsDate(loFrm.CheckDate) Then
                lsSQL = lsSQL & _
                    ", dCheckDte = " & dateParm(loFrm.CheckDate)
            End If

            p_oApp.Execute(lsSQL, "LR_Master_Release", Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

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
                getClient(2, 80, fsValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                Case "dtransact", "dexprelse"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dmodified", "smodified"
                Case "cpostedxx", "ctranstat"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "nloanamtx"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case "nintratex"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
    End Sub

    Private Function isEntryOk() As Boolean

        'Check validity of transaction date
        If p_oDTMstr(0).Item("dTransact") <= "2016-01-01" And p_oDTMstr(0).Item("dTransact") > p_oApp.SysDate Then
            MsgBox("Transaction release date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check if application has client
        If p_oDTMstr(0).Item("sClientID") = "" Then
            MsgBox("Client Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Val(p_oDTMstr(0).Item("nLoanAmtx")) <= 1000 Then
            MsgBox("Loan Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check when will be the exptected released of this loan
        If p_bValidCode Then
            If p_oDTMstr(0).Item("dExpRelse") < p_oDTMstr(0).Item("dTransact") Then
                MsgBox("Expected release date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If
        End If

        If p_oDTMstr(0).Item("cTranStat") = "2" Then
            MsgBox("Application was approved! Approved application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Application was cancelled! Cancelled application are no longer allowed to update!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    Private Function isValidCode() As Boolean
        Dim oApproval As CodeApproval
        oApproval = New CodeApproval
        oApproval.XSystem = CodeApproval.pxePreApproved
        oApproval.DateRequested = p_oDTMstr(0).Item("dTransact")
        oApproval.MiscInfo = p_oOthersx.sClientNm
        oApproval.IssuedBy = "6"        '6 is the Issuee Code of Telemarketing....

        If Not oApproval.Encode() Then
            Return False
        End If

        If Not oApproval.Equalx(p_oDTMstr(0).Item("sApprovCd"), oApproval.Result) Then
            Return False
        Else
            Return True
        End If
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sClientNm <> "" Then Exit Sub
        Else
            'Do not allow searching of value if fsValue is empty
            If (fsValue = p_oOthersx.sClientNm And fsValue <> "") Or fsValue = "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oApp)
        loClient.Parent = "LRApplication"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oClient = loClient
                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oOthersx.sClientNm = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")

                p_oOthersx.sAddressx = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            Else
                p_oDTMstr(0).Item("sClientID") = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sAddressx = ""
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                p_oClient = loClient
                p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")
                p_oOthersx.sClientNm = p_oClient.Master("sLastName") & ", " & _
                                       p_oClient.Master("sFrstName") & _
                                       IIf(p_oClient.Master("sSuffixNm") = "", "", " " & p_oClient.Master("sSuffixNm")) & " " & _
                                       p_oClient.Master("sMiddName")
                p_oOthersx.sAddressx = IIf(p_oClient.Master("sHouseNox") = "", "", p_oClient.Master("sHouseNox") & " ") & _
                                           p_oClient.Master("sAddressx") & ", " & _
                                           p_oClient.Master("sTownName")
            End If
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
                    ", a.sClientID" & _
                    ", a.sRemarks1" & _
                    ", a.sRemarks2" & _
                    ", a.nLoanAmtx" & _
                    ", a.nIntRatex" & _
                    ", a.dExpRelse" & _
                    ", a.sApproved" & _
                    ", a.sApprovCd" & _
                    ", a.sSourceCd" & _
                    ", a.sSourceNo" & _
                    ", a.cPostedxx" & _
                    ", a.sPostedxx" & _
                    ", a.cTranStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
              " WHERE a.sClientID = b.sClientID"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_oClient = New Client(foRider)
        p_oClient.Parent = "LRApplication"
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
    End Class
End Class
