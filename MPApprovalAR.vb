'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Credit Online Application Object
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
'  Kalyptus [ 06/04/2016 11:25 am ]
'      Started creating this object.
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class MPApprovalAR
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String
    Private p_sBranchCD As String

    Private p_oClient As ggcClient.Client

    Private Const p_sMasTable As String = "Credit_Online_Application"
    Private Const p_sMsgHeadr As String = "Credit Online Application"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Branch As String
        Get
            Return p_sBranchCD
        End Get
        Set(ByVal value As String)
            'If Product ID is LR then do allow changing of Branch
            If p_oApp.ProductID = "LRTrackr" Then
                p_sBranchCD = value
            End If
        End Set
    End Property

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
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
                    Case 1
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = Convert.ToDateTime(value)
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))

                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))

                End Select
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
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
                If fsValue = p_oDTMstr(0).Item("sClientNm") Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = getSQ_Browse()
        End If
        Debug.Print(lsSQL)

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "a.sClientNm LIKE " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , "Transact No»Client Name»Trans Date", _
                                        , "a.sTransNox»a.sClientNm»a.dTransact" _
                                        , IIf(fbByCode, 0, 1))
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

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            If Not p_oClient.SaveClient Then
                MsgBox("Unable to save client info!", vbOKOnly, p_sMsgHeadr)
                If p_sParent = "" Then p_oApp.RollBackTransaction()
                Return False
            End If

            p_oDTMstr(0).Item("sClientID") = p_oClient.Master("sClientID")

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                'Save master table 
                p_oDTMstr(0).Item("sAcctNmbr") = GetNextCode(p_sMasTable, "sAcctNmbr", True, p_oApp.Connection, True, "M" & Mid(p_sBranchCD, 2))
                p_oDTMstr(0).Item("sBranchCD") = p_sBranchCD
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)
                p_oApp.Execute(lsSQL, p_sMasTable)
            Else
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sAcctNmbr = " & strParm(p_oDTMstr(0).Item("sAcctNmbr")), p_oApp.UserID, Format(p_oApp.SysDate, "yyyy-MM-dd"), "")
                If lsSQL <> "" Then
                    p_oApp.Execute(lsSQL, p_sMasTable)
                End If
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
            Case 80 ' sClientNm
                getClient(2, 80, fsValue, False, True)
            Case 85 ' sCompnyNm
                getCompany(34, 85, fsValue, False, True)
        End Select
    End Sub
    'GetNextCode(p_sMasTable, "sacctnmbr", True, p_oApp.Connection, True, "M" & Mid(p_sBranchCD, 2))
    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "stransnox"
                    p_oDTMstr(0).Item(lnCtr) = ""
                Case "sbranchd"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.BranchCode
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.getSysDate
                Case "sclientnm"
                    p_oDTMstr(0).Item(lnCtr) = ""
                Case "sdetlinfo"
                    p_oDTMstr(0).Item(lnCtr) = ""
                Case "sqmatchno"
                    p_oDTMstr(0).Item(lnCtr) = ""
                Case "sqmappcde"
                    p_oDTMstr(0).Item(lnCtr) = ""
                Case "ctranstat"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
        p_oOthersx.sTownIDxx = ""
        p_oOthersx.sCompnyNm = ""
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

        Return True
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
            If fsValue = p_oOthersx.sClientNm And fsValue <> "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oApp)
        loClient.Parent = "MC_AR_Master"

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
                p_oOthersx.sTownIDxx = p_oClient.Master("sTownIDxx")
            Else
                p_oDTMstr(0).Item("sClientID") = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sAddressx = ""
                p_oOthersx.sTownIDxx = ""
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
                p_oOthersx.sTownIDxx = p_oClient.Master("sTownIDxx")
            End If
        End If
        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
    End Sub


    'This method implements a search master where id and desc are not joined.
    Private Sub getCompany(ByVal fnColIdx As Integer _
                         , ByVal fnColDsc As Integer _
                         , ByVal fsValue As String _
                         , ByVal fbIsCode As Boolean _
                         , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sCompnyNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sCompnyNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sCompnyID" & _
                       ", a.sCompnyNm" & _
               " FROM Company a" & _
               IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sCompnyID»sCompnyNm" _
                                             , "ID»Company", _
                                             , "a.sCompnyID»a.sCompnyNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sCompnyNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sCompnyID")
                p_oOthersx.sCompnyNm = loRow.Item("sCompnyNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCompnyNm)
            Exit Sub

        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sCompnyID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sCompnyNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sCompnyNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sCompnyID")
            p_oOthersx.sCompnyNm = loDta(0).Item("sCompnyNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCompnyNm)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getBranch(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sBranchNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sBranchNm And fsValue <> "" Then Exit Sub
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
                                             , "Code»Branch", _
                                             , "a.sBranchCD»a.sBranchNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBranchNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sBranchCD")
                p_oOthersx.sBranchNm = loRow.Item("sBranchNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBranchNm)
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
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sBranchNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sBranchCD")
            p_oOthersx.sBranchNm = loDta(0).Item("sBranchNm")
        End If
        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBranchNm)
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sBranchCd" & _
                    ", a.dTransact" & _
                    ", a.sClientNm" & _
                    ", a.sDetlInfo" & _
                    ", a.sQMatchNo" & _
                    ", a.sQMAppCde" & _
                    ", a.cTranStat" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCD = p_oApp.BranchCode
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sAddressx As String
        Public sTownIDxx As String
        Public sCompnyNm As String
        Public sBranchNm As String
    End Class

End Class

