'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC Model Object
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
'  Jovan [ 05/18/2019 11:11 am ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports ggcClient
Imports System.Windows.Forms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class McModel

    Private Const pxeTableName As String = "MC_Model"
    Private Const pxeTableName1 As String = "MC_Model_Features"
    Private Const pxeTableName2 As String = "MC_Model_Specs"
    Private Const p_sMsgHeadr As String = "MC_Model"
    Private Const pxeModuleName As String = "MC Model"

    Private p_nEditMode As Integer
    Private p_oApp As GRider
    Protected p_oDTMaster As DataTable
    Private p_sBranchCd As String
    Private p_oDTDetl As DataTable
    Private p_oOthers As DataTable
    Protected p_bInitTran As Boolean

    Private p_oDTMstr As DataTable

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Master(ByVal Index As Object) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0)(Index)
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

    Public ReadOnly Property ItemCount() As Integer
        Get
            If Not IsNothing(p_oDTDetl) Then
                Return p_oDTDetl.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property OthersCount() As Integer
        Get
            If Not IsNothing(p_oOthers) Then
                Return p_oOthers.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property BranchCode() As String
        Get
            Return p_sBranchCd
        End Get
        Set(ByVal value As String)
            p_sBranchCd = value
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

    Property Others(ByVal Row As Integer, ByVal Index As Object) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oOthers(Row)(Index)
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal Value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oOthers(Row)(Index) = Value
            End If
        End Set
    End Property

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                 ByVal Value As Object)
    Public Event DetailRetrieved(ByVal Row As Integer, _
                                 ByVal Index As Integer, _
                                 ByVal Value As Object)
    Public Event OtherRetrieved(ByVal Row As Integer, _
                                ByVal Index As Integer, _
                                ByVal Value As Object)

    Public Function NewTransaction() As Boolean
        Dim lsProcName As String = "NewTransaction"

        Try
            p_oDTMstr = New DataTable
            Debug.Print(AddCondition(getSQ_Master, "0=1"))
            p_oDTMstr = p_oApp.ExecuteQuery(AddCondition(getSQ_Master, "0=1"))
            Call initMaster()

            p_oDTDetl.Clear()
            p_oOthers.Clear()
            Call initDetail()
            Call initOthers()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, pxeModuleName & "-" & lsProcName)
            GoTo errProc
        Finally
            RaiseEvent MasterRetrieved(0, p_oDTMstr.Rows(0)("sModelIDx"))
            p_nEditMode = xeEditMode.MODE_ADDNEW
        End Try

endProc:
        Return True
        Exit Function
errProc:
        Return False
    End Function

    Private Sub initMaster()
        Dim lnCtr As Integer

        With p_oDTMstr
            .Rows.Add()
            For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
                Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                    Case "smodelidx"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "smodelcde"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "smodelnme"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "smodeldsc"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "crecdstat"
                        p_oDTMstr(0).Item(lnCtr) = strParm(xeRecordStat.RECORD_NEW)
                    Case Else
                        p_oDTMstr(0).Item(lnCtr) = ""
                End Select
            Next
        End With
    End Sub


    Private Function getSQ_Master() As String
        Return "SELECT a.sModelIDx" & _
                    ", a.sModelCde" & _
                    ", a.sModelNme" & _
                " FROM " & pxeTableName & " a"
    End Function


    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Private Function getSQL_Branch() As String
        Return "SELECT" & _
                    "  sBranchCd" & _
                    ", sBranchNm" & _
                " FROM Branch" & _
                " WHERE cRecdStat = " & strParm(xeRecordStat.RECORD_NEW)
    End Function

    Function SearchTransaction( _
                            ByVal fsValue As String) As Boolean

        Dim lsSQL As String
        Dim lsCondition As String

        'Check if already loaded base on edit mode
        'If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
        '    If fsValue = p_oDTMaster(0).Item("sModelNme") Then Return True
        'End If

        lsSQL = getSQ_Master()

        'create Kwiksearch filter
        Dim lsFilter As String

        lsFilter = "a.sModelNme LIKE " & strParm("%" & fsValue & "%")

        lsSQL = AddCondition(lsSQL, lsCondition)

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sModelCde»sModelNme" _
                                        , "Model Code»Model Name", _
                                        , "a.sModelCde»a.sModelNme" _
                                        , 1)
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            OpenTransaction(loDta.Item("sModelIDx"))
            Return True
        End If
    End Function

    Public Function OpenTransaction(ByVal fsModelIdx As String) As Boolean
        Dim lsSQL As String
        Dim lsSQL1 As String
        Dim loDetail As DataTable
        Dim loOthers As DataTable

        lsSQL = AddCondition(getSQ_Master, "a.sModelIDx = " & strParm(fsModelIdx))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        If p_oDTMstr.Rows.Count = 0 Then Return False

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        lsSQL = AddCondition(getSQL_Detail, "b.sModelIDx = " & strParm(fsModelIdx))
        loDetail = p_oApp.ExecuteQuery(lsSQL)

        lsSQL1 = AddCondition(getSQL_Others, "c.sModelIDx = " & strParm(fsModelIdx))
        loOthers = p_oApp.ExecuteQuery(lsSQL1)

        p_oDTDetl.Clear()
        With loDetail
            If loDetail.Rows.Count > 0 Then
                For nCtr As Integer = 0 To .Rows.Count - 1
                    p_oDTDetl.Rows.Add()
                    p_oDTDetl.Rows(nCtr)("sModelIDx") = .Rows(nCtr)("sModelIDx")
                    p_oDTDetl.Rows(nCtr)("nEntryNox") = .Rows(nCtr)("nEntryNox")
                    p_oDTDetl.Rows(nCtr)("sFeatrIDx") = .Rows(nCtr)("sFeatrIDx")
                    p_oDTDetl.Rows(nCtr)("sDescript") = .Rows(nCtr)("sDescript")
                Next nCtr
            End If
        End With

        p_oOthers.Clear()
        With loOthers
            If loOthers.Rows.Count > 0 Then
                For nCtr As Integer = 0 To .Rows.Count - 1
                    p_oOthers.Rows.Add()
                    p_oOthers.Rows(nCtr)("sModelIDx") = .Rows(nCtr)("sModelIDx")
                    p_oOthers.Rows(nCtr)("nEntryNox") = .Rows(nCtr)("nEntryNox")
                    p_oOthers.Rows(nCtr)("sSpecsIDx") = .Rows(nCtr)("sSpecsIDx")
                    p_oOthers.Rows(nCtr)("sDescript") = .Rows(nCtr)("sDescript")
                Next nCtr
            End If
        End With


        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    Private Function getSQL_Detail() As String
        Return "SELECT" & _
                    "  b.sModelIDx" & _
                    ", b.nEntryNox" & _
                    ", b.sFeatrIDx" & _
                    ", b.sDescript" & _
                " FROM " & pxeTableName1 & " b" & _
                " ORDER BY b.nEntryNox ASC"
    End Function

    Public Function InitTransaction() As Boolean
        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oApp.BranchCode
        createDetailTable()
        createOtherTable()

        p_sBranchCd = p_oApp.BranchCode
        p_bInitTran = True
        InitTransaction = True
    End Function

    Private Function getSQL_Others() As String
        Return "SELECT" & _
                    "  c.sModelIDx" & _
                    ", c.nEntryNox" & _
                    ", c.sSpecsIDx" & _
                    ", c.sDescript" & _
                " FROM " & pxeTableName2 & " c" & _
                " ORDER BY c.nEntryNox ASC"

    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub initDetail()
        Dim lnCtr As Integer

        With p_oDTDetl
            .Rows.Add()
            For lnCtr = 0 To p_oDTDetl.Columns.Count - 1
                Select Case LCase(p_oDTDetl.Columns(lnCtr).ColumnName)
                    Case "smodelidx"
                        .Rows(0)(lnCtr) = p_oDTMstr(0).Item("sModelIDx")
                    Case "nentrynox"
                        .Rows(0)(lnCtr) = 0
                    Case "sfeatridx"
                        .Rows(0).Item(lnCtr) = ""
                    Case "sDescript"
                        .Rows(0).Item(lnCtr) = ""
                End Select
            Next
        End With

    End Sub

    Private Sub initOthers()
        Dim lnCtr As Integer

        With p_oOthers
            .Rows.Add()
            For lnCtr = 0 To p_oOthers.Columns.Count - 1
                Select Case LCase(p_oOthers.Columns(lnCtr).ColumnName)
                    Case "smodelidx"
                        .Rows(0)(lnCtr) = p_oDTMstr(0).Item("sModelIDx")
                    Case "nentrynox"
                        .Rows(0)(lnCtr) = 0
                    Case "sspecsidx"
                        .Rows(0).Item(lnCtr) = ""
                    Case "sdescript"
                        .Rows(0).Item(lnCtr) = ""
                End Select
            Next
        End With

    End Sub

    Private Sub createDetailTable()
        p_oDTDetl = New DataTable
        With p_oDTDetl
            .Columns.Add("sModelIDx", GetType(String)).MaxLength = 9
            .Columns.Add("nEntryNox", GetType(Integer))
            .Columns.Add("sFeatrIDx", GetType(String)).MaxLength = 8
            .Columns.Add("sDescript", GetType(String)).MaxLength = 64
        End With
    End Sub

    Private Sub createOtherTable()
        p_oOthers = New DataTable
        With p_oOthers
            .Columns.Add("sModelIDx", GetType(String)).MaxLength = 9
            .Columns.Add("nEntryNox", GetType(Integer))
            .Columns.Add("sSpecsIDx", GetType(String)).MaxLength = 8
            .Columns.Add("sDescript", GetType(String)).MaxLength = 64
        End With
    End Sub

End Class