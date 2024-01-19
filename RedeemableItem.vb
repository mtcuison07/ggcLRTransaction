'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Redeemable Object
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
'  Jovan [ 07/10/2019 01:20 am ]
'     Start coding this object...
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports ggcClient
Imports System.Windows.Forms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class RedeemableItem

    Private Const pxeTableName As String = "G_Card_Promo_Master"
    Private Const p_sMsgHeadr As String = "Redeemable Items"
    Private Const pxeModuleName As String = "Redeemable Items"

    Private p_nEditMode As Integer
    Private p_oApp As GRider
    Protected p_oDTMaster As DataTable
    Private p_sBranchCd As String
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
            End If
        End Set
    End Property

    Public Property BranchCode() As String
        Get
            Return p_sBranchCd
        End Get
        Set(ByVal value As String)
            p_sBranchCd = value
        End Set
    End Property

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                 ByVal Value As Object)

    Public Function NewTransaction() As Boolean
        Dim lsProcName As String = "NewTransaction"

        Try
            p_oDTMstr = New DataTable
            Debug.Print(AddCondition(getSQ_Master, "0=1"))
            p_oDTMstr = p_oApp.ExecuteQuery(AddCondition(getSQ_Master, "0=1"))
            Call initMaster()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, pxeModuleName & "-" & lsProcName)
            GoTo errProc
        Finally
            RaiseEvent MasterRetrieved(0, p_oDTMstr.Rows(0)("sPromCode"))
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
                    Case "spromcode"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "spromdesc"
                        p_oDTMstr(0).Item(lnCtr) = ""
                    Case "npointsxx"
                        p_oDTMstr(0).Item(lnCtr) = 0
                    Case "cpreorder"
                        p_oDTMstr(0).Item(lnCtr) = "0"
                    Case Else
                        p_oDTMstr(0).Item(lnCtr) = ""
                End Select
            Next
        End With
    End Sub

    Private Function getSQ_Master() As String
        Return "SELECT   a.sPromCode" & _
                        ", a.sPromDesc" & _
                        ", a.nPointsxx" & _
                        ", a.cPreOrder" & _
                        " FROM " & pxeTableName & " a" & _
                        " WHERE a.cTranStat <> '3'" & _
                        " AND dDateFrom <= " & dateParm(p_oApp.getSysDate) & _
                        " AND dDateThru >= " & dateParm(p_oApp.getSysDate) & _
                        " ORDER BY a.sPromCode ASC"
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

    Public Function SaveUpdate(ByVal sPromCode) As Boolean
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnRow As Integer

        lsProcName = "SaveUpdate"
        'On Error GoTo errProc

        If Trim(sPromCode) = "" Then
            GoTo endProc
        End If

        'Retreiving of meal voucher
        With p_oDTMaster
            lsSQL = "UPDATE " & _
                " G_Card_Promo_Master SET " & _
                " cPreOrder = " & strParm(p_oDTMstr.Rows(0)("cPreOrder")) & _
                " WHERE sPromCode = " & strParm(p_oDTMstr.Rows(0)("sPromCode"))

            lnRow = p_oApp.Execute(lsSQL, "G_Card_Promo_Master")
        End With

        If lnRow = 0 Then
            Call initMaster()
            GoTo endProc
        End If

        Return True
endProc:
        lnRow = Nothing

        Return False

    End Function

    Public Function SearchTransaction( _
                      ByVal fsValue As String _
                    , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String
        Dim lsCondition As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fsValue = p_oDTMstr.Rows(0)("sPromCode") Then Return True
        End If

        lsSQL = getSQ_Master()

        'create Kwiksearch filter
        Dim lsFilter As String

        lsFilter = "a.sPromDesc LIKE " & strParm("%" & fsValue & "%")

        lsSQL = AddCondition(lsSQL, lsCondition)

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sPromCode»sPromDesc" _
                                        , "Promo Code»Descrption", _
                                        , "a.sPromCode»a.sPromDesc" _
                                        , 1)
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            OpenTransaction(loDta.Item("sPromCode"))
            Return True
        End If
    End Function

    Public Function OpenTransaction(ByVal fsPromCode As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sPromCode = " & strParm(fsPromCode))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        If p_oDTMstr.Rows.Count = 0 Then Return False

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function


    Public Function InitTransaction() As Boolean
        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oApp.BranchCode

        p_sBranchCd = p_oApp.BranchCode
        p_bInitTran = True
        InitTransaction = True
    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class