'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Application Selling Branch Transfer
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
'  Mac [ 06/25/2021 02:25 pm ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit On

Imports ADODB
Imports ggcAppDriver

Public Class LRApplicationTransfer
    Private Const pxeModuleName As String = "LRApplicationTransfer"

    Private p_oApp As GRider
    Private p_oDTDetx As DataTable
    Private p_nEditMode As xeEditMode

    Private p_sDestinat As String

    Public Event MasterRetreive(ByVal Index As Integer, ByVal Value As Object)

    Public Sub New()
        initRecord()
    End Sub

    Public Property AppDriver() As GRider
        Get
            Return p_oApp
        End Get
        Set(ByVal value As GRider)
            p_oApp = value
        End Set
    End Property

    Public ReadOnly Property ItemCount()
        Get
            Return p_oDTDetx.Rows.Count
        End Get
    End Property

    Public Property Detail(ByVal fnRow As Integer, ByVal fsIndex As String) As Object
        Get
            Return p_oDTDetx(fnRow)(fsIndex)
        End Get

        Set(ByVal value As Object)
            If fsIndex = "cUpdteRec" Then
                p_oDTDetx(fnRow)(fsIndex) = value
            End If
        End Set
    End Property

    Public Function SearchRecord(ByVal fsValue As String, ByVal fbByCode As Boolean) As Boolean
        Dim lsSQL As String
        Dim loDT As DataTable
        Dim lnCtr As Integer

        If TypeName(p_oApp) = "Nothing" Then
            MsgBox("Application driver is not set.", MsgBoxStyle.Critical, pxeModuleName)
            Return False
        End If

        lsSQL = getSQLSearch()

        If Not fbByCode Then
            lsSQL = AddCondition(lsSQL, "b.sCompnyNm LIKE " & strParm(fsValue & "%"))
        Else
            lsSQL = AddCondition(lsSQL, "a.sGOCASNox = " & strParm(fsValue))
        End If

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            MsgBox("No record found.")
            Return False
        Else
            initRecord()

            For lnCtr = 0 To loDT.Rows.Count - 1
                p_oDTDetx.Rows.Add()
                p_oDTDetx(lnCtr)("sTransNox") = loDT(lnCtr)("sTransNox")
                p_oDTDetx(lnCtr)("sBranchNm") = loDT(lnCtr)("sBranchNm")
                p_oDTDetx(lnCtr)("dAppliedx") = loDT(lnCtr)("dAppliedx")
                p_oDTDetx(lnCtr)("sCompnyNm") = loDT(lnCtr)("sCompnyNm")
                p_oDTDetx(lnCtr)("sQMatchNo") = loDT(lnCtr)("sQMatchNo")
                p_oDTDetx(lnCtr)("sGOCASNox") = loDT(lnCtr)("sGOCASNox")
                p_oDTDetx(lnCtr)("sReferNox") = loDT(lnCtr)("sReferNox")
                p_oDTDetx(lnCtr)("sBranchCd") = loDT(lnCtr)("sBranchCd")
                p_oDTDetx(lnCtr)("cUpdteRec") = xeLogical.NO
            Next
        End If

        p_nEditMode = xeEditMode.MODE_UPDATE
        Return True
    End Function

    Public Function SaveRecord() As Boolean
        Dim lsSQL As String
        Dim lnCtr As Integer

        If TypeName(p_oApp) = "Nothing" Then
            MsgBox("Application driver is not set.", MsgBoxStyle.Critical, pxeModuleName)
            Return False
        End If

        If p_nEditMode <> xeEditMode.MODE_UPDATE Then
            MsgBox("Invalid edit mode detected.", MsgBoxStyle.Critical, pxeModuleName)
            Return False
        End If

        If p_sDestinat = "" Then
            MsgBox("No destination branch loaded.", MsgBoxStyle.Critical, pxeModuleName)
            Return False
        End If

        Try
            p_oApp.BeginTransaction()

            For lnCtr = 0 To ItemCount - 1
                If p_oDTDetx(lnCtr)("cUpdteRec") = "1" And IFNull(p_oDTDetx(lnCtr)("sBranchCd"), "") <> p_sDestinat Then
                    lsSQL = "UPDATE MC_Credit_Application SET" & _
                                "  sBranchCd = " & strParm(p_sDestinat) & _
                                ", sModified = " & strParm(p_oApp.UserID) & _
                                ", dModified = " & datetimeParm(p_oApp.SysDate) & _
                            " WHERE sTransNox = " & strParm(p_oDTDetx(lnCtr)("sTransNox"))

                    If p_oApp.Execute(lsSQL, "MC_Credit_Application") <= 0 Then
                        MsgBox("Unable to update Credit Application Info.", MsgBoxStyle.Critical, pxeModuleName)
                        GoTo endwithRoll
                    End If

                    'If p_oDTDetx(lnCtr)("sReferNox") <> "" Then
                    '    lsSQL = "UPDATE Credit_Online_Application SET" & _
                    '            "  sBranchCd = " & strParm(p_sDestinat) & _
                    '            ", sModified = " & strParm(p_oApp.UserID) & _
                    '        " WHERE sTransNox = " & strParm(p_oDTDetx(lnCtr)("sReferNox"))

                    '    If p_oApp.Execute(lsSQL, "Credit_Online_Application") <= 0 Then
                    '        MsgBox("Unable to update Credit Online Application Info.", MsgBoxStyle.Critical, pxeModuleName)
                    '        GoTo endwithRoll
                    '    End If
                    'End If
                End If
            Next

            p_oApp.CommitTransaction()

            initRecord()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, pxeModuleName)
        End Try

endwithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function


    Private Sub initRecord()
        p_oDTDetx = New DataTable

        p_oDTDetx.Columns.Add("sTransNox", GetType(String))
        p_oDTDetx.Columns.Add("sBranchNm", GetType(String))
        p_oDTDetx.Columns.Add("dAppliedx", GetType(Date))
        p_oDTDetx.Columns.Add("sCompnyNm", GetType(String))
        p_oDTDetx.Columns.Add("sQMatchNo", GetType(String))
        p_oDTDetx.Columns.Add("sGOCASNox", GetType(String))
        p_oDTDetx.Columns.Add("sBranchCd", GetType(String))
        p_oDTDetx.Columns.Add("cUpdteRec", GetType(Char))
        p_oDTDetx.Columns.Add("sReferNox", GetType(String))

        p_sDestinat = ""
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Private Function getSQLSearch() As String
        Return "SELECT" & _
                    "  a.sTransNox" & _
                    ", c.sBranchNm" & _
                    ", a.dAppliedx" & _
                    ", b.sCompnyNm" & _
                    ", a.sQMatchNo" & _
                    ", IFNULL(a.sGOCASNox, '') sGOCASNox" & _
                    ", IFNULL(a.sReferNox, '') sReferNox" & _
                    ", a.sBranchCd" & _
                " FROM MC_Credit_Application a" & _
                    " LEFT JOIN Branch c" & _
                        " ON c.sBranchCd = IFNULL(a.sBranchCd, LEFT(a.sTransNox, 4))" & _
                    ", Client_Master b" & _
                " WHERE a.sClientID = b.sClientID" & _
                    " AND a.dAppliedx >= DATE_SUB(CURRENT_DATE(), INTERVAL 60 DAY)" & _
                    " AND a.cTranStat = '2'" &
                    " AND IFNULL(a.sBranchCd, '') = ''" & _
                " ORDER BY a.dAppliedx, b.sCompnyNm, c.sBranchNm"
    End Function

    Public Sub SearchBranch(ByVal fsValue As String, ByVal fbIsCode As Boolean)
        If p_nEditMode <> xeEditMode.MODE_UPDATE Then Exit Sub

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_sDestinat And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBranchCD" & _
                       ", a.sBranchNm" & _
               " FROM Branch a" & _
               IIf(Not fbIsCode, " WHERE a.cRecdStat = '1'", "")

        Dim loRow As DataRow = KwikSearch(p_oApp _
                                         , lsSQL _
                                         , True _
                                         , fsValue _
                                         , "sBranchCD»sBranchNm" _
                                         , "ID»Company", _
                                         , "a.sBranchCD»a.sBranchNm" _
                                         , IIf(fbIsCode, 0, 1))
        If IsNothing(loRow) Then
            p_sDestinat = ""
            RaiseEvent MasterRetreive(1, "")
        Else
            p_sDestinat = loRow.Item("sBranchCD")
            RaiseEvent MasterRetreive(1, loRow.Item("sBranchNm"))
        End If
    End Sub
End Class