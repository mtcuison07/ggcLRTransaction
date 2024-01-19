'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Car Serial Registration Object
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
'  Jheff [ 04/27/2018 08:44 am ]
'     Start coding this object...
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports ggcClient
Imports System.Windows.Forms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class LRCarSerialRegistration

#Region "Constant"
    Private Const xsSignature As String = "08220326"
    Private Const pxeMODULENAME As String = "LRSerial"
    Private Const pxeMasterTble As String = "Car_Serial_Registration"
#End Region

#Region "Protected Members"
    Protected p_oAppDrvr As GRider
    Protected p_oDTMaster As DataTable
    Protected p_nEditMode As xeEditMode
    Protected p_oSC As New MySqlCommand
    Protected p_oDT As DataTable

    Protected p_sBranchCd As String
    Protected p_sParent As String
    Protected p_bCancelled As Boolean
#End Region

#Region "Public Event"
    Public Event MasterRetrieved(ByVal Index As Integer, _
                              ByVal Value As Object)
#End Region

#Region "Private"
    Private p_oClientID As ggcClient.Client
    Private p_oCoCltID1 As ggcClient.Client
    Private p_oCoCltID2 As ggcClient.Client
    Private p_oRegCltID As ggcClient.Client
    Private p_oRgCltID1 As ggcClient.Client
    Private p_oRgCltID2 As ggcClient.Client
#End Region

#Region "Properties"
    ReadOnly Property AppDriver As GRider
        Get
            Return p_oAppDrvr
        End Get
    End Property

    ReadOnly Property Cancelled
        Get
            Return p_bCancelled
        End Get
    End Property

    Property Branch() As String
        Get
            Return p_sBranchCd
        End Get
        Set(ByVal Value As String)
            p_sBranchCd = Value
        End Set
    End Property

    Public ReadOnly Property ItemCount() As Integer
        Get
            If Not IsNothing(p_oDTMaster) Then
                Return p_oDTMaster.Rows.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Property Master(ByVal Index As Object) As Object
        Get
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sserialid" : Index = 0
                    Case "sengineno" : Index = 1
                    Case "sframenox" : Index = 2
                    Case "sbrandnme" : Index = 3
                    Case "smodelnme" : Index = 4
                    Case "scolornme" : Index = 5
                    Case "nyearmodl" : Index = 6
                    Case "sfilenoxx" : Index = 7
                    Case "screnoxxx" : Index = 8
                    Case "scrnoxxxx" : Index = 9
                    Case "splatenop" : Index = 10
                    Case "sregornox" : Index = 11
                    Case "sstickrno" : Index = 12
                    Case "dregister" : Index = 13
                    Case "sclientnm" : Index = 14
                    Case "scocltnm1" : Index = 15
                    Case "scocltnm2" : Index = 16
                    Case "sregcltNm" : Index = 17
                    Case "srgcltnm1" : Index = 18
                    Case "srgcltnm2" : Index = 19
                    Case "sclientid" : Index = 20
                    Case "scocltid1" : Index = 21
                    Case "scocltid2" : Index = 22
                    Case "sregcltid" : Index = 23
                    Case "srgcltid1" : Index = 24
                    Case "srgcltid2" : Index = 25
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                        Return DBNull.Value
                End Select
            End If
            Return p_oDTMaster(0)(Index)
        End Get

        Set(ByVal Value As Object)
            If Not IsNumeric(Index) Then
                Index = LCase(Index)
                Select Case Index
                    Case "sserialid" : Index = 0
                    Case "sengineno" : Index = 1
                    Case "sframenox" : Index = 2
                    Case "sbrandnme" : Index = 3
                    Case "smodelnme" : Index = 4
                    Case "scolornme" : Index = 5
                    Case "nyearmodl" : Index = 6
                    Case "sfilenoxx" : Index = 7
                    Case "screnoxxx" : Index = 8
                    Case "scrnoxxxx" : Index = 9
                    Case "splatenop" : Index = 10
                    Case "sregornox" : Index = 11
                    Case "sstickrno" : Index = 12
                    Case "dregister" : Index = 13
                    Case "sclientnm" : Index = 14
                        getClient(Index, Value, False)
                    Case "scocltnm1" : Index = 15
                        getClient(Index, Value, False)
                    Case "sCocltnm2" : Index = 16
                        getClient(Index, Value, False)
                    Case "sregcltnm" : Index = 17
                        getClient(Index, Value, False)
                    Case "srgcltnm1" : Index = 18
                        getClient(Index, Value, False)
                    Case "srgcltnm2" : Index = 19
                        getClient(Index, Value, False)
                    Case "sclientid" : Index = 20
                    Case "scocltid1" : Index = 21
                    Case "scocltid2" : Index = 22
                    Case "sregcltid" : Index = 23
                    Case "srgcltid1" : Index = 24
                    Case "srgcltid2" : Index = 25
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                End Select
            End If
            p_oDTMaster(0)(Index) = Value
        End Set
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property
#End Region

#Region "Private Function"
    Private Function getSQL_Master() As String
        Return "SELECT a.sSerialID" & _
                    ", a.sEngineNo" & _
                    ", a.sFrameNox" & _
                    ", a.nYearModl" & _
                    ", a.sFileNoxx" & _
                    ", a.sCRENoxxx" & _
                    ", a.sCRNoxxxx" & _
                    ", a.sPlateNoP" & _
                    ", a.sRegORNox" & _
                    ", a.sStickrNo" & _
                    ", a.dRegister" & _
                    ", b.sBrandNme" & _
                    ", c.sModelNme" & _
                    ", d.sColorNme" & _
                    ", a.sClientID" & _
                    ", a.sCoCltID1" & _
                    ", a.sCoCltID2" & _
                    ", a.sRegCltID" & _
                    ", a.sRgCltID1" & _
                    ", a.sRgCltID2" & _
            " FROM Car_Serial a" & _
                ", Car_Brand b" & _
                ", Car_Model c" & _
                ", Color d" & _
           " WHERE a.sBrandCde = b.sBrandIDx" & _
                " AND a.sModelCde = c.sModelIDx" & _
                " AND a.sColorCde = d.sColorIDx"
    End Function

    Private Function isEntryOk() As Boolean
        If p_oDTMaster(0).Item("sEngineNo") = "" Then
            MsgBox("Engino No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, pxeMODULENAME)
            Return False
        End If

        Return True
    End Function
#End Region

#Region "Public function"
    Function SaveTransaction() As Boolean
        Dim lsSQL As String
        Dim lnRow As Integer

        With p_oDTMaster
            If p_bCancelled Then Return False

            If Not isEntryOk() Then Return False

            Try
                If p_sParent = "" Then p_oAppDrvr.BeginTransaction()

                lsSQL = "INSERT INTO " & pxeMasterTble & " SET" & _
                        "  sSerialID = " & strParm(.Rows(0)("sSerialID")) & _
                        ", sFileNoxx = " & strParm(.Rows(0)("sFileNoxx")) & _
                        ", sCRENoxxx = " & strParm(.Rows(0)("sCRENoxxx")) & _
                        ", sCRNoxxxx = " & strParm(.Rows(0)("sCRNoxxxx")) & _
                        ", sPlateNoP = " & strParm(.Rows(0)("sPlateNoP")) & _
                        ", sRegORNox = " & strParm(.Rows(0)("sRegORNox")) & _
                        ", sStickrNo = " & strParm(.Rows(0)("sStickrNo")) & _
                        ", dRegister = " & dateParm(.Rows(0)("dRegister")) & _
                        ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                        ", sCoCltID1 = " & strParm(.Rows(0)("sCoCltID1")) & _
                        ", sCoCltID2 = " & strParm(.Rows(0)("sCoCltID2")) & _
                        ", sRegCltID = " & strParm(.Rows(0)("sRegCltID")) & _
                        ", sRgCltID1 = " & strParm(.Rows(0)("sRgCltID1")) & _
                        ", sRgCltID2 = " & strParm(.Rows(0)("sRgCltID2")) & _
                      " ON DUPLICATE KEY UPDATE" & _
                        "  sFileNoxx = " & strParm(.Rows(0)("sFileNoxx")) & _
                        ", sCRENoxxx = " & strParm(.Rows(0)("sCRENoxxx")) & _
                        ", sCRNoxxxx = " & strParm(.Rows(0)("sCRNoxxxx")) & _
                        ", sPlateNoP = " & strParm(.Rows(0)("sPlateNoP")) & _
                        ", sRegORNox = " & strParm(.Rows(0)("sRegORNox")) & _
                        ", sStickrNo = " & strParm(.Rows(0)("sStickrNo")) & _
                        ", dRegister = " & dateParm(.Rows(0)("dRegister")) & _
                        ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                        ", sCoCltID1 = " & strParm(.Rows(0)("sCoCltID1")) & _
                        ", sCoCltID2 = " & strParm(.Rows(0)("sCoCltID2")) & _
                        ", sRegCltID = " & strParm(.Rows(0)("sRegCltID")) & _
                        ", sRgCltID1 = " & strParm(.Rows(0)("sRgCltID1")) & _
                        ", sRgCltID2 = " & strParm(.Rows(0)("sRgCltID2"))
                Try
                    lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                    If lnRow <= 0 Then
                        MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                                "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                Catch ex As Exception
                    Throw ex
                End Try

                If .Rows(0)("sClientID") <> "" Then
                    If Not p_oClientID.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sCoCltID1") <> "" Then
                    If Not p_oCoCltID1.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sCoCltID2") <> "" Then
                    If Not p_oCoCltID2.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sRegCltID") <> "" Then
                    If Not p_oRegCltID.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sRgCltID1") <> "" Then
                    If Not p_oRgCltID1.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("ssRgCltID2") <> "" Then
                    If Not p_oRgCltID2.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If p_sParent = "" Then p_oAppDrvr.CommitTransaction()

                Return True
            Catch ex As Exception
                If p_sParent = "" Then p_oAppDrvr.RollBackTransaction()

                MsgBox(ex.Message)

                Return False
            End Try
        End With
    End Function

    Function OpenTransaction(ByVal fsSerialID As String, ByVal fdRegister As Date) As Boolean
        Dim loDT As New DataTable
        Dim lsSQL As String

        lsSQL = AddCondition(getSQL_Master, "a.sSerialID = " & strParm(fsSerialID) & _
                                                " AND a.dRegister = " & dateParm(fdRegister))

        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Return False

        Call createMasterTable()

        If loDT.Rows.Count > 0 Then
            With p_oDTMaster

                .Rows.Add()
                For nCtr As Integer = 0 To .Columns.Count - 1
                    Select Case .Columns.Item(nCtr).ColumnName
                        Case "sClientNm"
                            getClient(14, .Rows(0)("sClientID"), True)
                        Case "sCoCltNm1"
                            getClient(15, .Rows(0)("sCoCltID1"), True)
                        Case "sCoCltNm2"
                            getClient(16, .Rows(0)("sCoCltID2"), True)
                        Case "sRegCltNm"
                            getClient(17, .Rows(0)("sRegCltID"), True)
                        Case "sRgCltNm1"
                            getClient(18, .Rows(0)("sRgCltID1"), True)
                        Case "sRgCltNm2"
                            getClient(19, .Rows(0)("sRgCltID2"), True)
                        Case Else
                            .Rows(0)(.Columns.Item(nCtr).ColumnName) = loDT.Rows(0)(.Columns.Item(nCtr).ColumnName)
                    End Select
                Next nCtr
            End With
        Else
            Call initMaster()
        End If

        Return True
    End Function

    Function getRecord(ByVal fsSerialID As String) As DataTable
        Dim loDT As New DataTable
        Dim lsSQL As String

        lsSQL = "SELECT a.sSerialID" & _
                    ", a.sEngineNo" & _
                    ", a.sFrameNox" & _
                    ", a.nYearModl" & _
                    ", a.sFileNoxx" & _
                    ", a.sCRENoxxx" & _
                    ", a.sCRNoxxxx" & _
                    ", a.sPlateNoP" & _
                    ", a.sRegORNox" & _
                    ", a.sStickrNo" & _
                    ", a.dRegister" & _
                    ", b.sBrandNme" & _
                    ", c.sModelNme" & _
                    ", d.sColorNme" & _
                    ", e.sCompnyNm sClientNm" & _
                    ", a.sClientID" & _
                    ", a.sCoCltID1" & _
                    ", a.sCoCltID2" & _
                    ", a.sRegCltID" & _
                    ", a.sRgCltID1" & _
                    ", a.sRgCltID2" & _
            " FROM Car_Serial a" & _
                ", Car_Brand b" & _
                ", Car_Model c" & _
                ", Color d" & _
                ", Client_Master e" & _
            " WHERE a.sBrandCde = b.sBrandIDx" & _
                " AND a.sModelCde = c.sModelIDx" & _
                " AND a.sColorCde = d.sColorIDx" & _
                " AND a.sClientID = e.sClientID" & _
            " ORDER BY a.dRegister"

        lsSQL = AddCondition(getSQL_Master, "a.sSerialID = " & strParm(fsSerialID))

        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        Return loDT
    End Function
#End Region

#Region "Private Procedures"
    Private Sub createMasterTable()
        p_oDTMaster = New DataTable
        With p_oDTMaster
            .Columns.Add("sSerialID", GetType(String)).MaxLength = 12
            .Columns.Add("sEngineNo", GetType(String)).MaxLength = 20
            .Columns.Add("sFrameNox", GetType(String)).MaxLength = 20
            .Columns.Add("sBrandNme", GetType(String)).MaxLength = 30
            .Columns.Add("sModelNme", GetType(String)).MaxLength = 30
            .Columns.Add("sColorNme", GetType(String)).MaxLength = 30
            .Columns.Add("nYearModl", GetType(Integer))
            .Columns.Add("sFileNoxx", GetType(String)).MaxLength = 20
            .Columns.Add("sCRENoxxx", GetType(String)).MaxLength = 10
            .Columns.Add("sCRNoxxxx", GetType(String)).MaxLength = 10
            .Columns.Add("sPlateNoP", GetType(String)).MaxLength = 8
            .Columns.Add("sRegORNox", GetType(String)).MaxLength = 15
            .Columns.Add("sStickrNo", GetType(String)).MaxLength = 8
            .Columns.Add("dRegister", GetType(Date))
            .Columns.Add("sClientNm", GetType(String)).MaxLength = 128
            .Columns.Add("sCoCltNm1", GetType(String)).MaxLength = 128
            .Columns.Add("sCoCltNm2", GetType(String)).MaxLength = 128
            .Columns.Add("sRegCltNm", GetType(String)).MaxLength = 128
            .Columns.Add("sRgCltNm1", GetType(String)).MaxLength = 128
            .Columns.Add("sRgCltNm2", GetType(String)).MaxLength = 128
            .Columns.Add("sClientID", GetType(String)).MaxLength = 12
            .Columns.Add("sCoCltID1", GetType(String)).MaxLength = 12
            .Columns.Add("sCoCltID2", GetType(String)).MaxLength = 12
            .Columns.Add("sRegCltID", GetType(String)).MaxLength = 12
            .Columns.Add("sRgCltID1", GetType(String)).MaxLength = 12
            .Columns.Add("sRgCltID2", GetType(String)).MaxLength = 12
        End With
    End Sub

    Private Sub initMaster()
        With p_oDTMaster
            .Rows.Add()
            .Rows(0)("sSerialID") = ""
            .Rows(0)("sEngineNo") = ""
            .Rows(0)("sFrameNox") = ""
            .Rows(0)("sBrandNme") = ""
            .Rows(0)("sModelNme") = ""
            .Rows(0)("sColorNme") = ""
            .Rows(0)("nYearModl") = 0
            .Rows(0)("sFileNoxx") = ""
            .Rows(0)("sCRENoxxx") = ""
            .Rows(0)("sCRNoxxxx") = ""
            .Rows(0)("sPlateNoP") = ""
            .Rows(0)("sRegORNox") = ""
            .Rows(0)("sStickrNo") = ""
            .Rows(0)("dRegister") = p_oAppDrvr.SysDate
            .Rows(0)("sClientNm") = ""
            .Rows(0)("sCoCltNm1") = ""
            .Rows(0)("sCoCltNm2") = ""
            .Rows(0)("sRegCltNm") = ""
            .Rows(0)("sRgCltNm1") = ""
            .Rows(0)("sRgCltNm2") = ""
            .Rows(0)("sClientID") = ""
            .Rows(0)("sCoCltID1") = ""
            .Rows(0)("sCoCltID2") = ""
            .Rows(0)("sRegCltID") = ""
            .Rows(0)("sRgCltID1") = ""
            .Rows(0)("sRgCltID2") = ""
        End With
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fnIndex As String _
                            , ByVal fsValue As String _
                            , ByVal fbIsCode As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMaster(0).Item(fnIndex) Then Exit Sub
        Else
            'Do not allow searching of value if fsValue is empty
            If (fsValue = p_oDTMaster(0).Item(fnIndex) And fsValue <> "") Or fsValue = "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oAppDrvr)
        loClient.Parent = "LRCarSerial"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                Select Case fnIndex
                    Case 14
                        p_oClientID = loClient
                        p_oDTMaster(0).Item("sClientID") = p_oClientID.Master("sClientID")
                        p_oDTMaster(0).Item("sClientNm") = p_oClientID.Master("sLastName") & ", " & _
                                                           p_oClientID.Master("sFrstName") & _
                                                           IIf(p_oClientID.Master("sSuffixNm") = "", "", " " & p_oClientID.Master("sSuffixNm")) & " " & _
                                                           p_oClientID.Master("sMiddName")

                    Case 15
                        p_oCoCltID1 = loClient
                        p_oDTMaster(0).Item("sCoCltID1") = p_oCoCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm1") = p_oCoCltID1.Master("sLastName") & ", " & _
                                                           p_oCoCltID1.Master("sFrstName") & _
                                                           IIf(p_oCoCltID1.Master("sSuffixNm") = "", "", " " & p_oCoCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID1.Master("sMiddName")

                    Case 16
                        p_oCoCltID2 = loClient
                        p_oDTMaster(0).Item("sCoCltID2") = p_oCoCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm2") = p_oCoCltID2.Master("sLastName") & ", " & _
                                                           p_oCoCltID2.Master("sFrstName") & _
                                                           IIf(p_oCoCltID2.Master("sSuffixNm") = "", "", " " & p_oCoCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID2.Master("sMiddName")

                    Case 17
                        p_oRegCltID = loClient
                        p_oDTMaster(0).Item("sRegCltID") = p_oRegCltID.Master("sClientID")
                        p_oDTMaster(0).Item("sRegCltNm") = p_oRegCltID.Master("sLastName") & ", " & _
                                                           p_oRegCltID.Master("sFrstName") & _
                                                           IIf(p_oRegCltID.Master("sSuffixNm") = "", "", " " & p_oRegCltID.Master("sSuffixNm")) & " " & _
                                                           p_oRegCltID.Master("sMiddName")

                    Case 18
                        p_oRgCltID1 = loClient
                        p_oDTMaster(0).Item("sRgCltID1") = p_oRgCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm1") = p_oRgCltID1.Master("sLastName") & ", " & _
                                                           p_oRgCltID1.Master("sFrstName") & _
                                                           IIf(p_oRgCltID1.Master("sSuffixNm") = "", "", " " & p_oRgCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID1.Master("sMiddName")
                    Case 19
                        p_oRgCltID2 = loClient
                        p_oDTMaster(0).Item("sRgCltID2") = p_oRgCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm2") = p_oRgCltID2.Master("sLastName") & ", " & _
                                                           p_oRgCltID2.Master("sFrstName") & _
                                                           IIf(p_oRgCltID2.Master("sSuffixNm") = "", "", " " & p_oRgCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID2.Master("sMiddName")
                End Select


            Else
                Select Case fnIndex
                    Case 14
                        p_oDTMaster(0).Item("sClientID") = ""
                        p_oDTMaster(0).Item("sClientNm") = ""
                    Case 15
                        p_oDTMaster(0).Item("sCoCltID1") = ""
                        p_oDTMaster(0).Item("sCoCltNm1") = ""
                    Case 16
                        p_oDTMaster(0).Item("sCoCltID2") = ""
                        p_oDTMaster(0).Item("sCoCltNm1") = ""
                    Case 17
                        p_oDTMaster(0).Item("sRegCltID") = ""
                        p_oDTMaster(0).Item("sRegCltNm") = ""
                    Case 18
                        p_oDTMaster(0).Item("sRgCltID1") = ""
                        p_oDTMaster(0).Item("sRgCltNm2") = ""
                    Case 19
                        p_oDTMaster(0).Item("sRgCltID1") = ""
                        p_oDTMaster(0).Item("sRgCltNm2") = ""
                End Select
            End If

            RaiseEvent MasterRetrieved(fnIndex, p_oDTMaster(0).Item(fnIndex))
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                Select Case fnIndex
                    Case 14
                        p_oClientID = loClient
                        p_oDTMaster(0).Item("sClientID") = p_oClientID.Master("sClientID")
                        p_oDTMaster(0).Item("sClientNm") = p_oClientID.Master("sLastName") & ", " & _
                                                           p_oClientID.Master("sFrstName") & _
                                                           IIf(p_oClientID.Master("sSuffixNm") = "", "", " " & p_oClientID.Master("sSuffixNm")) & " " & _
                                                           p_oClientID.Master("sMiddName")

                    Case 15
                        p_oCoCltID1 = loClient
                        p_oDTMaster(0).Item("sCoCltID1") = p_oCoCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm1") = p_oCoCltID1.Master("sLastName") & ", " & _
                                                           p_oCoCltID1.Master("sFrstName") & _
                                                           IIf(p_oCoCltID1.Master("sSuffixNm") = "", "", " " & p_oCoCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID1.Master("sMiddName")

                    Case 16
                        p_oCoCltID2 = loClient
                        p_oDTMaster(0).Item("sCoCltID2") = p_oCoCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm2") = p_oCoCltID2.Master("sLastName") & ", " & _
                                                           p_oCoCltID2.Master("sFrstName") & _
                                                           IIf(p_oCoCltID2.Master("sSuffixNm") = "", "", " " & p_oCoCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID2.Master("sMiddName")

                    Case 17
                        p_oRegCltID = loClient
                        p_oDTMaster(0).Item("sRegCltID") = p_oRegCltID.Master("sClientID")
                        p_oDTMaster(0).Item("sRegCltNm") = p_oRegCltID.Master("sLastName") & ", " & _
                                                           p_oRegCltID.Master("sFrstName") & _
                                                           IIf(p_oRegCltID.Master("sSuffixNm") = "", "", " " & p_oRegCltID.Master("sSuffixNm")) & " " & _
                                                           p_oRegCltID.Master("sMiddName")

                    Case 18
                        p_oRgCltID1 = loClient
                        p_oDTMaster(0).Item("sRgCltID1") = p_oRgCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm1") = p_oRgCltID1.Master("sLastName") & ", " & _
                                                           p_oRgCltID1.Master("sFrstName") & _
                                                           IIf(p_oRgCltID1.Master("sSuffixNm") = "", "", " " & p_oRgCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID1.Master("sMiddName")
                    Case 19
                        p_oRgCltID2 = loClient
                        p_oDTMaster(0).Item("sRgCltID2") = p_oRgCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm2") = p_oRgCltID2.Master("sLastName") & ", " & _
                                                           p_oRgCltID2.Master("sFrstName") & _
                                                           IIf(p_oRgCltID2.Master("sSuffixNm") = "", "", " " & p_oRgCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID2.Master("sMiddName")
                End Select
            End If
        End If

        RaiseEvent MasterRetrieved(fnIndex, p_oDTMaster.Rows(0)(fnIndex))
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub New(ByVal foRider As GRider)
        p_oAppDrvr = foRider

        p_oClientID = New Client(foRider)
        p_oClientID.Parent = "LRCarSerialRegistration"

        p_oCoCltID1 = New Client(foRider)
        p_oCoCltID1.Parent = "LRCarSerialRegistration"

        p_oCoCltID2 = New Client(foRider)
        p_oCoCltID2.Parent = "LRCarSerialRegistration"

        p_oRegCltID = New Client(foRider)
        p_oRegCltID.Parent = "LRCarSerialRegistration"

        p_oRgCltID1 = New Client(foRider)
        p_oRgCltID1.Parent = "LRCarSerialRegistration"

        p_oRgCltID2 = New Client(foRider)
        p_oRgCltID2.Parent = "LRCarSerialRegistration"

        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oAppDrvr.BranchCode

        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 16 - 21
                getClient(fnIndex, fsValue, False)
        End Select
    End Sub
End Class