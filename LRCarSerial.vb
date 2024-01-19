'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Serial Object
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

Public Class LRCarSerial

#Region "Constant"
    Private Const xsSignature As String = "08220326"
    Private Const pxeMODULENAME As String = "LRSerial"
    Private Const pxeMasterTble As String = "Car_Serial"
#End Region

#Region "Protected Members"
    Protected p_oAppDrvr As GRider
    Protected p_oDTMaster As DataTable
    Protected p_nEditMode As xeEditMode
    Protected p_oSC As New MySqlCommand
    Protected p_oDT As DataTable

    Protected p_sBranchCd As String
    Protected p_bCancelled As Boolean
    Protected p_bHasParent As Boolean
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

    WriteOnly Property HasParent() As Boolean
        Set(Value As Boolean)
            p_bHasParent = Value
        End Set
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
                    Case "csoldstat" : Index = 14
                    Case "clocation" : Index = 15
                    Case "sclientnm" : Index = 16
                    Case "scocltnm1" : Index = 17
                    Case "scocltnm2" : Index = 18
                    Case "sregcltNm" : Index = 19
                    Case "srgcltnm1" : Index = 20
                    Case "srgcltnm2" : Index = 21
                    Case "sstockidx" : Index = 22
                    Case "sbranchcd" : Index = 23
                    Case "sbrandcde" : Index = 24
                    Case "smodelcde" : Index = 25
                    Case "scolorcde" : Index = 26
                    Case "sclientid" : Index = 27
                    Case "scocltid1" : Index = 28
                    Case "scocltid2" : Index = 29
                    Case "sregcltid" : Index = 30
                    Case "srgcltid1" : Index = 31
                    Case "srgcltid2" : Index = 32
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
                        getBrand(Value, True, False)
                    Case "smodelnme" : Index = 4
                        getModel(Value, True, False)
                    Case "scolornme" : Index = 5
                        getColor(Value, True, False)
                    Case "nyearmodl" : Index = 6
                    Case "sfilenoxx" : Index = 7
                    Case "screnoxxx" : Index = 8
                    Case "scrnoxxxx" : Index = 9
                    Case "splatenop" : Index = 10
                    Case "sregornox" : Index = 11
                    Case "sstickrno" : Index = 12
                    Case "dregister" : Index = 13
                    Case "csoldstat" : Index = 14
                    Case "clocation" : Index = 15
                    Case "sclientnm" : Index = 16
                        getClient(Index, Value, False, False)
                    Case "scocltnm1" : Index = 17
                        getClient(Index, Value, False, False)
                    Case "sCocltnm2" : Index = 18
                        getClient(Index, Value, False, False)
                    Case "sregcltnm" : Index = 19
                        getClient(Index, Value, False, False)
                    Case "srgcltnm1" : Index = 20
                        getClient(Index, Value, False, False)
                    Case "srgcltnm2" : Index = 21
                        getClient(Index, Value, False, False)
                    Case "sstockidx" : Index = 22
                    Case "sbranchcd" : Index = 23
                    Case "sbrandcde" : Index = 24
                    Case "smodelcde" : Index = 25
                    Case "scolorcde" : Index = 26
                    Case "sclientid" : Index = 27
                    Case "scocltid1" : Index = 28
                    Case "scocltid2" : Index = 29
                    Case "sregcltid" : Index = 30
                    Case "srgcltid1" : Index = 31
                    Case "srgcltid2" : Index = 32
                    Case Else
                        MsgBox("Invalid Field Detected!!!", MsgBoxStyle.Critical, "WARNING")
                End Select
            End If
            p_oDTMaster(0)(Index) = Value
        End Set
    End Property
#End Region

#Region "Private Function"
    Private Function isUserActive(ByRef loDT As DataTable) As Boolean
        Dim lnCtr As Integer = 0
        Dim lbMember As Boolean = False

        If loDT.Rows(0).Item("cUserType").Equals(0) Then
            For lnCtr = 0 To loDT.Rows.Count - 1
                If loDT.Rows(0).Item("sProdctID").Equals(p_oAppDrvr.ProductID) Then
                    Exit For
                    lbMember = True
                End If
            Next
        Else
            lbMember = True
        End If

        If Not lbMember Then
            MsgBox("User is not a member of this application!!!" & vbCrLf & _
               "Application used is not allowed!!!", vbCritical, "Warning")
        End If

        ' check user status
        If loDT.Rows(0).Item("cUserStat").Equals(xeUserStatus.SUSPENDED) Then
            MsgBox("User is currently suspended!!!" & vbCrLf & _
                     "Application used is not allowed!!!", vbCritical, "Warning")
            Return False
        End If
        Return True
    End Function

    Private Function getSQL_User() As String
        Return "SELECT sUserIDxx" & _
              ", sLogNamex" & _
              ", sPassword" & _
              ", sUserName" & _
              ", nUserLevl" & _
              ", cUserType" & _
              ", sProdctID" & _
              ", cUserStat" & _
              ", nSysError" & _
              ", cLogStatx" & _
              ", cLockStat" & _
              ", cAllwLock" & _
           " FROM xxxSysUser" & _
           " WHERE sLogNamex = ?sLogNamex" & _
              " AND sPassword = ?sPassword"
    End Function

    Private Function getSQL_Brand() As String
        Return "SELECT sBrandIDx" & _
                    ", sBrandNme" & _
                " FROM Car_Brand" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Model() As String
        Return "SELECT a.sModelIDx" & _
                    ", a.sModelNme" & _
                    ", b.sBrandNme" & _
                    ", b.sBrandIDx" & _
                " FROM Car_Model a" & _
                    ", Car_Brand b" & _
                " WHERE a.sBrandIDx = b.sBrandIDx" & _
                    " AND a.cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Color() As String
        Return "SELECT sColorIDx" & _
                    ", sColorNme" & _
                " FROM Color" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Browse() As String
        Return "SELECT a.sSerialID" & _
                    ", a.sEngineNo" & _
                    ", a.sFrameNox" & _
                    ", b.sBrandNme" & _
                    ", c.sModelNme" & _
                    ", d.sColorNme" & _
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
                If Not p_bHasParent Then p_oAppDrvr.BeginTransaction()

                If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                    lsSQL = "INSERT INTO " & pxeMasterTble & " SET" & _
                                "  sSerialID = " & strParm(.Rows(0)("sSerialID")) & _
                                ", sEngineNo = " & strParm(.Rows(0)("sEngineNo")) & _
                                ", sFrameNox = " & strParm(.Rows(0)("sFrameNox")) & _
                                ", nYearModl = " & CInt(IFNull(.Rows(0)("nYearModl"), Year(Now))) & _
                                ", sFileNoxx = " & strParm(IFNull(.Rows(0)("sFileNoxx"), "")) & _
                                ", sCRENoxxx = " & strParm(IFNull(.Rows(0)("sCRENoxxx"), "")) & _
                                ", sCRNoxxxx = " & strParm(IFNull(.Rows(0)("sCRNoxxxx"), "")) & _
                                ", sPlateNoP = " & strParm(IFNull(.Rows(0)("sPlateNoP"), "")) & _
                                ", sRegORNox = " & strParm(IFNull(.Rows(0)("sRegORNox"), "")) & _
                                ", sStickrNo = " & strParm(IFNull(.Rows(0)("sStickrNo"), "")) & _
                                ", dRegister = " & dateParm(IFNull(.Rows(0)("dRegister"), "")) & _
                                ", cSoldStat = " & strParm(IFNull(.Rows(0)("cSoldStat"), "")) & _
                                ", cLocation = " & strParm(IFNull(.Rows(0)("cLocation"), "")) & _
                                ", sStockIDx = " & strParm(IFNull(.Rows(0)("sStockIDx"), "")) & _
                                ", sBranchCD = " & strParm(IFNull(.Rows(0)("sBranchCD"), "")) & _
                                ", sBrandCde = " & strParm(IFNull(.Rows(0)("sBrandCde"), "")) & _
                                ", sModelCde = " & strParm(IFNull(.Rows(0)("sModelCde"), "")) & _
                                ", sColorCde = " & strParm(IFNull(.Rows(0)("sColorCde"), "")) & _
                                ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                                ", sCoCltID1 = " & strParm(.Rows(0)("sCoCltID1")) & _
                                ", sCoCltID2 = " & strParm(.Rows(0)("sCoCltID2")) & _
                                ", sRegCltID = " & strParm(.Rows(0)("sRegCltID")) & _
                                ", sRgCltID1 = " & strParm(.Rows(0)("sRgCltID1")) & _
                                ", sRgCltID2 = " & strParm(.Rows(0)("sRgCltID2")) & _
                                ", sModified = " & strParm(p_oAppDrvr.UserID) & _
                                ", dModified = " & dateParm(p_oAppDrvr.SysDate)

                Else
                    lsSQL = "UPDATE " & pxeMasterTble & " SET" & _
                                "  sEngineNo = " & strParm(.Rows(0)("sEngineNo")) & _
                                ", sFrameNox = " & strParm(.Rows(0)("sFrameNox")) & _
                                ", nYearModl = " & CInt(IFNull(.Rows(0)("nYearModl"), Year(Now))) & _
                                ", sFileNoxx = " & strParm(IFNull(.Rows(0)("sFileNoxx"), "")) & _
                                ", sCRENoxxx = " & strParm(IFNull(.Rows(0)("sCRENoxxx"), "")) & _
                                ", sCRNoxxxx = " & strParm(IFNull(.Rows(0)("sCRNoxxxx"), "")) & _
                                ", sPlateNoP = " & strParm(IFNull(.Rows(0)("sPlateNoP"), "")) & _
                                ", sRegORNox = " & strParm(IFNull(.Rows(0)("sRegORNox"), "")) & _
                                ", sStickrNo = " & strParm(IFNull(.Rows(0)("sStickrNo"), "")) & _
                                ", dRegister = " & dateParm(IFNull(.Rows(0)("dRegister"), "")) & _
                               ", cSoldStat = " & strParm(IFNull(.Rows(0)("cSoldStat"), "")) & _
                                ", cLocation = " & strParm(IFNull(.Rows(0)("cLocation"), "")) & _
                                ", sStockIDx = " & strParm(IFNull(.Rows(0)("sStockIDx"), "")) & _
                                ", sBranchCD = " & strParm(IFNull(.Rows(0)("sBranchCD"), "")) & _
                                ", sBrandCde = " & strParm(IFNull(.Rows(0)("sBrandCde"), "")) & _
                                ", sModelCde = " & strParm(IFNull(.Rows(0)("sModelCde"), "")) & _
                                ", sColorCde = " & strParm(IFNull(.Rows(0)("sColorCde"), "")) & _
                                ", sClientID = " & strParm(.Rows(0)("sClientID")) & _
                                ", sCoCltID1 = " & strParm(.Rows(0)("sCoCltID1")) & _
                                ", sCoCltID2 = " & strParm(.Rows(0)("sCoCltID2")) & _
                                ", sRegCltID = " & strParm(.Rows(0)("sRegCltID")) & _
                                ", sRgCltID1 = " & strParm(.Rows(0)("sRgCltID1")) & _
                                ", sRgCltID2 = " & strParm(.Rows(0)("sRgCltID2")) & _
                            " WHERE sSerialID = " & strParm(.Rows(0)("sSerialID"))
                End If

                lnRow = p_oAppDrvr.Execute(lsSQL, pxeMasterTble)
                If lnRow <= 0 Then
                    MsgBox("Unable to Save Transaction!!!" & vbCrLf & _
                            "Please contact GGC SSG/SEG for assistance!!!", MsgBoxStyle.Critical, "WARNING")
                    If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                    Return False
                End If

                If Not p_bHasParent Then
                    If .Rows(0)("sClientID") <> "" Then
                        If Not p_oClientID.SaveClient Then
                            MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                            If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                            Return False
                        End If
                    End If
                End If

                If .Rows(0)("sCoCltID1") <> "" Then
                    If Not p_oCoCltID1.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sCoCltID2") <> "" Then
                    If Not p_oCoCltID2.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sRegCltID") <> "" Then
                    If Not p_oRegCltID.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sRgCltID1") <> "" Then
                    If Not p_oRgCltID1.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                If .Rows(0)("sRgCltID2") <> "" Then
                    If Not p_oRgCltID2.SaveClient Then
                        MsgBox("Unable to save client info!", vbOKOnly, pxeMODULENAME)
                        If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()
                        Return False
                    End If
                End If

                Dim loSerialReg As LRCarSerialRegistration
                loSerialReg = New LRCarSerialRegistration(p_oAppDrvr)
                'loSerialReg.OpenTransaction(p_oDTMaster.Rows(0)("sSerialID"))

                With loSerialReg

                End With

                If Not p_bHasParent Then p_oAppDrvr.CommitTransaction()

                Return True
            Catch ex As Exception
                If Not p_bHasParent Then p_oAppDrvr.RollBackTransaction()

                MsgBox(ex.Message)

                Return False
            End Try
        End With
    End Function

    Public Function SearchTransaction( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMaster(0).Item("sSerialID") Then Return True
            Else
                If fsValue = p_oDTMaster(0).Item("sEngineNo") Then Return True
            End If
        End If

        lsSQL = getSQL_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sSerialID = " & strParm(fsValue)
        Else
            lsFilter = "a.sEngineNo LIKE " & strParm("%" & fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oAppDrvr _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sSerialID»sEngineNo»sFrameNox»sBrandNme»sModelNme»sColorNme" _
                                        , "SerialID»EngineNo»FrameNo»Brand»Model»Color", _
                                        , "a.sSerialID»a.sEngineNo»a.sFrameNox»b.sBrandNme»c.sModelNme»d.sColorNme" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sSerialID"))
        End If
    End Function

    Function OpenTransaction(ByVal fsSerialID As String) As Boolean
        Dim loDT As New DataTable
        Dim lsSQL As String

        lsSQL = AddCondition("SELECT * FROM " & pxeMasterTble, "sSerialID = " & strParm(fsSerialID))

        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Return False

        Call createMasterTable()
        If loDT.Rows.Count > 0 Then
            With p_oDTMaster
                .Rows.Add()
                For nCtr As Integer = 0 To .Columns.Count - 1
                    Select Case .Columns.Item(nCtr).ColumnName
                        Case "sBrandNme"
                            getBrand(loDT.Rows(0)("sBrandCde"), False, True)
                        Case "sModelNme"
                            getModel(loDT.Rows(0)("sModelCde"), False, True)
                        Case "sColorNme"
                            getColor(loDT.Rows(0)("sColorCde"), False, True)
                        Case "sClientNm"
                            getClient(16, loDT.Rows(0)("sClientID"), True, False)
                        Case "sCoCltNm1"
                            getClient(17, loDT.Rows(0)("sCoCltID1"), True, False)
                        Case "sCoCltNm2"
                            getClient(18, loDT.Rows(0)("sCoCltID2"), True, False)
                        Case "sRegCltNm"
                            getClient(19, loDT.Rows(0)("sRegCltID"), True, False)
                        Case "sRgCltNm1"
                            getClient(20, loDT.Rows(0)("sRgCltID1"), True, False)
                        Case "sRgCltNm2"
                            getClient(21, loDT.Rows(0)("sRgCltID2"), True, False)
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

    Function NewTransaction() As Boolean
        Call createMasterTable()
        p_oDTMaster.Rows.Add(p_oDTMaster.NewRow())

        Call initMaster()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Function UpdateTransaction() As Boolean
        If Not p_nEditMode = xeEditMode.MODE_READY Then
            If Not OpenTransaction(p_oDTMaster.Rows(0)("sSerialID")) Then Return False
        End If

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
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
            .Columns.Add("cSoldStat", GetType(Char))
            .Columns.Add("cLocation", GetType(Char))
            .Columns.Add("sClientNm", GetType(String)).MaxLength = 128
            .Columns.Add("sCoCltNm1", GetType(String)).MaxLength = 128
            .Columns.Add("sCoCltNm2", GetType(String)).MaxLength = 128
            .Columns.Add("sRegCltNm", GetType(String)).MaxLength = 128
            .Columns.Add("sRgCltNm1", GetType(String)).MaxLength = 128
            .Columns.Add("sRgCltNm2", GetType(String)).MaxLength = 128
            .Columns.Add("sStockIDx", GetType(String)).MaxLength = 12
            .Columns.Add("sBranchCD", GetType(String)).MaxLength = 4
            .Columns.Add("sBrandCde", GetType(String)).MaxLength = 9
            .Columns.Add("sModelCde", GetType(String)).MaxLength = 9
            .Columns.Add("sColorCde", GetType(String)).MaxLength = 7
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
            .Rows(0)("sSerialID") = GetNextCode(pxeMasterTble, "sSerialID", True, p_oAppDrvr.Connection, True, p_sBranchCd)
            .Rows(0)("sEngineNo") = ""
            .Rows(0)("sFrameNox") = ""
            .Rows(0)("sBrandNme") = ""
            .Rows(0)("sModelNme") = ""
            .Rows(0)("sColorNme") = ""
            .Rows(0)("nYearModl") = Year(Now)
            .Rows(0)("sFileNoxx") = ""
            .Rows(0)("sCRENoxxx") = ""
            .Rows(0)("sCRNoxxxx") = ""
            .Rows(0)("sPlateNoP") = ""
            .Rows(0)("sRegORNox") = ""
            .Rows(0)("sStickrNo") = ""
            .Rows(0)("dRegister") = p_oAppDrvr.SysDate
            .Rows(0)("cSoldStat") = "0"
            .Rows(0)("cLocation") = "1"
            .Rows(0)("sClientNm") = ""
            .Rows(0)("sCoCltNm1") = ""
            .Rows(0)("sCoCltNm2") = ""
            .Rows(0)("sRegCltNm") = ""
            .Rows(0)("sRgCltNm1") = ""
            .Rows(0)("sRgCltNm2") = ""
            .Rows(0)("sStockIDx") = ""
            .Rows(0)("sBranchCD") = p_sBranchCd
            .Rows(0)("sBrandCde") = ""
            .Rows(0)("sModelCde") = ""
            .Rows(0)("sColorCde") = ""
            .Rows(0)("sClientID") = ""
            .Rows(0)("sCoCltID1") = ""
            .Rows(0)("sCoCltID2") = ""
            .Rows(0)("sRegCltID") = ""
            .Rows(0)("sRgCltID1") = ""
            .Rows(0)("sRgCltID2") = ""
        End With
    End Sub

    Private Function getBrand(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getBrand"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If Not bSearch Then
                    lsCondition = "sBrandNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sBrandNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sBrandIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Brand, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMaster
                .Rows(0)("sBrandCde") = loDT(0)("sBrandIDx")
                .Rows(0)("sBrandNme") = loDT(0)("sBrandNme")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sBrandIDx»sBrandNme", _
                                "Brand ID»Brand", _
                                "", _
                                "", _
                                3)

            If Not IsNothing(loDataRow) Then
                With p_oDTMaster
                    .Rows(0)("sBrandCde") = loDataRow("sBrandIDx")
                    .Rows(0)("sBrandNme") = loDataRow("sBrandNme")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMaster
            RaiseEvent MasterRetrieved(3, .Rows(0)("sBrandNme"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMaster
            .Rows(0)("sBrandCde") = ""
            .Rows(0)("sBrandNme") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getModel(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Model, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMaster
                .Rows(0)("sModelCde") = loDT(0)("sModelIDx")
                .Rows(0)("sModelNme") = loDT(0)("sModelNme")
                .Rows(0)("sBrandCde") = loDT(0)("sBrandIDx")
                .Rows(0)("sBrandNme") = loDT(0)("sBrandNme")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sModelIDx»sModelNme»sBrandNme", _
                                "Model ID»Model»Brand", _
                                "", _
                                "a.sModelIDx»a.sModelNme»b.sBrandNme", _
                                2)

            If Not IsNothing(loDataRow) Then
                With p_oDTMaster
                    .Rows(0)("sModelCde") = loDataRow("sModelIDx")
                    .Rows(0)("sModelNme") = loDataRow("sModelNme")
                    .Rows(0)("sBrandCde") = loDataRow("sBrandIDx")
                    .Rows(0)("sBrandNme") = loDataRow("sBrandNme")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMaster
            RaiseEvent MasterRetrieved(3, .Rows(0)("sBrandNme"))
            RaiseEvent MasterRetrieved(4, .Rows(0)("sModelNme"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMaster
            .Rows(0)("sBrandCde") = ""
            .Rows(0)("sBrandNme") = ""
            .Rows(0)("sModelCde") = ""
            .Rows(0)("sModelNme") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getColor(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As Boolean
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getColor"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sColorNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sColorNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sColorIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Color, lsCondition)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oAppDrvr.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            With p_oDTMaster
                .Rows(0)("sColorCde") = loDT(0)("sColorIDx")
                .Rows(0)("sColorNme") = loDT(0)("sColorNme")
            End With
        Else
            loDataRow = KwikSearch(p_oAppDrvr, _
                                lsSQL, _
                                "", _
                                "sColorIDx»sColorNme", _
                                "Color ID»Color", _
                                "", _
                                "", _
                                5)

            If Not IsNothing(loDataRow) Then
                With p_oDTMaster
                    .Rows(0)("sColorCde") = loDataRow("sColorIDx")
                    .Rows(0)("sColorNme") = loDataRow("sColorNme")
                End With
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        With p_oDTMaster
            RaiseEvent MasterRetrieved(5, .Rows(0)("sColorNme"))
        End With

        Return True
        Exit Function
endWithClear:
        With p_oDTMaster
            .Rows(0)("sColorCde") = ""
            .Rows(0)("sColorNme") = ""
        End With
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fnIndex As Integer _
                            , ByVal fsValue As String _
                            , ByVal fbIsCode As Boolean _
                            , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = IFNull(p_oDTMaster(0).Item("sClientID"), "") And fsValue <> "" And IFNull(p_oDTMaster(0).Item("sClientNm"), "") <> "" Then Exit Sub
        Else
            'Do not allow searching of value if fsValue is empty
            If (fsValue = p_oDTMaster(0).Item(CInt(fnIndex)) And fsValue <> "") Or fsValue = "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oAppDrvr)
        loClient.Parent = "LRCarSerial"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                Select Case fnIndex
                    Case 16
                        p_oClientID = loClient
                        p_oDTMaster(0).Item("sClientID") = p_oClientID.Master("sClientID")
                        p_oDTMaster(0).Item("sClientNm") = p_oClientID.Master("sLastName") & ", " & _
                                                           p_oClientID.Master("sFrstName") & _
                                                           IIf(p_oClientID.Master("sSuffixNm") = "", "", " " & p_oClientID.Master("sSuffixNm")) & " " & _
                                                           p_oClientID.Master("sMiddName")
                    Case 17
                        p_oCoCltID1 = loClient
                        p_oDTMaster(0).Item("sCoCltID1") = p_oCoCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm1") = p_oCoCltID1.Master("sLastName") & ", " & _
                                                           p_oCoCltID1.Master("sFrstName") & _
                                                           IIf(p_oCoCltID1.Master("sSuffixNm") = "", "", " " & p_oCoCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID1.Master("sMiddName")
                    Case 18
                        p_oCoCltID2 = loClient
                        p_oDTMaster(0).Item("sCoCltID2") = p_oCoCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm2") = p_oCoCltID2.Master("sLastName") & ", " & _
                                                           p_oCoCltID2.Master("sFrstName") & _
                                                           IIf(p_oCoCltID2.Master("sSuffixNm") = "", "", " " & p_oCoCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID2.Master("sMiddName")
                    Case 19
                        p_oRegCltID = loClient
                        p_oDTMaster(0).Item("sRegCltID") = p_oRegCltID.Master("sClientID")
                        p_oDTMaster(0).Item("sRegCltNm") = p_oRegCltID.Master("sLastName") & ", " & _
                                                           p_oRegCltID.Master("sFrstName") & _
                                                           IIf(p_oRegCltID.Master("sSuffixNm") = "", "", " " & p_oRegCltID.Master("sSuffixNm")) & " " & _
                                                           p_oRegCltID.Master("sMiddName")
                    Case 20
                        p_oRgCltID1 = loClient
                        p_oDTMaster(0).Item("sRgCltID1") = p_oRgCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm1") = p_oRgCltID1.Master("sLastName") & ", " & _
                                                           p_oRgCltID1.Master("sFrstName") & _
                                                           IIf(p_oRgCltID1.Master("sSuffixNm") = "", "", " " & p_oRgCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID1.Master("sMiddName")
                    Case 21
                        p_oRgCltID2 = loClient
                        p_oDTMaster(0).Item("sRgCltID2") = p_oRgCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm2") = p_oRgCltID2.Master("sLastName") & ", " & _
                                                           p_oRgCltID2.Master("sFrstName") & _
                                                           IIf(p_oRgCltID2.Master("sSuffixNm") = "", "", " " & p_oRgCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID2.Master("sMiddName")
                End Select


            Else
                Select Case fnIndex
                    Case 16
                        p_oDTMaster(0).Item("sClientID") = ""
                        p_oDTMaster(0).Item("sClientNm") = ""
                    Case 17
                        p_oDTMaster(0).Item("sCoCltID1") = ""
                        p_oDTMaster(0).Item("sCoCltNm1") = ""
                    Case 18
                        p_oDTMaster(0).Item("sCoCltID2") = ""
                        p_oDTMaster(0).Item("sCoCltNm1") = ""
                    Case 19
                        p_oDTMaster(0).Item("sRegCltID") = ""
                        p_oDTMaster(0).Item("sRegCltNm") = ""
                    Case 20
                        p_oDTMaster(0).Item("sRgCltID1") = ""
                        p_oDTMaster(0).Item("sRgCltNm2") = ""
                    Case 21
                        p_oDTMaster(0).Item("sRgCltID1") = ""
                        p_oDTMaster(0).Item("sRgCltNm2") = ""
                End Select
            End If

            RaiseEvent MasterRetrieved(fnIndex, p_oDTMaster.Rows(0)(fnIndex))
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                Select Case fnIndex
                    Case 16
                        p_oClientID = loClient
                        p_oDTMaster(0).Item("sClientID") = p_oClientID.Master("sClientID")
                        p_oDTMaster(0).Item("sClientNm") = p_oClientID.Master("sLastName") & ", " & _
                                                           p_oClientID.Master("sFrstName") & _
                                                           IIf(p_oClientID.Master("sSuffixNm") = "", "", " " & p_oClientID.Master("sSuffixNm")) & " " & _
                                                           p_oClientID.Master("sMiddName")
                    Case 17
                        p_oCoCltID1 = loClient
                        p_oDTMaster(0).Item("sCoCltID1") = p_oCoCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm1") = p_oCoCltID1.Master("sLastName") & ", " & _
                                                           p_oCoCltID1.Master("sFrstName") & _
                                                           IIf(p_oCoCltID1.Master("sSuffixNm") = "", "", " " & p_oCoCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID1.Master("sMiddName")
                    Case 18
                        p_oCoCltID2 = loClient
                        p_oDTMaster(0).Item("sCoCltID2") = p_oCoCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sCoCltNm2") = p_oCoCltID2.Master("sLastName") & ", " & _
                                                           p_oCoCltID2.Master("sFrstName") & _
                                                           IIf(p_oCoCltID2.Master("sSuffixNm") = "", "", " " & p_oCoCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oCoCltID2.Master("sMiddName")
                    Case 19
                        p_oRegCltID = loClient
                        p_oDTMaster(0).Item("sRegCltID") = p_oRegCltID.Master("sClientID")
                        p_oDTMaster(0).Item("sRegCltNm") = p_oRegCltID.Master("sLastName") & ", " & _
                                                           p_oRegCltID.Master("sFrstName") & _
                                                           IIf(p_oRegCltID.Master("sSuffixNm") = "", "", " " & p_oRegCltID.Master("sSuffixNm")) & " " & _
                                                           p_oRegCltID.Master("sMiddName")
                    Case 20
                        p_oRgCltID1 = loClient
                        p_oDTMaster(0).Item("sRgCltID1") = p_oRgCltID1.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm1") = p_oRgCltID1.Master("sLastName") & ", " & _
                                                           p_oRgCltID1.Master("sFrstName") & _
                                                           IIf(p_oRgCltID1.Master("sSuffixNm") = "", "", " " & p_oRgCltID1.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID1.Master("sMiddName")
                    Case 21
                        p_oRgCltID2 = loClient
                        p_oDTMaster(0).Item("sRgCltID2") = p_oRgCltID2.Master("sClientID")
                        p_oDTMaster(0).Item("sRgCltNm2") = p_oRgCltID2.Master("sLastName") & ", " & _
                                                           p_oRgCltID2.Master("sFrstName") & _
                                                           IIf(p_oRgCltID2.Master("sSuffixNm") = "", "", " " & p_oRgCltID2.Master("sSuffixNm")) & " " & _
                                                           p_oRgCltID2.Master("sMiddName")
                End Select
            End If
        End If

        RaiseEvent MasterRetrieved(fnIndex, p_oDTMaster.Rows(0)(CInt(fnIndex)))
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub New(ByVal foRider As GRider)
        p_oAppDrvr = foRider

        p_oClientID = New Client(foRider)
        p_oClientID.Parent = "LRSerial"

        p_oCoCltID1 = New Client(foRider)
        p_oCoCltID1.Parent = "LRSerial"

        p_oCoCltID2 = New Client(foRider)
        p_oCoCltID2.Parent = "LRSerial"

        p_oRegCltID = New Client(foRider)
        p_oRegCltID.Parent = "LRSerial"

        p_oRgCltID1 = New Client(foRider)
        p_oRgCltID1.Parent = "LRSerial"

        p_oRgCltID2 = New Client(foRider)
        p_oRgCltID2.Parent = "LRSerial"

        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oAppDrvr.BranchCode

        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 3
                getBrand(fsValue, False, False)
            Case 4
                getModel(fsValue, True, False)
            Case 5
                getColor(fsValue, True, False)
            Case 16 To 21
                getClient(fnIndex, fsValue, False, False)
        End Select
    End Sub

    'ALTER TABLE `GGC_ISysDBF`.`Car_Serial` ADD COLUMN `sRegORNox` VARCHAR(15) CHARSET latin1 COLLATE latin1_swedish_ci NULL AFTER `dRegister`, ADD COLUMN `sStickrNo` VARCHAR(8) CHARSET latin1 COLLATE latin1_swedish_ci NULL AFTER `sRegORNox`;
    'ALTER TABLE `GGC_ISysDBF`.`Car_Serial_Registration` DROP PRIMARY KEY, ADD PRIMARY KEY (`sSerialID`, `dRegister`);
End Class