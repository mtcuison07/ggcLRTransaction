'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Master Object
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

Public Class LRMaster
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String
    Private p_sBranchCD As String

    Private p_oClient As ggcClient.Client

    Private Const p_sMasTable As String = "LR_Master"
    Private Const p_sMsgHeadr As String = "LR Master"

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
        Set(value As String)
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
                    Case 80 ' sClientNm
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case 82 ' sTownIDxx 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sTownIDxx
                    Case 83 'sCollatrl
                        If Trim(IFNull(p_oDTMstr(0).Item(37))) <> "" And Trim(p_oOthersx.sCollatrl) = "" Then
                            getCollateral(37, 83, p_oDTMstr(0).Item(37), True, False)
                        End If
                        Return p_oOthersx.sCollatrl
                    Case 84
                        If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                            Return p_oDTMstr(0).Item("nPrincipl") - _
                                  (p_oDTMstr(0).Item("nIntTotal") + _
                                   p_oDTMstr(0).Item("nSrvcChrg") + _
                                   p_oDTMstr(0).Item("nInsChrge") + _
                                   p_oDTMstr(0).Item("nOthChrg1") + _
                                   p_oDTMstr(0).Item("nOthChrg2") + _
                                   p_oDTMstr(0).Item("nOthChrg3"))
                        Else
                            Return 0.0
                        End If
                    Case 85 ' sCompnyNm
                        If Trim(IFNull(p_oDTMstr(0).Item(34))) <> "" And Trim(p_oOthersx.sCompnyNm) = "" Then
                            getCompany(34, 85, p_oDTMstr(0).Item(34), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case 86 ' sBranchNm
                        If Trim(IFNull(p_oDTMstr(0).Item(1))) <> "" And Trim(p_oOthersx.sBranchNm) = "" Then
                            getBranch(1, 86, p_oDTMstr(0).Item(1), True, False)
                        End If
                        Return p_oOthersx.sBranchNm
                    Case 87 ' sRouteNme 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sRouteNme
                    Case 88 ' sCollName
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCollName
                    Case 89 ' sCoBranch
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCoBranch
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
                    Case 82 ' sTownIDxx
                    Case 83 ' sCollatrl
                        getCollateral(37, 83, value, False, False)
                    Case 84 ' nTakeHome
                    Case 85 ' sCompnyNm
                        getCompany(34, 85, value, False, False)
                    Case 86 ' sBranchNm
                    Case 87 ' sRouteNme
                        getRoute(4, 87, value, False, False)
                    Case 88 ' sCollName
                    Case 89 ' sCoBranch

                    Case 9, 36 ' nInterest, nIntTotal
                    Case 8, 10 To 15 'Principal and charges
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = Convert.ToDecimal(value)
                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))

                        'If principal/interest rate/term then compute for amortization
                        If (Index = 8 Or Index = 17 Or Index = 15) And p_oDTMstr(0).Item("nAcctTerm") > 0 Then
                            'Set the principal as the initial balance
                            p_oDTMstr(0).Item("nABalance") = p_oDTMstr(0).Item("nPrincipl")
                            'Compute for the monthly amortization
                            p_oDTMstr(0).Item("nMonAmort") = Math.Round(p_oDTMstr(0).Item("nPrincipl") / p_oDTMstr(0).Item("nAcctTerm"), 2)
                            RaiseEvent MasterRetrieved(19, p_oDTMstr(0).Item(19))

                            p_oDTMstr(0).Item("nInterest") = p_oDTMstr(0).Item("nPrincipl") * p_oDTMstr(0).Item("nAcctTerm") * p_oDTMstr(0).Item("nIntRatex") / 100
                            RaiseEvent MasterRetrieved(9, p_oDTMstr(0).Item(9))
                        End If

                        If p_oDTMstr(0).Item("sCollatID") = "" Then
                            p_oDTMstr(0).Item("nIntTotal") = p_oDTMstr(0).Item("nInterest")
                        Else
                            p_oDTMstr(0).Item("nIntTotal") = 0
                        End If
                        RaiseEvent MasterRetrieved(36, p_oDTMstr(0).Item("nIntTotal"))

                    Case 7, 16, 22, 32
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = Convert.ToDateTime(value)
                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))

                        If Index = 16 Then
                            p_oDTMstr(0).Item("dDueDatex") = DateAdd(DateInterval.Month, p_oDTMstr(0).Item("nAcctTerm") - 1, p_oDTMstr(0).Item("dFirstPay"))
                            RaiseEvent MasterRetrieved(18, p_oDTMstr(0).Item(18))
                        End If

                        If Index = 7 Then
                            p_oDTMstr(0).Item("dLastPaym") = p_oDTMstr(0).Item("dTransact")
                        End If

                    Case 17
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = Convert.ToSingle(value)
                        End If

                        p_oDTMstr(0).Item("nInterest") = p_oDTMstr(0).Item("nPrincipl") * p_oDTMstr(0).Item("nAcctTerm") * p_oDTMstr(0).Item("nIntRatex") / 100

                        p_oDTMstr(0).Item("dDueDatex") = DateAdd(DateInterval.Month, p_oDTMstr(0).Item("nAcctTerm") - 1, p_oDTMstr(0).Item("dFirstPay"))
                        p_oDTMstr(0).Item("nMonAmort") = Math.Round(p_oDTMstr(0).Item("nPrincipl") / p_oDTMstr(0).Item("nAcctTerm"), 2)
                        RaiseEvent MasterRetrieved(19, p_oDTMstr(0).Item(19))
                        RaiseEvent MasterRetrieved(9, p_oDTMstr(0).Item(9))
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                        RaiseEvent MasterRetrieved(18, p_oDTMstr(0).Item(18))

                        If p_oDTMstr(0).Item("sCollatID") = "" Then
                            p_oDTMstr(0).Item("nIntTotal") = p_oDTMstr(0).Item("nInterest")
                        Else
                            p_oDTMstr(0).Item("nIntTotal") = 0
                        End If
                        RaiseEvent MasterRetrieved(36, p_oDTMstr(0).Item("nIntTotal"))

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
                    Case "stownidxx" ' 82 
                        If Trim(IFNull(p_oDTMstr(0).Item(2))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getClient(2, 80, p_oDTMstr(0).Item(2), True, False)
                        End If
                        Return p_oOthersx.sTownIDxx
                    Case "scollatrl" '83
                        If Trim(IFNull(p_oDTMstr(0).Item(37))) <> "" And Trim(p_oOthersx.sCollatrl) = "" Then
                            getCollateral(37, 83, p_oDTMstr(0).Item(37), True, False)
                        End If
                        Return p_oOthersx.sCollatrl
                    Case "ntakehome" ' 84
                        If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                            Return p_oDTMstr(0).Item("nPrincipl") - _
                                  (p_oDTMstr(0).Item("nIntTotal") + _
                                   p_oDTMstr(0).Item("nSrvcChrg") + _
                                   p_oDTMstr(0).Item("nInsChrge") + _
                                   p_oDTMstr(0).Item("nOthChrg1") + _
                                   p_oDTMstr(0).Item("nOthChrg2") + _
                                   p_oDTMstr(0).Item("nOthChrg3"))
                        Else
                            Return 0.0
                        End If
                    Case "scompnynm" ' 85
                        If Trim(IFNull(p_oDTMstr(0).Item(34))) <> "" And Trim(p_oOthersx.sCompnyNm) = "" Then
                            getCompany(34, 86, p_oDTMstr(0).Item(34), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case "sbranchnm" ' 86
                        If Trim(IFNull(p_oDTMstr(0).Item(1))) <> "" And Trim(p_oOthersx.sBranchNm) = "" Then
                            getBranch(1, 86, p_oDTMstr(0).Item(1), True, False)
                        End If
                        Return p_oOthersx.sBranchNm
                    Case "sroutenme" ' 87 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sRouteNme
                    Case "scollname" ' 88  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCollName
                    Case "scobranch" ' 89
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sRouteNme) = "" Then
                            getRoute(4, 87, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCoBranch

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
                    Case "sclientnm" '80
                        getClient(2, 80, value, False, False)
                    Case "saddressx" '81
                    Case "stownidxx" '82
                    Case "scollatrl" '83
                        getCollateral(37, 83, value, False, False)
                    Case "ntakehome" '84
                    Case "scompnynm" '85
                        getCompany(34, 85, value, False, False)
                    Case "sbranchnm" '86
                    Case "sroutenme" '87
                        getRoute(4, 87, value, False, False)
                    Case "scollname" '88  
                    Case "scobranch" '89

                    Case "ninterest", "ninttotal"
                    Case "nprincipl", "nsrvcchrg", "ninschrge", "nothchrg1", "nothchrg2", "nothchrg3", "nintratex"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "dtransact", "dfirstpay", "dlastpaym", "dclosedxx"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "nacctterm"
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

        lsSQL = AddCondition(getSQ_Master, "a.sAcctNmbr = " & strParm(fsTransNox))
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
            Return OpenTransaction(p_oDTMstr(0).Item("sAcctNmbr"))
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
                If fsValue = p_oDTMstr(0).Item("sAcctNmbr") Then Return True
            Else
                If fsValue = p_oOthersx.sClientNm Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQ_Browse, "a.cAcctstat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = getSQ_Browse()
        End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sAcctNmbr LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "b.sCompnyNm like " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sAcctNmbr»sClientNm»dTransact" _
                                        , "Acct No»Client»Date", _
                                        , "a.sAcctNmbr»b.sCompnyNm»a.dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenTransaction(loDta.Item("sAcctNmbr"))
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

        'If route is not assigned which is most probably true call getRoute
        If p_oDTMstr(0).Item("sRouteIDx") = "" Then
            Call getRoute(4, 83, "", False, False)
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
                p_oDTMstr(0).Item("sAcctNmbr") = GetNextCode(p_sMasTable, "sAcctNmbr", True, p_oApp.Connection, True, "L" & Mid(p_sBranchCD, 2))
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
            Case 83 ' sCollatrl
                getCollateral(37, 83, fsValue, False, True)
            Case 85 ' sCompnyNm
                getCompany(34, 85, fsValue, False, True)
            Case 87 ' sRouteNme
                getRoute(4, 87, fsValue, False, True)
        End Select
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "sacctnmbr"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sacctnmbr", True, p_oApp.Connection, True, "L" & Mid(p_sBranchCD, 2))
                Case "dtransact"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "dfirstpay"
                    p_oDTMstr(0).Item(lnCtr) = DateAdd(DateInterval.Month, 1, p_oApp.SysDate)
                Case "dmodified", "smodified", "dduedatex", "dlastpaym", "dclosedxx"
                Case "cratingxx", "cacctstat"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "cactivexx"
                    p_oDTMstr(0).Item(lnCtr) = "1"
                Case "nprincipl", "ninterest", "nsrvcchrg", "ninschrge", "nothchrg1", "nothchrg2", "nothchrg3", "nmonamort" _
                   , "npaymtotl", "npentotlx", "ndebttotl", "ncredtotl", "namtduexx", "nabalance", "nlastpaym"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case "nintratex", "npenltyrt", "ndelayavg", "ninttotal"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case "nacctterm"
                    p_oDTMstr(0).Item(lnCtr) = 1
                Case "nledgerno"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next

        p_oDTMstr(0).Item("dDueDatex") = DateAdd(DateInterval.Month, p_oDTMstr(0).Item("nAcctTerm") - 1, p_oDTMstr(0).Item("dFirstPay"))
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
        p_oOthersx.sRouteNme = ""
        p_oOthersx.sTownIDxx = ""
        p_oOthersx.nTakeHome = 0
        p_oOthersx.sCollatrl = ""
        p_oOthersx.sCompnyNm = ""
        p_oOthersx.sCollName = ""
        p_oOthersx.sCoBranch = ""
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
        If Val(p_oDTMstr(0).Item("nPrincipl")) <= 1000 Then
            MsgBox("Loan Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check when will be the exptected released of this loan
        If p_oDTMstr(0).Item("dFirstPay") < p_oDTMstr(0).Item("dTransact") Then
            MsgBox("Expected first pay date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
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
        loClient.Parent = "LRMaster"

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
    Private Sub getRoute(ByVal fnColIdx As Integer _
                       , ByVal fnColDsc As Integer _
                       , ByVal fsValue As String _
                       , ByVal fbIsCode As Boolean _
                       , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And Trim(fsValue) <> "" And p_oOthersx.sRouteNme <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sRouteNme And Trim(fsValue) <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT DISTINCT" & _
                       "  a.sRouteIDx" & _
                       ", a.sRouteNme" & _
                       ", c.sBranchNm" & _
                       ", d.sCompnyNm" & _
               " FROM Route_Area a" & _
                 " LEFT JOIN Route_Area_Town b ON a.sRouteIDx = b.sRouteIDx" & _
                 " LEFT JOIN Branch c ON a.sBranchCd = c.sBranchCd" & _
                 " LEFT JOIN Client_Master d ON a.sCollctID = d.sClientID"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sRouteIDx»sRouteNme" _
                                             , "ID»Route", _
                                             , "a.sRouteIDx»a.sRouteNme" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sRouteNme = ""
                p_oOthersx.sCoBranch = ""
                p_oOthersx.sCollName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sRouteIDx")
                p_oOthersx.sRouteNme = loRow.Item("sRouteNme")
                p_oOthersx.sCoBranch = loRow.Item("sBranchNm")
                p_oOthersx.sCollName = loRow.Item("sCompnyNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sRouteNme)
            Exit Sub

        End If

        If fsValue = "" Then
            lsSQL = AddCondition(lsSQL, "b.sTownIDxx = " & strParm(p_oOthersx.sTownIDxx))
        Else
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sRouteIDx = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sRouteNme = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sRouteNme = ""
            p_oOthersx.sCoBranch = ""
            p_oOthersx.sCollName = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sRouteIDx")
            p_oOthersx.sRouteNme = loDta(0).Item("sRouteNme")
            p_oOthersx.sCoBranch = loDta(0).Item("sBranchNm")
            p_oOthersx.sCollName = loDta(0).Item("sCompnyNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sRouteNme)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getCollateral(ByVal fnColIdx As Integer _
                            , ByVal fnColDsc As Integer _
                            , ByVal fsValue As String _
                            , ByVal fbIsCode As Boolean _
                            , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sCollatrl <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sCollatrl And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sCollatID" & _
                       ", a.sDescript" & _
               " FROM Collateral a" & _
               IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sCollatID»sDescript" _
                                             , "ID»Collateral", _
                                             , "a.sCollatID»a.sDescript" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sCollatrl = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sCollatID")
                p_oOthersx.sCollatrl = loRow.Item("sDescript")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCollatrl)

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                If p_oDTMstr(0).Item("sCollatID") = "" Then
                    p_oDTMstr(0).Item("nIntTotal") = p_oDTMstr(0).Item("nInterest")
                Else
                    p_oDTMstr(0).Item("nIntTotal") = 0
                End If
                RaiseEvent MasterRetrieved(36, p_oDTMstr(0).Item("nIntTotal"))
            End If

            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sCollatID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sDescript = " & strParm(fsValue))
            End If
        Else
            lsSQL = AddCondition(lsSQL, "a.sDescript = " & strParm(fsValue))
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sCollatrl = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sCollatID")
            p_oOthersx.sCollatrl = loDta(0).Item("sDescript")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCollatrl)

        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            If p_oDTMstr(0).Item("sCollatID") = "" Then
                p_oDTMstr(0).Item("nIntTotal") = p_oDTMstr(0).Item("nInterest")
            Else
                p_oDTMstr(0).Item("nIntTotal") = 0
            End If
        End If

        RaiseEvent MasterRetrieved(36, p_oDTMstr(0).Item("nIntTotal"))

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
        Return "SELECT a.sAcctNmbr" & _
                    ", a.sBranchCd" & _
                    ", a.sClientID" & _
                    ", a.sMCActNox" & _
                    ", a.sRouteIDx" & _
                    ", a.sRemarksx" & _
                    ", a.sExAcctNo" & _
                    ", a.dTransact" & _
                    ", a.nPrincipl" & _
                    ", a.nInterest" & _
                    ", a.nSrvcChrg" & _
                    ", a.nInsChrge" & _
                    ", a.nOthChrg1" & _
                    ", a.nOthChrg2" & _
                    ", a.nOthChrg3" & _
                    ", a.nIntRatex" & _
                    ", a.dFirstPay" & _
                    ", a.nAcctTerm" & _
                    ", a.dDueDatex" & _
                    ", a.nMonAmort" & _
                    ", a.nPenltyRt" & _
                    ", a.nLastPaym" & _
                    ", a.dLastPaym" & _
                    ", a.nPaymTotl" & _
                    ", a.nPenTotlx" & _
                    ", a.nDebtTotl" & _
                    ", a.nCredTotl" & _
                    ", a.nAmtDuexx" & _
                    ", a.nABalance" & _
                    ", a.nDelayAvg" & _
                    ", a.cRatingxx" & _
                    ", a.cAcctstat" & _
                    ", a.dClosedxx" & _
                    ", a.cActivexx" & _
                    ", a.sCompnyID" & _
                    ", a.nLedgerNo" & _
                    ", a.nIntTotal" & _
                    ", a.sCollatID" & _
                    ", a.sCollNote" & _
                    ", a.sApplicNo" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sAcctNmbr" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
                    ", a.nPrincipl" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
              " WHERE a.sClientID = b.sClientID"
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
        Public sCollatrl As String
        Public nTakeHome As Decimal
        Public sCompnyNm As String
        Public sBranchNm As String
        Public sRouteNme As String
        Public sCollName As String
        Public sCoBranch As String
    End Class
End Class
