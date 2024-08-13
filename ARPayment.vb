'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     LR Payment Object
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
'  Kalyptus [ 07/08/2016 02:40 pm ]
'      Started creating this object.
'  Concerns:
'      1. Printing 
'      Transaction Closing
'      GCard Online/Offline Points
'      Cancellation 
'      1. Check Transaction
'      1. Check Clearing

'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports ggcClient
Imports System.Drawing
Imports Newtonsoft.Json.Linq

Public Class ARPayment
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String
    Private p_oPaidBy As ggcClient.Client

    Private p_cTranType As String

    Private Const p_sSystemCd As String = "ARPy"
    Private Const p_sMasTable As String = "LR_Payment_Master"
    Private Const p_sMsgHeadr As String = "LR Payment"
    Private Const p_sSourceCd As String = "M02910000004" 'G_Card_Points_Basis code for Monthly Payment

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
                    'getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                    Case 80 ' sClientNm
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If

                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case 82 ' nPNValuex
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nPNValuex
                    Case 83 ' nDownPaym
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nDownPaym
                    Case 84 ' nGrossPrc
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nGrossPrc
                    Case 85 ' nMonAmort
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case 86 ' nCashBalx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nCashBalx
                    Case 87 ' nAcctTerm
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case 88 ' nABalance
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case 89 ' nAmtDuexx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case 90 ' xRebatesx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.xRebatesx
                    Case 91 ' sEngineNo
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sEngineNo
                    Case 92 ' sFrameNox
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sFrameNox
                    Case 93 ' sModelNme
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sModelNme
                    Case 94 ' sColorNme
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sColorNme

                    Case 95 ' sCompnyNm 
                        Return p_oOthersx.sCompnyNm
                    Case 96 ' sCompnyID 
                        Return p_oOthersx.sCompnyID

                    Case 97 ' xPaidByxx
                        If Trim(IFNull(p_oDTMstr(0).Item(6))) <> "" And Trim(p_oOthersx.xPaidByxx) = "" Then
                            getPaidBy(6, 97, p_oDTMstr(0).Item(6), True, False)
                        End If
                        Return p_oOthersx.xPaidByxx

                    Case 98 ' sCollName
                        If Trim(IFNull(p_oDTMstr(0).Item(12))) <> "" And Trim(p_oOthersx.sCollName) = "" Then
                            getCollector(12, 98, p_oDTMstr(0).Item(12), True, False)
                        End If
                        Return p_oOthersx.sCollName
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
                        getAccount(4, 80, value, False, False)
                    Case 81 To 94
                    Case 95 To 96
                    Case 97 ' xPaidByxx
                        getPaidBy(6, 97, value, False, False)
                    Case 98 ' sCollName
                        getCollector(12, 98, value, False, False)
                    Case 1
                        If IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 3 'sReferNox
                        p_oDTMstr(0).Item(Index) = ""

                        If isValidReceipt(value) Then
                            p_oDTMstr(0).Item(Index) = CStr(value)
                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 8 To 11
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
                    Case "sclientnm" '80  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" '81  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case "npnvaluex" '82 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nPNValuex
                    Case "ndownpaym" '83 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nDownPaym
                    Case "ngrossprc" '84  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nGrossPrc
                    Case "nmonamort" '85 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case "ncashbalx" '86  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nCashBalx
                    Case "nacctterm" '87  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case "nabalance" '88  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case "namtduexx" '89  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case "xrebatesx" '90  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.xRebatesx
                    Case "sengineno" '91  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sEngineNo
                    Case "sframenox" '92 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sFrameNox
                    Case "smodelnme" '93  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sModelNme
                    Case "scolornme" '94  
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sColorNme

                    Case "scompnynm" '95  
                        Return p_oOthersx.sCompnyNm
                    Case "scompnyid" '96  
                        Return p_oOthersx.sCompnyID

                    Case "xpaidbyxx" '97  
                        If Trim(IFNull(p_oDTMstr(0).Item(6))) <> "" And Trim(p_oOthersx.xPaidByxx) = "" Then
                            getPaidBy(6, 97, p_oDTMstr(0).Item(6), True, False)
                        End If
                        Return p_oOthersx.xPaidByxx

                    Case "scollname" '98  
                        If Trim(IFNull(p_oDTMstr(0).Item(12))) <> "" And Trim(p_oOthersx.sCollName) = "" Then
                            getCollector(12, 90, p_oDTMstr(0).Item(12), True, False)
                        End If
                        Return p_oOthersx.sCollName
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
                        getAccount(4, 80, value, False, False)
                    Case "xpaidbyxx" '97  
                        getPaidBy(6, 97, value, False, False)
                    Case "scollname" '98  
                        getCollector(12, 98, value, False, False)

                    Case "sclientnm", "saddressx", "npnvaluex", "ndownpaym", "ngrossprc", "nmonamort", "ncashbalx", "nacctterm", "nabalance", _
                         "namtduexx", "xrebatesx", "sengineno", "sframenox", "smodelnme", "scolornme"

                    Case "dtransact"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = getValidDate(p_oApp, value)
                    Case "srefernox"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "namountxx", "nintamtxx", "nrebatesx", "npenaltyx"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    Public Property CheckInfo(ByVal Index As String)
        Get
            Select Case LCase(Index)
                Case "schecknox"
                    Return p_oOthersx.sCheckNox
                Case "sacctnoxx"
                    Return p_oOthersx.sAcctNoxx
                Case "sbankidxx"
                    Return p_oOthersx.sBankIDxx
                Case "sbankname"
                    Return p_oOthersx.sBankName
                Case "scheckdte"
                    Return p_oOthersx.sCheckDte
                Case "ncheckamt"
                    Return p_oOthersx.nCheckAmt
                Case Else
                    Return ""
            End Select
        End Get
        Set(value)
            Select Case LCase(Index)
                Case "schecknox"
                    p_oOthersx.sCheckNox = value
                Case "sacctnoxx"
                    p_oOthersx.sAcctNoxx = value
                Case "sbankidxx"
                    p_oOthersx.sBankIDxx = value
                Case "sbankname"
                    p_oOthersx.sBankName = value
                Case "scheckdte"
                    p_oOthersx.sCheckDte = value
                Case "ncheckamt"
                    p_oOthersx.nCheckAmt = value
            End Select
        End Set
    End Property

    'EPayInfo
    Public Property EPayInfo(ByVal Index As String)
        Get
            Select Case LCase(Index)
                Case "sclientid"
                    Return p_oOthersx.sEPCltIDx
                Case "sclientnm"
                    Return p_oOthersx.sEPCltNme
                Case "srefernox"
                    Return p_oOthersx.sEPReferN
                Case "namtpaidx"
                    Return p_oOthersx.nEPAmount
                Case "stermcode"
                    Return p_oOthersx.sEPTermID
                Case "stermname"
                    Return p_oOthersx.sEPTermNm
                Case "sremarksx"
                    Return p_oOthersx.sEPRemrks
                Case Else
                    Return ""
            End Select
        End Get
        Set(ByVal value)
            Select Case LCase(Index)
                Case "sclientid"
                    p_oOthersx.sEPCltIDx = value
                Case "sclientnm"
                    p_oOthersx.sEPCltNme = value
                Case "srefernox"
                    p_oOthersx.sEPReferN = value
                Case "namtpaidx"
                    p_oOthersx.nEPAmount = value
                Case "stermcode"
                    p_oOthersx.sEPTermID = value
                Case "stermname"
                    p_oOthersx.sEPTermNm = value
                Case "sremarksx"
                    p_oOthersx.sEPRemrks = value
            End Select
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
                If fsValue = p_oDTMstr(0).Item("sReferNox") Then Return True
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
        Debug.Print(lsSQL)
        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sReferNox LIKE " & strParm(fsValue)
        Else
            lsFilter = "b.sCompnyNm LIKE " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sReferNox»sClientNm»dTransact»sTransNox" _
                                        , "Refer No»Client»Date»Trans No", _
                                        , "a.sReferNox»b.sCompnyNm»a.dTransact»a.sTransNox" _
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

        If p_cTranType = "2" Then
            If Not isRebateOk() Then Return False
        End If

        Dim lsSQL As String = ""

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            'Save master table 
            'Note: Update is not allowed!!!
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                If Trim(p_oOthersx.sCheckNox) <> "" Then

                    MsgBox("This payment is using a check! Please use the PR Module...", MsgBoxStyle.Critical, "Payment Validation")
                    Return False

                    'Dim lsSQx As String
                    'lsSQx = "INSERT INTO Check_Payments" & _
                    '       " SET sTransNox = " & strParm(GetNextCode("Check_Payments", "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)) & _
                    '          ", dTransact = " & dateParm(p_oDTMstr(0).Item("dTransact")) & _
                    '          ", sBankIDxx = " & strParm(p_oOthersx.sBankIDxx) & _
                    '          ", sBranchxx = " & strParm(p_oOthersx.sBranchXX) & _
                    '          ", sCheckNox = " & strParm(p_oOthersx.sCheckNox) & _
                    '          ", dCheckDte = " & dateParm(p_oOthersx.sCheckDte) & _
                    '          ", sPayorIDx = " & strParm(p_oDTMstr(0).Item("sClientID")) & _
                    '          ", sPayeeIDx = " & strParm(p_oOthersx.sCompnyID) & _
                    '          ", sAcctCode = " & strParm("") & _
                    '          ", sReferNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                    '          ", nAmountxx = " & p_oDTMstr(0).Item("nAmountxx") & _
                    '          ", sRemarksx = " & strParm("ARPy»" & p_oOthersx.sBankAcct & "»" & p_oOthersx.sClientNm) & _
                    '          ", nClearDay = " & 0 & _
                    '          ", cTranStat = " & strParm("0") & _
                    '          ", sModified = " & strParm(p_oApp.UserID) & _
                    '          ", dModified = " & dateParm(p_oApp.getSysDate)
                    'p_oApp.Execute(lsSQx, "Check_Payments")
                    'p_oDTMstr(0).Item("cPaymForm") = xeLogical.YES
                End If
            End If

            p_oDTMstr(0).Item("cTranType") = p_cTranType
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)

            If lsSQL <> "" Then
                'mac 2020.11.19
                '   added validation, rollback changes if rows affected is <= 0
                If p_oApp.Execute(lsSQL, p_sMasTable) <= 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()

                    Return False
                End If
            End If

            'check if it is an e-payment 
            If p_oOthersx.sEPCltIDx <> "" And p_oOthersx.nEPAmount > 0.0# Then
                lsSQL = "INSERT INTO Other_Payment_Received SET" & _
                        "  sTransNox = " & strParm(GetNextCode("Other_Payment_Received", "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)) & _
                        ", sClientID = " & strParm(p_oOthersx.sEPCltIDx) & _
                        ", sReferNox = " & strParm(p_oOthersx.sEPReferN) & _
                        ", nTotlAmnt = " & Master("namountxx") + Master("nintamtxx") + Master("nPenaltyX") & _
                        ", nAmtPaidx = " & p_oOthersx.nEPAmount & _
                        ", sTermCode = " & strParm(p_oOthersx.sEPTermID) & _
                        ", sRemarksx = " & strParm(p_oOthersx.sEPRemrks) & _
                        ", sSourceCd = " & strParm(p_sSystemCd) & _
                        ", sSourceNo = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                        ", cTranStat = '1'" & _
                        ", dModified = " & dateParm(p_oApp.SysDate)

                If p_oApp.Execute(lsSQL, "Other_Payment_Received", p_sBranchCd) <= 0 Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    MsgBox("Unable to save AR Payment.", MsgBoxStyle.Critical, "Warning")
                    Return False
                End If
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            p_nEditMode = xeEditMode.MODE_READY

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

        If p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_POSTED) Then
            MsgBox("Request was already posted!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED) Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cGCrdPstd") = "1" Then
            MsgBox("GCard point was already posted! Please void the GCard transaction before continuing...", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED)
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

            'mac 2020.11.19
            '   added validation, rollback changes if rows affected is <= 0
            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                If p_sParent = "" Then p_oApp.RollBackTransaction()

                Return False
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()

            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function PrintTrans() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPrintedx") = xeLogical.YES Then
            MsgBox("Receipt was already printed!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'If p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_OPEN) Then
        '    If Not PostTransaction() Then
        '        MsgBox("Payment cannot be posted. Please inform MIS/SEG for assistance!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
        '        Return False
        '    End If
        'End If

        If p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED) Then
            MsgBox("Receipt was already CANCELLED!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        p_oDTMstr(0).Item("cPrintedx") = xeLogical.YES

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim lsSQL As String
            lsSQL = "UPDATE " & p_sMasTable & _
                   " SET cPrintedx = " & strParm(xeLogical.YES) & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))

            p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

            If p_sParent = "" Then p_oApp.CommitTransaction()
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            MsgBox(ex.Message)

            Return False
        End Try
        

        Dim loPrint As ggcLRReports.clsDirectPrintSF
        loPrint = New ggcLRReports.clsDirectPrintSF
        loPrint.PrintFont = New Font("Arial", 9)
        loPrint.PrintBegin()

        Dim lnTotlSale As Decimal = p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nPenaltyx")
        Dim lnVatSales As Decimal = lnTotlSale / 1.12
        Dim lnLessVatx As Decimal = lnVatSales * 0.12

        'Total Sales(VAT Inclusive)
        'loPrint.Print(5.8, 2.0, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(6.5, 2.3, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)

        'Print transaction Date
        'loPrint.Print(6.5, 6.0, Format(p_oDTMstr(0).Item("dTransact"), "MMM dd, yyyy"))
        loPrint.Print(7.5, 6.5, Format(p_oDTMstr(0).Item("dTransact"), "MMM dd, yyyy"))

        'Less VAT
        'loPrint.Print(7, 2.0, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(8, 2.3, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)

        'Total
        'loPrint.Print(8.4, 2.0, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(9.4, 2.3, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)

        'Print Name
        'loPrint.Print(8.4, 3.2, Master(80) & " / " & p_oDTMstr(0).Item("sAcctNmbr"))
        loPrint.Print(9.4, 3.4, Master(80) & " / " & p_oDTMstr(0).Item("sAcctNmbr"))

        'Print Address
        'loPrint.Print(10, 3.2, Master(81))
        loPrint.Print(11, 3.5, Master(81))

        'Amount in words
        'loPrint.Print(14, 3.2, AmountInWords(lnTotlSale))
        loPrint.Print(15, 3.5, AmountInWords(lnTotlSale))

        'Amount in number
        'loPrint.Print(15.5, 6.7, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(16.5, 7.4, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)

        'Model/Color/Engine No
        'loPrint.Print(17.3, 3.7, "Model/Color: " & p_oOthersx.sModelNme & " / " & p_oOthersx.sColorNme & " / " & "Engine No.: " & p_oOthersx.sEngineNo)
        loPrint.Print(19, 4.3, "Model/Color: " & p_oOthersx.sModelNme & " / " & p_oOthersx.sColorNme & " / " & "Engine No.: " & p_oOthersx.sEngineNo)

        'VATable Sales
        'loPrint.Print(17.7, 2.0, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(20, 2.3, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)

        'VAT Exempt
        'loPrint.Print(19.2, 2.0, "0.00", StringAlignment.Far)
        loPrint.Print(21.5, 2.3, "0.00", StringAlignment.Far)

        'zero rated Sales
        'loPrint.Print(20.7, 2.0, "0.00", StringAlignment.Far)
        loPrint.Print(23, 2.3, "0.00", StringAlignment.Far)

        'VAT amount
        'loPrint.Print(22, 2.0, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(24.4, 2.3, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)

        'Total Sales 
        'loPrint.Print(23.5, 2.0, Format(lnVatSales + lnLessVatx, "#,##0.00"), StringAlignment.Far)
        loPrint.Print(25.2, 2.3, Format(lnVatSales + lnLessVatx, "#,##0.00"), StringAlignment.Far)


        'Principal
        'loPrint.Print(20, 4.6, Format(p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx"), "#,##0.00"), StringAlignment.Far)
        loPrint.Print(21.1, 5.2, Format(p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx"), "#,##0.00"), StringAlignment.Far)

        'Penalty
        loPrint.Print(23.5, 7.5, Format(p_oDTMstr(0).Item("nPenaltyx"), "#,##0.00"), StringAlignment.Far) '+3

        'Rebate
        loPrint.Print(24.9, 7.5, Format(p_oDTMstr(0).Item("nRebatesx"), "#,##0.00"), StringAlignment.Far) '+1.4

        ''Print transaction Date
        'loPrint.Print(9, 2.9, Format(p_oDTMstr(0).Item("dTransact"), "MMM dd, yyyy"))

        ''Print Name
        'loPrint.Print(10.5, 0.9, Master(80) & " / " & p_oDTMstr(0).Item("sAcctNmbr"))
        'loPrint.Print(11.5, 0.9, IIf(p_oDTMstr(0).Item("sPaidByID") <> "", "PAID BY: " & p_oOthersx.xPaidByxx, ""))

        ''Print Address
        'loPrint.Print(12.5, 0.9, Master(81))

        ''Model/Color
        'loPrint.Print(20, 1.2, "Model/Color: " & p_oOthersx.sModelNme & " / " & p_oOthersx.sColorNme)
        ''Engine No
        'loPrint.Print(21, 1.2, "Engine No.: " & p_oOthersx.sEngineNo)

        ''Principal
        'loPrint.Print(18, 3.55, Format(p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx"), "#,##0.00"), StringAlignment.Far)
        ''Penalty
        'loPrint.Print(28.5, 3.55, Format(p_oDTMstr(0).Item("nPenaltyx"), "#,##0.00"), StringAlignment.Far)
        ''Rebate
        'loPrint.Print(30, 3.55, Format(p_oDTMstr(0).Item("nRebatesx"), "#,##0.00"), StringAlignment.Far)

        'Dim lnTotlSale As Decimal = p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nPenaltyx")
        'Dim lnVatSales As Decimal = lnTotlSale / 1.12
        'Dim lnLessVatx As Decimal = lnVatSales * 0.12

        ''Total Sales(VAT Inclusive)
        'loPrint.Print(32.5, 3.55, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)
        ''Less VAT
        'loPrint.Print(34, 3.55, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        ''Total
        'loPrint.Print(35.5, 3.55, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)

        loPrint.PrintEnd()

        Return True
    End Function

    'mac 2020-07-24
    Private Function OnlineEntry() As Boolean
        'run command
        Dim lsVal = p_oApp.ProductID & " " & p_oApp.UserID & " " & p_oOthersx.sGCardNox & " " & p_oOthersx.cDigitalx & " " & p_oDTMstr(0)("sReferNox") & " " & p_sSourceCd & " " & Format(p_oDTMstr(0)("dTransact"), "yyyy-MM-dd")
        RMJExecute("D:\GGC_Java_Systems\", "gcard-online-points-entry.bat", lsVal)

        Return True
    End Function

    'mac 2024-02.22
    Private Function SendTDS() As Boolean
        Dim lsResultPath As String = "D:\GGC_Java_Systems\temp\res.tmp"

        'delete result file
        DeleteFile(lsResultPath)

        'run command
        RMJExecute("D:\GGC_Java_Systems\", "gcard-online-points-entry.bat", p_oApp.ProductID & " " & p_oApp.UserID & " " & p_oOthersx.sGCardNox & " " & p_oOthersx.cDigitalx & " " & p_oDTMstr(0)("sReferNox") & " " & p_sSourceCd)

        'get response from result file
        Dim lsResponse As String = ReadFile(lsResultPath)
        'delete result file
        DeleteFile(lsResultPath)

        If lsResponse = "" Then
            MsgBox("No response from JAVA API.", MsgBoxStyle.Exclamation, "Warning")
            Return False
        End If

        'extract data
        Dim loJSON As JObject = JObject.Parse(lsResponse)

        If loJSON.GetValue("result") = "success" Then
            MsgBox(loJSON.GetValue("message"), MsgBoxStyle.Information, "Notice")
        Else
            MsgBox(CStr(loJSON.GetValue("message")), MsgBoxStyle.Exclamation, "Notice")
            Return False
        End If

        Return True
    End Function

    'mac 2020-07-20
    Private Function OfflineEntry() As Boolean
        Dim lsResultPath As String = "D:\GGC_Java_Systems\temp\res.tmp"

        'delete result file
        DeleteFile(lsResultPath)

        'run command
        RMJExecute("D:\GGC_Java_Systems\", "gcard-offline-points-inquiry.bat", p_oApp.ProductID & " " & p_oApp.UserID & " " & p_oOthersx.sGCardNox & " " & p_oDTMstr(0)("sReferNox") & " " & Format(p_oDTMstr(0)("dTransact"), "yyyy-MM-dd"))

        'get response from result file
        Dim lsResponse As String = ReadFile(lsResultPath)
        'delete result file
        DeleteFile(lsResultPath)

        If lsResponse = "" Then
            MsgBox("No response from JAVA API.", MsgBoxStyle.Exclamation, "Warning")
            Return False
        End If

        'extract data
        Dim loJSON As JObject = JObject.Parse(lsResponse)

        If loJSON.GetValue("result") = "success" Then
            If CDbl(loJSON.GetValue("nPointsxx")) > 0.0 Then
                Dim lsTransNo As String
                Dim lsSQL As String

                lsTransNo = GetNextCode("G_Card_Detail_Offline", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)

                lsSQL = "INSERT INTO G_Card_Detail_Offline SET" & _
                             "  sTransNox = " & strParm(lsTransNo) & _
                             ", sGCardNox = " & strParm(CStr(loJSON.GetValue("sGCardNox"))) & _
                             ", sCompnyID = ''" & _
                             ", dTransact = " & strParm(CStr(loJSON.GetValue("dTransact"))) & _
                             ", sSourceNo = " & strParm(CStr(loJSON.GetValue("sSourceNo"))) & _
                             ", sSourceCd = " & strParm(CStr(loJSON.GetValue("sSourceCd"))) & _
                             ", nTranAmtx = " & CDbl(loJSON.GetValue("nTranAmtx")) & _
                             ", nPointsxx = " & CDbl(loJSON.GetValue("nPointsxx")) & _
                             ", sRemarksx = 'Payment to payment center partners.'" & _
                             ", sOTPasswd = ''" & _
                             ", cTranStat = '0'" & _
                             ", sPostedxx = ''" & _
                             ", sModified = " & strParm(p_oApp.UserID) & _
                             ", dModified = " & datetimeParm(p_oApp.SysDate)

                If p_oApp.Execute(lsSQL, "G_Card_Detail_Offline") <= 0 Then GoTo endWithRoll

                lsSQL = "UPDATE G_Card_Detail_Offline SET" & _
                             "  cTranStat = '1'" & _
                             ", sPostedxx = " & strParm(p_oApp.UserID) & _
                             ", dPostedxx = " & datetimeParm(p_oApp.SysDate) & _
                        " WHERE sTransNox = " & strParm(lsTransNo)

                If p_oApp.Execute(lsSQL, "G_Card_Detail_Offline") <= 0 Then GoTo endWithRoll

                lsSQL = "UPDATE G_Card_Master SET" & _
                            " nTotPoint = nTotPoint + " & CDbl(loJSON.GetValue("nPointsxx")) & _
                        " WHERE sGCardNox = " & strParm(CStr(loJSON.GetValue("sGCardNox")))

                If p_oApp.Execute(lsSQL, "G_Card_Master") <= 0 Then GoTo endWithRoll

                lsSQL = "UPDATE LR_Payment_Master SET" & _
                            " cGCrdPstd = '1'" & _
                        " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox"))

                If p_oApp.Execute(lsSQL, "LR_Payment_Master") <= 0 Then GoTo endWithRoll

                Return True
            Else
                MsgBox("This transaction maybe encoded on OFFLINE/ONLINE POINTS ENTRY.", vbInformation, "Notice")
                Return False
            End If
        Else
            MsgBox(loJSON.GetValue("message"), MsgBoxStyle.Exclamation, "Warning")
            Return False
        End If

endWithRoll:
        MsgBox("Unable to update table for OFFLINE ENTRY.", vbExclamation, "Warning")
        Return False
    End Function

    'Public Function PostTransaction()
    Public Function PostTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_POSTED) Then
            MsgBox("Application was already posted!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED) Then
            MsgBox("Application was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'kalyptus - 2017.03.10 03:53pm
        'Check if there are unposted payment for this account...
        Dim lsSQL As String
        lsSQL = "SELECT sTransNox" & _
               " FROM " & p_sMasTable & _
               " WHERE sTransNox <> " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                 " AND sAcctNmbr = " & strParm(p_oDTMstr(0).Item("sAcctNmbr")) & _
                 " AND dTransact < " & dateParm(p_oDTMstr(0).Item("dTransact")) & _
                 " AND cPostedxx = '0'" & _
               " UNION" & _
               " SELECT sTransNox" & _
               " FROM LR_Payment_Master_PR" & _
               " WHERE sAcctNmbr = " & strParm(p_oDTMstr(0).Item("sAcctNmbr")) & _
                 " AND dTransact < " & dateParm(p_oDTMstr(0).Item("dTransact")) & _
                 " AND cPostedxx = '0'" & _
                 " AND cPaymForm = '0'"
        Debug.Print(lsSQL)
        'she 2017-03-27 2:52 pm 
        'Add date filter to check all unposted payment < than sa current na pinopost.
        '" AND dTransact < " & dateParm(p_oDTMstr(0).Item("dTransact"))
        Dim loDta As DataTable = p_oApp.ExecuteQuery(lsSQL)
        If loDta.Rows.Count > 0 Then
            MsgBox("There are unposted payment for this account!" & vbCrLf & _
                   "Please post the transaction first...", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If



        'mac 2020-03-09
        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim loTrans As ARTrans

            loTrans = New ARTrans(p_oApp)
            loTrans.Master("sAcctNmbr") = p_oDTMstr(0).Item("sAcctNmbr")
            loTrans.Master("dTransact") = p_oDTMstr(0).Item("dTransact")
            loTrans.Master("nTranAmtx") = p_oDTMstr(0).Item("nAmountxx")
            loTrans.Master("nRebatesx") = p_oDTMstr(0).Item("nRebatesx")
            loTrans.Master("nPenaltyx") = p_oDTMstr(0).Item("nPenaltyx")
            loTrans.Master("sRemarksx") = p_oDTMstr(0).Item("sRemarksx")
            loTrans.Master("sReferNox") = p_oDTMstr(0).Item("sReferNox")
            loTrans.Master("sCollIDxx") = p_oDTMstr(0).Item("sCollIDxx")

            Select Case p_oDTMstr(0).Item("cTranType")
                Case "2"
                    If Not loTrans.MonthlyPayment(p_oDTMstr(0).Item("sTransNox"), Trim(p_oDTMstr(0).Item("sCollIDxx")) = "") Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
                Case "3"
                    If Not loTrans.CashBalance(p_oDTMstr(0).Item("sTransNox"), Trim(p_oDTMstr(0).Item("sCollIDxx")) = "") Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
                Case "4"
                    If Not loTrans.DownPayment(p_oDTMstr(0).Item("sTransNox"), Trim(p_oDTMstr(0).Item("sCollIDxx")) = "") Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        Return False
                    End If
            End Select

            p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_POSTED)
            p_oDTMstr(0).Item("dPostedxx") = p_oApp.getSysDate

            lsSQL = "UPDATE " & p_sMasTable & _
                   " SET cPostedxx = " & strParm(CStr(xeTranStat.TRANS_POSTED)) & _
                      ", dPostedxx = " & dateParm(p_oDTMstr(0).Item("dPostedxx")) & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))

            'mac 2020.11.19
            '   added validation, rollback changes if rows affected is <= 0
            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                If p_sParent = "" Then p_oApp.RollBackTransaction()

                Return False
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            'mac 2024.04.10
            '   implementation of TDS
            If Not OnlineEntry() Then
                MsgBox("Unable to AUTO ENTRY POINTS." & vbCrLf & vbCrLf & "You may ENCODE the client's POINTS on GCARD SYSTEM.", MsgBoxStyle.Exclamation, "Notice")
            End If

            'mac 2020-07-23
            '   added auto encode of online and offline points
            'If p_oOthersx.sGCardNox <> "" And p_oOthersx.cDigitalx <> "" Then 'the card must be activated to use this feature
            '    'validate date and product id
            '    If Format(p_oApp.SysDate, "yyyy-MM-dd") = Format(p_oDTMstr(0)("dTransact"), "yyyy-MM-dd") And p_oApp.ProductID.ToLower = "integsys" Then
            '        If p_oDTMstr(0)("cGCrdPstd") = xeLogical.NO Then 'transaction must not be used on G-Card
            '            If OnlineEntry() Then
            '                If p_oOthersx.cDigitalx = "1" Then
            '                    If Not SendTDS() Then
            '                        MsgBox("Unable to AUTO UPLOAD TDS." & vbCrLf & vbCrLf & "Load GCARD SYSTEM and go to ONLINE POINTS ENTRY HISTORY to get the QR Code for customer's POINTS UPDATE.", MsgBoxStyle.Exclamation, "Notice")
            '                    End If
            '                End If
            '            Else
            '                MsgBox("Unable to AUTO ENTRY POINTS." & vbCrLf & vbCrLf & "You may ENCODE the client's POINTS on GCARD SYSTEM.", MsgBoxStyle.Exclamation, "Notice")
            '            End If
            '        End If
            '    Else 'else use OFFLINE ENTRY
            '        If p_oDTMstr(0)("cGCrdPstd") = xeLogical.NO Then 'transaction must not be used on G-Card
            '            If Not OfflineEntry() Then
            '                MsgBox("Unable to AUTO ENTRY POINTS." & vbCrLf & vbCrLf & "Please inform MIS Department immediately.", MsgBoxStyle.Exclamation, "Notice")
            '            End If
            '        End If
            '    End If
            'End If
            'end - mac 2020-07-23

            Return True
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & _
                   "Please inform MIS Department immediately.", , "Unable to AUTO ENTRY GCARD POINTS")
        End Try

        Return True
    End Function

    Public Sub SearchMaster(ByVal fnIndex As Integer, ByVal fsValue As String)
        Select Case fnIndex
            Case 4  ' sClientNm
                getAccount(4, 80, fsValue, True, True)
            Case 80 ' sClientNm
                getAccount(4, 80, fsValue, False, True)
            Case 97 ' xpaidbyxx  
                getPaidBy(6, 97, fsValue, False, True)
            Case 98 ' sCollName 
                getCollector(12, 98, fsValue, False, True)
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
                Case "cpostedxx", "cpaymform", "cprintedx", "cgcrdpstd"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "ctrantype"
                    p_oDTMstr(0).Item(lnCtr) = p_cTranType
                Case "namountxx", "nintamtxx", "nrebatesx", "npenaltyx"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub InitOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sAddressx = ""
        p_oOthersx.nPNValuex = 0.0
        p_oOthersx.nDownPaym = 0.0
        p_oOthersx.nGrossPrc = 0.0
        p_oOthersx.nMonAmort = 0.0
        p_oOthersx.nCashBalx = 0.0
        p_oOthersx.nAcctTerm = 0
        p_oOthersx.nABalance = 0.0
        p_oOthersx.nAmtDuexx = 0.0
        p_oOthersx.xRebatesx = 0.0
        p_oOthersx.sEngineNo = ""
        p_oOthersx.sFrameNox = ""
        p_oOthersx.sModelNme = ""
        p_oOthersx.sColorNme = ""

        p_oOthersx.sCompnyNm = "Northpoint Excelsior Credit Corp."
        p_oOthersx.sCompnyID = "M005"

        p_oOthersx.sCollName = ""

        'kalyptus - 2017.07.12 03:32pm
        'Change structure
        p_oOthersx.sCheckNox = ""
        p_oOthersx.sAcctNoxx = ""
        p_oOthersx.sBankIDxx = ""
        p_oOthersx.sBankName = ""
        p_oOthersx.sCheckDte = ""
        p_oOthersx.nCheckAmt = 0.0

        p_oOthersx.sEPCltIDx = ""
        p_oOthersx.sEPCltNme = ""
        p_oOthersx.sEPReferN = ""
        p_oOthersx.nEPAmount = 0.0#
        p_oOthersx.sEPTermID = ""
        p_oOthersx.sEPTermNm = ""
        p_oOthersx.sEPRemrks = ""
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

        'Check how much does he intends to pay
        If Val(p_oDTMstr(0).Item("nAmountxx")) < 0 Then
            MsgBox("Transaction Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much is the customer's rebate
        If Val(p_oDTMstr(0).Item("nRebatesx")) < 0 Then
            MsgBox("Rebate Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Val(p_oDTMstr(0).Item("nPenaltyx")) < 0 Then
            MsgBox("Penalty Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Trim(p_oDTMstr(0).Item("sReferNox")) = "" Then
            MsgBox("Document/Reference No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'If Bank Account has info then assume this payment is check and should look for the information of the check
        If Trim(p_oOthersx.sAcctNoxx) <> "" Then
            If Trim(p_oOthersx.sBankIDxx) = "" Then
                MsgBox("Bank Info seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            If Trim(p_oOthersx.sCheckNox) = "" Then
                MsgBox("Check No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            If Not IsDate(p_oOthersx.sCheckDte) Then
                MsgBox("Check Date seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If

            If p_oOthersx.nCheckAmt <= 0 Then
                MsgBox("Check Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
                Return False
            End If
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

    '2021-09-03
    'to validate the duplicate reference number
    Private Function isValidReceipt(ByVal fsValue As String) As Boolean
        Dim lsSQL As String

        lsSQL = "SELECT * " & _
            " FROM LR_Payment_Master" & _
            " WHERE sTransNox LIKE " & strParm(Left(p_oDTMstr(0).Item("sTransNox"), 6) + "%") & _
            " AND sReferNox = " & strParm(fsValue) & _
            " AND cPostedxx <> " & strParm(xeTranStat.TRANS_CANCELLED)

        Dim loRec As DataTable

        loRec = p_oApp.ExecuteQuery(lsSQL)

        If loRec.Rows.Count > 0 Then
            MsgBox("Duplicate Receipt Number Detected!!!" & vbCrLf & _
                    "Verify your entry then try again!", vbCritical, "Warning")
            isValidReceipt = False
        Else
            isValidReceipt = True
        End If

        loRec = Nothing

        Return isValidReceipt
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

        If fsValue = "" Then
            MsgBox("Please enter a value to search!", MsgBoxStyle.Information, "Notification")
            Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sAcctNmbr" & _
                       ", b.sCompnyNm sClientNm" & _
                       ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" & _
                       ", a.nPNValuex" & _
                       ", a.nDownPaym" & _
                       ", a.nGrossPrc" & _
                       ", a.nMonAmort" & _
                       ", a.nCashBalx" & _
                       ", a.nAcctTerm" & _
                       ", a.nABalance" & _
                       ", a.nAmtDuexx" & _
                       ", a.nRebatesx" & _
                       ", a.sClientID" & _
                       ", IFNULL(e.sEngineNo, '') sEngineNo" & _
                       ", IFNULL(e.sFrameNox, '') sFrameNox" & _
                       ", IFNULL(f.sModelNme, '') sModelNme" & _
                       ", IFNULL(g.sColorNme, '') sColorNme" & _
                       ", IFNULL(h.sGCardNox, '') sGCardNox" & _
                       ", IFNULL(i.cDigitalx, '') cDigitalx" & _
                       ", IFNULL(j.nAcctTerm, 0) nPromTerm" & _
                       ", IFNULL(j.nRebatesx, 0) nPromRebt" & _
               " FROM MC_AR_Master a" & _
                " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" & _
                " LEFT JOIN TownCity c ON b.sTownIDxx = c.sTownIDxx" & _
                " LEFT JOIN Province d ON c.sProvIDxx = d.sProvIDxx" & _
                " LEFT JOIN MC_Serial e ON a.sSerialID = e.sSerialID" & _
                " LEFT JOIN MC_Model f ON e.sModelIDx = f.sModelIDx" & _
                " LEFT JOIN Color g ON e.sColorIDx = g.sColorIDx" & _
                " LEFT JOIN MC_Serial_Service h ON h.sSerialID = a.sSerialID" & _
                " LEFT JOIN G_Card_Master i ON h.sGCardNox = i.sGCardNox AND i.cCardStat = '4'" & _
                " LEFT JOIN MC_AR_Rebate j ON a.sAcctNmbr = j.sAcctNmbr"

        'Salahin na agad ang mga account sa paghahanap pa lang ng Account Number para sa transaction 
        If p_cTranType = "2" Or p_cTranType = "4" Then
            'Monthly Payment
            lsSQL = AddCondition(lsSQL, "a.nAcctTerm > 0")
        ElseIf p_cTranType = "3" Then
            'Cash Balance
            lsSQL = AddCondition(lsSQL, "a.nAcctTerm = 0")
        End If

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loDT As DataTable
            Dim loRow As DataRow

            'mac 2020-07-18
            If fbIsCode Then
                loDT = p_oApp.ExecuteQuery(AddCondition(lsSQL, "a.sAcctNmbr = " & strParm(fsValue)))
            Else
                loDT = p_oApp.ExecuteQuery(AddCondition(lsSQL, "b.sCompnyNm LIKE " & strParm(fsValue & "%")))
            End If


            If loDT.Rows.Count = 1 Then
                'auto-load record when the result is 1
                loRow = loDT(0)
                loDT = Nothing
            Else
                'user kwiksearch for multiple result
                loDT = Nothing
                loRow = KwikSearch(p_oApp _
                                , lsSQL _
                                , True _
                                , fsValue _
                                , "sAcctNmbr»sClientNm»xAddressx" _
                                , "Account No»Client»Address", _
                                , "a.sAcctNmbr»b.sCompnyNm»CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode)" _
                                , IIf(fbIsCode, 0, 1))
            End If
            'end - mac 2020-07-18

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oDTMstr(0).Item("sClientID") = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sAddressx = ""
                p_oOthersx.nPNValuex = 0.0
                p_oOthersx.nDownPaym = 0.0
                p_oOthersx.nGrossPrc = 0.0
                p_oOthersx.nMonAmort = 0.0
                p_oOthersx.nCashBalx = 0.0
                p_oOthersx.nAcctTerm = 0
                p_oOthersx.nABalance = 0.0
                p_oOthersx.nAmtDuexx = 0.0
                p_oOthersx.xRebatesx = 0.0

                'kalyptus - 2022.10.22 09:50am
                'Set promo rebate and promo term to zero
                p_oOthersx.nPromRebt = 0
                p_oOthersx.nPromTerm = 0

                p_oOthersx.sEngineNo = ""
                p_oOthersx.sFrameNox = ""
                p_oOthersx.sModelNme = ""
                p_oOthersx.sColorNme = ""
                p_oOthersx.sGCardNox = ""
                p_oOthersx.cDigitalx = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sAcctNmbr")
                p_oDTMstr(0).Item("sClientID") = loRow.Item("sClientID")
                p_oOthersx.sClientNm = loRow.Item("sClientNm")
                p_oOthersx.sAddressx = loRow.Item("xAddressx")
                p_oOthersx.nPNValuex = loRow.Item("nPNValuex")
                p_oOthersx.nDownPaym = loRow.Item("nDownPaym")
                p_oOthersx.nGrossPrc = loRow.Item("nGrossPrc")
                p_oOthersx.nMonAmort = loRow.Item("nMonAmort")
                p_oOthersx.nCashBalx = loRow.Item("nCashBalx")
                p_oOthersx.nAcctTerm = loRow.Item("nAcctTerm")
                p_oOthersx.nABalance = loRow.Item("nABalance")

                'kalyptus - 2022.10.22 09:50am
                'Set promo rebate and promo term to zero
                p_oOthersx.nPromRebt = loRow.Item("nPromRebt")
                p_oOthersx.nPromTerm = loRow.Item("nPromTerm")

                p_oOthersx.xRebatesx = loRow.Item("nRebatesx")
                p_oOthersx.sEngineNo = loRow.Item("sEngineNo")
                p_oOthersx.sFrameNox = loRow.Item("sFrameNox")
                p_oOthersx.sModelNme = loRow.Item("sModelNme")
                p_oOthersx.sColorNme = loRow.Item("sColorNme")
                p_oOthersx.sGCardNox = loRow.Item("sGCardNox")
                p_oOthersx.cDigitalx = loRow.Item("cDigitalx")

                Dim loLR As New ARTrans(p_oApp)
                loLR.Master("sAcctNmbr") = p_oDTMstr(0).Item("sAcctNmbr")
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
                lsSQL = AddCondition(lsSQL, "b.sCompnyNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oDTMstr(0).Item("sClientID") = ""
            p_oOthersx.sClientNm = ""
            p_oOthersx.sAddressx = ""
            p_oOthersx.nPNValuex = 0.0
            p_oOthersx.nDownPaym = 0.0
            p_oOthersx.nGrossPrc = 0.0
            p_oOthersx.nMonAmort = 0.0
            p_oOthersx.nCashBalx = 0.0
            p_oOthersx.nAcctTerm = 0
            p_oOthersx.nABalance = 0.0
            p_oOthersx.nAmtDuexx = 0.0
            p_oOthersx.xRebatesx = 0.0

            'kalyptus - 2022.10.22 09:50am
            'Set promo rebate and promo term to zero
            p_oOthersx.nPromRebt = 0
            p_oOthersx.nPromTerm = 0

            p_oOthersx.sEngineNo = ""
            p_oOthersx.sFrameNox = ""
            p_oOthersx.sModelNme = ""
            p_oOthersx.sColorNme = ""
            p_oOthersx.sGCardNox = ""
            p_oOthersx.cDigitalx = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sAcctNmbr")
            p_oDTMstr(0).Item("sClientID") = loDta(0).Item("sClientID")
            p_oOthersx.sClientNm = loDta(0).Item("sClientNm")
            p_oOthersx.sAddressx = loDta(0).Item("xAddressx")
            p_oOthersx.nPNValuex = loDta(0).Item("nPNValuex")
            p_oOthersx.nDownPaym = loDta(0).Item("nDownPaym")
            p_oOthersx.nGrossPrc = loDta(0).Item("nGrossPrc")
            p_oOthersx.nMonAmort = loDta(0).Item("nMonAmort")
            p_oOthersx.nCashBalx = loDta(0).Item("nCashBalx")
            p_oOthersx.nAcctTerm = loDta(0).Item("nAcctTerm")
            p_oOthersx.nABalance = loDta(0).Item("nABalance")
            'p_oOthersx.nAmtDuexx = loDta(0).Item("nAmtDuexx")
            p_oOthersx.xRebatesx = loDta(0).Item("nRebatesx")

            'kalyptus - 2022.10.22 09:50am
            'Set promo rebate and promo term to zero
            p_oOthersx.nPromRebt = loDta(0).Item("nPromRebt")
            p_oOthersx.nPromTerm = loDta(0).Item("nPromTerm")

            p_oOthersx.sEngineNo = loDta(0).Item("sEngineNo")
            p_oOthersx.sFrameNox = loDta(0).Item("sFrameNox")
            p_oOthersx.sModelNme = loDta(0).Item("sModelNme")
            p_oOthersx.sColorNme = loDta(0).Item("sColorNme")
            p_oOthersx.sGCardNox = loDta(0).Item("sGCardNox")
            p_oOthersx.cDigitalx = loDta(0).Item("cDigitalx")

            Dim loLR As New ARTrans(p_oApp)
            loLR.Master("sAcctNmbr") = p_oDTMstr(0).Item("sAcctNmbr")
            Dim loLRMstr = loLR.GetMaster()
            p_oOthersx.nAmtDuexx = loLR.getDelay(loLRMstr, p_oDTMstr(0).Item("dTransact")) * p_oOthersx.nMonAmort
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getPaidBy(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.xPaidByxx <> "" Then Exit Sub
        Else
            'Do not allow searching of value if fsValue is empty
            If (fsValue = p_oOthersx.xPaidByxx And fsValue <> "") Or fsValue = "" Then Exit Sub
        End If

        Dim loClient As ggcClient.Client
        loClient = New ggcClient.Client(p_oApp)
        loClient.Parent = "ARPayment"

        'Assume that a call to this module using CLIENT ID is open
        If fbIsCode Then
            If loClient.OpenClient(fsValue) Then
                p_oPaidBy = loClient
                p_oDTMstr(0).Item("sPaidByID") = p_oPaidBy.Master("sClientID")
                p_oOthersx.xPaidByxx = p_oPaidBy.Master("sLastName") & ", " & _
                                       p_oPaidBy.Master("sFrstName") & _
                                       IIf(p_oPaidBy.Master("sSuffixNm") = "", "", " " & p_oPaidBy.Master("sSuffixNm")) & " " & _
                                       p_oPaidBy.Master("sMiddName")
            Else
                p_oDTMstr(0).Item("sPaidByID") = ""
                p_oOthersx.xPaidByxx = ""
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.xPaidByxx)
            Exit Sub
        End If

        'A call to this module using client name is search
        If loClient.SearchClient(fsValue, False) Then
            If loClient.ShowClient Then
                p_oPaidBy = loClient
                p_oDTMstr(0).Item("sPaidByID") = p_oPaidBy.Master("sClientID")
                p_oOthersx.xPaidByxx = p_oPaidBy.Master("sLastName") & ", " & _
                                       p_oPaidBy.Master("sFrstName") & _
                                       IIf(p_oPaidBy.Master("sSuffixNm") = "", "", " " & p_oPaidBy.Master("sSuffixNm")) & " " & _
                                       p_oPaidBy.Master("sMiddName")
            End If
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.xPaidByxx)
    End Sub


    'This method implements a search master where id and desc are not joined.
    Private Sub getCollector(ByVal fnColIdx As Integer _
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
                       "  b.sClientID" & _
                       ", b.sCompnyNm sCollName" & _
               " FROM Employee_Master001 a" & _
                " LEFT JOIN Client_Master b ON a.sEmployID = b.sClientID" & _
               " WHERE a.cCollectr = '1'" & _
                 " AND a.sBranchCD = " & strParm(p_sBranchCd) & _
        IIf(p_nEditMode = xeEditMode.MODE_ADDNEW, " AND a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sClientID»sCollName" _
                                             , "Coll ID»Collector", _
                                             , "b.sClientID»b.sCompnyNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sCollName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sClientID")
                p_oOthersx.sCollName = loRow.Item("sCollName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCollName)
            Exit Sub
        End If

        If fsValue = "" Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sCollName = ""
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "b.sClientID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "b.sCompnyNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sCollName = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sClientID")
            p_oOthersx.sCollName = loDta(0).Item("sCollName")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCollName)
    End Sub

    Private Function isRebateOk() As Boolean
        Dim lnRebates As Double
        Dim lnAllowReb As Double
        Dim lnExcess As Integer

        Dim lsApprovedCD As String = "", lsApproveID As String = "", lsApproveName As String = ""

        lnRebates = getRebates(lnExcess, lnAllowReb)

        If lnRebates > p_oDTMstr(0).Item("nRebatesx") Then
            If MsgBox("Rebate given is LESSER than the supposed rebate." & vbCrLf & _
                  "Continue Anyway?", vbQuestion + vbYesNo) <> vbYes Then
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("nRebatesx") > lnRebates Then
            MsgBox("Rebate given to " & p_oDTMstr(0).Item("sReferNox") & " is GREATER than the supposed rebate." & vbCrLf & _
               "You will be asked to enter the APPROVAL CODE given by an authorized personnel!", vbCritical, "Warning")
            If Not GetCodeApproval(p_oApp, lsApprovedCD, lsApproveID, lsApproveName) Then
                MsgBox("Rebate given is GREATER than the supposed rebate." & vbCrLf & _
                   "Verify entry then try again!", vbCritical, "Warning")
                Return False
            Else
                If isValidApproveCode( _
                   IIf(p_oDTMstr(0).Item("sCollIDxx") <> "", CodeApproval.pxeFieldRebate, CodeApproval.pxeOfficeRebate), _
                   p_oApp.BranchCode, _
                   Mid(lsApprovedCD, 4, 1), _
                   p_oDTMstr(0).Item("dTransact"), _
                   p_oDTMstr(0).Item("sReferNox"), _
                   lsApprovedCD) Then

                    p_oDTMstr(0).Item("sApproved") = lsApproveID
                    p_oDTMstr(0).Item("sAPprCode") = lsApprovedCD

                Else
                    MsgBox("Invalid APPROVAL CODE detected." & vbCrLf & _
                       "Verify entry then try again!", vbCritical, "Warning")
                    Return False
                End If
            End If
        End If

        isRebateOk = True
    End Function

    Public Sub showRebate()
        Dim loLRMaster As ARTrans
        Dim ldDueDate As Date
        Dim lnActTerm As Long
        Dim lnAmtDuex As Double

        Dim lnExcessDay As Integer
        Dim lnRebates As Double

        Dim lnTotlTerm As Long
        Dim lnPaymTerm As Long

        loLRMaster = New ARTrans(p_oApp)
        With loLRMaster
            Dim loDta As DataTable
            .Master("sAcctNmbr") = p_oDTMstr(0).Item("sAcctNmbr")
            loDta = .GetMaster()

            ldDueDate = p_oDTMstr(0).Item("dTransact")
            If ldDueDate > loDta(0).Item("dDueDatex") Then ldDueDate = loDta(0).Item("dDueDatex")

            lnActTerm = .getMonthTerm(loDta(0).Item("dFirstPay"), ldDueDate)

            'kalyptus - 2020.06.06 03:49pm
            'Replace the logic below
            'Freeze the term for 2 months for sales prior to the lockdown period and payments from the lockdown period...
            lnActTerm = lnActTerm - getFreezeMonth(loDta(0).Item("sAcctNmbr"), ldDueDate)

            'if account payment has no/or beyond promo term then perform the regular computation of rebate
            If lnAmtDuex < 0 Then
                If p_oDTMstr(0).Item("dTransact") < loDta(0).Item("dFirstPay") Then
                    'mac 2021.09.22
                    lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort"))) * loDta(0).Item("nRebatesx")
                Else
                    'orig
                    'lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort")) + 1) * loDta(0).Item("nRebatesx")
                    'mac 2022.12.28
                    lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort"))) * loDta(0).Item("nRebatesx")
                End If
            ElseIf lnAmtDuex = 0 Then
                lnRebates = loDta(0).Item("nRebatesx")
            End If

            'If lnExcessDay < 30 Then
            '    If lnAmtDuex <= loDta(0).Item("nMonAmort") Then
            '        lnRebates = lnRebates + loDta(0).Item("nRebatesx")
            '    End If
            'End If

            If p_oOthersx.nPromTerm <> 0 And lnActTerm <= p_oOthersx.nPromTerm Then 'pasok pa sa promo
                'Compute rebate based on the promo
                lnTotlTerm = (loDta(0).Item("nPaymTotl") + loDta(0).Item("nRebTotlx")) \ loDta(0).Item("nMonAmort")
                lnPaymTerm = (p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx")) \ loDta(0).Item("nMonAmort")

                If lnAmtDuex <= 0 Then
                    'Compute for the remaining term that has a promo rebate
                    lnTotlTerm = p_oOthersx.nPromTerm - lnTotlTerm

                    'Does this payment has promo rebate
                    If lnTotlTerm > 0 Then
                        'Is remaining promo term less than or equal the current monthly amortization payment
                        If lnTotlTerm >= lnPaymTerm Then
                            lnRebates = lnPaymTerm * p_oOthersx.nPromRebt
                        Else
                            lnRebates = lnTotlTerm * p_oOthersx.nPromRebt
                            lnRebates = lnRebates + ((lnPaymTerm - lnTotlTerm) * loDta(0).Item("nRebatesx"))
                        End If
                    Else
                        'Since it has no promo rebate then use the default rebate
                        lnRebates = lnPaymTerm * loDta(0).Item("nRebatesx")
                    End If
                End If

                If lnRebates > 0.0# Then
                    MsgBox("Ms/Mr " & p_oOthersx.sClientNm & " has rebate of " & lnRebates & " from PROMO.", vbInformation, "Notice")
                Else
                    MsgBox("Ms/Mr " & p_oOthersx.sClientNm & " is not entitled for a rebate.", vbInformation, "Notice")
                End If
            Else
                If lnRebates > 0.0# Then
                    MsgBox("Ms/Mr " & p_oOthersx.sClientNm & " has regular rebate of " & lnRebates & ".", vbInformation, "Notice")
                Else
                    MsgBox("Ms/Mr " & p_oOthersx.sClientNm & " is not entitled for a rebate.", vbInformation, "Notice")
                End If
            End If
        End With
    End Sub

    Private Function getRebates(ByRef lnExcessDay As Integer, ByRef lnRebates As Double) As Double
        Dim loLRMaster As ARTrans
        Dim ldDueDate As Date
        Dim lnActTerm As Long
        Dim lnAmtDuex As Double

        Dim lnTotlTerm As Long
        Dim lnPaymTerm As Long

        getRebates = 0

        loLRMaster = New ARTrans(p_oApp)
        With loLRMaster
            Dim loDta As DataTable
            .Master("sAcctNmbr") = p_oDTMstr(0).Item("sAcctNmbr")
            loDta = .GetMaster()

            ldDueDate = p_oDTMstr(0).Item("dTransact")
            If ldDueDate > loDta(0).Item("dDueDatex") Then ldDueDate = loDta(0).Item("dDueDatex")

            lnActTerm = .getMonthTerm(loDta(0).Item("dFirstPay"), ldDueDate)

            'kalyptus - 2020.06.06 03:49pm
            'Replace the logic below
            'Freeze the term for 2 months for sales prior to the lockdown period and payments from the lockdown period...
            lnActTerm = lnActTerm - getFreezeMonth(loDta(0).Item("sAcctNmbr"), ldDueDate)

            ''kalyptus - 2020.05.18 11:28am
            ''Freeze the term for 2 months for sales prior to the lockdown period and payments from the lockdown period...
            'If loDta(0).Item("dPurchase") < CDate("2020-04-01") And p_oDTMstr(0).Item("dTransact") >= CDate("2020-05-16") Then
            '    lnActTerm = lnActTerm - 2
            'End If

            ' compute the excess days for validation of rebates by user
            If Day(p_oDTMstr(0).Item("dTransact")) > Day(loDta(0).Item("dFirstPay")) Then
                lnExcessDay = DateDiff("d", DateSerial(Year(p_oDTMstr(0).Item("dTransact")), _
                               Month(p_oDTMstr(0).Item("dTransact")), Day(loDta(0).Item("dFirstPay"))), _
                               p_oDTMstr(0).Item("dTransact"))
            Else
                ldDueDate = DateSerial(Year(p_oDTMstr(0).Item("dTransact")), _
                               Month(p_oDTMstr(0).Item("dTransact")) + 1, Day(loDta(0).Item("dFirstPay")))
                If Month(DateAdd("m", 1, p_oDTMstr(0).Item("dTransact"))) <> Month(ldDueDate) Then
                    ldDueDate = DateAdd("d", Day(ldDueDate) * -1, ldDueDate)
                End If

                lnExcessDay = DateDiff("d", p_oDTMstr(0).Item("dTransact"), ldDueDate)
            End If

            lnAmtDuex = (lnActTerm * loDta(0).Item("nMonAmort")) + _
                        loDta(0).Item("nDownPaym") + _
                        loDta(0).Item("nCashBalx") + _
                        loDta(0).Item("nDebtTotl")

            lnAmtDuex = lnAmtDuex - (loDta(0).Item("nPaymTotl") + loDta(0).Item("nRebTotlx") + _
                        p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx") + _
                        loDta(0).Item("nDownTotl") + loDta(0).Item("nCashTotl") + _
                        loDta(0).Item("nCredTotl"))


            'kalyptus-2022.10.22 10:08am
            'incorporate computation of rebate with promo
            If p_oOthersx.nPromTerm <> 0 And lnActTerm > p_oOthersx.nPromTerm Then
                'if account payment has no/or beyond promo term then perform the regular computation of rebate
                If lnAmtDuex < 0 Then
                    If p_oDTMstr(0).Item("dTransact") < loDta(0).Item("dFirstPay") Then
                        'mac 2021.09.22
                        lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort"))) * loDta(0).Item("nRebatesx")
                    Else
                        'orig
                        'lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort")) + 1) * loDta(0).Item("nRebatesx")
                        'mac 2022.12.28
                        lnRebates = (((lnAmtDuex * -1) \ loDta(0).Item("nMonAmort"))) * loDta(0).Item("nRebatesx")
                    End If

                    getRebates = lnRebates
                ElseIf lnAmtDuex = 0 Then
                    lnRebates = loDta(0).Item("nRebatesx")
                    getRebates = lnRebates
                End If

                If lnExcessDay < 30 Then
                    If lnAmtDuex <= loDta(0).Item("nMonAmort") Then
                        lnRebates = lnRebates + loDta(0).Item("nRebatesx")
                    End If
                End If
            Else
                'Compute rebate based on the promo
                lnTotlTerm = (loDta(0).Item("nPaymTotl") + loDta(0).Item("nRebTotlx")) \ loDta(0).Item("nMonAmort")
                lnPaymTerm = (p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nRebatesx")) \ loDta(0).Item("nMonAmort")

                If lnAmtDuex <= 0 Then
                    'Compute for the remaining term that has a promo rebate
                    lnTotlTerm = p_oOthersx.nPromTerm - lnTotlTerm

                    'Does this payment has promo rebate
                    If lnTotlTerm > 0 Then
                        'Is remaining promo term less than or equal the current monthly amortization payment
                        If lnTotlTerm >= lnPaymTerm Then
                            lnRebates = lnPaymTerm * p_oOthersx.nPromRebt
                        Else
                            lnRebates = lnTotlTerm * p_oOthersx.nPromRebt
                            lnRebates = lnRebates + ((lnPaymTerm - lnTotlTerm) * loDta(0).Item("nRebatesx"))
                        End If
                    Else
                        'Since it has no promo rebate then use the default rebate
                        lnRebates = lnPaymTerm * loDta(0).Item("nRebatesx")
                    End If

                    getRebates = lnRebates
                End If
            End If
        End With

        Return getRebates
    End Function

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
                    ", a.cPaymForm" & _
                    ", a.sReferNox" & _
                    ", a.sAcctNmbr" & _
                    ", a.sClientID" & _
                    ", a.sPaidByID" & _
                    ", a.sRemarksx" & _
                    ", a.nAmountxx" & _
                    ", a.nIntAmtxx" & _
                    ", a.nRebatesx" & _
                    ", a.nPenaltyx" & _
                    ", a.sCollIDxx" & _
                    ", a.sApproved" & _
                    ", a.sAPprCode" & _
                    ", a.cTranType" & _
                    ", a.cPostedxx" & _
                    ", a.dPostedxx" & _
                    ", a.sSourceCd" & _
                    ", a.sSourceNo" & _
                    ", a.cPrintedx" & _
                    ", a.cGCrdPstd" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a" & _
                " WHERE a.cTranType = " & strParm(p_cTranType)
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sReferNox" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
              " WHERE a.sClientID = b.sClientID" & _
                " AND a.cTranType = " & strParm(p_cTranType)
    End Function

    Public Function AmountInWords(ByVal nAmount As String, Optional ByVal wAmount _
                 As String = vbNullString, Optional ByVal nSet As Object = Nothing) As String
        'Let's make sure entered value is numeric
        If Not IsNumeric(nAmount) Then Return "Please enter numeric values only."

        Dim tempDecValue As String = String.Empty : If InStr(nAmount, ".") Then _
            tempDecValue = nAmount.Substring(nAmount.IndexOf("."))
        nAmount = Replace(nAmount, tempDecValue, String.Empty)

        Try
            Dim intAmount As Long = nAmount
            If intAmount > 0 Then
                nSet = IIf((intAmount.ToString.Trim.Length / 3) _
                 > (CLng(intAmount.ToString.Trim.Length / 3)), _
                  CLng(intAmount.ToString.Trim.Length / 3) + 1, _
                   CLng(intAmount.ToString.Trim.Length / 3))
                Dim eAmount As Long = Microsoft.VisualBasic.Left(intAmount.ToString.Trim, _
                  (intAmount.ToString.Trim.Length - ((nSet - 1) * 3)))
                Dim multiplier As Long = 10 ^ (((nSet - 1) * 3))

                Dim Ones() As String = _
                {"", "One", "Two", "Three", _
                  "Four", "Five", _
                  "Six", "Seven", "Eight", "Nine"}
                Dim Teens() As String = {"", _
                "Eleven", "Twelve", "Thirteen", _
                  "Fourteen", "Fifteen", _
                  "Sixteen", "Seventeen", "Eighteen", "Nineteen"}
                Dim Tens() As String = {"", "Ten", _
                "Twenty", "Thirty", _
                  "Forty", "Fifty", "Sixty", _
                  "Seventy", "Eighty", "Ninety"}
                Dim HMBT() As String = {"", "", _
                "Thousand", "Million", _
                  "Billion", "Trillion", _
                  "Quadrillion", "Quintillion"}

                intAmount = eAmount

                Dim nHundred As Integer = intAmount \ 100 : intAmount = intAmount Mod 100
                Dim nTen As Integer = intAmount \ 10 : intAmount = intAmount Mod 10
                Dim nOne As Integer = intAmount \ 1

                If nHundred > 0 Then wAmount = wAmount & _
                Ones(nHundred) & " Hundred " 'This is for hundreds                
                If nTen > 0 Then 'This is for tens and teens
                    If nTen = 1 And nOne > 0 Then 'This is for teens 
                        wAmount = wAmount & Teens(nOne) & " "
                    Else 'This is for tens, 10 to 90
                        wAmount = wAmount & Tens(nTen) & IIf(nOne > 0, "-", " ")
                        If nOne > 0 Then wAmount = wAmount & Ones(nOne) & " "
                    End If
                Else 'This is for ones, 1 to 9
                    If nOne > 0 Then wAmount = wAmount & Ones(nOne) & " "
                End If
                wAmount = wAmount & HMBT(nSet) & " "
                wAmount = AmountInWords(CStr(CLng(nAmount) - _
                  (eAmount * multiplier)).Trim & tempDecValue, wAmount, nSet - 1)
            Else
                If Val(nAmount) = 0 Then nAmount = nAmount & _
                tempDecValue : tempDecValue = String.Empty
                If (Math.Round(Val(nAmount), 2) * 100) > 0 Then wAmount = Trim(CStr(wAmount.Trim & " Pesos " + "& " + (nAmount * 100).ToString + "/100"))
            End If
        Catch ex As Exception
            MsgBox(ex.Message())
            Return ""
        End Try

        'Trap null values
        If IsNothing(wAmount) = True Then wAmount = String.Empty Else wAmount = _
          IIf(InStr(wAmount.Trim.ToLower, "pesos"), _
          wAmount.Trim, wAmount.Trim & " Pesos")

        'Display the result
        Return wAmount
    End Function

    Private Function getFreezeMonth(ByVal fsAcctNmbr As String, ByVal fdTransact As Date) As Integer
        Dim lsSQL As String
        Dim loDta As DataTable

        getFreezeMonth = 0

        lsSQL = "SELECT sAcctNmbr, b.*" & _
               " FROM MC_AR_Master a" & _
                   " LEFT JOIN Branch_Lockdown_History b ON a.sBranchCd = b.sBranchCD" & _
               " WHERE sAcctNmbr = " & strParm(fsAcctNmbr) & _
                 " AND b.dDateFrom > a.dFirstPay" & _
                 " AND b.dDateThru < " & dateParm(fdTransact)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        Dim lnCtr As Integer
        For lnCtr = 0 To loDta.Rows.Count - 1
            getFreezeMonth = getFreezeMonth + loDta(lnCtr).Item("nMonthxxx")
        Next

        'kalyptus - 2020.08.06 09:18am
        'Remove Unfreezed month if client grab the promo
        lsSQL = "SELECT DISTINCT a.sOthrInfo nUnfreezd, b.sMainInfo, b.sOthrInfo" & _
               " FROM CCS_Promo_Master a" & _
                    " LEFT JOIN CCS_Promo_Detail b ON a.sTransNox = b.sTransNox" & _
               " WHERE a.sProgrmCD = '0001'" & _
                 " AND b.sOthrInfo BETWEEN a.dDateFrom AND a.dDateThru" & _
                 " AND b.sMainInfo = " & strParm(fsAcctNmbr) & _
                 " AND b.sOthrInfo < " & dateParm(fdTransact)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        For lnCtr = 0 To loDta.Rows.Count - 1
            getFreezeMonth = getFreezeMonth - Val(loDta(lnCtr).Item("nUnfreezd"))
        Next

    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        p_sBranchCd = p_oApp.BranchCode
        p_sBranchNm = p_oApp.BranchName

        p_nTranStat = -1
        p_cTranType = "2"   'Set default transaction type to Monthly Payment
        p_sParent = ""
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal trantype As String)
        Me.New(foRider)
        p_cTranType = trantype
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32, ByVal fctrantype As String)
        Me.New(foRider)
        p_nTranStat = fnStatus
        p_cTranType = fctrantype

        If fctrantype = "2" Or fctrantype = "3" Or fctrantype = "4" Then
            p_cTranType = fctrantype
        End If
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sAddressx As String
        Public nPNValuex As Decimal
        Public nDownPaym As Integer
        Public nGrossPrc As Decimal
        Public nMonAmort As Decimal
        Public nCashBalx As Decimal
        Public nAcctTerm As Decimal
        Public nABalance As Decimal
        Public nAmtDuexx As Decimal
        Public xRebatesx As Decimal
        Public sEngineNo As String
        Public sFrameNox As String
        Public sModelNme As String
        Public sColorNme As String
        Public sGCardNox As String
        Public cDigitalx As String

        Public sCompnyNm As String
        Public sCompnyID As String
        Public xPaidByxx As String
        Public sCollName As String

        'kalyptus - 2022.10.22 09:54am
        'Add these fields to store promo term and rebate
        Public nPromTerm As Decimal
        Public nPromRebt As Decimal

        'kalyptus - 2017.07.12 03:32pm
        'Change structure
        Public sCheckNox As String
        Public sAcctNoxx As String
        Public sBankIDxx As String
        Public sBankName As String
        Public sCheckDte As String
        Public nCheckAmt As Decimal

        'mac 2022.10.05 10:30am
        Public sEPCltIDx As String
        Public sEPCltNme As String
        Public sEPReferN As String
        Public sEPTermID As String
        Public sEPTermNm As String
        Public sEPRemrks As String
        Public nEPAmount As Decimal
    End Class
End Class
