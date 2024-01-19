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
'
'   Mac 2020.03.09
'       Added Try/Catch statement on insert/update statements
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports System.Drawing

Public Class LRPayment
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_oOthersx As New Others
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String

    Private Const p_sMasTable As String = "LR_Payment_Master"
    Private Const p_sMsgHeadr As String = "LR Payment"
    Private Const p_sSourceCd As String = "ARRc"

    Private p_cLoanType As String

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
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case 81 ' sAddressx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case 82 ' nABalance 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case 83 ' nInterest 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nInterest
                    Case 84 ' nIntTotal 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nIntTotal
                    Case 85 ' nAcctTerm
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case 86 ' nMonAmort 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case 87 ' nAmtDuexx
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case 88 ' sCompnyNm 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case 89 ' sCompnyID 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCompnyID
                    Case 90 ' sCollName
                        If Trim(IFNull(p_oDTMstr(0).Item(10))) <> "" And Trim(p_oOthersx.sCollName) = "" Then
                            getCollector(10, 90, p_oDTMstr(0).Item(10), True, False)
                        End If
                        Return p_oOthersx.sCollName

                    Case 91 ' xTranAmtx
                        Return p_oOthersx.xTranAmtx
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
                    Case 81 To 89
                    Case 90 ' sCollName
                        getCollector(10, 90, value, False, False)
                    Case 91
                        If Trim(p_oDTMstr(0).Item("sAcctNmbr")) = "" Then
                            Exit Property
                        End If

                        If IsNumeric(value) Then
                            p_oOthersx.xTranAmtx = value

                            Dim loLR As New LRTrans(p_oApp)
                            loLR.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")

                            Dim loDta As DataTable = loLR.GetMaster()
                            Dim lnPrincipl As Decimal = loDta(0).Item("nPrincipl")
                            Dim lnInterest As Decimal = loDta(0).Item("nInterest")
                            Dim lnAcctTerm As Integer = loDta(0).Item("nAcctTerm")
                            Dim lnPaymTotl As Decimal = loDta(0).Item("nPaymTotl")
                            Dim lnIntTotal As Decimal = loDta(0).Item("nIntTotal")
                            Dim lnTranAmtx As Decimal = p_oOthersx.xTranAmtx
                            Dim lnPaidAmtx As Decimal = 0
                            Dim lnIntAmtxx As Decimal = 0

                            Call SplitPayment(lnPrincipl, lnInterest, lnAcctTerm, lnPaymTotl, lnIntTotal, lnTranAmtx, lnPaidAmtx, lnIntAmtxx)

                            p_oDTMstr(0).Item("nAmountxx") = lnPaidAmtx
                            p_oDTMstr(0).Item("nIntAmtxx") = lnIntAmtxx
                        End If

                        RaiseEvent MasterRetrieved(7, p_oDTMstr(0).Item("nAmountxx"))
                        RaiseEvent MasterRetrieved(8, p_oDTMstr(0).Item("nIntAmtxx"))
                        RaiseEvent MasterRetrieved(91, p_oOthersx.xTranAmtx)
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
                    Case 9  'Penalty
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
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sClientNm
                    Case "saddressx" ' 81 
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sAddressx
                    Case "nabalance" ' 82
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nABalance
                    Case "ninterest" ' 83
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nInterest
                    Case "ninttotal" ' 84
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nIntTotal
                    Case "nacctterm" ' 85
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAcctTerm
                    Case "nmonamort" ' 86
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nMonAmort
                    Case "namtduexx" ' 87
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.nAmtDuexx
                    Case "scompnynm" '88
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCompnyNm
                    Case "scompnyid" '89
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sClientNm) = "" Then
                            getAccount(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCompnyID
                    Case "scollname" '90
                        If Trim(IFNull(p_oDTMstr(0).Item(10))) <> "" And Trim(p_oOthersx.sCollName) = "" Then
                            getCollector(10, 90, p_oDTMstr(0).Item(10), True, False)
                        End If
                        Return p_oOthersx.sCollName
                    Case "xtranamtx" '91  
                        Return p_oOthersx.xTranAmtx

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
                        getAccount(4, 80, value, False, False)
                    Case "scollname"
                        getCollector(10, 90, value, False, False)
                    Case "xtranamtx"
                        Master(91) = value
                    Case "saddressx", "nabalance", "ninterest", "ninttotal", "nacctterm", "nmonamort", "namtduexx", "scompnynm"
                    Case "dtransact"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = getValidDate(p_oApp, value)
                    Case "npenaltyx"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "srefernox"
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

    Public Property LoanType() As String
        Get
            Return p_cLoanType
        End Get
        Set(ByVal value As String)
            p_cLoanType = value
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

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sReferNox LIKE " & strParm(fsValue)
        Else
            lsFilter = "b.sCompnyNm like " & strParm(fsValue & "%")
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

        Dim lsSQL As String = ""

        Try
            'Save master table 
            'Note: Update is not allowed!!!
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
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
                    '          ", nAmountxx = " & p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nIntAmtxx") + p_oDTMstr(0).Item("nPenaltyx") & _
                    '          ", sRemarksx = " & strParm("LRPy»" & p_oOthersx.sBankAcct & "»" & p_oOthersx.sClientNm) & _
                    '          ", nClearDay = " & 0 & _
                    '          ", cTranStat = " & strParm("0") & _
                    '          ", sModified = " & strParm(p_oApp.UserID) & _
                    '          ", dModified = " & dateParm(p_oApp.getSysDate)
                    'p_oApp.Execute(lsSQx, "Check_Payments", p_sBranchCd)
                End If

                If p_sParent = "" Then p_oApp.BeginTransaction()


                p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
                lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)

                If lsSQL <> "" Then
                    If p_oApp.Execute(lsSQL, p_sMasTable, p_sBranchCd) <= 0 Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        MsgBox("Unable to save LR Payment.", MsgBoxStyle.Critical, "Warning")
                        Return False
                    End If
                End If

                'check if it is an e-payment 
                If p_oOthersx.sEPCltIDx <> "" Then
                    lsSQL = "INSERT INTO Other_Payment_Received SET" & _
                            "  sTransNox = " & strParm(GetNextCode("Other_Payment_Received", "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)) & _
                            ", sClientID = " & strParm(p_oOthersx.sEPCltIDx) & _
                            ", sReferNox = " & strParm(p_oOthersx.sEPReferN) & _
                            ", nTotlAmnt = " & Master(91) + Master("nPenaltyX") & _
                            ", nAmtPaidx = " & p_oOthersx.nEPAmount & _
                            ", sTermCode = " & strParm(p_oOthersx.sEPTermID) & _
                            ", sRemarksx = " & strParm(p_oOthersx.sEPRemrks) & _
                            ", sSourceCd = " & strParm(p_sSourceCd) & _
                            ", sSourceNo = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                            ", cTranStat = '1'" & _
                            ", dModified = " & dateParm(p_oApp.SysDate)

                    If p_oApp.Execute(lsSQL, "Other_Payment_Received", p_sBranchCd) <= 0 Then
                        If p_sParent = "" Then p_oApp.RollBackTransaction()
                        MsgBox("Unable to save LR Payment.", MsgBoxStyle.Critical, "Warning")
                        Return False
                    End If
                End If

                If p_sParent = "" Then p_oApp.CommitTransaction()
            End If

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
            MsgBox("Request was already posted!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED) Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED)
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

        Dim lnTotlSale As Decimal = p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nIntAmtxx") + p_oDTMstr(0).Item("nPenaltyx")
        Dim lnVatSales As Decimal = lnTotlSale / 1.12
        Dim lnLessVatx As Decimal = lnVatSales * 0.12

        'Total Sales(VAT Inclusive)
        loPrint.Print(5.8, 2.0, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)

        'Print transaction Date
        loPrint.Print(6.5, 6.0, Format(p_oDTMstr(0).Item("dTransact"), "MMM dd, yyyy"))

        'Less VAT
        loPrint.Print(7, 2.0, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        'Total
        loPrint.Print(8.4, 2.0, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)

        'Print Name
        loPrint.Print(8.4, 3.2, Master(80) & " / " & p_oDTMstr(0).Item("sAcctNmbr"))

        'Print Address
        loPrint.Print(10, 3.2, Master(81))

        'Amount in words
        loPrint.Print(14, 3.2, AmountInWords(lnTotlSale))
        'Amount in number
        loPrint.Print(15.5, 6.7, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)

        'VATable Sales
        loPrint.Print(17.7, 2.0, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)
        'VAT Exempt
        loPrint.Print(19.2, 2.0, "0.00", StringAlignment.Far)
        'zero rated Sales
        loPrint.Print(20.7, 2.0, "0.00", StringAlignment.Far)
        'VAT amount
        loPrint.Print(22, 2.0, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        'Total Sales 
        loPrint.Print(23.5, 2.0, Format(lnVatSales + lnLessVatx, "#,##0.00"), StringAlignment.Far)

        'Principal
        loPrint.Print(20, 4.6, Format(p_oDTMstr(0).Item("nAmountxx"), "#,##0.00"), StringAlignment.Far)
        'Interest
        loPrint.Print(21.5, 4.6, Format(p_oDTMstr(0).Item("nIntAmtxx"), "#,##0.00"), StringAlignment.Far)
        'Penalty
        loPrint.Print(23.5, 7, Format(p_oDTMstr(0).Item("nPenaltyx"), "#,##0.00"), StringAlignment.Far)
        

        ''Print transaction Date
        'loPrint.Print(9, 2.9, Format(p_oDTMstr(0).Item("dTransact"), "MMM dd, yyyy"))

        ''Print Name
        'loPrint.Print(10.5, 0.9, Master(80) & " / " & p_oDTMstr(0).Item("sAcctNmbr"))

        ''Print Address
        'loPrint.Print(12.5, 0.9, Master(81))

        ''Principal
        'loPrint.Print(18, 3.55, Format(p_oDTMstr(0).Item("nAmountxx"), "#,##0.00"), StringAlignment.Far)
        ''Interest
        'loPrint.Print(19.3, 3.55, Format(p_oDTMstr(0).Item("nIntAmtxx"), "#,##0.00"), StringAlignment.Far)
        ''Penalty
        'loPrint.Print(28.5, 3.55, Format(p_oDTMstr(0).Item("nPenaltyx"), "#,##0.00"), StringAlignment.Far)

        'Dim lnTotlSale As Decimal = p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nIntAmtxx") + p_oDTMstr(0).Item("nPenaltyx")
        'Dim lnVatSales As Decimal = lnTotlSale / 1.12
        'Dim lnLessVatx As Decimal = lnVatSales * 0.12

        ''Total Sales(VAT Inclusive)
        ''loPrint.Print(30, 40, Format(lnTotlSale, "#,##0.00"))
        'loPrint.Print(32.5, 3.55, Format(lnTotlSale, "#,##0.00"), StringAlignment.Far)
        ''Less VAT
        ''loPrint.Print(31, 40, Format(lnLessVatx, "#,##0.00"))
        'loPrint.Print(34, 3.55, Format(lnLessVatx, "#,##0.00"), StringAlignment.Far)
        ''Total
        ''loPrint.Print(32, 40, Format(lnVatSales, "#,##0.00"))
        'loPrint.Print(35.5, 3.55, Format(lnVatSales, "#,##0.00"), StringAlignment.Far)

        loPrint.PrintEnd()

        Return True

    End Function

    'Public Function PostTransaction()
    Public Function PostTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_POSTED) Then
            MsgBox("Payment was already posted!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        ElseIf p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_CANCELLED) Then
            MsgBox("Payment was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
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

        'she 2017-03-27 2:52 pm 
        'Add date filter to check all unposted payment < than sa current na pinopost.
        '" AND dTransact < " & dateParm(p_oDTMstr(0).Item("dTransact"))
        Dim loDta As DataTable = p_oApp.ExecuteQuery(lsSQL)
        If loDta.Rows.Count > 0 Then
            MsgBox("There are unposted payment for this account!" & vbCrLf & _
                   "Please post the transaction first...", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Try
            If p_sParent = "" Then p_oApp.BeginTransaction()

            Dim loTrans As LRTrans
            loTrans = New LRTrans(p_oApp)
            loTrans.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")
            loTrans.Transact_Date = p_oDTMstr(0).Item("dTransact")
            loTrans.Amount = p_oDTMstr(0).Item("nAmountxx")
            loTrans.Penalty = p_oDTMstr(0).Item("nPenaltyx")
            loTrans.Interest = p_oDTMstr(0).Item("nIntAmtxx")
            loTrans.SourceNo = p_oDTMstr(0).Item("sTransNox")
            loTrans.Remarks = p_oDTMstr(0).Item("sRemarksx")

            loTrans.ReferNo = p_oDTMstr(0).Item("sReferNox")
            loTrans.Collector = p_oDTMstr(0).Item("sCollIDxx")
            loTrans.isOffice = Trim(p_oDTMstr(0).Item("sCollIDxx")) = ""

            'Check Transaction Type
            If p_oDTMstr(0).Item("cTranType") = "0" Then
                'Payment?
                If Not loTrans.Payment Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            ElseIf p_oDTMstr(0).Item("cTranType") = "1" Then
                'Penalty?
                If Not loTrans.Penalty Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If

            p_oDTMstr(0).Item("cPostedxx") = CStr(xeTranStat.TRANS_POSTED)
            p_oDTMstr(0).Item("dPostedxx") = p_oApp.getSysDate

            lsSQL = "UPDATE " & p_sMasTable & _
                   " SET cPostedxx = " & strParm(CStr(xeTranStat.TRANS_POSTED)) & _
                      ", dPostedxx = " & dateParm(p_oDTMstr(0).Item("dPostedxx")) & _
                   " WHERE sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
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
            Case 4  ' sClientNm
                getAccount(4, 80, fsValue, True, True)
            Case 80 ' sClientNm
                getAccount(4, 80, fsValue, False, True)
            Case 90 ' sCollName
                getCollector(10, 90, fsValue, False, True)
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
                Case "cpostedxx", "ctrantype", "cpaymform", "cprintedx"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "namountxx", "nintamtxx", "npenaltyx"
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
        p_oOthersx.sCompnyID = ""

        p_oOthersx.sCollName = ""
        p_oOthersx.xTranAmtx = p_oDTMstr(0).Item("nAmountxx") + p_oDTMstr(0).Item("nIntAmtxx")

        'kalyptus - 2017.07.12 03:32pm
        'Change structure
        p_oOthersx.sCheckNox = ""
        p_oOthersx.sAcctNoxx = ""
        p_oOthersx.sBankIDxx = ""
        p_oOthersx.sBankName = ""
        p_oOthersx.sCheckDte = ""
        p_oOthersx.nCheckAmt = 0.0
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
        If Val(p_oDTMstr(0).Item("nAmountxx")) + Val(p_oDTMstr(0).Item("nIntAmtxx")) + Val(p_oDTMstr(0).Item("nPenaltyx")) <= 0 Then
            MsgBox("Transaction Amount seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
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
                       ", a.nInterest" & _
                       ", a.nAcctTerm" & _
                       ", a.nMonAmort" & _
                       ", a.nAmtDuexx" & _
                       ", e.sCompnyNm" & _
                       ", a.sCompnyID" & _
                       ", a.sClientID" & _
                       ", a.nIntTotal" & _
               " FROM LR_Master a" & _
                " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" & _
                " LEFT JOIN TownCity c ON b.sTownIDxx = c.sTownIDxx" & _
                " LEFT JOIN Province d ON c.sProvIDxx = d.sProvIDxx" & _
                " LEFT JOIN Company e ON a.sCompnyID = e.sCompnyID" & _
               " WHERE a.cLoanType = " & strParm(p_cLoanType)

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
                p_oOthersx.nIntTotal = loRow.Item("nIntTotal")
                p_oOthersx.nMonAmort = loRow.Item("nMonAmort")
                p_oOthersx.sCompnyNm = IFNull(loRow.Item("sCompnyNm"), "")
                p_oOthersx.sCompnyID = loRow.Item("sCompnyID")

                Dim loLR As New LRTrans(p_oApp)
                loLR.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")
                Dim loLRMstr = loLR.GetMaster()
                'for testing oly
                'p_oOthersx.nAmtDuexx = loLR.getDelay(loLRMstr, p_oDTMstr(0).Item("dTransact")) * p_oOthersx.nMonAmort
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
            Exit Sub
        End If

        If fsValue = "" Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oDTMstr(0).Item("sClientID") = ""
            Call InitOthers()
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
            p_oOthersx.nIntTotal = loDta(0).Item("nIntTotal")
            p_oOthersx.nMonAmort = loDta(0).Item("nMonAmort")
            p_oOthersx.sCompnyNm = IFNull(loDta(0).Item("sCompnyNm"))
            p_oOthersx.sCompnyID = IFNull(loDta(0).Item("sCompnyID"))

            Dim loLR As New LRTrans(p_oApp)
            loLR.AccountNo = p_oDTMstr(0).Item("sAcctNmbr")
            Dim loLRMstr = loLR.GetMaster()
            p_oOthersx.nAmtDuexx = loLR.getDelay(loLRMstr, p_oDTMstr(0).Item("dTransact")) * p_oOthersx.nMonAmort
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sClientNm)
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

        If fbIsCode Then
            lsSQL = AddCondition(lsSQL, "b.sClientID = " & strParm(fsValue))
        Else
            lsSQL = AddCondition(lsSQL, "b.sCompnyNm = " & strParm(fsValue))
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

    Private Sub SplitPayment( _
            ByVal fnPrincipl As Decimal _
          , ByVal fnInterest As Decimal _
          , ByVal fnAcctTerm As Integer _
          , ByRef fnPaymTotl As Decimal _
          , ByRef fnIntTotal As Decimal _
          , ByRef fnTranAmtx As Decimal _
          , ByRef fnPaidAmtx As Decimal _
          , ByRef fnIntAmtxx As Decimal)

        'Compute for the monthly amortization for the principal and interest
        Dim lnPayAmort As Decimal = fnPrincipl / fnAcctTerm
        Dim lnIntAmort As Decimal = fnInterest / fnAcctTerm

        'Compute for the number of terms paid for the principal and interest
        Dim lnPayTermx As Single = fnPaymTotl / lnPayAmort
        Dim lnIntTermx As Single = fnIntTotal / lnIntAmort

        If fnTranAmtx <= 0 Then Exit Sub

        If lnPayTermx = lnIntTermx Then
            'Distribute payment to interest payment
            If fnTranAmtx < lnIntAmort Then
                'Get the actual interest deducted
                lnIntAmort = fnTranAmtx

                fnIntAmtxx = fnIntAmtxx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnIntAmtxx = fnIntAmtxx + lnIntAmort
                fnTranAmtx = fnTranAmtx - lnIntAmort
            End If

            'Distribute payment to monthly payment
            If fnTranAmtx < lnPayAmort Then
                'Get the actual monthly amortization deducted
                lnPayAmort = fnTranAmtx

                fnPaidAmtx = fnPaidAmtx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnPaidAmtx = fnPaidAmtx + lnPayAmort
                fnTranAmtx = fnTranAmtx - lnPayAmort
            End If
            fnPaymTotl = fnPaymTotl + lnPayAmort
            fnIntTotal = fnIntTotal + lnIntAmort
        ElseIf lnPayTermx < lnIntTermx Then
            'Compute for the amount to be distributed for monthly payment
            'Dim lnDiff As Decimal = (lnPayTermx - lnIntTermx) * lnPayAmort
            Dim lnDiff As Decimal = (lnIntTermx - lnPayTermx) * lnPayAmort
            lnPayAmort = lnDiff
            If fnTranAmtx < lnDiff Then
                'Get the actual monthly amortization
                lnPayAmort = fnTranAmtx

                fnPaidAmtx = fnPaidAmtx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnPaidAmtx = fnPaidAmtx + lnDiff
                fnTranAmtx = fnTranAmtx - lnDiff
            End If
            fnPaymTotl = fnPaymTotl + lnPayAmort
        Else
            'Compute for the amount to be distributed for interest payment
            Dim lnDiff As Decimal = (lnPayTermx - lnIntTermx) * lnIntAmort
            lnIntAmort = lnDiff
            If fnTranAmtx < lnDiff Then
                lnIntAmort = fnTranAmtx
                fnIntAmtxx = fnIntAmtxx + fnTranAmtx
                fnTranAmtx = 0
            Else
                fnIntAmtxx = fnIntAmtxx + lnDiff
                fnTranAmtx = fnTranAmtx - lnDiff
            End If
            fnIntTotal = fnIntTotal + lnIntAmort
        End If

        'Execute a recursive function if fnTranAmtx is not yet 0
        If fnTranAmtx > 0 Then
            SplitPayment(fnPrincipl, fnInterest, fnAcctTerm, fnPaymTotl, fnIntTotal, fnTranAmtx, fnPaidAmtx, fnIntAmtxx)
        End If
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
                    ", a.cPaymForm" & _
                    ", a.sReferNox" & _
                    ", a.sAcctNmbr" & _
                    ", a.sClientID" & _
                    ", a.sRemarksx" & _
                    ", a.nAmountxx" & _
                    ", a.nIntAmtxx" & _
                    ", a.nPenaltyx" & _
                    ", a.sCollIDxx" & _
                    ", a.sApproved" & _
                    ", a.sAPprCode" & _
                    ", a.cTranType" & _
                    ", a.cPostedxx" & _
                    ", a.dPostedxx" & _
                    ", a.sPaidByID" & _
                    ", a.sSourceCD" & _
                    ", a.sSourceNo" & _
                    ", a.cPrintedx" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a" & _
                " WHERE a.cTranType IN ('0', '1')"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sReferNox" & _
                    ", b.sCompnyNm sClientNm" & _
                    ", a.dTransact" & _
              " FROM " & p_sMasTable & " a" & _
                    ", Client_Master b" & _
                    ", LR_Master c" & _
              " WHERE a.sClientID = b.sClientID" & _
                " AND a.sAcctNmbr = c.sAcctNmbr" & _
                " AND c.cLoanType = " & strParm(p_cLoanType) & _
                " AND a.cTranType IN ('0', '1')"


    End Function

    Private Function getSQ_OtherPaym() As String
        Return "SELECT" & _
                   "  a.sTransNox" & _
                   ", a.sClientID" & _
                   ", a.nTotlAmnt" & _
                   ", a.sRemarksx" & _
                   ", a.nAmtPaidx" & _
                   ", b.sCompnyNm" & _
                   ", a.dModified" & _
                   ", c.nTranTotl" & _
                   ", a.sReferNox" & _
                   ", a.sTermCode" & _
                   ", e.sTermName" & _
                " FROM Other_Payment_Received a" & _
                      " LEFT JOIN Term e" & _
                         " ON a.sTermCode = e.sTermIDxx" & _
                   ", Client_Master b" & _
                   ", CP_SO_Master c" & _
                   ", Payment_Processor d" & _
                " WHERE a.sClientID = b.sClientID" & _
                   " AND b.sClientID = d.sClientID" & _
                   " AND a.sSourceCd = " & strParm("CPSl") & _
                   " AND a.sSourceNo = c.sTransNox"
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
        Public nIntTotal As Decimal
        Public nAcctTerm As Integer
        Public nMonAmort As Decimal
        Public nAmtDuexx As Decimal
        Public sCompnyNm As String
        Public sCompnyID As String
        Public sCollName As String

        Public xTranAmtx As Decimal

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


