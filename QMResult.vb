'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Create Quickmatch Result based on Basic Personal Information of borrower
'
' Copyright 2007 and Beyond
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
'  XerSys [ 01/07/2010 01:17 pm ]
'     Start coding this object...
'  XerSys [ 01/11/2010 02:07 pm ]
'     Continue creating this object
'  XerSys [ 05/09/2012 01:34 pm ]
'     Integrate the blacklist of other customer to our quickmatch, maintaining our
'        existing policy in credit investigation.
'  XerSys [ 08/02/2013 01:48 pm ]
'     Incorporate blacklisted town and unallowed branch to QM
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports ggcAppDriver
Imports MySql.Data.MySqlClient
Imports ADODB

Public Class QMResult
    Private Const pxeMODULENAME As String = "clsQMResult"
    Private Const pxeEmptyDate As Date = "1900-01-01"
    Private p_oApp As GRider

    Private Enum compResult
        pxeEqual = 1
        pxeUncertain = 2
        pxeDifferent = 3
    End Enum

    Private Enum resultRelevance
        pxeRelevanceHi = 1
        pxeRelevanceMi = 2
        pxeRelevanceLo = 3
        pxeIrrelevant = 4
    End Enum

    Private p_sBranchCd As String
    Private p_oResult As DataTable
    Private p_oTmpRst As DataTable

    Private p_sTransNox As String
    Private p_sApplicNo As String
    Private p_sApplRslt As String
    Private p_sSpseRslt As String
    Private p_sResltCde As String

    Private p_sClientID As String
    Private p_sLastName As String
    Private p_sFrstName As String
    Private p_sMiddName As String
    Private p_cGenderCd As String
    Private p_cCvilStat As String
    Private p_dBirthDte As String
    Private p_sBirthPlc As String
    Private p_sAddressx As String
    Private p_sBrgyIDxx As String
    Private p_sTownIDxx As String

    Private p_sSpouseID As String
    Private p_sSLastNme As String
    Private p_sSFrstNme As String
    Private p_sSMiddNme As String
    Private p_cSGendrCd As String
    Private p_cSCvlStat As String
    Private p_dSBrthDte As String
    Private p_sSBrthPlc As String
    Private p_sSAddress As String
    Private p_sSBrgyIDx As String
    Private p_sSTownIDx As String

    Private p_sModelIDx As String
    Private p_nDownPaym As Double
    Private p_nAcctTerm As Integer
    Private p_bQMAllowed As Boolean

    Private pbInitTran As Boolean

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public ReadOnly Property ApplicationNo() As String
        Get
            ApplicationNo = p_sApplicNo
        End Get
    End Property

    Public ReadOnly Property TransNo() As String
        Get
            TransNo = p_sTransNox
        End Get
    End Property

    Public WriteOnly Property ApplicationNo() As String
        Set(Value As String)
            p_sApplicNo = Value
        End Set
    End Property

    Public Property Branch As String
        Get
            Return p_sBranchCd
        End Get
        Set(value As String)
            'If Product ID is LR then do allow changing of Branch
            If p_oApp.ProductID = "LRTrackr" Then
                p_sBranchCD = value
            End If
        End Set
    End Property

    Public ReadOnly Property Applicant(ByVal Index As String) As Object
        Get
            If pbInitTran = False Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    Applicant = p_sClientID
                Case "slastname"
                    Applicant = p_sLastName
                Case "sfrstname"
                    Applicant = p_sFrstName
                Case "smiddname"
                    Applicant = p_sMiddName
                Case "cgendercd"
                    Applicant = p_cGenderCd
                Case "ccvilstat"
                    Applicant = p_cCvilStat
                Case "dbirthdte"
                    Applicant = p_dBirthDte
                Case "sbirthplc"
                    Applicant = p_sBirthPlc
                Case "saddressx"
                    Applicant = p_sAddressx
                Case "sbrgyidxx"
                    Applicant = p_sBrgyIDxx
                Case "stownidxx"
                    Applicant = p_sTownIDxx
            End Select
        End Get
    End Property

    Public WriteOnly Property Applicant(ByVal Index As String) As Object
        Set(value As Object)
            If pbInitTran = False Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    p_sClientID = value
                Case "slastname"
                    p_sLastName = value
                Case "sfrstname"
                    p_sFrstName = value
                Case "smiddname"
                    p_sMiddName = value
                Case "cgendercd"
                    If Not (value = xeYes Or value = xeNo) Then Exit Property

                    p_cGenderCd = value
                Case "ccvilstat"
                    p_cCvilStat = value
                Case "dbirthdte"
                    If Not IsDate(Value) Or CDate(Value) > Date Then Exit Property

                    p_dBirthDte = value
                Case "sbirthplc"
                    p_sBirthPlc = value
                Case "saddressx"
                    p_sAddressx = value
                Case "sbrgyidxx"
                    p_sBrgyIDxx = value
                Case "stownidxx"
                    p_sTownIDxx = value
            End Select
        End Set
    End Property

    Public ReadOnly Property Spouse(ByVal Index As String) As Object
        Get
            If pbInitTran = False Then Exit Property
            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    Spouse = p_sSpouseID
                Case "slastname"
                    Spouse = p_sSLastNme
                Case "sfrstname"
                    Spouse = p_sSFrstNme
                Case "smiddname"
                    Spouse = p_sSMiddNme
                Case "cgendercd"
                    Spouse = p_cSGendrCd
                Case "ccvilstat"
                    Spouse = p_cSCvlStat
                Case "dbirthdte"
                    Spouse = p_dSBrthDte
                Case "sbirthplc"
                    Spouse = p_sSBrthPlc
                Case "saddressx"
                    Spouse = p_sSAddress
                Case "sbrgyidxx"
                    Spouse = p_sBrgyIDxx
                Case "stownidxx"
                    Spouse = p_sSTownIDx
            End Select
        End Get
    End Property

    Public WriteOnly Property Spouse(ByVal Index As String) As Object
        Set(Value As Object)
            If pbInitTran = False Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    p_sSpouseID = value
                Case "slastname"
                    p_sSLastNme = value
                Case "sfrstname"
                    p_sSFrstNme = value
                Case "smiddname"
                    p_sSMiddNme = value
                Case "cgendercd"
                    If Not (Value = xeLogical.YES Or Value = xeLogical.NO) Then Exit Property

                    p_cSGendrCd = value
                Case "ccvilstat"
                    p_cSCvlStat = value
                Case "dbirthdte"
                    If Not IsDate(Value) Or CDate(Value) > Now Then Exit Property

                    p_dSBrthDte = value
                Case "sbirthplc"
                    p_sSBrthPlc = value
                Case "saddressx"
                    p_sSAddress = value
                Case "sbrgyidxx"
                    p_sSBrgyIDx = value
                Case "stownidxx"
                    p_sSTownIDx = value
            End Select
        End Set
    End Property

    Public ReadOnly Property QuickMatchResult(ByVal Index As String) As String  
        Get
            If p_sResltCde = "" Then Exit Property
            Index = LCase(Index)
            Select Case Index
                Case "stransnox"
                    QuickMatchResult = p_sApplRslt
                Case "applicant"
                    QuickMatchResult = p_sApplRslt
                Case "spouse"
                    QuickMatchResult = p_sSpseRslt
                Case "application"
                    QuickMatchResult = p_sResltCde
            End Select
        End Get
    End Property

    'To implement this logic, always call the InitTransaction
    Public ReadOnly Property Result() As DataTable
        Get
            If p_sApplRslt = "" Then
                Result = Nothing
            Else
                Result = p_oResult
            End If
        End Get
    End Property

    Public ReadOnly Property ResultDetail(ByVal Row As Integer, ByVal Index As String) As String
        Get
            If p_sResltCde = "" Then Exit Property

            If Row > p_oResult.Rows.Count - 1 Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sfullname"
                    ResultDetail = p_oResult.Rows(Row)("sFullName")
                Case "sresltcde"
                    ResultDetail = p_oResult.Rows(Row)("sResltCde")
                Case "sacctnmbr"
                    ResultDetail = p_oResult.Rows(Row)("sAcctNmbr")
                Case "smcsonmbr"
                    ResultDetail = p_oResult.Rows(Row)("sMCSONmbr")
                Case "sapplnmbr"
                    ResultDetail = p_oResult.Rows(Row)("sApplNmbr")
            End Select
        End Get
    End Property

    Public ReadOnly Property Term(ByVal Index As String) As Object
        Get
            Index = LCase(Index)
            Select Case Index
                Case "smodelidx"
                    Term = p_sModelIDx
                Case "ndownpaym"
                    Term = p_nDownPaym
                Case "nacctterm"
                    Term = p_nAcctTerm
            End Select
        End Get
    End Property

    Public WriteOnly Property Term(ByVal Index As String) As Object
        Set(value As Object)
            Index = LCase(Index)
            Select Case Index
                Case "smodelidx"
                    p_sModelIDx = Value
                Case "ndownpaym"
                    If Not IsNumeric(Value) Then Exit Property

                    p_nDownPaym = Value
                Case "nacctterm"
                    If Not IsNumeric(Value) Then Exit Property

                    p_nAcctTerm = Value
            End Select
        End Set
    End Property

    Function InitTransaction() As Boolean
        Dim lsOldProc As String
        Dim lsSQL As String

        lsOldProc = "InitTransaction"
        'On Error GoTo errProc

        If p_sBranchCd = String.Empty Then p_sBranchCd = p_oApp.BranchCode

        p_bQMAllowed = isBranchAllowed()

        pbInitTran = True
        InitTransaction = True

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    ' Return value
    ' * 2  Result
    ' * 2  Term
    ' * 2  Rating
    ' * 2  Year
    ' * 1  +/-
    ' * 2  Blacklisted Year to other Dealer
    Function QuickMatch() As String
        Dim lsOldProc As String
        Dim lsAppBList As String
        Dim lsSpsBList As String

        lsOldProc = "QuickMatch"
        'On Error GoTo errProc

        p_sResltCde = ""

        Call initResult()

        p_sApplRslt = MatchApplicant(p_sClientID, p_sLastName, p_sFrstName, p_sMiddName, _
                             p_dBirthDte, p_sBirthPlc, p_sTownIDxx, p_sBrgyIDxx, p_sAddressx)

        ' »»» XerSys 2012-05-09
        '  Match to blacklisted account of other dealer
        lsAppBList = match2Other(p_sLastName, p_sFrstName, p_sMiddName, _
                             p_dBirthDte, p_sTownIDxx, p_sAddressx)

        If p_sSLastNme <> "" Then
            p_sSpseRslt = MatchApplicant(p_sSpouseID, p_sSLastNme, p_sSFrstNme, p_sSMiddNme, _
                                 p_dSBrthDte, p_sSBrthPlc, p_sSTownIDx, p_sSBrgyIDx, p_sSAddress)

            ' »»» XerSys 2012-05-09
            '  Match to blacklisted account of other dealer
            lsSpsBList = match2Other(p_sSLastNme, p_sSFrstNme, p_sSMiddNme, _
                                 p_dSBrthDte, p_sSTownIDx, p_sSAddress)

            ' after getting the individual quickmatch process the application result
            If Left(p_sApplRslt, 2) = Left(p_sSpseRslt, 2) Then
                p_sResltCde = p_sApplRslt
            Else
                Select Case Left(p_sApplRslt, 2)
                    Case "AP"
                        Select Case Left(p_sSpseRslt, 2)
                            Case "DA", "SA"
                                p_sResltCde = "SA" & Mid(p_sApplRslt, 3)
                            Case Else
                                p_sResltCde = p_sApplRslt
                        End Select
                    Case "CI"
                        Select Case Left(p_sSpseRslt, 2)
                            Case "DA", "SA", "SV"
                                p_sResltCde = p_sSpseRslt
                            Case Else
                                p_sResltCde = p_sApplRslt
                        End Select
                    Case "SA", "SV"
                        p_sResltCde = p_sApplRslt
                    Case "DA"
                        p_sResltCde = p_sApplRslt
                    Case "BA"
                        p_sResltCde = p_sApplRslt
                End Select
            End If
        Else
            p_sResltCde = p_sApplRslt
        End If

        ' »»» XerSys 2012-05-09
        '  Process the other dealers result then add it to the result of our database
        If Left(lsAppBList, 1) = "P" Then
        ElseIf Left(lsSpsBList, 1) = "P" Then
            lsAppBList = lsSpsBList
        ElseIf Left(lsAppBList, 1) = "U" Then
        ElseIf Left(lsSpsBList, 1) = "U" Then
            lsAppBList = lsSpsBList
        End If

        Select Case Left(p_sResltCde, 2)
            Case "DA", "SA", "SV", "PA"
                ' any result from other dealer, don't affect the result
                p_sResltCde = p_sResltCde & Mid(lsAppBList, 2)
            Case "CI"
                ' other dealers result is relevant to our result, seek help from collection department
                If Left(lsAppBList, 1) = "P" Then
                    p_sResltCde = "SA" & Mid(p_sResltCde, 3) & Mid(lsAppBList, 2)
                ElseIf Left(lsAppBList, 1) = "U" Then
                    p_sResltCde = "SV" & Mid(p_sResltCde, 3) & Mid(lsAppBList, 2)
                End If
            Case "AP"
                If Left(lsAppBList, 1) = "P" Then
                    ' approve customer means repeat customer, follow the latest result
                    If CInt(Mid(lsAppBList, 2)) > CInt(Mid(p_sResltCde, 7, 2)) Then
                        p_sResltCde = "SA" & Mid(p_sResltCde, 3) & Mid(lsAppBList, 2)
                    Else
                        p_sResltCde = p_sResltCde & Mid(lsAppBList, 2)
                    End If
                ElseIf Left(lsAppBList, 1) = "U" Then
                    p_sResltCde = p_sResltCde & Mid(lsAppBList, 2)
                End If
        End Select

        ' save the retrieve result
        If Not saveResult() Then
            p_sResltCde = ""
            GoTo endProc
        End If

        QuickMatch = p_sResltCde

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function MatchApplicant( _
          ByVal lsClientID As String, _
          ByVal lsLastName As String, _
          ByVal lsFrstName As String, _
          ByVal lsMiddName As String, _
          ByVal ldBirthDte As String, _
          ByVal lsBirthPlc As String, _
          ByVal lsTownIDxx As String, _
          ByVal lsBrgyIDxx As String, _
          ByVal lsAddressx As String) As String
        Dim loRS As Recordset
        Dim lsOldProc As String
        Dim lsSQL As String

        Dim lnMiddName As compResult
        Dim lnBirthDte As compResult
        Dim lnBirthPlc As compResult
        Dim lnTownIDxx As compResult
        Dim lnAddressx As compResult

        Dim lsResult As String
        Dim lsRating As String
        Dim lnTerm As Integer
        Dim ldTransact As Date
        Dim lnRecExist As compResult
        Dim lsAcctNmbr As String
        Dim lsTransNox As String
        Dim lnRelevance As resultRelevance

        Dim lbRecExist As Boolean
        Dim lbRecUnsre As Boolean
        Dim lbBLAddress As Boolean

        lsOldProc = "MatchApplicant"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", a.sLastName" & _
                    ", CONCAT(a.sFrstName, IF(TRIM(IFNULL(a.sSuffixNm, '')) = '', '', CONCAT(' ', a.sSuffixNm))) sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.sAddressx" & _
                    ", a.sTownIDxx" & _
                    ", a.dBirthDte" & _
                    ", a.sBirthPlc" & _
                    ", b.sAcctNmbr" & _
                    ", b.nAcctTerm" & _
                    ", b.cRatingxx" & _
                    ", b.dClosedxx" & _
                    ", b.dPurchase dTransact" & _
                    ", b.cAcctStat cTranStat" & _
                    ", c.sResltCde" & _
                 " FROM Client_Master a" & _
                    " LEFT JOIN MC_AR_Master b" & _
                       " ON a.sClientID = b.sClientID" & _
                    " LEFT JOIN MC_LR_QuickMatch_Result c" & _
                       " ON b.sAcctNmbr = c.sAcctNmbr" & _
                 " WHERE a.sLastName LIKE " & strParm(lsLastName) & _
                    " AND a.sFrstName LIKE " & strParm(lsFrstName)
        lsSQL = lsSQL & _
                 " UNION SELECT" & _
                    "  a.sClientID" & _
                    ", a.sLastName" & _
                    ", CONCAT(a.sFrstName, IF(TRIM(IFNULL(a.sSuffixNm, '')) = '', '', CONCAT(' ', a.sSuffixNm))) sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.sAddressx" & _
                    ", a.sTownIDxx" & _
                    ", a.dBirthDte" & _
                    ", a.sBirthPlc" & _
                    ", b.sTransNox sAcctNmbr" & _
                    ", 0 nAcctTerm" & _
                    ", 'NB' cRatingxx" & _
                    ", CAST('1900/01/01' AS DATE) dClosedxx" & _
                    ", b.dTransact" & _
                    ", b.cTranStat" & _
                    ", d.sResltCde"
        lsSQL = lsSQL & _
                 " FROM Client_Master a" & _
                    " LEFT JOIN MC_SO_Master b" & _
                       " ON a.sClientID = b.sClientID" & _
                          " AND b.cTranType = " & strParm("0") & _
                    " LEFT JOIN MC_SO_Detail c" & _
                       " ON b.sTransNox = c.sTransNox" & _
                    " LEFT JOIN MC_LR_QuickMatch_Result d" & _
                       " ON b.sTransNox = d.sMCSONmbr" & _
                          " AND b.sClientID = d.sClientID" & _
                 " WHERE a.sLastName LIKE " & strParm(lsLastName) & _
                    " AND a.sFrstName LIKE " & strParm(lsFrstName)
        lsSQL = lsSQL & _
                 " UNION SELECT" & _
                    "  a.sClientID" & _
                    ", a.sLastName" & _
                    ", CONCAT(a.sFrstName, IF(TRIM(IFNULL(a.sSuffixNm, '')) = '', '', CONCAT(' ', a.sSuffixNm))) sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.sAddressx" & _
                    ", a.sTownIDxx" & _
                    ", a.dBirthDte" & _
                    ", a.sBirthPlc" & _
                    ", b.sTransNox sAcctNmbr" & _
                    ", b.nAcctTerm" & _
                    ", 'PA' cRatingxx" & _
                    ", CAST('1900/01/01' AS DATE) dClosedxx" & _
                    ", b.dAppliedx dTransact" & _
                    ", b.cTranStat" & _
                    ", c.sResltCde"
        lsSQL = lsSQL & _
                 " FROM Client_Master a" & _
                    " LEFT JOIN MC_Credit_Application b" & _
                       " ON a.sClientID = b.sClientID" & _
                          " AND b.cTranStat <> " & strParm(xeTranStat.TRANS_UNKNOWN) & _
                    " LEFT JOIN MC_LR_QuickMatch_Result c" & _
                       " ON b.sTransNox = c.sApplNmbr" & _
                          " AND b.sClientID = c.sClientID" & _
                 " WHERE a.sLastName LIKE " & strParm(lsLastName) & _
                    " AND a.sFrstName LIKE " & strParm(lsFrstName) & _
                 " ORDER BY dTransact DESC"

        ' XerSys - 2013-08-07
        '  Check if address is blacklisted area
        lbBLAddress = isAddBlackList(lsTownIDxx, lsBrgyIDxx)

        loRS = New Recordset
        With loRS
            Debug.Print(lsSQL)
            .Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

            If .EOF Then
                If lbBLAddress Then
                    lsResult = "BA"
                Else
                    lsResult = "CI"
                End If
                lsRating = "00"

                lsResult = getResult(lsResult, 0, lsRating, CDate("1900-01-01"), pxeDifferent)
                Call addResult(lsClientID, lsLastName & ", " & lsFrstName & " " & lsMiddName, _
                      lsResult, pxeIrrelevant, "", "", "", "1900-01-01", "")

                MatchApplicant = lsResult
                GoTo endProc
            End If

            Do While .EOF = False
                If compareName(.Fields("sLastName"), lsLastName) = False Then GoTo moveToNext

                If compareName(.Fields("sFrstName"), lsFrstName) = False Then GoTo moveToNext

                lnMiddName = compareMiddName(IFNull(.Fields("sMiddName"), ""), lsMiddName)
                lnBirthDte = compareBirthDate(IFNull(.Fields("dBirthDte"), "1900-01-01"), ldBirthDte)
                lnBirthPlc = compareBirthPlace(IFNull(.Fields("sBirthPlc")), lsBirthPlc)
                lnTownIDxx = compareTown(IFNull(.Fields("sTownIDxx")), lsTownIDxx)
                lnAddressx = compareAddress(IFNull(.Fields("sAddressx"), ""), lsAddressx)

                If (lnMiddName + lnBirthDte + lnBirthPlc) <= 4 Then
                    lnRecExist = pxeEqual
                ElseIf (lnMiddName + lnBirthDte + lnBirthPlc) <= 6 Then
                    lnRecExist = pxeUncertain
                Else
                    lnRecExist = pxeDifferent
                End If

                lsResult = "CI"
                lsRating = "00"
                lnTerm = IIf(IsNull(.Fields("nAcctTerm")), 0, .Fields("nAcctTerm"))
                ldTransact = IIf(IsNull(.Fields("dTransact")), pxeEmptyDate, .Fields("dTransact"))
                lsAcctNmbr = ""
                lsTransNox = ""
                lnRelevance = pxeRelevanceLo

                Select Case lnRecExist
                    Case pxeDifferent
                        If Not (lbRecExist Or lbRecUnsre) Then
                            If lbBLAddress Then
                                lsResult = "BA"
                            Else
                                lsResult = "CI"
                            End If
                        End If
                    Case pxeEqual, pxeUncertain
                        If lnRecExist = pxeEqual Then
                            lbRecExist = True
                        Else
                            If lbRecExist Then
                                GoTo moveToNext
                            Else
                                lbRecUnsre = True
                            End If
                        End If

                        If IsNull(.Fields("sAcctNmbr")) Then
                            lsRating = "00"
                            If lbBLAddress Then
                                lsResult = "BA"
                            Else
                                lsResult = "CI"
                            End If
                            '               lsResult = getResult(lsResult, 0, "00", CDate("1900-01-01"), lnRecExist)
                            '               Call addResult(.Fields("sClientID"), _
                            '                     .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                            '                     lsResult, pxeRelevanceLo, "", "", "", "1900-01-01", "")
                        Else
                            Select Case .Fields("cRatingxx")
                                Case "NB" ' Cash Sales
                                    lsRating = "NB"
                                    ldTransact = .Fields("dTransact")

                                    '                  lsResult = getResult("CI", 0, "NB", .Fields("dTransact"), lnRecExist)
                                    '                  Call addResult(.Fields("sClientID"), _
                                    '                        .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                                    '                        lsResult, pxeRelevanceLo, "", _
                                    '                        .Fields("sAcctNmbr"), "", .Fields("dTransact"), .Fields("sResltCde"))
                                Case "PA" ' With Pending Application
                                    'If we are dealing with this application no then skip
                                    If .Fields("sAcctNmbr") = p_sApplicNo Then GoTo moveToNext
                                    lnRelevance = pxeRelevanceMi
                                    lsTransNox = .Fields("sAcctNmbr")
                                    lsResult = prcPendingApplication(.Fields("cTranStat"), .Fields("dTransact"), .Fields("nAcctTerm"), lsRating, lnRecExist)
                                    '                  Call addResult(.Fields("sClientID"), _
                                    '                        .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                                    '                        lsResult, pxeRelevanceMi, "", "", _
                                    '                        .Fields("sAcctNmbr"), .Fields("dTransact"), .Fields("sResltCde"))
                                Case Else
                                    lsRating = .Fields("cRatingxx")
                                    lsAcctNmbr = .Fields("sAcctNmbr")
                                    ldTransact = IIf(.Fields("cTranStat") = xeActStatActive, .Fields("dTransact"), IFNull(.Fields("dClosedxx"), pxeEmptyDate))
                                    If .Fields("cTranStat") = xeActStatActive Then
                                        lsResult = prcRepeatAccount(.Fields("cTranStat"), _
                                                       .Fields("dTransact"), _
                                                       .Fields("nAcctTerm"), _
                                                       lsRating, _
                                                       p_sModelIDx, _
                                                       p_nDownPaym, _
                                                       lnRecExist)
                                    Else
                                        lsResult = prcRepeatAccount(.Fields("cTranStat"), _
                                                       .Fields("dClosedxx"), _
                                                       .Fields("nAcctTerm"), _
                                                       lsRating, _
                                                       p_sModelIDx, _
                                                       p_nDownPaym, _
                                                       lnRecExist)

                                    End If
                                    '                  Call addResult(.Fields("sClientID"), _
                                    '                           .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                                    '                           lsResult, _
                                    '                           IIf(lnRecExist = pxeEqual, pxeRelevanceHi, pxeRelevanceMi), _
                                    '                           .Fields("sAcctNmbr"), _
                                    '                           "", "", _
                                    '                           IIf(.Fields("cTranStat") = xeActStatActive, .Fields("dTransact"), IFNull(.Fields("dClosedxx"), "1900-01-01")), _
                                    '                           IFNull(.Fields("sResltCde")))
                            End Select
                        End If
                End Select
                If Not p_bQMAllowed Then
                    ' XerSys - 2013-08-14
                    '  Only approved account can be given to branches that are not allowed to issue QM Number
                    If lsResult <> "AP" Then
                        lsResult = "SA"
                    End If
                End If

                lsResult = getResult(lsResult, lnTerm, lsRating, ldTransact, lnRecExist)
                Call addResult(.Fields("sClientID"), _
                         .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                         lsResult, _
                         IIf(lnRecExist = pxeEqual, pxeRelevanceHi, pxeRelevanceMi), _
                         lsAcctNmbr, _
                         "", lsTransNox, _
                         IIf(.Fields("cTranStat") = xeActStatActive, .Fields("dTransact"), IFNull(.Fields("dClosedxx"), pxeEmptyDate)), _
                         IFNull(.Fields("sResltCde")))

moveToNext:
                .MoveNext()
            Loop
        End With

        With p_oTmpRst
            .Sort = "nRelevnce, dTransact DESC"

            If lbRecExist Or lbRecUnsre Then
                If lbRecExist Then
                    .Filter = "nComparsn = " & pxeEqual
                Else
                    .Filter = "nComparsn = " & pxeUncertain
                End If
                '      Else
                '         MatchApplicant = lsResult
                '         Call addResult(lsClientID, lsLastName & ", " & lsFrstName & " " & lsMiddName, _
                '               lsResult, pxeIrrelevant, "", "", "", "1900-01-01", "")
            End If
            MatchApplicant = .Fields("sResltCde")
        End With

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function match2Other( _
          lsLastName As String, _
          lsFrstName As String, _
          lsMiddName As String, _
          ldBirthDte As String, _
          lsTownIDxx As String, _
          lsAddressx As String) As String
        Dim loRS As Recordset
        Dim lsOldProc As String
        Dim lsSQL As String

        Dim lnMiddName As compResult
        Dim lnBirthDte As compResult
        Dim lnTownIDxx As compResult
        Dim lnAddressx As compResult

        Dim lsResult As String
        Dim lsTmpReslt As String

        lsOldProc = "match2Other"
        'On Error GoTo errProc

        lsResult = ""

        lsSQL = "SELECT" & _
                    "  a.sClientID" & _
                    ", a.sLastName" & _
                    ", a.sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.sAddressx" & _
                    ", IFNULL(a.sTownIDxx, '') sTownIDxx" & _
                    ", a.sProvIDxx" & _
                    ", c.sProvIDxx xProvIDxx" & _
                    ", a.dBirthDte" & _
                    ", a.nBListdYr" & _
                 " FROM Client_Blacklist a" & _
                       " LEFT JOIN TownCity b" & _
                          " ON b.sTownIDxx = " & strParm(lsTownIDxx) & _
                       " LEFT JOIN Province c" & _
                          " ON b.sProvIDxx = c.sProvIDxx" & _
                 " WHERE a.sLastName LIKE " & strParm(lsLastName) & _
                    " AND a.sFrstName LIKE " & strParm(lsFrstName) & _
                 " ORDER BY nBListdYr DESC"

        loRS = New Recordset
        With loRS
            Debug.Print(lsSQL)
            .Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

            If .EOF Then GoTo endProc

            Do Until .EOF
                If compareName(.Fields("sLastName"), lsLastName) = False Then GoTo moveToNext

                If compareName(.Fields("sFrstName"), lsFrstName) = False Then GoTo moveToNext

                lnMiddName = compareMiddName(IFNull(.Fields("sMiddName"), ""), lsMiddName)
                lnBirthDte = compareBirthDate(IFNull(.Fields("dBirthDte"), "1900-01-01"), ldBirthDte)

                ' if no town exist, then use province
                If .Fields("sTownIDxx") = "" Then
                    lnTownIDxx = compareTown(.Fields("sProvIDxx"), .Fields("xProvIDxx"))
                Else
                    lnTownIDxx = compareTown(.Fields("sTownIDxx"), lsTownIDxx)
                End If
                lnAddressx = compareAddress(IFNull(.Fields("sAddressx"), ""), lsAddressx)

                '         Debug.Print lnMiddName, lnBirthDte, lnBirthPlc, lnTownIDxx, lnAddressx

                If (lnMiddName + lnBirthDte + lnTownIDxx) <= 4 Then
                    lsTmpReslt = "P" 'pxeEqual
                ElseIf (lnMiddName + lnBirthDte + lnTownIDxx) <= 6 Then
                    lsTmpReslt = "U" 'pxeUncertain
                Else
                    lsTmpReslt = pxeDifferent
                End If
                lsTmpReslt = lsTmpReslt & Format(CDate(.Fields("nBListdYr") & "/01/01"), "yy")

                If lsResult = "" Then
                    lsResult = lsTmpReslt
                Else
                    If Left(lsResult, 1) = "P" Then
                        If Left(lsTmpReslt, 1) = "P" Then
                            If CInt(Mid(lsResult, 2)) < CInt(Mid(lsTmpReslt, 2)) Then
                                lsResult = lsTmpReslt
                            End If
                        End If
                    ElseIf Left(lsTmpReslt, 1) = "P" Then
                        lsResult = lsTmpReslt
                    ElseIf Left(lsTmpReslt, 1) = "U" Then
                        If CInt(Mid(lsResult, 2)) < CInt(Mid(lsTmpReslt, 2)) Then
                            lsResult = lsTmpReslt
                        End If
                    End If
                End If

moveToNext:
                .MoveNext()
            Loop
        End With

        match2Other = lsResult

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function isAddBlackList(ByVal lsTownIDxx As String, ByVal lsBrgyIDxx As String) As Boolean
        Dim lsOldProc As String
        Dim lsSQL As String
        Dim loRS As Recordset

        lsSQL = "SELECT cBlackLst FROM TownCity" & _
                 " WHERE sTownIDxx = " & strParm(lsTownIDxx)
        loRS = New Recordset
        loRS.Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

        If Not loRS.EOF Then
            If loRS("cBlackLst") = xeYes Then
                isAddBlackList = True
                GoTo endProc
            End If
        End If

        lsSQL = "SELECT cBlackLst FROM Barangay" & _
                 " WHERE sBrgyIDxx = " & strParm(lsBrgyIDxx)
        loRS = New Recordset
        loRS.Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

        If Not loRS.EOF Then
            If loRS("cBlackLst") = xeYes Then
                isAddBlackList = True
                GoTo endProc
            End If
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function isBranchAllowed() As Boolean
        Dim lsOldProc As String
        Dim lsSQL As String
        Dim loRS As Recordset

        lsSQL = "SELECT cAllowQMx FROM Branch_Others" & _
                 " WHERE sBranchCd = " & strParm(p_sBranchCd)
        loRS = New Recordset
        loRS.Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

        If Not loRS.EOF Then
            If loRS("cAllowQMx") = xeYes Then
                isBranchAllowed = True
                GoTo endProc
            End If
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function getResult( _
          ByVal lsResult As String, _
          ByVal lnTerm As Integer, _
          ByVal lsRating As String, _
          ByVal ldTransact As Date, _
          ByVal lnExist As compResult) As String
        getResult = lsResult & Format(lnTerm, "00") & _
                       lsRating & Format(ldTransact, "yy") & _
                       IIf(lnExist = pxeDifferent, "N", IIf(lnExist = pxeEqual, "P", "U"))
    End Function

    Private Function initResult() As Boolean
        Dim lsOldProc As String
        Dim lsSQL As String

        lsOldProc = "initResult"
        'On Error GoTo errProc

        p_oTmpRst = New Recordset
        With p_oTmpRst
            .Fields.Append("sClientID", adVarChar, 12)
            .Fields.Append("sFullName", adVarChar, 512)
            .Fields.Append("nRelevnce", adInteger, 4)
            .Fields.Append("dTransact", adDate)
            .Fields.Append("nComparsn", adInteger, 4)
            .Fields.Append("sResltCde", adVarChar, 11)
            .Fields.Append("sAcctNmbr", adVarChar, 10)
            .Fields.Append("sMCSONmbr", adVarChar, 12)
            .Fields.Append("sApplNmbr", adVarChar, 12)
            .Fields.Append("cAddRecrd", adChar, 1)

            .Open()
        End With

        p_oResult = New Recordset
        With p_oResult
            .Fields.Append("nEntryNox", adInteger, 4)
            .Fields.Append("sFullName", adVarChar, 512)
            .Fields.Append("sResltCde", adVarChar, 11)
            .Fields.Append("sAcctNmbr", adVarChar, 10)
            .Fields.Append("sMCSONmbr", adVarChar, 12)
            .Fields.Append("sApplNmbr", adVarChar, 12)
            .Open()
        End With
        initResult = True

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function addResult( _
          lsClientID As String, _
          lsFullName As String, _
          lsResltCde As String, _
          lnRelevnce As resultRelevance, _
          lsAcctNmbr As String, _
          lsMCSONmbr As String, _
          lsApplNmbr As String, _
          ldTransact As Date, _
          lsPrevRslt As Object) As Boolean
        Dim lsOldProc As String
        Dim lcAddRecord As Integer

        lsOldProc = "addResult"
        'On Error GoTo errProc

        With p_oTmpRst
            .Filter = "sResltCde = " & strParm(lsResltCde) & " AND nRelevnce = " & lnRelevnce
            If .EOF Then
                lcAddRecord = xeYes
            Else
                lcAddRecord = xeNo
            End If
            .Filter = ""

            If Not IsNull(lsPrevRslt) Then
                If lsPrevRslt = lsResltCde Then
                    lcAddRecord = xeNo
                End If
            End If

            .AddNew()
            .Fields("sClientID") = lsClientID
            .Fields("sFullName") = lsFullName
            .Fields("nRelevnce") = lnRelevnce
            .Fields("sResltCde") = lsResltCde
            .Fields("dTransact") = ldTransact
            .Fields("sAcctNmbr") = lsAcctNmbr
            .Fields("sMCSONmbr") = lsMCSONmbr
            .Fields("sApplNmbr") = lsApplNmbr
            .Fields("cAddRecrd") = lcAddRecord

            '      If Not IsNull(lsPrevRslt) Then
            '         If lsPrevRslt = lsResltCde Then
            '            .Fields("cAddRecrd") = xeNo
            '         End If
            '      End If

            Select Case Right(lsResltCde, 1)
                Case "P"
                    .Fields("nComparsn") = pxeEqual
                Case "N"
                    .Fields("nComparsn") = pxeDifferent
                Case Else
                    .Fields("nComparsn") = pxeUncertain
            End Select
        End With

        addResult = True

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function saveResult() As Boolean
        Dim lsOldProc As String
        Dim lsTransNox As String
        Dim lsSQL As String
        Dim lnCtr As Integer

        lsOldProc = "saveResult"
        'On Error GoTo errProc

        With p_oTmpRst
            .MoveFirst()

            lsTransNox = GetNextCode("MC_LR_QuickMatch", "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
            lnCtr = 1

            If p_sApplicNo = "" Then p_oAppDrivr.BeginTrans()
            Do While .EOF = False
                If .Fields("cAddRecrd") = xeYes Then
                    lsSQL = "INSERT INTO MC_LR_QuickMatch_Result SET" & _
                                "  sTransNox = " & strParm(lsTransNox) & _
                                ", sClientID = " & strParm(.Fields("sClientID")) & _
                                ", nEntryNox = " & lnCtr & _
                                ", sResltCde = " & strParm(.Fields("sResltCde")) & _
                                ", sAcctNmbr = " & strParm(.Fields("sAcctNmbr")) & _
                                ", sMCSONmbr = " & strParm(.Fields("sMCSONmbr")) & _
                                ", sApplNmbr = " & strParm(.Fields("sApplNmbr")) & _
                                ", dModified = " & dateParm(p_oAppDrivr.ServerDate)

                    If p_oAppDrivr.Execute(lsSQL, "MC_LR_QuickMatch_Result") <= 0 Then
                        MsgBox("Unable to Save Changes")
                        GoTo endWithRoll
                    End If

                    lnCtr = lnCtr + 1
                End If

                p_oResult.AddNew()
                p_oResult("nEntryNox") = p_oResult.RecordCount
                p_oResult("sFullName") = .Fields("sFullName")
                p_oResult("sResltCde") = .Fields("sResltCde")
                p_oResult("sAcctNmbr") = .Fields("sAcctNmbr")
                p_oResult("sMCSONmbr") = .Fields("sMCSONmbr")
                p_oResult("sApplNmbr") = .Fields("sApplNmbr")

                .MoveNext()
            Loop

            lsSQL = "INSERT INTO MC_LR_QuickMatch SET" & _
                        "  sTransNox = " & strParm(lsTransNox) & _
                        ", sApplicNo = " & strParm(p_sApplicNo) & _
                        ", sLastName = " & strParm(p_sLastName) & _
                        ", sFrstName = " & strParm(p_sFrstName) & _
                        ", sMiddName = " & strParm(p_sMiddName) & _
                        ", dBirthDte = " & dateParm(p_dBirthDte) & _
                        ", sBirthPlc = " & strParm(p_sBirthPlc) & _
                        ", sTownIDxx = " & strParm(p_sTownIDxx) & _
                        ", sSLastNme = " & strParm(p_sSLastNme) & _
                        ", sSFrstNme = " & strParm(p_sSFrstNme) & _
                        ", sSMiddNme = " & strParm(p_sSMiddNme)

            If p_dSBrthDte = "" Then
                lsSQL = lsSQL & ", dSBrthDte = NULL"
            Else
                lsSQL = lsSQL & ", dSBrthDte = " & dateParm(p_dSBrthDte)
            End If

            lsSQL = lsSQL & _
                        ", sSBrthPlc = " & strParm(p_sSBrthPlc) & _
                        ", sSTownIDx = " & strParm(p_sSTownIDx) & _
                        ", sModelIDx = " & strParm(p_sModelIDx) & _
                        ", nDownPaym = " & p_nDownPaym & _
                        ", nAcctTerm = " & p_nAcctTerm & _
                        ", sResltCde = " & strParm(p_sResltCde) & _
                        ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate)

            If p_oAppDrivr.Execute(lsSQL, "MC_LR_QuickMatch") <= 0 Then
                MsgBox("Unable to Save Changes")
                GoTo endWithRoll
            End If

            If p_sApplicNo = "" Then p_oAppDrivr.CommitTrans()
        End With

        p_sTransNox = lsTransNox

        saveResult = True

endProc:
        Exit Function
endWithRoll:
        If p_sApplicNo = "" Then p_oAppDrivr.RollbackTrans()
        GoTo endProc
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareName( _
          lsNameNRec As String, _
          lsNameNew As String) As compResult
        Dim lsOldProc As String
        Dim lasTemp() As String
        Dim lnCtr As Integer

        lsOldProc = "compareName"
        'On Error GoTo errProc
        compareName = pxeDifferent

        If StrComp(Trim(lsNameNRec), Trim(lsNameNew), vbTextCompare) <> 0 Then
            ' if not exactly equal, remove the conjunction
            If InStr(1, lsNameNRec, "&/") > 0 Then
                lasTemp = Split(lsNameNRec, "&/")
            ElseIf InStr(1, lsNameNRec, "/") > 0 Then
                lasTemp = Split(lsNameNRec, "/")
            Else
                GoTo endProc
            End If

            ' now check if name is exactly equal
            For lnCtr = 0 To UBound(lasTemp)
                If StrComp(Trim(lasTemp(lnCtr)), Trim(lsNameNew), vbTextCompare) <> 0 Then
                    compareName = pxeEqual
                    GoTo endProc
                End If
            Next
        Else
            compareName = pxeEqual
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareMiddName( _
          ByVal lsNameNRec As String, _
          ByVal lsNameNew As String) As compResult
        Dim lsOldProc As String
        Dim lasTemp() As String
        Dim lnCtr As Integer

        lsOldProc = "compareMiddName"
        'On Error GoTo errProc
        compareMiddName = pxeDifferent

        Call IFNull(lsNameNRec, "")
        If StrComp(Trim(lsNameNRec), Trim(lsNameNew), vbTextCompare) <> 0 Then
            ' if not exactly equal, remove the conjunction
            If InStr(1, lsNameNRec, "&/") > 0 Then
                lasTemp = Split(lsNameNRec, "&/")
            ElseIf InStr(1, lsNameNRec, "/") > 0 Then
                lasTemp = Split(lsNameNRec, "/")
            Else
                If lsNameNRec = "" Then
                    compareMiddName = pxeUncertain
                Else
                    lnCtr = InStr(lsNameNRec, ".")
                    lsNameNRec = IIf(lnCtr = 0, lsNameNRec, lnCtr - 1)
                    'Remove this process and replace it with the above
                    '            lsNameNRec = Mid(lsNameNRec, 1, IIf(lnCtr = 0, 0, lnCtr - 1))

                    If StrComp(Trim(lsNameNRec), Left(Trim(lsNameNew), Len(Trim(lsNameNRec))), vbTextCompare) = 0 Then
                        compareMiddName = pxeUncertain
                    End If
                End If
                GoTo endProc
            End If

            ' now check if name is exactly equal
            For lnCtr = 0 To UBound(lasTemp)
                If StrComp(Trim(lasTemp(lnCtr)), Trim(lsNameNew), vbTextCompare) <> 0 Then
                    compareMiddName = pxeEqual
                    GoTo endProc
                End If
            Next
        Else
            compareMiddName = pxeEqual
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareBirthDate( _
          ByVal ldInfoNRec As Date, _
          ByVal ldNewInfo As Date) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareBirthDate"
        'On Error GoTo errProc

        compareBirthDate = pxeDifferent
        Debug.Print(ldNewInfo, ldInfoNRec)
        If ldInfoNRec <> ldNewInfo Then
            '      If DateDiff("d", ldInfoNRec, CDate("1900-01-01")) = 0 Then
            '         compareBirthDate = pxeUncertain
            '      End If
            If DateDiff("d", ldInfoNRec, CDate("1900-01-01")) = 0 _
             Or DateDiff("d", ldInfoNRec, CDate("1901-01-01")) = 0 _
             Or DateDiff("d", ldInfoNRec, CDate("1999-01-01")) = 0 Then
                compareBirthDate = pxeUncertain
            End If

        Else
            compareBirthDate = pxeEqual
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareBirthPlace( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareBirthPlace"
        'On Error GoTo errProc
        compareBirthPlace = pxeDifferent

        If lsInfoNRec <> lsNewInfo Then
            If lsInfoNRec = "" Then
                compareBirthPlace = pxeUncertain
            End If
        Else
            compareBirthPlace = pxeEqual
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareTown( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareTown"
        'On Error GoTo errProc

        If lsInfoNRec = lsNewInfo Then
            compareTown = pxeEqual
        Else
            compareTown = pxeDifferent
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function compareAddress( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareAddress"
        'On Error GoTo errProc
        compareAddress = pxeDifferent

        If StrComp(Trim(lsInfoNRec), Trim(lsNewInfo), vbTextCompare) <> 0 Then
            If lsInfoNRec = "" Then
                compareAddress = pxeUncertain
            End If
        Else
            compareAddress = pxeEqual
        End If

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Function prcPendingApplication( _
          lcTranStat As xeTransactionStatus, _
          ldTransact As Date, _
          lnAcctTerm As Integer, _
          lsRating As String, _
          lnRecExist As compResult) As String
        Dim lsResult As String
        '   Dim lsRating As String

        Select Case lcTranStat
            Case xeStatePosted
                If DateDiff("m", ldTransact, p_oAppDrivr.ServerDate) > 3 Then
                    lsResult = "CI"
                Else
                    lsResult = "AP"
                End If
                lsRating = "AA"
            Case xeStateCancelled
                If DateDiff("m", ldTransact, p_oAppDrivr.ServerDate) > 24 Then
                    lsResult = "CI"
                Else
                    lsResult = "DA"
                End If
                lsRating = "DA"
            Case Else
                lsResult = "PA"
                lsRating = "AA"
        End Select

        prcPendingApplication = lsResult 'getResult(lsResult, lnAcctTerm, lsRating, ldTransact, lnRecExist)
    End Function

    Private Function prcRepeatAccount( _
          ByVal lcAcctStat As AccoutStat, _
          ByVal ldTransact As Date, _
          ByVal lnAcctTerm As Integer, _
          ByRef lsRating As String, _
          ByVal lsModelIDx As String, _
          ByVal lnDownPaym As Double, _
          ByVal lnRecExist As compResult) As String
        Dim lsResult As String
        'Dim lsRating As String
        Dim lnYear As Integer

        'lsRating = lcRatingxx
        Select Case lcAcctStat
            Case xeActStatActive
                If chkDownPaym(lsModelIDx, lnDownPaym, 70) Then
                    'Check lnRecExist Value
                    lsResult = IIf(lnRecExist = pxeUncertain, "SV", "AP")
                    '        lsResult = "AP"
                Else
                    lsResult = "SA"
                End If
                lsRating = "Aa"
            Case xeActStatDiscarded
                lsResult = "CI"
                If lsRating = "n" Then
                    lsRating = "AS"
                Else
                    lsRating = "Pr"
                End If
            Case xeActStatImpounded
                'Check lnRecExist Value
                lsResult = IIf(lnRecExist = pxeUncertain, "SV", "DA")
                lsRating = "BL"
                '     lsResult = "DA"
            Case xeActStatDead
                lsResult = "DA"
                lsRating = "Dd"
            Case 5 'Rejected
                lsResult = "DA"
                lsRating = "RJ"
            Case xeActStatClosed
                Select Case lsRating
                    Case "x", "g"
                        ' XerSys - 2013-08-14
                        '  Rating is valid only for a year for branches that are not allowed to issu QM #
                        If Not p_bQMAllowed Then
                            If DateDiff("m", ldTransact, p_oAppDrivr.ServerDate) <= 12 Then
                                'Check lnRecExist Value
                                lsResult = IIf(lnRecExist = pxeUncertain, "SV", "AP")
                            Else
                                lsResult = "SV"
                            End If
                        Else
                            If DateDiff("m", ldTransact, p_oAppDrivr.ServerDate) <= 24 Then
                                'Check lnRecExist Value
                                lsResult = IIf(lnRecExist = pxeUncertain, "SV", "AP")
                                '            lsResult = "AP"
                            Else
                                If chkDownPaym(lsModelIDx, lnDownPaym, 30) Then
                                    'Check lnRecExist Value
                                    lsResult = IIf(lnRecExist = pxeUncertain, "SV", "AP")
                                    '               lsResult = "AP"
                                Else
                                    lsResult = "CI"
                                End If
                            End If
                        End If

                        If lsRating = "x" Then
                            lsRating = "Ex"
                        Else
                            lsRating = "Gd"
                        End If
                    Case "f"
                        If chkDownPaym(lsModelIDx, lnDownPaym, 40) Then
                            'Check lnRecExist Value
                            lsResult = IIf(lnRecExist = pxeUncertain, "SV", "AP")
                            '            lsResult = "AP"
                        Else
                            lsResult = "CI"
                        End If
                        lsRating = "Fr"
                    Case "p"
                        If chkDownPaym(lsModelIDx, lnDownPaym, 40) Then
                            lsResult = "SA"
                        Else
                            'Check lnRecExist Value
                            lsResult = IIf(lnRecExist = pxeUncertain, "SV", "DA")
                            '            lsResult = "DA"
                        End If
                        lsRating = "Pr"
                    Case "b"
                        If chkDownPaym(lsModelIDx, lnDownPaym, 70) Then
                            lsResult = "SA"
                        Else
                            'Check lnRecExist Value
                            lsResult = IIf(lnRecExist = pxeUncertain, "SV", "DA")
                            '            lsResult = "DA"
                        End If
                        lsRating = "BP"
                    Case "l"
                        lsResult = "SA"
                        lsRating = "BL"
                    Case "n"
                        lsResult = "CI"
                        lsRating = "NB"
                    Case Else 'No rating available
                        lsResult = "CI"
                        lsRating = "NR"
                End Select
        End Select

        prcRepeatAccount = lsResult 'getResult(lsResult, lnAcctTerm, lsRating, ldTransact, lnRecExist)
    End Function

    Private Function chkDownPaym( _
          lsModelIDx As String, _
          lnDownPaym As Double, _
          lnPercentg As Long) As Boolean
        Dim loRS As Recordset
        Dim lsOldProc As String
        Dim lsSQL As String

        lsOldProc = "chkDownPaym"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  sModelIDx" & _
                    ", nSelPrice" & _
                 " FROM MC_Model_Price" & _
                 " WHERE sModelIDx = " & strParm(lsModelIDx)
        loRS = New Recordset
        loRS.Open(lsSQL, p_oAppDrivr.Connection, , , adCmdText)

        If loRS.EOF Then GoTo endProc

        chkDownPaym = Round(loRS("nSelPrice") * lnPercentg / 100, 0) <= lnDownPaym

endProc:
        Exit Function
errProc:
        ShowError(lsOldProc & "( " & " )")
    End Function

    Private Sub ShowError(ByVal lsOldProc As String)
        With p_oAppDrivr
            .xLogError(Err.Number, Err.Description, pxeMODULENAME, lsOldProc, Erl)
        End With
        With Err()
            .Raise.Number, .Source, .Description
        End With
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnStatus As Int32)
        Me.New(foRider)
        p_nTranStat = fnStatus
    End Sub
End Class

