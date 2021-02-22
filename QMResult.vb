'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     QM Result Object
'
' Copyright 2020 and Beyond
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
'  jeep [ 12/05/2019 15:19 ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class QMResult
    Private Const pxeMODULENAME As String = "clsQMResult"
    Private Const pxeEmptyDate As String = "1900-01-01"

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

    Private p_oAppDrivr As GRider
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

    Public WriteOnly Property AppDriver As GRider
        Set(Value As GRider)
            p_oAppDrivr = Value
        End Set
    End Property

    Public Property ApplicationNo() As String
        Get
            Return p_sApplicNo
        End Get
        Set(ByVal Value As String)
            p_sApplicNo = Value
        End Set
    End Property

    Public ReadOnly Property TransNo() As String
        Get
            Return p_sTransNox
        End Get
    End Property

    Public Property Branch() As String
        Get
            Return p_sBranchCd
        End Get
        Set(ByVal Value As String)
            p_sBranchCd = Value
        End Set
    End Property

    Public Property Applicant(ByVal Index As String) As Object
        Get
            On Error Resume Next

            If pbInitTran = False Then
                Return ""
                Exit Property
            End If

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

        Set(Value)
            On Error Resume Next

            If pbInitTran = False Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    p_sClientID = Value
                Case "slastname"
                    p_sLastName = Value
                Case "sfrstname"
                    p_sFrstName = Value
                Case "smiddname"
                    p_sMiddName = Value
                Case "cgendercd"
                    If Not (Value = xeLogical.YES Or Value = xeLogical.NO) Then Exit Property

                    p_cGenderCd = Value
                Case "ccvilstat"
                    p_cCvilStat = Value
                Case "dbirthdte"
                    If Not IsDate(Value) Or CDate(Value) > Now Then Exit Property

                    p_dBirthDte = Value
                Case "sbirthplc"
                    p_sBirthPlc = Value
                Case "saddressx"
                    p_sAddressx = Value
                Case "sbrgyidxx"
                    p_sBrgyIDxx = Value
                Case "stownidxx"
                    p_sTownIDxx = Value
            End Select
        End Set
    End Property

    Public Property Spouse(ByVal Index As String) As Object
        Get
            On Error Resume Next

            If pbInitTran = False Then
                Return ""
                Exit Property
            End If

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

        Set(Value)
            On Error Resume Next

            If pbInitTran = False Then Exit Property

            Index = LCase(Index)
            Select Case Index
                Case "sclientid"
                    p_sSpouseID = Value
                Case "slastname"
                    p_sSLastNme = Value
                Case "sfrstname"
                    p_sSFrstNme = Value
                Case "smiddname"
                    p_sSMiddNme = Value
                Case "cgendercd"
                    If Not (Value = xeLogical.YES Or Value = xeLogical.NO) Then Exit Property

                    p_cSGendrCd = Value
                Case "ccvilstat"
                    p_cSCvlStat = Value
                Case "dbirthdte"
                    If Not IsDate(Value) Or CDate(Value) > Now Then Exit Property

                    p_dSBrthDte = Value
                Case "sbirthplc"
                    p_sSBrthPlc = Value
                Case "saddressx"
                    p_sSAddress = Value
                Case "sbrgyidxx"
                    p_sSBrgyIDx = Value
                Case "stownidxx"
                    p_sSTownIDx = Value
            End Select
        End Set
    End Property

    Public ReadOnly Property QuickMatchResult(ByVal Index As String) As String
        Get
            On Error Resume Next

            If p_sResltCde = "" Then
                Return ""
                Exit Property
            End If

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
            On Error Resume Next

            If p_sResltCde = "" Then
                Return ""
                Exit Property
            End If


            If Row > p_oResult.Rows.Count - 1 Then
                Return ""
                Exit Property
            End If

            Index = LCase(Index)
            Select Case Index
                Case "sfullname"
                    ResultDetail = p_oResult(Row)("sFullName")
                Case "sresltcde"
                    ResultDetail = p_oResult(Row)("sResltCde")
                Case "sacctnmbr"
                    ResultDetail = p_oResult(Row)("sAcctNmbr")
                Case "smcsonmbr"
                    ResultDetail = p_oResult(Row)("sMCSONmbr")
                Case "sapplnmbr"
                    ResultDetail = p_oResult(Row)("sApplNmbr")
            End Select
        End Get
    End Property

    Public Property Term(ByVal Index As String) As Object
        Get
            On Error Resume Next

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

        Set(Value)
            On Error Resume Next

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

        lsOldProc = "InitTransaction"
        'On Error GoTo errProc

        If p_sBranchCd = "" Then p_sBranchCd = p_oAppDrivr.BranchCode

        p_bQMAllowed = isBranchAllowed()

        pbInitTran = True
        Return True
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
                    Case "PA"
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
            Return p_sResltCde
        End If

        Return p_sResltCde
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
        Dim loRS As DataTable
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
        Dim lnRow As Integer
        Dim dtRow() As DataRow

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

        loRS = New DataTable
        Debug.Print(lsSQL)
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)

        With loRS
            If loRS.Rows.Count = 0 Then
                If lbBLAddress Then
                    lsResult = "BA"
                Else
                    lsResult = "CI"
                End If
                lsRating = "00"

                lsResult = getResult(lsResult, 0, lsRating, CDate("1900-01-01"), compResult.pxeDifferent)
                Call addResult(lsClientID, lsLastName & ", " & lsFrstName & " " & lsMiddName, _
                      lsResult, resultRelevance.pxeIrrelevant, "", "", "", "1900-01-01", "")

                Return lsResult
            End If

            For lnRow = 0 To .Rows.Count - 1
                If compareName(.Rows(lnRow)("sLastName"), lsLastName) = False Then GoTo moveToNext

                If compareName(.Rows(lnRow)("sFrstName"), lsFrstName) = False Then GoTo moveToNext

                lnMiddName = compareMiddName(IFNull(.Rows(lnRow)("sMiddName"), ""), lsMiddName)
                lnBirthDte = compareBirthDate(IFNull(.Rows(lnRow)("dBirthDte"), "1900-01-01"), ldBirthDte)

                lnBirthPlc = compareBirthPlace(IFNull(.Rows(lnRow)("sBirthPlc")), lsBirthPlc)
                lnTownIDxx = compareTown(IFNull(.Rows(lnRow)("sTownIDxx")), lsTownIDxx)
                lnAddressx = compareAddress(IFNull(.Rows(lnRow)("sAddressx"), ""), lsAddressx)

                If (lnMiddName + lnBirthDte + lnBirthPlc) <= 4 Then
                    lnRecExist = compResult.pxeEqual
                ElseIf (lnMiddName + lnBirthDte + lnBirthPlc) <= 6 Then
                    lnRecExist = compResult.pxeUncertain
                Else
                    lnRecExist = compResult.pxeDifferent
                End If

                lsResult = "CI"
                lsRating = "00"
                lnTerm = IIf(IsDBNull(.Rows(lnRow)("nAcctTerm")), 0, .Rows(lnRow)("nAcctTerm"))
                ldTransact = IIf(IsDBNull(.Rows(lnRow)("dTransact")), pxeEmptyDate, .Rows(lnRow)("dTransact"))
                lsAcctNmbr = ""
                lsTransNox = ""
                lnRelevance = resultRelevance.pxeRelevanceLo

                Select Case lnRecExist
                    Case compResult.pxeDifferent
                        If Not (lbRecExist Or lbRecUnsre) Then
                            If lbBLAddress Then
                                lsResult = "BA"
                            Else
                                lsResult = "CI"
                            End If
                        End If
                    Case compResult.pxeEqual, compResult.pxeUncertain
                        If lnRecExist = compResult.pxeEqual Then
                            lbRecExist = True
                        Else
                            If lbRecExist Then
                                GoTo moveToNext
                            Else
                                lbRecUnsre = True
                            End If
                        End If

                        If IsDBNull(.Rows(lnRow)("sAcctNmbr")) Then
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
                            Select Case .Rows(lnRow)("cRatingxx")
                                Case "NB" ' Cash Sales
                                    lsRating = "NB"
                                    ldTransact = .Rows(lnRow)("dTransact")

                                    '                  lsResult = getResult("CI", 0, "NB", .Fields("dTransact"), lnRecExist)
                                    '                  Call addResult(.Fields("sClientID"), _
                                    '                        .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                                    '                        lsResult, pxeRelevanceLo, "", _
                                    '                        .Fields("sAcctNmbr"), "", .Fields("dTransact"), .Fields("sResltCde"))
                                Case "PA" ' With Pending Application
                                    'If we are dealing with this application no then skip
                                    If .Rows(lnRow)("sAcctNmbr") = p_sApplicNo Then GoTo moveToNext
                                    lnRelevance = resultRelevance.pxeRelevanceMi
                                    lsTransNox = .Rows(lnRow)("sAcctNmbr")
                                    lsResult = prcPendingApplication(.Rows(lnRow)("cTranStat"), .Rows(lnRow)("dTransact"), .Rows(lnRow)("nAcctTerm"), lsRating, lnRecExist)
                                    '                  Call addResult(.Fields("sClientID"), _
                                    '                        .Fields("sLastName") & ", " & .Fields("sFrstName") & " " & IFNull(.Fields("sMiddName"), ""), _
                                    '                        lsResult, pxeRelevanceMi, "", "", _
                                    '                        .Fields("sAcctNmbr"), .Fields("dTransact"), .Fields("sResltCde"))
                                Case Else
                                    lsRating = .Rows(lnRow)("cRatingxx")
                                    lsAcctNmbr = .Rows(lnRow)("sAcctNmbr")
                                    ldTransact = IIf(.Rows(lnRow)("cTranStat") = xeAccountStat.ACTIVE, .Rows(lnRow)("dTransact"), IFNull(.Rows(lnRow)("dClosedxx"), pxeEmptyDate))
                                    If .Rows(lnRow)("cTranStat") = xeAccountStat.ACTIVE Then
                                        lsResult = prcRepeatAccount(.Rows(lnRow)("cTranStat"), _
                                                       .Rows(lnRow)("dTransact"), _
                                                       .Rows(lnRow)("nAcctTerm"), _
                                                       lsRating, _
                                                       p_sModelIDx, _
                                                       p_nDownPaym, _
                                                       lnRecExist)
                                    Else
                                        lsResult = prcRepeatAccount(.Rows(lnRow)("cTranStat"), _
                                                       .Rows(lnRow)("dClosedxx"), _
                                                       .Rows(lnRow)("nAcctTerm"), _
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
                Dim ldTemp As Date = pxeEmptyDate

                If Not IsDBNull(.Rows(lnRow)("cTranStat")) Then
                    ldTemp = IIf(.Rows(lnRow)("cTranStat") = xeAccountStat.ACTIVE, .Rows(lnRow)("dTransact"), IFNull(.Rows(lnRow)("dClosedxx"), pxeEmptyDate))
                End If

                Call addResult(.Rows(lnRow)("sClientID"), _
                         .Rows(lnRow)("sLastName") & ", " & .Rows(lnRow)("sFrstName") & " " & IFNull(.Rows(lnRow)("sMiddName"), ""), _
                         lsResult, _
                         IIf(lnRecExist = compResult.pxeEqual, resultRelevance.pxeRelevanceHi, resultRelevance.pxeRelevanceMi), _
                         lsAcctNmbr, _
                         "", lsTransNox, _
                         ldTemp, _
                         IFNull(.Rows(lnRow)("sResltCde")))
                'MsgBox(lnRow & .Rows(lnRow)("sClientID"), , lsResult)
                'Debug.Print(lnRow & " " & .Rows(lnRow)("sClientID") & " " & lsResult)

moveToNext:
                'lnRow = lnRow + 1
            Next
        End With

        With p_oTmpRst
            .DefaultView.Sort = "nRelevnce, dTransact DESC"
            Dim dtResult As DataTable = .DefaultView.ToTable()

            If lbRecExist Or lbRecUnsre Then
                If lbRecExist Then
                    dtRow = dtResult.Select("nComparsn = " & compResult.pxeEqual)
                Else
                    dtRow = dtResult.Select("nComparsn = " & compResult.pxeUncertain)
                End If
                '      Else
                '         MatchApplicant = lsResult
                '         Call addResult(lsClientID, lsLastName & ", " & lsFrstName & " " & lsMiddName, _
                '               lsResult, pxeIrrelevant, "", "", "", "1900-01-01", "")

                Return IFNull(dtRow(0)("sResltCde"), "")
            Else
                Return IFNull(dtResult.Rows(0)("sResltCde"), "")
            End If
        End With
    End Function

    Private Function match2Other( _
          ByVal lsLastName As String, _
          ByVal lsFrstName As String, _
          ByVal lsMiddName As String, _
          ByVal ldBirthDte As String, _
          ByVal lsTownIDxx As String, _
          ByVal lsAddressx As String) As String
        Dim loRS As DataTable
        Dim lsOldProc As String
        Dim lsSQL As String

        Dim lnMiddName As compResult
        Dim lnBirthDte As compResult
        Dim lnTownIDxx As compResult
        Dim lnAddressx As compResult

        Dim lsResult As String
        Dim lsTmpReslt As String
        Dim lnRow As Integer

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

        loRS = New DataTable
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)
        With loRS
            Debug.Print(lsSQL)
            If .Rows.Count = 0 Then Return ""

            For lnRow = 0 To loRS.Rows.Count - 1
                If compareName(.Rows(lnRow)("sLastName"), lsLastName) = False Then GoTo moveToNext

                If compareName(.Rows(lnRow)("sFrstName"), lsFrstName) = False Then GoTo moveToNext

                lnMiddName = compareMiddName(IFNull(.Rows(lnRow)("sMiddName"), ""), lsMiddName)
                lnBirthDte = compareBirthDate(IFNull(.Rows(lnRow)("dBirthDte"), "1900-01-01"), ldBirthDte)

                ' if no town exist, then use province
                If .Rows(lnRow)("sTownIDxx") = "" Then
                    lnTownIDxx = compareTown(.Rows(lnRow)("sProvIDxx"), .Rows(lnRow)("xProvIDxx"))
                Else
                    lnTownIDxx = compareTown(.Rows(lnRow)("sTownIDxx"), lsTownIDxx)
                End If
                lnAddressx = compareAddress(IFNull(.Rows(lnRow)("sAddressx"), ""), lsAddressx)

                '         Debug.Print lnMiddName, lnBirthDte, lnBirthPlc, lnTownIDxx, lnAddressx

                If (lnMiddName + lnBirthDte + lnTownIDxx) <= 4 Then
                    lsTmpReslt = "P" 'pxeEqual
                ElseIf (lnMiddName + lnBirthDte + lnTownIDxx) <= 6 Then
                    lsTmpReslt = "U" 'pxeUncertain
                Else
                    lsTmpReslt = compResult.pxeDifferent
                End If
                lsTmpReslt = lsTmpReslt & Format(CDate(.Rows(lnRow)("nBListdYr") & "/01/01"), "yy")

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
                lnRow = lnRow + 1
            Next
        End With

        Return lsResult
    End Function

    Private Function isAddBlackList(ByVal lsTownIDxx As String, ByVal lsBrgyIDxx As String) As Boolean
        Dim lsSQL As String
        Dim loRS As DataTable

        lsSQL = "SELECT cBlackLst FROM TownCity" & _
                 " WHERE sTownIDxx = " & strParm(lsTownIDxx)
        loRS = New DataTable
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)

        If loRS.Rows.Count > 0 Then
            If loRS.Rows(0)("cBlackLst") = xeLogical.YES Then
                Return True
            End If
        End If

        lsSQL = "SELECT cBlackLst FROM Barangay" & _
                 " WHERE sBrgyIDxx = " & strParm(lsBrgyIDxx)
        loRS = New DataTable
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)

        If loRS.Rows.Count > 0 Then
            If loRS.Rows(0)("cBlackLst") = xeLogical.YES Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Function isBranchAllowed() As Boolean
        Dim lsSQL As String
        Dim loRS As DataTable

        lsSQL = "SELECT cAllowQMx FROM Branch_Others" & _
                 " WHERE sBranchCd = " & strParm(p_sBranchCd)
        loRS = New DataTable
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)

        If loRS.Rows.Count > 0 Then
            If loRS.Rows(0)("cAllowQMx") = xeLogical.YES Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Function getResult( _
          ByVal lsResult As String, _
          ByVal lnTerm As Integer, _
          ByVal lsRating As String, _
          ByVal ldTransact As Date, _
          ByVal lnExist As compResult) As String
        Return lsResult & Format(lnTerm, "00") & _
                       lsRating & Format(ldTransact, "yy") & _
                       IIf(lnExist = compResult.pxeDifferent, "N", IIf(lnExist = compResult.pxeEqual, "P", "U"))
    End Function

    Private Function initResult() As Boolean
        Dim lsOldProc As String

        lsOldProc = "initResult"
        'On Error GoTo errProc

        p_oTmpRst = New DataTable
        With p_oTmpRst
            .Columns.Add("sClientID", GetType(String))
            .Columns.Add("sFullName", GetType(String))
            .Columns.Add("nRelevnce", GetType(Integer))
            .Columns.Add("dTransact", GetType(Date))
            .Columns.Add("nComparsn", GetType(Integer))
            .Columns.Add("sResltCde", GetType(String))
            .Columns.Add("sAcctNmbr", GetType(String))
            .Columns.Add("sMCSONmbr", GetType(String))
            .Columns.Add("sApplNmbr", GetType(String))
            .Columns.Add("cAddRecrd", GetType(String))
        End With

        p_oResult = New DataTable
        With p_oResult
            .Columns.Add("nEntryNox", GetType(Integer))
            .Columns.Add("sFullName", GetType(String))
            .Columns.Add("sResltCde", GetType(String))
            .Columns.Add("sAcctNmbr", GetType(String))
            .Columns.Add("sMCSONmbr", GetType(String))
            .Columns.Add("sApplNmbr", GetType(String))
        End With
        Return True
    End Function

    Private Function addResult( _
          ByVal lsClientID As String, _
          ByVal lsFullName As String, _
          ByVal lsResltCde As String, _
          ByVal lnRelevnce As resultRelevance, _
          ByVal lsAcctNmbr As String, _
          ByVal lsMCSONmbr As String, _
          ByVal lsApplNmbr As String, _
          ByVal ldTransact As Date, _
          ByVal lsPrevRslt As Object) As Boolean
        Dim lsOldProc As String
        Dim lcAddRecord As Integer
        Dim lnComparisn As compResult
        Dim dtRow() As DataRow

        lsOldProc = "addResult"
        'On Error GoTo errProc

        With p_oTmpRst
            dtRow = .Select("sResltCde = " & strParm(lsResltCde) & " AND nRelevnce = " & lnRelevnce)


            If UBound(dtRow) = -1 Then
                lcAddRecord = xeLogical.YES
            Else
                lcAddRecord = xeLogical.NO
            End If
            '.Select("")

            If Not IsDBNull(lsPrevRslt) Then
                If lsPrevRslt = lsResltCde Then
                    lcAddRecord = xeLogical.NO
                End If
            End If

            Select Case Right(lsResltCde, 1)
                Case "P"
                    lnComparisn = compResult.pxeEqual
                Case "N"
                    lnComparisn = compResult.pxeDifferent
                Case Else
                    lnComparisn = compResult.pxeUncertain
            End Select
            .Rows.Add(lsClientID, lsFullName, lnRelevnce, ldTransact, lnComparisn, lsResltCde, lsAcctNmbr, lsMCSONmbr, lsApplNmbr, lcAddRecord)
        End With

        Return True
    End Function

    Private Function saveResult() As Boolean
        Dim lsOldProc As String
        Dim lsTransNox As String
        Dim lsSQL As String
        Dim lnCtr As Integer
        Dim lnRow As Integer

        lsOldProc = "saveResult"
        'On Error GoTo errProc

        With p_oTmpRst
            .DefaultView.Sort = "nRelevnce, dTransact DESC"
            Dim dtResult As DataTable = p_oTmpRst.DefaultView.ToTable()

            lsTransNox = GetNextCode("MC_LR_QuickMatch", "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)

            If p_sApplicNo = "" Then p_oAppDrivr.BeginTransaction()
            lnCtr = 0
            For lnRow = 0 To dtResult.Rows.Count - 1
                If dtResult.Rows(lnRow)("cAddRecrd") = xeLogical.YES Then
                    lsSQL = "INSERT INTO MC_LR_QuickMatch_Result SET" & _
                                "  sTransNox = " & strParm(lsTransNox) & _
                                ", sClientID = " & strParm(dtResult.Rows(lnRow)("sClientID")) & _
                                ", nEntryNox = " & lnCtr + 1 & _
                                ", sResltCde = " & strParm(dtResult.Rows(lnRow)("sResltCde")) & _
                                ", sAcctNmbr = " & strParm(dtResult.Rows(lnRow)("sAcctNmbr")) & _
                                ", sMCSONmbr = " & strParm(dtResult.Rows(lnRow)("sMCSONmbr")) & _
                                ", sApplNmbr = " & strParm(dtResult.Rows(lnRow)("sApplNmbr")) & _
                                ", dModified = " & dateParm(p_oAppDrivr.SysDate)

                    If p_oAppDrivr.Execute(lsSQL, "MC_LR_QuickMatch_Result") <= 0 Then
                        MsgBox("Unable to Save Changes")
                        Return False
                    End If

                    lnCtr = lnCtr + 1
                End If

                p_oResult.Rows.Add(p_oResult.Rows.Count, dtResult.Rows(lnRow)("sFullName"), dtResult.Rows(lnRow)("sResltCde"), dtResult.Rows(lnRow)("sAcctNmbr"), dtResult.Rows(lnRow)("sMCSONmbr"), dtResult.Rows(lnRow)("sApplNmbr"))
            Next

            lsSQL = "INSERT INTO MC_LR_QuickMatch SET" & _
                            "  sTransNox = " & strParm(lsTransNox) & _
                            ", sApplicNo = " & strParm(p_sApplicNo) & _
                            ", sLastName = " & strParm(p_sLastName) & _
                            ", sFrstName = " & strParm(p_sFrstName) & _
                            ", sMiddName = " & strParm(p_sMiddName) & _
                            ", dBirthDte = " & dateParm(IIf(p_dBirthDte <> "", p_dBirthDte, p_oAppDrivr.getSysDate)) & _
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
                        ", dModified = " & dateParm(p_oAppDrivr.SysDate)

            If p_oAppDrivr.Execute(lsSQL, "MC_LR_QuickMatch") <= 0 Then
                MsgBox("Unable to Save Changes")
                Return False
            End If

            If p_sApplicNo = "" Then p_oAppDrivr.CommitTransaction()
        End With

        p_sTransNox = lsTransNox

        Return True
    End Function

    Private Function compareName( _
          lsNameNRec As String, _
          lsNameNew As String) As compResult
        Dim lsOldProc As String
        Dim lasTemp() As String
        Dim lnCtr As Integer

        lsOldProc = "compareName"
        'On Error GoTo errProc

        If StrComp(Trim(lsNameNRec), Trim(lsNameNew), vbTextCompare) <> 0 Then
            ' if not exactly equal, remove the conjunction
            If InStr(1, lsNameNRec, "&/") > 0 Then
                lasTemp = Split(lsNameNRec, "&/")
            ElseIf InStr(1, lsNameNRec, "/") > 0 Then
                lasTemp = Split(lsNameNRec, "/")
            Else
                Return compResult.pxeDifferent
            End If

            ' now check if name is exactly equal
            For lnCtr = 0 To UBound(lasTemp)
                If StrComp(Trim(lasTemp(lnCtr)), Trim(lsNameNew), vbTextCompare) <> 0 Then
                    Return compResult.pxeEqual
                End If
            Next
        Else
            Return compResult.pxeEqual
        End If

        Return compResult.pxeDifferent
    End Function

    Private Function compareMiddName( _
          ByVal lsNameNRec As String, _
          ByVal lsNameNew As String) As compResult
        Dim lsOldProc As String
        Dim lasTemp() As String
        Dim lnCtr As Integer

        lsOldProc = "compareMiddName"
        'On Error GoTo errProc

        Call IFNull(lsNameNRec, "")
        If StrComp(Trim(lsNameNRec), Trim(lsNameNew), vbTextCompare) <> 0 Then
            ' if not exactly equal, remove the conjunction
            If InStr(1, lsNameNRec, "&/") > 0 Then
                lasTemp = Split(lsNameNRec, "&/")
            ElseIf InStr(1, lsNameNRec, "/") > 0 Then
                lasTemp = Split(lsNameNRec, "/")
            Else
                If lsNameNRec = "" Then
                    Return compResult.pxeUncertain
                Else
                    lnCtr = InStr(lsNameNRec, ".")
                    lsNameNRec = IIf(lnCtr = 0, lsNameNRec, lnCtr - 1)
                    'Remove this process and replace it with the above
                    '            lsNameNRec = Mid(lsNameNRec, 1, IIf(lnCtr = 0, 0, lnCtr - 1))

                    If StrComp(Trim(lsNameNRec), Left(Trim(lsNameNew), Len(Trim(lsNameNRec))), vbTextCompare) = 0 Then
                        Return compResult.pxeUncertain
                    End If
                End If
                Return compResult.pxeDifferent
            End If

            ' now check if name is exactly equal
            For lnCtr = 0 To UBound(lasTemp)
                If StrComp(Trim(lasTemp(lnCtr)), Trim(lsNameNew), vbTextCompare) <> 0 Then
                    Return compResult.pxeEqual
                End If
            Next
        Else
            Return compResult.pxeEqual
        End If

        Return compResult.pxeDifferent
    End Function

    Private Function compareBirthDate( _
          ByVal ldInfoNRec As Date, _
          ByVal ldNewInfo As Date) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareBirthDate"
        'On Error GoTo errProc

        compareBirthDate = compResult.pxeDifferent
        Debug.Print(ldNewInfo, ldInfoNRec)
        If ldInfoNRec <> ldNewInfo Then
            '      If DateDiff("d", ldInfoNRec, CDate("1900-01-01")) = 0 Then
            '         compareBirthDate = pxeUncertain
            '      End If
            If DateDiff("d", ldInfoNRec, CDate("1900-01-01")) = 0 _
             Or DateDiff("d", ldInfoNRec, CDate("1901-01-01")) = 0 _
             Or DateDiff("d", ldInfoNRec, CDate("1999-01-01")) = 0 Then
                Return compResult.pxeUncertain
            End If

        Else
            Return compResult.pxeEqual
        End If
    End Function

    Private Function compareBirthPlace( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareBirthPlace"
        'On Error GoTo errProc
        compareBirthPlace = compResult.pxeDifferent

        If lsInfoNRec <> lsNewInfo Then
            If lsInfoNRec = "" Then
                Return compResult.pxeUncertain
            End If
        Else
            Return compResult.pxeEqual
        End If
    End Function

    Private Function compareTown( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareTown"
        'On Error GoTo errProc

        If lsInfoNRec = lsNewInfo Then
            Return compResult.pxeEqual
        Else
            Return compResult.pxeDifferent
        End If
    End Function

    Private Function compareAddress( _
          ByVal lsInfoNRec As String, _
          ByVal lsNewInfo As String) As compResult
        Dim lsOldProc As String

        lsOldProc = "compareAddress"
        'On Error GoTo errProc
        compareAddress = compResult.pxeDifferent

        If StrComp(Trim(lsInfoNRec), Trim(lsNewInfo), vbTextCompare) <> 0 Then
            If lsInfoNRec = "" Then
                Return compResult.pxeUncertain
            End If
        Else
            Return compResult.pxeEqual
        End If
    End Function

    Private Function prcPendingApplication( _
          ByVal lcTranStat As xeTranStat, _
          ByVal ldTransact As Date, _
          ByVal lnAcctTerm As Integer, _
          ByRef lsRating As String, _
          ByVal lnRecExist As compResult) As String
        Dim lsResult As String
        '   Dim lsRating As String

        Select Case lcTranStat
            Case xeTranStat.TRANS_POSTED
                If DateDiff("m", ldTransact, p_oAppDrivr.SysDate) > 3 Then
                    lsResult = "CI"
                Else
                    lsResult = "AP"
                End If
                lsRating = "AA"
            Case xeTranStat.TRANS_CANCELLED
                If DateDiff("m", ldTransact, p_oAppDrivr.SysDate) > 24 Then
                    lsResult = "CI"
                Else
                    lsResult = "DA"
                End If
                lsRating = "DA"
            Case Else
                lsResult = "PA"
                lsRating = "AA"
        End Select

        Return lsResult 'getResult(lsResult, lnAcctTerm, lsRating, ldTransact, lnRecExist)
    End Function

    Private Function prcRepeatAccount( _
          ByVal lcAcctStat As xeAccountStat, _
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
            Case xeAccountStat.ACTIVE
                If chkDownPaym(lsModelIDx, lnDownPaym, 70) Then
                    'Check lnRecExist Value
                    lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "AP")
                    '        lsResult = "AP"
                Else
                    lsResult = "SA"
                End If
                lsRating = "Aa"
            Case xeAccountStat.DISCARDED
                lsResult = "CI"
                If lsRating = "n" Then
                    lsRating = "AS"
                Else
                    lsRating = "Pr"
                End If
            Case xeAccountStat.IMPOUNDED
                'Check lnRecExist Value
                lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "DA")
                lsRating = "BL"
                '     lsResult = "DA"
            Case xeAccountStat.DEAD
                lsResult = "DA"
                lsRating = "Dd"
            Case 5 'Rejected
                lsResult = "DA"
                lsRating = "RJ"
            Case xeTranStat.TRANS_CLOSED
                Select Case lsRating
                    Case "x", "g"
                        ' XerSys - 2013-08-14
                        '  Rating is valid only for a year for branches that are not allowed to issu QM #
                        If Not p_bQMAllowed Then
                            If DateDiff("m", ldTransact, p_oAppDrivr.SysDate) <= 12 Then
                                'Check lnRecExist Value
                                lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "AP")
                            Else
                                lsResult = "SV"
                            End If
                        Else
                            If DateDiff("m", ldTransact, p_oAppDrivr.SysDate) <= 24 Then
                                'Check lnRecExist Value
                                lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "AP")
                                '            lsResult = "AP"
                            Else
                                If chkDownPaym(lsModelIDx, lnDownPaym, 30) Then
                                    'Check lnRecExist Value
                                    lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "AP")
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
                            lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "AP")
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
                            lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "DA")
                            '            lsResult = "DA"
                        End If
                        lsRating = "Pr"
                    Case "b"
                        If chkDownPaym(lsModelIDx, lnDownPaym, 70) Then
                            lsResult = "SA"
                        Else
                            'Check lnRecExist Value
                            lsResult = IIf(lnRecExist = compResult.pxeUncertain, "SV", "DA")
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

        Return lsResult 'getResult(lsResult, lnAcctTerm, lsRating, ldTransact, lnRecExist)
    End Function

    Private Function chkDownPaym( _
          lsModelIDx As String, _
          lnDownPaym As Double, _
          lnPercentg As Long) As Boolean
        Dim loRS As DataTable
        Dim lsOldProc As String
        Dim lsSQL As String

        lsOldProc = "chkDownPaym"
        'On Error GoTo errProc

        lsSQL = "SELECT" & _
                    "  sModelIDx" & _
                    ", nSelPrice" & _
                 " FROM MC_Model_Price" & _
                 " WHERE sModelIDx = " & strParm(lsModelIDx)
        loRS = New DataTable
        loRS = ExecuteQuery(lsSQL, p_oAppDrivr.Connection)

        If loRS.Rows.Count = 0 Then Return False

        Return Math.Round(loRS(0)("nSelPrice") * lnPercentg / 100, 0) <= lnDownPaym
    End Function

    Private Sub ShowError(ByVal lsOldProc As String)
        With p_oAppDrivr
            .ErrorLog(Err.Description)
        End With
        With Err()
            .Raise(.Number, .Source, .Description)
        End With
    End Sub
End Class