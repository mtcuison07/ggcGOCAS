'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MP Application Object
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
'  jeep [ 11/10/2020 10:06 ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports rmjGOCAS
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MCApplication
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String

    Private Const p_sMasTable As String = "Credit_Online_Application"
    Private Const p_sMsgHeadr As String = "Credit_Online_Application"

    Private jsonDet As String
    Private jsonObjDet As New App_DetParam
    Private jsonObjCat As New GOCAS_Param

    Private p_oFrmResult As frmQuickMatch

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
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTMstr(0).Item(Index) = value
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Return p_oDTMstr(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
        Set(ByVal Value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                p_oDTMstr(0).Item(Index) = Value
            End If
        End Set
    End Property

    Public ReadOnly Property Detail() As App_DetParam
        Get
            Return jsonObjDet
        End Get
    End Property

    Public Property Category() As GOCAS_Param
        Get
            Return jsonObjCat
        End Get

        Set(ByVal Value As GOCAS_Param)
            jsonObjCat = Value
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

        jsonObjDet = New App_DetParam
        jsonObjCat = New GOCAS_Param

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

        jsonObjDet = New App_DetParam
        Call populateJSONObjectDetail(jsonObjDet, IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), ""))

        jsonObjCat = New GOCAS_Param
        Call populateJSONObject(jsonObjCat, IFNull(p_oDTMstr.Rows(0)("sCatInfox"), ""))

        'jovan temporary remove this procedute
        If IsDBNull(p_oDTMstr(0)("cTranStat")) Then
            If Not saveHistory() Then Return False
        Else
            If p_oDTMstr(0)("cTranStat") = xeTranStat.TRANS_OPEN Then
                If Not saveHistory() Then Return False
            End If
        End If

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
                      , Optional ByVal fbByCode As Boolean = False _
                      , Optional ByVal fbEvaluate As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sTransNox") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sClientNm") Then Return True
            End If
        End If

        'Initialize SQL filter
        If p_nTranStat >= 0 Then
            lsSQL = AddCondition(getSQ_Browse, "a.cTranStat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = AddCondition(getSQ_Browse, "NOT a.cTranStat IS NULL")
        End If

        'jovan remove this condition 2020-11-11
        'lsSQL = AddCondition(lsSQL, "a.cEvaluatr = " & strParm(IIf(fbEvaluate, 1, 0)))

        'If p_sBranchCd <> "" Then
        '    lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
        'End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "a.sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "a.sClientNm LIKE " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , "Trans No»Client»Date", _
                                        , "a.sTransNox»a.sClientNm»a.dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            Return False
        Else
            'jovan remove this condition 2020-11-11
            '            If IsDBNull(loDta.Item("sLoadedBy")) Then GoTo moveNext
            '            If loDta.Item("sLoadedBy") <> p_oApp.UserID Then
            '                Dim lnRep As Integer

            '                lnRep = MsgBox("Transaction was already loaded by other evaluator..." & vbCrLf & _
            '                                "Do you want to open transaction", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "CONFIRMATION")

            '                If lnRep = vbNo Then Return False
            '            End If
            'MoveNext:

            Return OpenTransaction(loDta.Item("sTransNox"))
        End If
    End Function

    'Public Function CancelTransaction
    Public Function CancelTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        '1 = pre-approved
        '2 = approved
        If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
            If MsgBox("Request was already approved! Do you continue?", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr) = MsgBoxResult.Cancel Then
                Return False
            End If
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Request was already cancelled!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "4"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))
        p_oApp.Execute(lsSQL, p_sMasTable)

        If p_sParent = "" Then p_oApp.CommitTransaction()
        Return True
    End Function

    Function PostQuickMatch() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        Select Case Trim(Left(p_oDTMstr.Rows(0)("sQMatchNo"), 2))
            Case "DA", "BA"
                p_oDTMstr(0).Item("cTranStat") = "1"
                p_oDTMstr(0).Item("cWithCIxx") = xeLogical.YES
            Case Else
                p_oDTMstr(0).Item("cTranStat") = "0"
        End Select

        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Function Approved(ByRef fnMonPaymx As Long) As Boolean
        Dim lsSQLBranch As String
        Dim loDTBranch As DataTable
        Dim loDT As DataTable

        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        'If p_sParent = "" Then p_oApp.BeginTransaction()

        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'If p_sParent = "" Then p_oApp.BeginTransaction()
        p_oDTMstr(0).Item("cTranStat") = "1"
        p_oDTMstr(0).Item("cEvaluatr") = "0"

        'jovan 01-28-2020' per sic mac need to populate this 2 field
        p_oDTMstr(0).Item("sVerified") = p_oApp.UserID
        p_oDTMstr(0).Item("dVerified") = p_oApp.getSysDate.ToString

        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        lsSQLBranch = "SELECT *" & _
                           " FROM Branch_Mobile" & _
                           " WHERE sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd"))

        loDTBranch = New DataTable
        loDTBranch = ExecuteQuery(lsSQLBranch, p_oApp.Connection)

        'For Branch
        For lnCtr As Integer = 0 To loDTBranch.Rows.Count - 1
            Call createReply(p_oDTMstr(0).Item("sTransNox") & vbCrLf & _
                             "Application of Mr/Ms. " & p_oDTMstr(0).Item("sClientNm") & " was APPROVED" & vbCrLf & _
                             "Valid until 90 days upon application." & vbCrLf & _
                             "REF. # " & p_oDTMstr(0).Item("sTransNox") & vbCrLf & _
                             "-GUANZON Group-", loDTBranch(lnCtr)("sMobileNo"), p_oDTMstr(0).Item("sTransNox"))
        Next

        'For Customer
        'createReply("", "", "")

        'If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Function TruncateDecimal(ByVal value As Decimal, ByVal precision As Integer) As Decimal
        Dim stepper As Decimal = Math.Pow(10, precision)
        Dim tmp As Decimal = Math.Truncate(stepper * value)
        Return tmp / stepper
    End Function

    Function DisApproved() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cTranStat") = "3"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Function Evaluate() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()
        p_oDTMstr(0).Item("cEvaluatr") = "1"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)
        p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4))

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function getNextCustomer() As String
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = AddCondition(getSQ_Master, " sCatInfox IS NULL" & _
                                            " AND sLoadedBy IS NULL" & _
                                            " AND sSourceCd = 'APP'" & _
                                            " AND cTranStat = '0'") & _
                " ORDER BY dTimeStmp ASC" & _
                " LIMIT 1"

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            Return ""
        Else
            Return loDT(0)("sTransNox")
        End If
    End Function

    Public Function getNextReference(ByVal fsMobileNo As String, ByRef fsRefNamex As String) As String
        'For nCtr As Integer = 0 To Detail.other_info.personal_reference.Count - 1
        '    If Detail.other_info.personal_reference(nCtr).sRefrMPNx = fsMobileNo Then
        '        If nCtr + 1 = Detail.other_info.personal_reference.Count Then
        '            fsRefNamex = Detail.other_info.personal_reference(0).sRefrNmex
        '            Return Detail.other_info.personal_reference(0).sRefrMPNx
        '        Else
        '            fsRefNamex = Detail.other_info.personal_reference(nCtr + 1).sRefrNmex
        '            Return Detail.other_info.personal_reference(nCtr + 1).sRefrMPNx
        '        End If
        '    End If
        'Next
    End Function

    Private Function getSQ_Master() As String
        Return "SELECT" & _
                    "  sTransNox" & _
                    ", sBranchCd" & _
                    ", dTransact" & _
                    ", dTargetDt" & _
                    ", sClientNm" & _
                    ", sGOCASNox" & _
                    ", sGOCASNoF" & _
                    ", cUnitAppl" & _
                    ", sSourceCD" & _
                    ", sDetlInfo" & _
                    ", sCatInfox" & _
                    ", sDesInfox" & _
                    ", sQMatchNo" & _
                    ", sQMAppCde" & _
                    ", nCrdtScrx" & _
                    ", nDownPaym" & _
                    ", nDownPayF" & _
                    ", sRemarksx" & _
                    ", sCreatedx" & _
                    ", dReceived" & _
                    ", sVerified" & _
                    ", dVerified" & _
                    ", cWithCIxx" & _
                    ", cTranStat" & _
                    ", cDivision" & _
                    ", cEvaluatr" & _
                    ", sLoadedBy" & _
                    ", dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sClientNm" & _
                    ", a.dTransact" & _
                    ", a.sLoadedBy" & _
                  " FROM " & p_sMasTable & " a" & _
                  " WHERE a.sSourceCD = 'WEB'" & _
                  " AND a.cDivision = '0'"
    End Function

    Private Function getSQL_Branch() As String
        Return "SELECT" & _
                    "  sBranchCd" & _
                    ", sBranchNm" & _
                " FROM Branch" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_TownCity() As String
        Return "SELECT" & _
                    "  a.sTownIDxx" & _
                    ", CONCAT(a.sTownName, ', ', b.sProvName) xTownCity" & _
                " FROM TownCity a" & _
                    ", Province b" & _
                " WHERE a.sProvIDxx = b.sProvIDxx" & _
                    " AND a.cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Country() As String
        Return "SELECT" & _
                    "  sCntryCde" & _
                    ", sCntryNme" & _
                " FROM Country" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Model() As String
        Return "SELECT" & _
                    "  sModelIDx" & _
                    ", sModelNme" & _
                    ", cMotorTyp" & _
                " FROM MC_Model" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Occupation() As String
        Return "SELECT" & _
                    "  sOccptnID" & _
                    ", sOccptnNm" & _
                " FROM Occupation" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function getSQL_Barangay() As String
        Return "SELECT" & _
                    "  a.sBrgyIDxx" & _
                    ", a.sBrgyName" & _
                    ", b.sTownName" & _
                    ", c.sProvName" & _
                " FROM Barangay a" & _
                    ", TownCity b" & _
                    ", Province c" & _
                " WHERE a.cRecdStat = " & strParm(xeLogical.YES) & _
                    " AND a.sTownIDxx = b.sTownIDxx" & _
                    " AND b.sProvIDxx = c.sProvIDxx"
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

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Function getBranch(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sBranchCd As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getBranch"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sBranchNm LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sBranchNm = " & strParm(sValue)
                End If
            Else
                lsCondition = "sBranchCd = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Branch, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getBranch = loDT(0)("sBranchNm")
            sBranchCd = loDT(0)("sBranchCd")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sBranchCd»sBranchNm", _
                                "ID»Branch", _
                                "", _
                                "sBranchCd»sBranchNm", _
                                1)

            If Not IsNothing(loDataRow) Then
                getBranch = loDataRow("sBranchNm")
                sBranchCd = loDataRow("sBranchCd")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getBranch = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getTownCity(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sTownIDxx As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getTownCity"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sTownName LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "a.sTownName = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sTownIDxx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_TownCity, lsCondition)
        Debug.Print(lsSQL)


        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getTownCity = loDT(0)("xTownCity")
            sTownIDxx = loDT(0)("sTownIDxx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sTownIDxx»xTownCity", _
                                "ID»Town", _
                                "", _
                                "a.sTownIDxx»CONCAT(a.sTownName, ', ', b.sProvName)", _
                                1)

            If Not IsNothing(loDataRow) Then
                getTownCity = loDataRow("xTownCity")
                sTownIDxx = loDataRow("sTownIDxx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getTownCity = ""
        sTownIDxx = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getCountry(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sCntryCde As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getCountry"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sCntryNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sCntryNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sCntryCde = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Country, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getCountry = loDT(0)("sCntryNme")
            sCntryCde = loDT(0)("sCntryCde")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sCntryCde»sCntryNme", _
                                "ID»Country", _
                                "", _
                                "sCntryCde»sCntryNme", _
                                1)

            If Not IsNothing(loDataRow) Then
                getCountry = loDataRow("sCntryNme")
                sCntryCde = loDataRow("sCntryCde")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getCountry = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getModel(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sModelIDx As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Model, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getModel = loDT(0)("sModelNme")
            sModelIDx = loDT(0)("sModelIDx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sModelIDx»sModelNme", _
                                "ID»Model", _
                                "", _
                                "sModelIDx»sModelNme", _
                                1)

            If Not IsNothing(loDataRow) Then
                getModel = loDataRow("sModelNme")
                sModelIDx = loDataRow("sModelIDx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getModel = ""
        sModelIDx = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getModel(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sModelIDx As String, ByRef cUnitType As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Model, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getModel = loDT(0)("sModelNme")
            sModelIDx = loDT(0)("sModelIDx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sModelIDx»sModelNme", _
                                "ID»Model", _
                                "", _
                                "sModelIDx»sModelNme", _
                                1)

            If Not IsNothing(loDataRow) Then
                getModel = loDataRow("sModelNme")
                sModelIDx = loDataRow("sModelIDx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getModel = ""
        sModelIDx = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getOccupation(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sOccptnID As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getOccupation"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sOccptnNm LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sOccptnNm = " & strParm(sValue)
                End If
            Else
                lsCondition = "sOccptnID = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Occupation, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getOccupation = loDT(0)("sOccptnNm")
            sOccptnID = loDT(0)("sOccptnID")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sOccptnID»sOccptnNm", _
                                "ID»Occupation", _
                                "", _
                                "sOccptnID»sOccptnNm", _
                                1)

            If Not IsNothing(loDataRow) Then
                getOccupation = loDataRow("sOccptnNm")
                sOccptnID = loDataRow("sOccptnID")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getOccupation = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Function getBarangay(ByVal sValue As String _
                                , ByVal bSearch As Boolean _
                                , ByVal bByCode As Boolean _
                                , ByRef sBrgyIDxx As String _
                                , Optional sTownIDxx As String = "") As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getBarangay"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "a.sBrgyName LIKE " & strParm("%" & sValue & "%") & _
                                    IIf(sTownIDxx = "", "", " AND a.sTownIDxx = " & strParm(sTownIDxx))
                Else
                    lsCondition = "a.sBrgyName = " & strParm(sValue) & _
                                    IIf(sTownIDxx = "", "", " AND a.sTownIDxx = " & strParm(sTownIDxx))
                End If
            Else
                lsCondition = "a.sBrgyIDxx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Barangay, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getBarangay = loDT(0)("sBrgyName")
            sBrgyIDxx = loDT(0)("sBrgyIDxx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sBrgyIDxx»sBrgyName»sTownName»sProvName", _
                                "ID»Barangay»Town»Province", _
                                "", _
                                "a.sBrgyIDxx»a.sBrgyName»b.sTownName»c.sProvName", _
                                1)

            If Not IsNothing(loDataRow) Then
                getBarangay = loDataRow("sBrgyName")
                sBrgyIDxx = loDataRow("sBrgyIDxx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getBarangay = ""
        sBrgyIDxx = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function saveHistory() As Boolean
        Dim lsSQL As String
        Dim loDT As DataTable

        loDT = New DataTable
        loDT = ExecuteQuery("SELECT * FROM Credit_Online_Application_Verification_History" & _
                                " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")), p_oApp.Connection)

        lsSQL = "INSERT INTO Credit_Online_Application_Verification_History SET" & _
                    "  sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                    ", nEntryNox = " & CDbl(loDT.Rows.Count + 1) & _
                    ", sModified = " & strParm(p_oApp.UserID) & _
                    ", dModified = " & dateParm(p_oApp.SysDate)

        If p_oApp.Execute(lsSQL, "Credit_Online_Application_Verification_History", p_sBranchCd) = 0 Then
            MsgBox("Unable to Save History Info!!!", vbCritical, "Warning")
            Return False
        End If

        Return True
    End Function

    Function confirmTransaction() As Boolean
        Dim lsSQL As String

        lsSQL = "UPDATE Credit_Online_Application SET" & _
                    "  sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
                    ", nDownPaym = " & CDbl(0) & _
                    ", sVerified = " & strParm(p_oApp.UserID) & _
                    ", dVerified = " & dateParm(p_oApp.getSysDate()) & _
                    ", cTranStat = " & strParm(xeTranStat.TRANS_CLOSED) & _
                    ",  sBranchCd = " & "'" & (p_oDTMstr.Rows(0)("sBranchCd")) & "'" & _
                    ",  sQMatchNo = " & "'" & (p_oDTMstr.Rows(0)("sQMatchNo")) & "'" & _
                " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox"))

        If p_oApp.Execute(lsSQL, "Credit_Online_Application", p_sBranchCd) = 0 Then
            MsgBox("Unable to Confirm Transaction!!!", vbCritical, "Warning")
            Return False
        End If

        Return OpenTransaction(p_oDTMstr(0)("sTransNox"))
    End Function

    Function JSONObjCategory() As String
        'Dim loValue As String

        'loValue = JsonConvert.SerializeObject(jsonObjCat)

        'Dim token As JToken = RemoveEmptyChildren(JToken.Parse(loValue))
        'Return token.ToString(Newtonsoft.Json.Formatting.None)
        Dim loSettings As New JsonSerializerSettings

        loSettings.NullValueHandling = NullValueHandling.Ignore
        loSettings.DefaultValueHandling = DefaultValueHandling.Ignore

        Return JsonConvert.SerializeObject(jsonObjCat, loSettings)
    End Function

    Function saveReference(ByVal fsMobileNo As String) As Boolean
        'Dim lsSQL As String

        'lsSQL = "INSERT INTO Credit_Online_Application_Reference SET" & _
        '            "  sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
        '            ", sMobileNo = " & strParm(fsMobileNo) & _
        '            ", sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
        '            ", sDesInfox = " & "'" & (JSONObjDescription()) & "'" & _
        '        " ON DUPLICATE KEY UPDATE" & _
        '            "  sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
        '            ", sDesInfox = " & "'" & (JSONObjDescription()) & "'"

        'If p_oApp.Execute(lsSQL, "Credit_Online_Application_Reference", p_sBranchCd) = 0 Then
        '    MsgBox("Unable to Save Reference Transaction!!!", vbCritical, "Warning")
        '    Return False
        'End If

        'Return True
    End Function

    Public Function SaveTransaction() As Boolean
        Dim lsSQL As String = ""
        Try

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
            lsSQL = "INSERT INTO Credit_Online_Application SET" & _
                        "  sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                        ", sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd")) & _
                        ", dTransact = " & dateParm(p_oDTMstr(0).Item("dTransact")) & _
                        ", sClientNm = " & strParm(Category.applicant_info.sLastName & ", " & Category.applicant_info.sFrstName & " " & Category.applicant_info.sMiddName) & _
                        ", sSourceCD = " & strParm(p_oDTMstr(0).Item("sSourceCd")) & _
                        ", sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
                        ", nDownPaym = " & strParm(IFNull(p_oDTMstr(0).Item("nDownPaym"), 0)) & _
                        ", sCreatedx = " & strParm(p_oApp.UserID) & _
                        ", dCreatedx = " & dateParm(p_oApp.getSysDate()) & _
                        ", cWithCIxx = " & strParm(xeTranStat.TRANS_CLOSED) & _
                        ", cTranStat = " & strParm(xeTranStat.TRANS_OPEN) & _
                        ", cDivision = " & strParm("2") & _
                        ", cEvaluatr = " & strParm("0") & _
                        ", dModified = " & dateParm(p_oApp.getSysDate()) & _
                    " ON DUPLICATE KEY UPDATE" & _
                        "  sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd")) & _
                        ", cTranStat = " & strParm(p_oApp.UserID) & _
                        ", sCatInfox = " & "'" & (JSONObjCategory()) & "'"
            Debug.Print(lsSQL)
            If lsSQL <> "" Then
                p_oApp.Execute(lsSQL, p_sMasTable)
            End If

            p_nEditMode = xeEditMode.MODE_READY

            Return True
        Catch ex As Exception
            If p_sParent = "" Then p_oApp.RollBackTransaction()
            MsgBox(ex.Message)

            Return False
        End Try
    End Function

    Function callApplicant() As String
        Dim lsMobile As String

        lsMobile = Detail.sMobileNo

        p_oApp.Execute("UPDATE Credit_Online_Application SET" & _
                            " sLoadedBY = " & strParm(p_oApp.UserID) & _
                        " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")), "Credit_Online_Application", p_sBranchCd)

        Return lsMobile
    End Function

    'Function callApplicant() As String()
    '    Dim lsMobile(Detail.applicant_info.mobile_number.Count - 1) As String

    '    For nCtr As Integer = 0 To Detail.applicant_info.mobile_number.Count - 1
    '        lsMobile(nCtr) = Detail.applicant_info.mobile_number(nCtr).sMobileNo
    '    Next nCtr

    '    p_oApp.Execute("UPDATE Credit_Online_Application SET" & _
    '                        " sLoadedBY = " & strParm(p_oApp.UserID) & _
    '                    " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")), "Credit_Online_Application", p_sBranchCd)

    '    Return lsMobile
    'End Function

    Function callReference(ByRef fsRefNamex As String) As String
        '        Dim loDT As DataTable
        '        Dim lsMobile As String
        '        Dim lbIsEqual As Boolean

        '        loDT = New DataTable
        '        loDT = p_oApp.ExecuteQuery("SELECT *" & _
        '                                        " FROM Credit_Online_Application_Reference" & _
        '                                        " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")))
        '        If loDT.Rows.Count = 0 Then
        '            If Detail.other_info.personal_reference.Count > 0 Then
        '                fsRefNamex = Detail.other_info.personal_reference(0).sRefrNmex
        '                Return Detail.other_info.personal_reference(0).sRefrMPNx
        '            End If
        '        Else
        '            For nCtr As Integer = 0 To Detail.other_info.personal_reference.Count - 1
        '                lbIsEqual = False
        '                For nCtr1 As Integer = 0 To loDT.Rows.Count - 1
        '                    If Not lbIsEqual Then
        '                        If Detail.other_info.personal_reference(nCtr).sRefrMPNx = loDT(nCtr1)("sMobileNo") Then
        '                            lbIsEqual = True
        '                            Exit For
        '                        End If
        '                    End If
        '                Next nCtr1

        '                If Not lbIsEqual Then
        '                    fsRefNamex = Detail.other_info.personal_reference(nCtr).sRefrNmex
        '                    Return Detail.other_info.personal_reference(nCtr).sRefrMPNx
        '                Else
        '                    GoTo movenext
        '                End If
        'movenext:
        '            Next nCtr

        '            Return ""
        '        End If

    End Function

    Function GenerateQM() As String
        Dim loQMResult As QMResult
        Dim loFrm As frmQuickMatch

        loQMResult = New QMResult
        loFrm = New frmQuickMatch

        With loQMResult
            .AppDriver = p_oApp
            .Branch = p_sBranchCd
            .ApplicationNo = p_oDTMstr.Rows(0)("sTransNox")

            .InitTransaction()
            'Set the Applicant info
            .Applicant("sClientID") = ""

            .Applicant("sLastName") = Category.applicant_info.sLastName & IIf(IFNull(Category.applicant_info.sSuffixNm) = "", "", " " & Category.applicant_info.sSuffixNm)
            .Applicant("sFrstName") = Category.applicant_info.sFrstName
            .Applicant("sMiddName") = Category.applicant_info.sMiddName
            .Applicant("dBirthDte") = Category.applicant_info.dBirthDte
            .Applicant("sBirthPlc") = Category.applicant_info.sBirthPlc
            .Applicant("sTownIDxx") = ""

            'Set the spouse info
            If Not IsNothing(Category.spouse_info.personal_info.sLastName) Then
                If IFNull(Category.spouse_info.personal_info.sLastName) <> "" Then
                    .Spouse("sClientID") = ""

                    .Spouse("sLastName") = Category.spouse_info.personal_info.sLastName
                    .Spouse("sFrstName") = Category.spouse_info.personal_info.sFrstName & IIf(IFNull(Category.spouse_info.personal_info.sSuffixNm) = "", "", " " & Category.spouse_info.personal_info.sSuffixNm)
                    .Spouse("sMiddName") = Category.spouse_info.personal_info.sMiddName
                    .Spouse("dBirthDte") = Category.spouse_info.personal_info.dBirthDte
                    .Spouse("sBirthPlc") = Category.spouse_info.personal_info.sBirthPlc
                    .Spouse("sTownIDxx") = Category.spouse_info.residence_info.present_address.sTownIDxx
                End If
            End If

            .Term("sModelIDx") = Category.sModelIDx
            .Term("nDownPaym") = Category.nDownPaym
            .Term("nAcctTerm") = Category.nAcctTerm

            'Execute quickmatch here
            GenerateQM = .QuickMatch

            If GenerateQM = "" Then
                Exit Function
            End If

            loFrm = New frmQuickMatch
            loFrm.Appdriver = p_oApp

            loFrm.txtField00.Text = .TransNo
            loFrm.txtField04.Text = Category.applicant_info.sLastName & _
                         ", " & Category.applicant_info.sFrstName & IIf(IFNull(Category.applicant_info.sSuffixNm) = "", "", " " & Category.applicant_info.sSuffixNm) & _
                          " " & Category.applicant_info.sMiddName
            loFrm.txtField20.Text = Category.residence_info.present_address.sAddress1 & _
                         ", " & getTownCity(Category.residence_info.present_address.sTownIDxx, True, True, "")

            If Not IsNothing(Category.spouse_info.personal_info.sLastName) Then
                'Display spouse info
                If IFNull(Category.spouse_info.personal_info.sLastName) = "" Then
                    loFrm.txtField06.Text = "N-O-N-E"
                    loFrm.txtField07.Text = "N-O-N-E"
                Else
                    loFrm.txtField06.Text = Category.spouse_info.personal_info.sLastName & _
                                 ", " & Category.spouse_info.personal_info.sFrstName & IIf(IFNull(Category.spouse_info.personal_info.sSuffixNm) = "", "", " " & Category.spouse_info.personal_info.sSuffixNm) & _
                                  " " & Category.spouse_info.personal_info.sMiddName
                    loFrm.txtField07.Text = Category.spouse_info.residence_info.present_address.sAddress1 & _
                                 ", " & getTownCity(Category.spouse_info.residence_info.present_address.sTownIDxx, True, True, "")
                End If
            End If

            p_oDTMstr.Rows(0)("sQMatchNo") = GenerateQM
            loFrm.txtField08.Text = p_oDTMstr.Rows(0)("sQMatchNo")
            loFrm.txtField09.Text = p_oDTMstr.Rows(0)("sTransNox")
            loFrm.txtField05.Text = Format(Master("dTransact"), "Mmmm DD, YYYY")

            loFrm.Result = .Result
            p_oFrmResult = loFrm
            loFrm.ShowDialog()
        End With
    End Function

    Private Sub showQMResult()
        p_oFrmResult.ShowDialog()
    End Sub

    Sub showQMResult(ByVal fsTransNox As String, ByVal fdTransact As Date)
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = "SELECT" & _
                    "  a.sTransNox" & _
                    ", a.sApplicNo" & _
                    ", a.sLastName" & _
                    ", a.sFrstName" & _
                    ", a.sMiddName" & _
                    ", a.dBirthDte" & _
                    ", a.sBirthPlc" & _
                    ", b.sTownName" & _
                    ", a.sSLastNme" & _
                    ", a.sSFrstNme" & _
                    ", a.sSMiddNme" & _
                    ", a.dSBrthDte" & _
                    ", a.sSBrthPlc" & _
                    ", c.sTownName sSTownNme" & _
                    ", a.sModelIDx" & _
                    ", a.nDownPaym" & _
                    ", a.nAcctTerm" & _
                    ", a.sResltCde" & _
                " FROM MC_LR_QuickMatch a" & _
                    " LEFT JOIN TownCity b" & _
                        " ON a.sTownIDxx = b.sTownIDxx" & _
                    " LEFT JOIN TownCity c" & _
                        " ON a.sSTownIDx = c.sTownIDxx" & _
                " WHERE a.sApplicNo = " & strParm(fsTransNox)

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)
        If loDT.Rows.Count = 0 Then Exit Sub

        p_oFrmResult = New frmQuickMatch
        With p_oFrmResult
            .txtField00.Text = loDT(0)("sTransNox")
            .txtField04.Text = loDT(0)("sLastName") & ", " & loDT(0)("sFrstName") & " " & loDT(0)("sMiddName")
            .txtField20.Text = IFNull(loDT(0)("sTownName"), "")
            .txtField06.Text = IFNull(loDT(0)("sSLastNme"), "") & ", " & IFNull(loDT(0)("sFrstName"), "") & " " & IFNull(loDT(0)("sMiddName"), "")
            .txtField07.Text = IFNull(loDT(0)("sSTownNme"), "")
            .txtField05.Text = Format(fdTransact, xsDATE_MEDIUM)
            .txtField09.Text = loDT(0)("sApplicNo")
            .txtField08.Text = loDT(0)("sResltCde")

            loDT = New DataTable
            loDT = p_oApp.ExecuteQuery("SELECT" & _
                                            "  CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) sFullName" & _
                                            ", a.sResltCde" & _
                                            ", a.sAcctNmbr" & _
                                            ", a.sMCSONmbr" & _
                                            ", a.sApplNmbr" & _
                                        " FROM MC_LR_QuickMatch_Result a" & _
                                            " LEFT JOIN Client_Master b" & _
                                                " ON a.sClientID = b.sClientID" & _
                                        " WHERE sTransNox = " & strParm(.txtField00.Text))
            .Result = loDT
            .ShowDialog()
        End With
    End Sub

    Private Function FromClass(Of T)(ByVal data As T,
                                   Optional ByVal isEmptyToNull As Boolean = False,
                                   Optional ByVal jsonSettings As JsonSerializerSettings = Nothing) As String

        Dim response As String = String.Empty

        If Not EqualityComparer(Of T).Default.Equals(data, Nothing) Then
            response = JsonConvert.SerializeObject(data, jsonSettings)
        End If

        Return If(isEmptyToNull, (If(response = "{}", "null", response)), response)
    End Function

    Private Function ToClass(Of T)(ByVal data As String,
                                 Optional ByVal jsonSettings As JsonSerializerSettings = Nothing) As T

        Dim response = Nothing

        If Not String.IsNullOrEmpty(data) Then
            response = If(jsonSettings Is Nothing,
                JsonConvert.DeserializeObject(Of T)(data),
                JsonConvert.DeserializeObject(Of T)(data, jsonSettings))
        End If

        Return response
    End Function

    Private Sub populateJSONObjectDetail(ByVal foJSONObject As App_DetParam, _
                                    ByVal fsJSONValue As String)
        Dim loSettings As New JsonSerializerSettings
        loSettings.DefaultValueHandling = DefaultValueHandling.Populate
        Dim loJSONObject() As App_DetParam = JsonConvert.DeserializeObject(Of App_DetParam())(fsJSONValue, loSettings)
        If fsJSONValue = "" Then Exit Sub

        With foJSONObject
            .sLoanType = loJSONObject(0).sLoanType
            .MCModelx = loJSONObject(0).MCModelx
            .nLoanTerm = loJSONObject(0).nLoanTerm
            .sBranchCd = loJSONObject(0).sBranchCd

            .sFrstName = loJSONObject(1).sFrstName
            .sMiddName = loJSONObject(1).sMiddName
            .sLastName = loJSONObject(1).sLastName
            .sSuffixNm = loJSONObject(1).sSuffixNm
            .sNickName = loJSONObject(1).sNickName
            .sGenderxx = loJSONObject(1).sGenderxx
            .sCvilStat = loJSONObject(1).sCvilStat
            .sPresAddr = loJSONObject(1).sPresAddr
            .sPrevAddr = loJSONObject(1).sPrevAddr
            .sLenStayx = loJSONObject(1).sLenStayx
            .sMobileNo = loJSONObject(1).sMobileNo
            .sEmailAdd = loJSONObject(1).sEmailAdd
            .dBirthDte = loJSONObject(1).dBirthDte
            .sBrtPlace = loJSONObject(1).sBrtPlace
            .nAgexxxxx = loJSONObject(1).nAgexxxxx
            .sMotherNm = loJSONObject(1).sMotherNm
            .sFatherNm = loJSONObject(1).sFatherNm
            .sParentAd = loJSONObject(1).sParentAd

            .sEmplType = loJSONObject(2).sEmplType
            .sCompnyNm = loJSONObject(2).sCompnyNm
            .sCompnyAd = loJSONObject(2).sCompnyAd
            .sCompTele = loJSONObject(2).sCompTele
            .sLenServe = loJSONObject(2).sLenServe
            .sGrIncome = loJSONObject(2).sGrIncome
            .sEmplStat = loJSONObject(2).sEmplStat
            .sEmpPostn = loJSONObject(2).sEmpPostn
            .sBusiness = loJSONObject(2).sBusiness
            .sBusiAddr = loJSONObject(2).sBusiAddr
            .sBusiTele = loJSONObject(2).sBusiTele
            .sBusIncom = loJSONObject(2).sBusIncom
            .sYrInBusi = loJSONObject(2).sYrInBusi
            .sSourceIn = loJSONObject(2).sSourceIn

            .sBankNme1 = loJSONObject(3).sBankNme1
            .sBankBrh1 = loJSONObject(3).sBankBrh1
            .sBankAcc1 = loJSONObject(3).sBankAcc1
            .sBankNme2 = loJSONObject(3).sBankNme2
            .sBankBrh2 = loJSONObject(3).sBankBrh2
            .sBankAcc2 = loJSONObject(3).sBankAcc2

            .sRefName1 = loJSONObject(4).sRefName1
            .sRefAddr1 = loJSONObject(4).sRefAddr1
            .sRefName2 = loJSONObject(4).sRefName2
            .sRefAddr2 = loJSONObject(4).sRefAddr2

            .sSpFrstNm = loJSONObject(5).sSpFrstNm
            .sSpMiddNm = loJSONObject(5).sSpMiddNm
            .sSpLastNm = loJSONObject(5).sSpLastNm
            .sSpSuffNm = loJSONObject(5).sSpSuffNm
            .sSpNickNm = loJSONObject(5).sSpNickNm
            .sSpPresAd = loJSONObject(5).sSpPresAd
            .sSpPrevAd = loJSONObject(5).sSpPrevAd
            .sSpLenSty = loJSONObject(5).sSpLenSty
            .sSpMobiNo = loJSONObject(5).sSpMobiNo
            .sSpEmailx = loJSONObject(5).sSpEmailx
            .dSpBrtDte = loJSONObject(5).dSpBrtDte
            .sSpBrtPlc = loJSONObject(5).sSpBrtPlc
            .nSpAgexxx = loJSONObject(5).nSpAgexxx

            .sSpCompNm = loJSONObject(6).sSpCompNm
            .sSpCompAd = loJSONObject(6).sSpCompAd
            .sSpComTel = loJSONObject(6).sSpComTel
            .sSpLenSrv = loJSONObject(6).sSpLenSrv
            .sSpMonPay = loJSONObject(6).sSpMonPay
            .sSpEmpSta = loJSONObject(6).sSpEmpSta
            .sSpEmpPos = loJSONObject(6).sSpEmpPos
            .sSpBusins = loJSONObject(6).sSpBusins
            .sSpBusiAd = loJSONObject(6).sSpBusiAd
            .sSpBusTel = loJSONObject(6).sSpBusTel
            .sSpBusInc = loJSONObject(6).sSpBusInc
            .sSpYrsBus = loJSONObject(6).sSpYrsBus
            .sSpSrcInc = loJSONObject(6).sSpSrcInc

            .sChldNme1 = loJSONObject(7).sChldNme1
            .sChldAge1 = loJSONObject(7).sChldAge1
            .sChldSch1 = loJSONObject(7).sChldSch1
            .sChldNme2 = loJSONObject(7).sChldNme2
            .sChldAge2 = loJSONObject(7).sChldAge2
            .sChldSch2 = loJSONObject(7).sChldSch2
            .sChldNme3 = loJSONObject(7).sChldNme3
            .sChldAge3 = loJSONObject(7).sChldAge3
            .sChldSch3 = loJSONObject(7).sChldSch3

            .sRentalxx = loJSONObject(8).sRentalxx
            .sElectric = loJSONObject(8).sElectric
            .sWaterBil = loJSONObject(8).sWaterBil
            .sOthrLoan = loJSONObject(8).sOthrLoan
            .sCredtCrd = loJSONObject(8).sCredtCrd
            .sCredtLmt = loJSONObject(8).sCredtLmt
            .sEducAttn = loJSONObject(8).sEducAttn
            .sNumOfCPs = loJSONObject(8).sNumOfCPs
            .sCPNumbr1 = loJSONObject(8).sCPNumbr1
            .sCPTypex1 = loJSONObject(8).sCPTypex1
            .sCPNumbr2 = loJSONObject(8).sCPNumbr2
            .sCPTypex2 = loJSONObject(8).sCPTypex2
            .sCPNumbr3 = loJSONObject(8).sCPNumbr3
            .sCPTypex3 = loJSONObject(8).sCPTypex3
            .sLandmark = loJSONObject(8).sLandmark

            .sCOFrstNm = loJSONObject(9).sCOFrstNm
            .sCOMiddNm = loJSONObject(9).sCOMiddNm
            .sCOLastNm = loJSONObject(9).sCOLastNm
            .sCORelatn = loJSONObject(9).sCORelatn
            .sCOOccptn = loJSONObject(9).sCOOccptn
            .sCONation = loJSONObject(9).sCONation
            .sCORemitt = loJSONObject(9).sCORemitt
            .sCOContct = loJSONObject(9).sCOContct
            .sCORoamNo = loJSONObject(9).sCORoamNo
            .sCOEmailx = loJSONObject(9).sCOEmailx

            .sCLFrstNm = loJSONObject(10).sCLFrstNm
            .sCLMiddNm = loJSONObject(10).sCLMiddNm
            .sCLLastNm = loJSONObject(10).sCLLastNm
            .sCLRelatn = loJSONObject(10).sCLRelatn
            .sCLAddres = loJSONObject(10).sCLAddres
            .sCLEmploy = loJSONObject(10).sCLEmploy
            .sCLContct = loJSONObject(10).sCLContct
            .sCLBrtPlc = loJSONObject(10).sCLBrtPlc
            .sCLBrtDte = loJSONObject(10).sCLBrtDte
            .sCLEmailx = loJSONObject(10).sCLEmailx
        End With
    End Sub

    Private Sub populateJSONObject(ByVal foJSONObject As GOCAS_Param, _
                                    ByVal fsJSONValue As String)
        Dim loSettings As New JsonSerializerSettings
        loSettings.DefaultValueHandling = DefaultValueHandling.Populate
        Dim loJSONObject As GOCAS_Param = JsonConvert.DeserializeObject(Of GOCAS_Param)(fsJSONValue, loSettings)

        With foJSONObject
            .sBranchCd = If(fsJSONValue = "", jsonObjDet.sBranchCd, loJSONObject.sBranchCd)
            .dAppliedx = p_oApp.getSysDate
            .sClientNm = ""
            .cUnitAppl = "0"
            .nDownPaym = If(fsJSONValue = "", jsonObjDet.nDownPaym, loJSONObject.nDownPaym)
            .dCreatedx = p_oApp.getSysDate
            .cApplType = If(fsJSONValue = "", "0", loJSONObject.cApplType)
            .sUnitAppl = ""
            .sModelIDx = If(fsJSONValue = "", jsonObjDet.MCModelx, loJSONObject.sModelIDx)
            .nAcctTerm = If(fsJSONValue = "", jsonObjDet.nLoanTerm, loJSONObject.nAcctTerm)
            .nMonAmort = 0
            .dTargetDt = p_oApp.getSysDate

            With .applicant_info
                .sLastName = If(fsJSONValue = "", jsonObjDet.sLastName, loJSONObject.applicant_info.sLastName)
                .sFrstName = If(fsJSONValue = "", jsonObjDet.sFrstName, loJSONObject.applicant_info.sFrstName)
                .sSuffixNm = If(fsJSONValue = "", jsonObjDet.sSuffixNm, loJSONObject.applicant_info.sSuffixNm)
                .sMiddName = If(fsJSONValue = "", jsonObjDet.sMiddName, loJSONObject.applicant_info.sMiddName)
                .sNickName = If(fsJSONValue = "", jsonObjDet.sNickName, loJSONObject.applicant_info.sNickName)
                .dBirthDte = If(fsJSONValue = "", jsonObjDet.dBirthDte, loJSONObject.applicant_info.dBirthDte)
                .sBirthPlc = If(fsJSONValue = "", jsonObjDet.sBrtPlace, loJSONObject.applicant_info.sBirthPlc)
                .sCitizenx = If(fsJSONValue = "", "", loJSONObject.applicant_info.sCitizenx)

                If fsJSONValue = "" Then
                    .mobile_number.Add(New GOCASConst.mobileno_param)
                    .mobile_number(0).sMobileNo = jsonObjDet.sMobileNo
                    .mobile_number(0).cPostPaid = ""
                    .mobile_number(0).nPostYear = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.mobile_number.Count - 1
                        .mobile_number.Add(New GOCASConst.mobileno_param)
                        .mobile_number(nCtr).sMobileNo = loJSONObject.applicant_info.mobile_number(nCtr).sMobileNo
                        .mobile_number(nCtr).cPostPaid = loJSONObject.applicant_info.mobile_number(nCtr).cPostPaid
                        .mobile_number(nCtr).nPostYear = loJSONObject.applicant_info.mobile_number(nCtr).nPostYear
                    Next
                End If

                If fsJSONValue = "" Then
                    .landline.Add(New GOCASConst.landline_param)
                    .landline(0).sPhoneNox = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.landline.Count - 1
                        .landline.Add(New GOCASConst.landline_param)
                        .landline(nCtr).sPhoneNox = loJSONObject.applicant_info.landline(nCtr).sPhoneNox
                    Next
                End If

                .cCvilStat = If(fsJSONValue = "", "", loJSONObject.applicant_info.cCvilStat)
                .cGenderCd = If(fsJSONValue = "", "", loJSONObject.applicant_info.cGenderCd)
                .sMaidenNm = If(fsJSONValue = "", jsonObjDet.sMotherNm, loJSONObject.applicant_info.sMaidenNm)

                If fsJSONValue = "" Then
                    .email_address.Add(New GOCASConst.email_param)
                    .email_address(0).sEmailAdd = jsonObjDet.sEmailAdd
                Else
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.email_address.Count - 1
                        .email_address.Add(New GOCASConst.email_param)
                        .email_address(nCtr).sEmailAdd = loJSONObject.applicant_info.email_address(nCtr).sEmailAdd
                    Next
                End If

                .facebook.sFBAcctxx = ""
                .facebook.cAcctStat = ""
                .facebook.nNoFriend = 0
                .facebook.nYearxxxx = 0
                .sVibeAcct = ""
            End With

            With .residence_info
                .cOwnershp = If(fsJSONValue = "", "", loJSONObject.residence_info.cOwnershp)
                .cOwnOther = If(fsJSONValue = "", "", loJSONObject.residence_info.cOwnOther)

                .rent_others.cRntOther = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.cRntOther)
                .rent_others.nLenStayx = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.nLenStayx)
                .rent_others.nRentExps = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.nRentExps)

                .sCtkReltn = If(fsJSONValue = "", "", loJSONObject.residence_info.sCtkReltn)
                .cHouseTyp = If(fsJSONValue = "", "", loJSONObject.residence_info.cHouseTyp)
                .cGaragexx = If(fsJSONValue = "", "", loJSONObject.residence_info.cGaragexx)

                .present_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sLandMark)
                .present_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sHouseNox)
                .present_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sAddress1)
                .present_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sAddress2)
                .present_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sTownIDxx)
                .present_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sBrgyIDxx)

                .permanent_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sLandMark)
                .permanent_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sHouseNox)
                .permanent_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sAddress1)
                .permanent_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sAddress2)
                .permanent_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sTownIDxx)
                .permanent_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sBrgyIDxx)
            End With

            With .means_info
                .cIncmeSrc = If(fsJSONValue = "", "", loJSONObject.means_info.cIncmeSrc)
                .employed.cEmpSectr = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cEmpSectr)
                .employed.cUniforme = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cUniforme)
                .employed.cMilitary = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cMilitary)
                .employed.cGovtLevl = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cGovtLevl)
                .employed.cCompLevl = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cCompLevl)
                .employed.cEmpLevlx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cEmpLevlx)
                .employed.cOcCatgry = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cOcCatgry)
                .employed.cOFWRegnx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cOFWRegnx)
                .employed.sOFWNatnx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.sOFWNatnx)
                .employed.sIndstWrk = ""
                .employed.sEmployer = If(fsJSONValue = "", jsonObjDet.sCompnyNm, loJSONObject.means_info.employed.sEmployer)
                .employed.sWrkAddrx = ""
                .employed.sWrkTownx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.sWrkTownx)
                .employed.sPosition = If(fsJSONValue = "", "", loJSONObject.means_info.employed.sPosition)
                .employed.sFunction = ""
                .employed.cEmpStatx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cEmpStatx)
                .employed.nLenServc = If(fsJSONValue = "", jsonObjDet.sLenServe, loJSONObject.means_info.employed.nLenServc)
                .employed.nSalaryxx = If(fsJSONValue = "", jsonObjDet.sGrIncome, loJSONObject.means_info.employed.nSalaryxx)
                .employed.sWrkTelno = If(fsJSONValue = "", jsonObjDet.sCompTele, loJSONObject.means_info.employed.sWrkTelno)
                .self_employed.sIndstBus = If(fsJSONValue = "", jsonObjDet.sBusiness, loJSONObject.means_info.self_employed.sIndstBus)
                .self_employed.sBusiness = ""
                .self_employed.sBusAddrx = ""
                .self_employed.sBusTownx = If(fsJSONValue = "", jsonObjDet.sBusiAddr, loJSONObject.means_info.self_employed.sBusTownx)
                .self_employed.cBusTypex = ""
                .self_employed.nBusLenxx = If(fsJSONValue = "", jsonObjDet.sYrInBusi, loJSONObject.means_info.self_employed.nBusLenxx)
                .self_employed.nBusIncom = If(fsJSONValue = "", jsonObjDet.sBusIncom, loJSONObject.means_info.self_employed.nBusIncom)
                .self_employed.nMonExpns = 0
                .self_employed.cOwnTypex = ""
                .self_employed.cOwnSizex = ""
                .pensioner.cPenTypex = ""
                .pensioner.nPensionx = 0
                .pensioner.nRetrYear = 0
                .financed.sReltnCde = ""
                .financed.sFinancer = ""
                .financed.nEstIncme = 0
                .financed.sNatnCode = ""
                .financed.sMobileNo = ""
                .financed.sFBAcctxx = ""
                .financed.sEmailAdd = ""

                .other_income.sOthrIncm = If(fsJSONValue = "", jsonObjDet.sSourceIn, loJSONObject.means_info.other_income.sOthrIncm)
                .other_income.nOthrIncm = ""
            End With

            With .other_info
                .sUnitUser = ""
                .sUsr2Buyr = ""
                .sPurposex = ""
                .sUnitPayr = ""
                .sPyr2Buyr = ""
                .sSrceInfo = ""
                If fsJSONValue = "" Then
                    .personal_reference.Add(New GOCASConst.personal_reference_param)
                    .personal_reference(0).sRefrNmex = jsonObjDet.sRefName1
                    .personal_reference(0).sRefrMPNx = ""
                    .personal_reference(0).sRefrAddx = jsonObjDet.sRefAddr1
                    .personal_reference(0).sRefrTown = ""

                    .personal_reference.Add(New GOCASConst.personal_reference_param)
                    .personal_reference(1).sRefrNmex = jsonObjDet.sRefName2
                    .personal_reference(1).sRefrMPNx = ""
                    .personal_reference(1).sRefrAddx = jsonObjDet.sRefAddr2
                    .personal_reference(1).sRefrTown = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.other_info.personal_reference.Count - 1
                        .personal_reference.Add(New GOCASConst.personal_reference_param)
                        .personal_reference(nCtr).sRefrNmex = loJSONObject.other_info.personal_reference(nCtr).sRefrNmex
                        .personal_reference(nCtr).sRefrMPNx = loJSONObject.other_info.personal_reference(nCtr).sRefrMPNx
                        .personal_reference(nCtr).sRefrAddx = loJSONObject.other_info.personal_reference(nCtr).sRefrAddx
                        .personal_reference(nCtr).sRefrTown = loJSONObject.other_info.personal_reference(nCtr).sRefrTown
                    Next
                End If
            End With

            With .comaker_info
                .sLastName = If(fsJSONValue = "", jsonObjDet.sCLLastNm, loJSONObject.comaker_info.sLastName)
                .sFrstName = If(fsJSONValue = "", jsonObjDet.sCLFrstNm, loJSONObject.comaker_info.sFrstName)
                .sSuffixNm = ""
                .sMiddName = If(fsJSONValue = "", jsonObjDet.sCLMiddNm, loJSONObject.comaker_info.sMiddName)
                .sNickName = ""
                .dBirthDte = If(fsJSONValue = "", jsonObjDet.sCLBrtDte, loJSONObject.comaker_info.dBirthDte)
                .sBirthPlc = If(fsJSONValue = "", jsonObjDet.sCLBrtPlc, loJSONObject.comaker_info.sBirthPlc)
                .cIncmeSrc = ""
                .sReltnCde = If(fsJSONValue = "", "", loJSONObject.comaker_info.sReltnCde)
                If fsJSONValue = "" Then
                    .mobile_number.Add(New GOCASConst.mobileno_param)
                    .mobile_number(0).sMobileNo = jsonObjDet.sCLContct
                    .mobile_number(0).cPostPaid = ""
                    .mobile_number(0).nPostYear = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.comaker_info.mobile_number.Count - 1
                        .mobile_number.Add(New GOCASConst.mobileno_param)
                        .mobile_number(nCtr).sMobileNo = loJSONObject.comaker_info.mobile_number(nCtr).sMobileNo
                        .mobile_number(nCtr).cPostPaid = loJSONObject.comaker_info.mobile_number(nCtr).cPostPaid
                        .mobile_number(nCtr).nPostYear = loJSONObject.comaker_info.mobile_number(nCtr).nPostYear
                    Next
                End If
                .sFBAcctxx = ""
            End With

            With .spouse_info
                .personal_info.sLastName = If(fsJSONValue = "", jsonObjDet.sSpLastNm, loJSONObject.spouse_info.personal_info.sLastName)
                .personal_info.sFrstName = If(fsJSONValue = "", jsonObjDet.sSpFrstNm, loJSONObject.spouse_info.personal_info.sFrstName)
                .personal_info.sSuffixNm = If(fsJSONValue = "", jsonObjDet.sSpSuffNm, loJSONObject.spouse_info.personal_info.sSuffixNm)
                .personal_info.sMiddName = If(fsJSONValue = "", jsonObjDet.sSpMiddNm, loJSONObject.spouse_info.personal_info.sMiddName)
                .personal_info.sNickName = If(fsJSONValue = "", jsonObjDet.sSpNickNm, loJSONObject.spouse_info.personal_info.sNickName)
                .personal_info.dBirthDte = If(fsJSONValue = "", jsonObjDet.dSpBrtDte, loJSONObject.spouse_info.personal_info.dBirthDte)
                .personal_info.sBirthPlc = If(fsJSONValue = "", jsonObjDet.sSpBrtPlc, loJSONObject.spouse_info.personal_info.sBirthPlc)
                .personal_info.sCitizenx = If(fsJSONValue = "", "", loJSONObject.spouse_info.personal_info.sCitizenx)

                If fsJSONValue = "" Then
                    .personal_info.mobile_number.Add(New GOCASConst.mobileno_param)
                    .personal_info.mobile_number(0).sMobileNo = jsonObjDet.sSpMobiNo
                    .personal_info.mobile_number(0).cPostPaid = ""
                    .personal_info.mobile_number(0).nPostYear = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.mobile_number.Count - 1
                        .personal_info.mobile_number.Add(New GOCASConst.mobileno_param)
                        .personal_info.mobile_number(nCtr).sMobileNo = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).sMobileNo
                        .personal_info.mobile_number(nCtr).cPostPaid = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).cPostPaid
                        .personal_info.mobile_number(nCtr).nPostYear = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).nPostYear
                    Next
                End If

                If fsJSONValue = "" Then
                    .personal_info.landline.Add(New GOCASConst.landline_param)
                    .personal_info.landline(0).sPhoneNox = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.landline.Count - 1
                        .personal_info.landline.Add(New GOCASConst.landline_param)
                        .personal_info.landline(nCtr).sPhoneNox = loJSONObject.spouse_info.personal_info.landline(nCtr).sPhoneNox
                    Next
                End If

                .personal_info.cCvilStat = ""
                .personal_info.cGenderCd = ""
                .personal_info.sMaidenNm = If(fsJSONValue = "", jsonObjDet.sSpMiddNm, loJSONObject.spouse_info.personal_info.sMaidenNm)
                If fsJSONValue = "" Then
                    .personal_info.email_address.Add(New GOCASConst.email_param)
                    .personal_info.email_address(0).sEmailAdd = jsonObjDet.sSpEmailx
                Else
                    For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.email_address.Count - 1
                        .personal_info.email_address.Add(New GOCASConst.email_param)
                        .personal_info.email_address(nCtr).sEmailAdd = loJSONObject.spouse_info.personal_info.email_address(nCtr).sEmailAdd
                    Next
                End If

                .personal_info.facebook.sFBAcctxx = ""
                .personal_info.facebook.cAcctStat = ""
                .personal_info.facebook.nNoFriend = 0
                .personal_info.facebook.nYearxxxx = 0
                .personal_info.sVibeAcct = ""

                .residence_info.cOwnershp = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cOwnershp)
                .residence_info.cOwnOther = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cOwnOther)

                .residence_info.rent_others.cRntOther = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.cRntOther)
                .residence_info.rent_others.nLenStayx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.nLenStayx)
                .residence_info.rent_others.nRentExps = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.nRentExps)

                .residence_info.sCtkReltn = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.sCtkReltn)
                .residence_info.cHouseTyp = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cHouseTyp)
                .residence_info.cGaragexx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cGaragexx)

                .residence_info.present_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sLandMark)
                .residence_info.present_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sHouseNox)
                .residence_info.present_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sAddress1)
                .residence_info.present_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sAddress2)
                .residence_info.present_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sTownIDxx)
                .residence_info.present_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sBrgyIDxx)

                .residence_info.permanent_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sLandMark)
                .residence_info.permanent_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sHouseNox)
                .residence_info.permanent_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sAddress1)
                .residence_info.permanent_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sAddress2)
                .residence_info.permanent_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sTownIDxx)
                .residence_info.permanent_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sBrgyIDxx)

            End With
            With .spouse_means
                .cIncmeSrc = If(fsJSONValue = "", "", loJSONObject.spouse_means.cIncmeSrc)
                .employed.cEmpSectr = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cEmpSectr)
                .employed.cUniforme = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cUniforme)
                .employed.cMilitary = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cMilitary)
                .employed.cGovtLevl = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cGovtLevl)
                .employed.cCompLevl = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cCompLevl)
                .employed.cEmpLevlx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cEmpLevlx)
                .employed.cOcCatgry = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cOcCatgry)
                .employed.cOFWRegnx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cOFWRegnx)
                .employed.sOFWNatnx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.sOFWNatnx)
                .employed.sIndstWrk = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.sIndstWrk)
                .employed.sEmployer = If(fsJSONValue = "", jsonObjDet.sSpCompNm, loJSONObject.spouse_means.employed.sEmployer)
                .employed.sWrkAddrx = ""
                .employed.sWrkTownx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.sWrkTownx)
                .employed.sPosition = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.sPosition)
                .employed.sFunction = ""
                .employed.cEmpStatx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cEmpStatx)
                .employed.nLenServc = If(fsJSONValue = "", jsonObjDet.sSpLenSrv, loJSONObject.spouse_means.employed.nLenServc)
                .employed.nSalaryxx = If(fsJSONValue = "", jsonObjDet.sSpMonPay, loJSONObject.spouse_means.employed.nSalaryxx)
                .employed.sWrkTelno = If(fsJSONValue = "", jsonObjDet.sSpComTel, loJSONObject.spouse_means.employed.sWrkTelno)

                .self_employed.sIndstBus = If(fsJSONValue = "", jsonObjDet.sSpBusins, loJSONObject.spouse_means.self_employed.sIndstBus)
                .self_employed.sBusiness = ""
                .self_employed.sBusAddrx = ""
                .self_employed.sBusTownx = If(fsJSONValue = "", jsonObjDet.sSpBusiAd, loJSONObject.spouse_means.self_employed.sBusTownx)
                .self_employed.cBusTypex = ""
                .self_employed.nBusLenxx = If(fsJSONValue = "", jsonObjDet.sSpYrsBus, loJSONObject.spouse_means.self_employed.nBusLenxx)
                .self_employed.nBusIncom = If(fsJSONValue = "", jsonObjDet.sSpBusInc, loJSONObject.spouse_means.self_employed.nBusIncom)
                .self_employed.nMonExpns = ""
                .self_employed.cOwnTypex = ""
                .self_employed.cOwnSizex = ""

                .pensioner.cPenTypex = ""
                .pensioner.nPensionx = ""
                .pensioner.nRetrYear = ""

                .financed.sReltnCde = ""
                .financed.sFinancer = ""
                .financed.nEstIncme = ""
                .financed.sNatnCode = ""
                .financed.sMobileNo = ""
                .financed.sFBAcctxx = ""
                .financed.sEmailAdd = ""

                .other_income.sOthrIncm = If(fsJSONValue = "", jsonObjDet.sSpSrcInc, loJSONObject.spouse_means.other_income.sOthrIncm)
                .other_income.nOthrIncm = ""
            End With

            With .disbursement_info
                .dependent_info.nHouseHld = ""
                If fsJSONValue = "" Then
                    .dependent_info.children.Add(New GOCASConst.children_param)
                    .dependent_info.children(0).sFullName = jsonObjDet.sChldNme1
                    .dependent_info.children(0).sRelatnCD = ""
                    .dependent_info.children(0).nDepdAgex = jsonObjDet.sChldAge1
                    .dependent_info.children(0).cIsPupilx = ""
                    .dependent_info.children(0).sSchlName = jsonObjDet.sChldSch1
                    .dependent_info.children(0).sSchlAddr = ""
                    .dependent_info.children(0).sSchlTown = ""
                    .dependent_info.children(0).cIsPrivte = ""
                    .dependent_info.children(0).sEducLevl = ""
                    .dependent_info.children(0).cIsSchlrx = ""
                    .dependent_info.children(0).cHasWorkx = ""
                    .dependent_info.children(0).cWorkType = ""
                    .dependent_info.children(0).sCompanyx = ""
                    .dependent_info.children(0).cHouseHld = ""
                    .dependent_info.children(0).cDependnt = ""
                    .dependent_info.children(0).cIsChildx = ""
                    .dependent_info.children(0).cIsMarrdx = ""

                    .dependent_info.children.Add(New GOCASConst.children_param)
                    .dependent_info.children(1).sFullName = jsonObjDet.sChldNme2
                    .dependent_info.children(1).sRelatnCD = ""
                    .dependent_info.children(1).nDepdAgex = jsonObjDet.sChldAge2
                    .dependent_info.children(1).cIsPupilx = ""
                    .dependent_info.children(1).sSchlName = jsonObjDet.sChldSch2
                    .dependent_info.children(1).sSchlAddr = ""
                    .dependent_info.children(1).sSchlTown = ""
                    .dependent_info.children(1).cIsPrivte = ""
                    .dependent_info.children(1).sEducLevl = ""
                    .dependent_info.children(1).cIsSchlrx = ""
                    .dependent_info.children(1).cHasWorkx = ""
                    .dependent_info.children(1).cWorkType = ""
                    .dependent_info.children(1).sCompanyx = ""
                    .dependent_info.children(1).cHouseHld = ""
                    .dependent_info.children(1).cDependnt = ""
                    .dependent_info.children(1).cIsChildx = ""
                    .dependent_info.children(1).cIsMarrdx = ""

                    .dependent_info.children.Add(New GOCASConst.children_param)
                    .dependent_info.children(2).sFullName = jsonObjDet.sChldNme3
                    .dependent_info.children(2).sRelatnCD = ""
                    .dependent_info.children(2).nDepdAgex = jsonObjDet.sChldAge3
                    .dependent_info.children(2).cIsPupilx = ""
                    .dependent_info.children(2).sSchlName = jsonObjDet.sChldSch3
                    .dependent_info.children(2).sSchlAddr = ""
                    .dependent_info.children(2).sSchlTown = ""
                    .dependent_info.children(2).cIsPrivte = ""
                    .dependent_info.children(2).sEducLevl = ""
                    .dependent_info.children(2).cIsSchlrx = ""
                    .dependent_info.children(2).cHasWorkx = ""
                    .dependent_info.children(2).cWorkType = ""
                    .dependent_info.children(2).sCompanyx = ""
                    .dependent_info.children(2).cHouseHld = ""
                    .dependent_info.children(2).cDependnt = ""
                    .dependent_info.children(2).cIsChildx = ""
                    .dependent_info.children(2).cIsMarrdx = ""
                Else
                    For nCtr As Integer = 0 To loJSONObject.disbursement_info.dependent_info.children.Count - 1
                        .dependent_info.children.Add(New GOCASConst.children_param)
                        .dependent_info.children(nCtr).sFullName = loJSONObject.disbursement_info.dependent_info.children(nCtr).sFullName
                        .dependent_info.children(nCtr).sRelatnCD = loJSONObject.disbursement_info.dependent_info.children(nCtr).sRelatnCD
                        .dependent_info.children(nCtr).nDepdAgex = loJSONObject.disbursement_info.dependent_info.children(nCtr).nDepdAgex
                        .dependent_info.children(nCtr).cIsPupilx = loJSONObject.disbursement_info.dependent_info.children(nCtr).cIsPupilx
                        .dependent_info.children(nCtr).sSchlName = loJSONObject.disbursement_info.dependent_info.children(nCtr).sSchlName
                        .dependent_info.children(nCtr).sSchlAddr = loJSONObject.disbursement_info.dependent_info.children(nCtr).sSchlAddr
                        .dependent_info.children(nCtr).sSchlTown = loJSONObject.disbursement_info.dependent_info.children(nCtr).sSchlTown
                        .dependent_info.children(nCtr).cIsPrivte = loJSONObject.disbursement_info.dependent_info.children(nCtr).cIsPrivte
                        .dependent_info.children(nCtr).sEducLevl = loJSONObject.disbursement_info.dependent_info.children(nCtr).sEducLevl
                        .dependent_info.children(nCtr).cIsSchlrx = loJSONObject.disbursement_info.dependent_info.children(nCtr).cIsSchlrx
                        .dependent_info.children(nCtr).cHasWorkx = loJSONObject.disbursement_info.dependent_info.children(nCtr).cHasWorkx
                        .dependent_info.children(nCtr).cWorkType = loJSONObject.disbursement_info.dependent_info.children(nCtr).cWorkType
                        .dependent_info.children(nCtr).sCompanyx = loJSONObject.disbursement_info.dependent_info.children(nCtr).sCompanyx
                        .dependent_info.children(nCtr).cHouseHld = loJSONObject.disbursement_info.dependent_info.children(nCtr).cHouseHld
                        .dependent_info.children(nCtr).cDependnt = loJSONObject.disbursement_info.dependent_info.children(nCtr).cDependnt
                        .dependent_info.children(nCtr).cIsChildx = loJSONObject.disbursement_info.dependent_info.children(nCtr).cIsChildx
                        .dependent_info.children(nCtr).cIsMarrdx = loJSONObject.disbursement_info.dependent_info.children(nCtr).cIsMarrdx
                    Next
                End If

                .properties.sProprty1 = ""
                .properties.sProprty2 = ""
                .properties.sProprty3 = ""
                .properties.cWith4Whl = ""
                .properties.cWith3Whl = ""
                .properties.cWith2Whl = ""
                .properties.cWithRefx = ""
                .properties.cWithTVxx = ""
                .properties.cWithACxx = ""

                .monthly_expenses.nElctrcBl = If(fsJSONValue = "", jsonObjDet.sElectric, loJSONObject.disbursement_info.monthly_expenses.nElctrcBl)
                .monthly_expenses.nWaterBil = If(fsJSONValue = "", jsonObjDet.sWaterBil, loJSONObject.disbursement_info.monthly_expenses.nWaterBil)
                .monthly_expenses.nFoodAllw = ""
                .monthly_expenses.nLoanAmtx = ""
                .bank_account.sBankName = If(fsJSONValue = "", jsonObjDet.sBankNme1, loJSONObject.disbursement_info.bank_account.sBankName)
                .bank_account.sAcctType = If(fsJSONValue = "", "", loJSONObject.disbursement_info.bank_account.sAcctType)
                .credit_card.sBankName = If(fsJSONValue = "", jsonObjDet.sCredtCrd, loJSONObject.disbursement_info.credit_card.sBankName)
                .credit_card.nCrdLimit = If(fsJSONValue = "", jsonObjDet.sCredtLmt, loJSONObject.disbursement_info.credit_card.nCrdLimit)
                .credit_card.nSinceYrx = ""
            End With
        End With
    End Sub

    'Private Sub populateJSONObjectOld(ByVal foJSONObject As GOCAS_Param, _
    '                                ByVal fsJSONValue As String)
    '    Dim loSettings As New JsonSerializerSettings

    '    loSettings.DefaultValueHandling = DefaultValueHandling.Populate

    '    Dim loJSONObject As GOCAS_Param = JsonConvert.DeserializeObject(Of GOCAS_Param)(fsJSONValue, loSettings)

    '    With foJSONObject
    '        .sBranchCd = If(fsJSONValue = "", jsonObjDet.sBranchCd, loJSONObject.sBranchCd)
    '        .nDownPaym = If(fsJSONValue = "", jsonObjDet.nDownPaym, loJSONObject.nDownPaym)
    '        .cApplType = 1
    '        .sModelIDx = If(fsJSONValue = "", jsonObjDet.MCModelx, loJSONObject.sModelIDx)
    '        .nAcctTerm = If(fsJSONValue = "", jsonObjDet.nLoanTerm, loJSONObject.nAcctTerm)

    '        With .applicant_info
    '            .sLastName = If(fsJSONValue = "", jsonObjDet.sLastName, loJSONObject.applicant_info.sLastName)
    '            .sFrstName = If(fsJSONValue = "", jsonObjDet.sFrstName, loJSONObject.applicant_info.sFrstName)
    '            .sSuffixNm = If(fsJSONValue = "", jsonObjDet.sSuffixNm, loJSONObject.applicant_info.sSuffixNm)
    '            .sMiddName = If(fsJSONValue = "", jsonObjDet.sMiddName, loJSONObject.applicant_info.sMiddName)
    '            .sNickName = If(fsJSONValue = "", jsonObjDet.sNickName, loJSONObject.applicant_info.sNickName)
    '            .dBirthDte = If(fsJSONValue = "", jsonObjDet.dBirthDte, loJSONObject.applicant_info.dBirthDte)
    '            .sBirthPlc = If(fsJSONValue = "", jsonObjDet.sBrtPlace, loJSONObject.applicant_info.sBirthPlc)

    '            If fsJSONValue = "" Then
    '                .mobile_number.Add(New GOCASConst.mobileno_param)
    '                .mobile_number(0).sMobileNo = jsonObjDet.sMobileNo
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.applicant_info.mobile_number.Count - 1
    '                    .mobile_number.Add(New GOCASConst.mobileno_param)
    '                    .mobile_number(nCtr).sMobileNo = loJSONObject.applicant_info.mobile_number(nCtr).sMobileNo
    '                Next
    '            End If

    '            .cCvilStat = If(fsJSONValue = "", "", loJSONObject.applicant_info.cCvilStat)
    '            .cGenderCd = If(fsJSONValue = "", "", loJSONObject.applicant_info.cGenderCd)
    '            .sMaidenNm = If(fsJSONValue = "", jsonObjDet.sMotherNm, loJSONObject.applicant_info.sMaidenNm)

    '            If fsJSONValue = "" Then
    '                .email_address.Add(New GOCASConst.email_param)
    '                .email_address(0).sEmailAdd = jsonObjDet.sEmailAdd
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.applicant_info.email_address.Count - 1
    '                    .email_address.Add(New GOCASConst.email_param)
    '                    .email_address(nCtr).sEmailAdd = loJSONObject.applicant_info.email_address(nCtr).sEmailAdd
    '                Next
    '            End If
    '        End With

    '        With .residence_info
    '            .cOwnershp = If(fsJSONValue = "", "", loJSONObject.residence_info.cOwnershp)
    '            .cOwnOther = If(fsJSONValue = "", "", loJSONObject.residence_info.cOwnOther)

    '            .rent_others.cRntOther = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.cRntOther)
    '            .rent_others.nLenStayx = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.nLenStayx)
    '            .rent_others.nRentExps = If(fsJSONValue = "", "", loJSONObject.residence_info.rent_others.nRentExps)

    '            .sCtkReltn = If(fsJSONValue = "", "", loJSONObject.residence_info.sCtkReltn)
    '            .cHouseTyp = If(fsJSONValue = "", "", loJSONObject.residence_info.cHouseTyp)
    '            .cGaragexx = If(fsJSONValue = "", "", loJSONObject.residence_info.cGaragexx)

    '            .present_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sLandMark)
    '            .present_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sHouseNox)
    '            .present_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sAddress1)
    '            .present_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sAddress2)
    '            .present_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sTownIDxx)
    '            .present_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.present_address.sBrgyIDxx)

    '            .permanent_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sLandMark)
    '            .permanent_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sHouseNox)
    '            .permanent_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sAddress1)
    '            .permanent_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sAddress2)
    '            .permanent_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sTownIDxx)
    '            .permanent_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.residence_info.permanent_address.sBrgyIDxx)
    '        End With

    '        With .means_info
    '            .cIncmeSrc = If(fsJSONValue = "", "", loJSONObject.means_info.cIncmeSrc)
    '            .employed.cEmpSectr = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cEmpSectr)
    '            .employed.sEmployer = If(fsJSONValue = "", jsonObjDet.sCompnyNm, loJSONObject.means_info.employed.sEmployer)
    '            .employed.sWrkTownx = If(fsJSONValue = "", jsonObjDet.sCompnyAd, loJSONObject.means_info.employed.sWrkTownx)
    '            .employed.sPosition = If(fsJSONValue = "", jsonObjDet.sEmpPostn, loJSONObject.means_info.employed.sPosition)
    '            .employed.nSalaryxx = If(fsJSONValue = "", jsonObjDet.sGrIncome, loJSONObject.means_info.employed.nSalaryxx)
    '            .employed.cEmpStatx = If(fsJSONValue = "", "", loJSONObject.means_info.employed.cEmpStatx)
    '            .employed.nLenServc = If(fsJSONValue = "", jsonObjDet.sLenServe, loJSONObject.means_info.employed.nLenServc)
    '            .employed.sWrkTelno = If(fsJSONValue = "", jsonObjDet.sCompTele, loJSONObject.means_info.employed.sWrkTelno)
    '            .self_employed.sBusiness = If(fsJSONValue = "", jsonObjDet.sBusiness, loJSONObject.means_info.self_employed.sBusiness)
    '            .self_employed.sBusAddrx = If(fsJSONValue = "", jsonObjDet.sBusiAddr, loJSONObject.means_info.self_employed.sBusAddrx)
    '            .self_employed.nBusLenxx = If(fsJSONValue = "", jsonObjDet.sYrInBusi, loJSONObject.means_info.self_employed.nBusLenxx)
    '            .self_employed.nBusIncom = If(fsJSONValue = "", jsonObjDet.sBusIncom, loJSONObject.means_info.self_employed.nBusIncom)
    '        End With

    '        With .other_info
    '            If fsJSONValue = "" Then
    '                .personal_reference.Add(New GOCASConst.personal_reference_param)
    '                .personal_reference(0).sRefrNmex = jsonObjDet.sRefName1
    '                .personal_reference(0).sRefrAddx = jsonObjDet.sRefAddr1

    '                .personal_reference.Add(New GOCASConst.personal_reference_param)
    '                .personal_reference(1).sRefrNmex = jsonObjDet.sRefName2
    '                .personal_reference(1).sRefrAddx = jsonObjDet.sRefAddr2
    '            Else
    '                For nCtr As Integer = 0 To .personal_reference.Count - 1
    '                    .personal_reference.Add(New GOCASConst.personal_reference_param)
    '                    .personal_reference(nCtr).sRefrNmex = loJSONObject.other_info.personal_reference(nCtr).sRefrNmex
    '                    .personal_reference(nCtr).sRefrAddx = loJSONObject.other_info.personal_reference(nCtr).sRefrAddx
    '                Next
    '            End If
    '        End With

    '        With .comaker_info
    '            .sLastName = If(fsJSONValue = "", jsonObjDet.sCOLastNm, loJSONObject.comaker_info.sLastName)
    '            .sFrstName = If(fsJSONValue = "", jsonObjDet.sCOFrstNm, loJSONObject.comaker_info.sFrstName)
    '            .sMiddName = If(fsJSONValue = "", jsonObjDet.sCOMiddNm, loJSONObject.comaker_info.sMiddName)

    '            If fsJSONValue = "" Then
    '                .mobile_number.Add(New GOCASConst.mobileno_param)
    '                .mobile_number(0).sMobileNo = jsonObjDet.sCOContct
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.comaker_info.mobile_number.Count - 1
    '                    .mobile_number.Add(New GOCASConst.mobileno_param)
    '                    .mobile_number(nCtr).sMobileNo = loJSONObject.comaker_info.mobile_number(nCtr).sMobileNo
    '                Next
    '            End If
    '        End With

    '        With .spouse_info
    '            .personal_info.sLastName = If(fsJSONValue = "", jsonObjDet.sSpLastNm, loJSONObject.spouse_info.personal_info.sLastName)
    '            .personal_info.sFrstName = If(fsJSONValue = "", jsonObjDet.sSpFrstNm, loJSONObject.spouse_info.personal_info.sFrstName)
    '            .personal_info.sSuffixNm = If(fsJSONValue = "", jsonObjDet.sSpSuffNm, loJSONObject.spouse_info.personal_info.sSuffixNm)
    '            .personal_info.sMiddName = If(fsJSONValue = "", jsonObjDet.sSpMiddNm, loJSONObject.spouse_info.personal_info.sMiddName)
    '            .personal_info.sNickName = If(fsJSONValue = "", jsonObjDet.sSpNickNm, loJSONObject.spouse_info.personal_info.sNickName)
    '            .personal_info.dBirthDte = If(fsJSONValue = "", jsonObjDet.dSpBrtDte, loJSONObject.spouse_info.personal_info.dBirthDte)
    '            .personal_info.sBirthPlc = If(fsJSONValue = "", jsonObjDet.sSpBrtPlc, loJSONObject.spouse_info.personal_info.sBirthPlc)

    '            If fsJSONValue = "" Then
    '                .personal_info.mobile_number.Add(New GOCASConst.mobileno_param)
    '                .personal_info.mobile_number(0).sMobileNo = jsonObjDet.sSpMobiNo
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.mobile_number.Count - 1
    '                    .personal_info.mobile_number.Add(New GOCASConst.mobileno_param)
    '                    .personal_info.mobile_number(nCtr).sMobileNo = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).sMobileNo
    '                Next
    '            End If

    '            If fsJSONValue = "" Then
    '                .personal_info.email_address.Add(New GOCASConst.email_param)
    '                .personal_info.email_address(0).sEmailAdd = jsonObjDet.sSpEmailx
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.email_address.Count - 1
    '                    .personal_info.email_address.Add(New GOCASConst.email_param)
    '                    .personal_info.email_address(nCtr).sEmailAdd = loJSONObject.spouse_info.personal_info.email_address(nCtr).sEmailAdd
    '                Next
    '            End If
    '        End With

    '        With .spouse_info.residence_info
    '            .cOwnershp = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cOwnershp)
    '            .cOwnOther = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cOwnOther)

    '            .rent_others.cRntOther = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.cRntOther)
    '            .rent_others.nLenStayx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.nLenStayx)
    '            .rent_others.nRentExps = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.rent_others.nRentExps)

    '            .sCtkReltn = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.sCtkReltn)
    '            .cHouseTyp = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cHouseTyp)
    '            .cGaragexx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.cGaragexx)

    '            .present_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sLandMark)
    '            .present_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sHouseNox)
    '            .present_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sAddress1)
    '            .present_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sAddress2)
    '            .present_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sTownIDxx)
    '            .present_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.present_address.sBrgyIDxx)

    '            .permanent_address.sLandMark = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sLandMark)
    '            .permanent_address.sHouseNox = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sHouseNox)
    '            .permanent_address.sAddress1 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sAddress1)
    '            .permanent_address.sAddress2 = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sAddress2)
    '            .permanent_address.sTownIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sTownIDxx)
    '            .permanent_address.sBrgyIDxx = If(fsJSONValue = "", "", loJSONObject.spouse_info.residence_info.permanent_address.sBrgyIDxx)
    '        End With

    '        With .spouse_means
    '            .cIncmeSrc = If(fsJSONValue = "", "", loJSONObject.spouse_means.cIncmeSrc)
    '            .employed.sEmployer = If(fsJSONValue = "", jsonObjDet.sSpCompNm, loJSONObject.spouse_means.employed.sEmployer)
    '            .employed.sWrkTownx = If(fsJSONValue = "", jsonObjDet.sSpCompAd, loJSONObject.spouse_means.employed.sWrkTownx)
    '            .employed.sPosition = If(fsJSONValue = "", jsonObjDet.sSpEmpPos, loJSONObject.spouse_means.employed.sPosition)
    '            .employed.cEmpStatx = If(fsJSONValue = "", "", loJSONObject.spouse_means.employed.cEmpStatx)
    '            .employed.nLenServc = If(fsJSONValue = "", jsonObjDet.sSpLenSrv, loJSONObject.spouse_means.employed.nLenServc)
    '            .employed.nSalaryxx = If(fsJSONValue = "", jsonObjDet.sSpMonPay, loJSONObject.spouse_means.employed.nSalaryxx)
    '            .employed.sWrkTelno = If(fsJSONValue = "", jsonObjDet.sSpComTel, loJSONObject.spouse_means.employed.sWrkTelno)

    '            .self_employed.sBusiness = If(fsJSONValue = "", jsonObjDet.sSpBusins, loJSONObject.spouse_means.self_employed.sBusiness)
    '            .self_employed.sBusAddrx = If(fsJSONValue = "", jsonObjDet.sSpBusiAd, loJSONObject.spouse_means.self_employed.sBusAddrx)
    '            .self_employed.nBusLenxx = If(fsJSONValue = "", jsonObjDet.sSpYrsBus, loJSONObject.spouse_means.self_employed.nBusLenxx)
    '            .self_employed.nBusIncom = If(fsJSONValue = "", jsonObjDet.sSpBusInc, loJSONObject.spouse_means.self_employed.nBusIncom)
    '        End With

    '        With .disbursement_info
    '            If fsJSONValue = "" Then
    '                .dependent_info.children.Add(New GOCASConst.children_param)
    '                .dependent_info.children(0).sFullName = jsonObjDet.sChldNme1
    '                .dependent_info.children(0).nDepdAgex = jsonObjDet.sChldAge1
    '                .dependent_info.children(0).sSchlName = jsonObjDet.sChldSch1

    '                .dependent_info.children.Add(New GOCASConst.children_param)
    '                .dependent_info.children(1).sFullName = jsonObjDet.sChldNme2
    '                .dependent_info.children(1).nDepdAgex = jsonObjDet.sChldAge2
    '                .dependent_info.children(1).sSchlName = jsonObjDet.sChldSch2

    '                .dependent_info.children.Add(New GOCASConst.children_param)
    '                .dependent_info.children(2).sFullName = jsonObjDet.sChldNme3
    '                .dependent_info.children(2).nDepdAgex = jsonObjDet.sChldAge3
    '                .dependent_info.children(2).sSchlName = jsonObjDet.sChldSch3
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.disbursement_info.dependent_info.children.Count - 1
    '                    .dependent_info.children.Add(New GOCASConst.children_param)
    '                    .dependent_info.children(nCtr).sFullName = loJSONObject.disbursement_info.dependent_info.children(nCtr).sFullName
    '                    .dependent_info.children(nCtr).nDepdAgex = loJSONObject.disbursement_info.dependent_info.children(nCtr).nDepdAgex
    '                    .dependent_info.children(nCtr).sSchlName = loJSONObject.disbursement_info.dependent_info.children(nCtr).sSchlName
    '                Next
    '            End If

    '            .monthly_expenses.nElctrcBl = If(fsJSONValue = "", jsonObjDet.sElectric, loJSONObject.disbursement_info.monthly_expenses.nElctrcBl)
    '            .monthly_expenses.nWaterBil = If(fsJSONValue = "", jsonObjDet.sWaterBil, loJSONObject.disbursement_info.monthly_expenses.nWaterBil)
    '            .bank_account.sBankName = If(fsJSONValue = "", jsonObjDet.sBankNme1, loJSONObject.disbursement_info.bank_account.sBankName)
    '            .bank_account.sAcctType = If(fsJSONValue = "", jsonObjDet.sBankNme1, loJSONObject.disbursement_info.bank_account.sBankName)
    '            .credit_card.sBankName = If(fsJSONValue = "", jsonObjDet.sCredtCrd, loJSONObject.disbursement_info.credit_card.sBankName)
    '            .credit_card.nCrdLimit = If(fsJSONValue = "", jsonObjDet.sCredtLmt, loJSONObject.disbursement_info.credit_card.nCrdLimit)
    '        End With
    '        With .comaker_info
    '            .sFrstName = If(fsJSONValue = "", jsonObjDet.sCLFrstNm, loJSONObject.comaker_info.sFrstName)
    '            .sMiddName = If(fsJSONValue = "", jsonObjDet.sCLMiddNm, loJSONObject.comaker_info.sMiddName)
    '            .sLastName = If(fsJSONValue = "", jsonObjDet.sCLLastNm, loJSONObject.comaker_info.sLastName)
    '            .sBirthPlc = If(fsJSONValue = "", jsonObjDet.sCLBrtPlc, loJSONObject.comaker_info.sBirthPlc)
    '            .sReltnCde = If(fsJSONValue = "", jsonObjDet.sCLRelatn, loJSONObject.comaker_info.sReltnCde)
    '            If fsJSONValue = "" Then
    '                .mobile_number.Add(New GOCASConst.mobileno_param)
    '                .mobile_number(0).sMobileNo = jsonObjDet.sSpMobiNo
    '            Else
    '                For nCtr As Integer = 0 To loJSONObject.comaker_info.mobile_number.Count - 1
    '                    .mobile_number.Add(New GOCASConst.mobileno_param)
    '                    .mobile_number(nCtr).sMobileNo = loJSONObject.comaker_info.mobile_number(nCtr).sMobileNo
    '                Next
    '            End If
    '            .dBirthDte = If(fsJSONValue = "", jsonObjDet.sCLBrtDte, loJSONObject.comaker_info.dBirthDte)
    '        End With
    '    End With
    'End Sub

    Class App_DetParam
        Property sLoanType As String
        Property nDownPaym As Decimal
        Property MCModelx As String
        Property sBranchCd As String
        Property nLoanTerm As Integer
        Property sFrstName As String
        Property sMiddName As String
        Property sLastName As String
        Property sSuffixNm As String
        Property sNickName As String
        Property sGenderxx As String
        Property sCvilStat As String
        Property sPresAddr As String
        Property sPrevAddr As String
        Property sLenStayx As String
        Property sMobileNo As String
        Property sEmailAdd As String
        Property dBirthDte As String
        Property sBrtPlace As String
        Property nAgexxxxx As String
        Property sMotherNm As String
        Property sFatherNm As String
        Property sParentAd As String

        Property sEmplType As String
        Property sCompnyNm As String
        Property sCompnyAd As String
        Property sCompTele As String
        Property sLenServe As String
        Property sGrIncome As String
        Property sEmplStat As String
        Property sEmpPostn As String
        Property sBusiness As String
        Property sBusiAddr As String
        Property sBusiTele As String
        Property sBusIncom As String
        Property sYrInBusi As String
        Property sSourceIn As String

        Property sBankNme1 As String
        Property sBankBrh1 As String
        Property sBankAcc1 As String
        Property sBankNme2 As String
        Property sBankBrh2 As String
        Property sBankAcc2 As String

        Property sRefName1 As String
        Property sRefAddr1 As String
        Property sRefName2 As String
        Property sRefAddr2 As String

        Property sSpFrstNm As String
        Property sSpMiddNm As String
        Property sSpLastNm As String
        Property sSpSuffNm As String
        Property sSpNickNm As String
        Property sSpPresAd As String
        Property sSpPrevAd As String
        Property sSpLenSty As String
        Property sSpMobiNo As String
        Property sSpEmailx As String
        Property dSpBrtDte As String
        Property sSpBrtPlc As String
        Property nSpAgexxx As String
        Property sSpCompNm As String
        Property sSpCompAd As String
        Property sSpComTel As String
        Property sSpLenSrv As String
        Property sSpMonPay As String
        Property sSpEmpSta As String
        Property sSpEmpPos As String
        Property sSpBusins As String
        Property sSpBusiAd As String
        Property sSpBusTel As String
        Property sSpBusInc As String
        Property sSpYrsBus As String
        Property sSpSrcInc As String

        Property sChldNme1 As String
        Property sChldAge1 As String
        Property sChldSch1 As String
        Property sChldNme2 As String
        Property sChldAge2 As String
        Property sChldSch2 As String
        Property sChldNme3 As String
        Property sChldAge3 As String
        Property sChldSch3 As String
        Property sRentalxx As String
        Property sElectric As String
        Property sWaterBil As String
        Property sOthrLoan As String
        Property sCredtCrd As String
        Property sCredtLmt As String
        Property sEducAttn As String
        Property sNumOfCPs As String
        Property sCPNumbr1 As String
        Property sCPTypex1 As String
        Property sCPNumbr2 As String
        Property sCPTypex2 As String
        Property sCPNumbr3 As String
        Property sCPTypex3 As String
        Property sLandmark As String
        Property sCOFrstNm As String
        Property sCOMiddNm As String
        Property sCOLastNm As String
        Property sCORelatn As String
        Property sCOOccptn As String
        Property sCONation As String
        Property sCORemitt As String
        Property sCOContct As String
        Property sCORoamNo As String
        Property sCOEmailx As String
        Property sCLFrstNm As String
        Property sCLMiddNm As String
        Property sCLLastNm As String
        Property sCLRelatn As String
        Property sCLAddres As String
        Property sCLEmploy As String
        Property sCLContct As String
        Property sCLBrtPlc As String
        Property sCLBrtDte As String
        Property sCLEmailx As String
    End Class

    Class GOCAS_Param
        Property sBranchCd As String
        Property dAppliedx As String
        Property sClientNm As String
        Property cUnitAppl As String
        Property sUnitAppl As String
        Property nDownPaym As String
        Property dCreatedx As String
        Property cApplType As String
        Property sModelIDx As String
        Property nAcctTerm As String
        Property nMonAmort As String
        Property dTargetDt As String
        Property applicant_info As New GOCASConst.applicant_param
        Property residence_info As New GOCASConst.client_param
        Property means_info As New GOCASConst.means_param
        Property other_info As New GOCASConst.other_param
        Property comaker_info As New GOCASConst.comaker_param
        Property spouse_info As New GOCASConst.spouse_param
        Property spouse_means As New GOCASConst.spouse__means_param
        Property disbursement_info As New GOCASConst.disbursement_param
    End Class

    Private Sub createReply(ByVal fsMessages As String, _
                                  ByVal fsMobileNo As String, _
                                  ByVal fsTransNox As String)
        Dim lsSQL As String

        lsSQL = "INSERT INTO HotLine_Outgoing SET" & _
                    "  sTransNox = " & strParm(GetNextCode("HotLine_Outgoing", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                    ", dTransact = " & dateParm(p_oApp.SysDate) & _
                    ", sDivision = " & strParm("MC") & _
                    ", sMobileNo = " & strParm(fsMobileNo) & _
                    ", sMessagex = " & strParm(fsMessages) & _
                    ", cSubscrbr = " & strParm(classifyMobileNo(fsMobileNo)) & _
                    ", dDueUntil = " & dateParm(DateAdd(DateInterval.Day, 10, p_oApp.SysDate)) & _
                    ", cSendStat = " & strParm("0") & _
                    ", nNoRetryx = " & strParm("0") & _
                    ", sUDHeader = " & strParm("") & _
                    ", sReferNox = " & strParm(fsTransNox) & _
                    ", sSourceCd = " & strParm("APP1") & _
                    ", cTranStat = " & strParm("0") & _
                    ", nPriority = 0" & _
                    ", sModified = " & strParm(p_oApp.UserID) & _
                    ", dModified = " & dateParm(p_oApp.SysDate)

        p_oApp.ExecuteActionQuery(lsSQL)
    End Sub

    Private Function classifyMobileNo(ByVal MobileNo As String) As Integer
        '0 = GLOBE
        '1 = SMART
        Select Case Left(MobileNo, 4)
            Case "0817", "0917", "0994", "0904", "0905", "0906", "0915", "0916", "0917", "0973"
                classifyMobileNo = 0
            Case "0925", "0926", "0927", "0935", "0978", "0979", "0936", "0996", "0997", "0999"
                classifyMobileNo = 0
            Case "0956", "0975", "0965", "0976", "0937", "0966", "0977", "0995", "0945", "0967"
                classifyMobileNo = 0
            Case Else
                classifyMobileNo = 1
        End Select
    End Function

    Function RemoveEmptyChildren(ByVal token As JToken) As JToken
        If token.Type = JTokenType.Object Then
            Dim copy As JObject = New JObject()

            For Each prop As JProperty In token.Children(Of JProperty)()
                Dim child As JToken = prop.Value
                If child.HasValues Then
                    child = RemoveEmptyChildren(child)
                End If

                If Not IsEmpty(child) Then
                    copy.Add(prop.Name, child)
                End If
            Next

            Return copy
        ElseIf token.Type = JTokenType.Array Then
            Dim copy As JArray = New JArray()

            For Each child As JToken In token.Children()
                If child.HasValues Then
                    child = RemoveEmptyChildren(child)
                End If

                If Not IsEmpty(child) Then
                    copy.Add(child)
                End If
            Next

            Return copy
        End If

        Return token
    End Function

    Function IsEmpty(ByVal token As JToken) As Boolean
        'Return (token.Type = JTokenType.Array And Not token.HasValues) Or
        '       (token.Type = JTokenType.Object And Not token.HasValues) Or
        '       (token.Type = JTokenType.String And token.ToString() = String.Empty) Or
        '       (token.Type = JTokenType.Null)

        Return (token.Type = JTokenType.Array And Not token.HasValues) Or
               (token.Type = JTokenType.Object And Not token.HasValues) Or
               (token.Type = JTokenType.Null)
    End Function
End Class