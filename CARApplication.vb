'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Car Application Object
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
'  jep [ 11/45/2020 11:45 ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports rmjGOCAS
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class CARApplication
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
    Private jsonObjDet As New GOCAS_Param

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

    Public ReadOnly Property Detail() As GOCAS_Param
        Get
            Return jsonObjDet
        End Get
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

    Public Property TranStatus() As String
        Get
            Return p_nTranStat
        End Get
        Set(ByVal value As String)
            p_nTranStat = value
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

        p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)
        p_oDTMstr(0).Item("dTransact") = p_oApp.SysDate
        p_oDTMstr(0).Item("sBranchCD") = ""
        p_oDTMstr(0).Item("sClientNm") = ""
        p_oDTMstr(0).Item("sSourceCD") = "CAR"
        p_oDTMstr(0).Item("sDetlInfo") = ""
        p_oDTMstr(0).Item("sQMatchNo") = ""
        p_oDTMstr(0).Item("nDownPaym") = 0
        p_oDTMstr(0).Item("cTranStat") = xeTranStat.TRANS_OPEN

        jsonObjDet = New GOCAS_Param

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Public Function UpdateTransaction() As Boolean
        If p_oDTMstr(0).Item("sTransNox") = "" Then
            MsgBox("Transaction is empty... Please verify your entry the try again!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Public Function SaveTransaction() As Boolean
        Dim lsSQL As String = ""
        Try

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then p_oDTMstr(0).Item("sTransNox") = GetNextCode(p_sMasTable, "sTransNox", True, p_oApp.Connection, True, p_sBranchCd)

            lsSQL = "INSERT INTO Credit_Online_Application SET" & _
                        "  sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                        ", sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd")) & _
                        ", dTransact = " & dateParm(p_oDTMstr(0).Item("dTransact")) & _
                        ", sClientNm = " & strParm(Detail.applicant_info.sLastName & ", " & Detail.applicant_info.sFrstName & " " & Detail.applicant_info.sMiddName) & _
                        ", sSourceCD = " & strParm(p_oDTMstr(0).Item("sSourceCd")) & _
                        ", sDetlInfo = " & "'" & (JSONObjDetail()) & "'" & _
                        ", nDownPaym = " & strParm(p_oDTMstr(0).Item("nDownPaym")) & _
                        ", sCreatedx = " & strParm(p_oApp.UserID) & _
                        ", dCreatedx = " & dateParm(p_oApp.getSysDate()) & _
                        ", cWithCIxx = " & strParm(xeTranStat.TRANS_CLOSED) & _
                        ", cTranStat = " & strParm(xeTranStat.TRANS_OPEN) & _
                        ", cDivision = " & strParm("2") & _
                        ", cEvaluatr = " & strParm("0") & _
                        ", dModified = " & dateParm(p_oApp.getSysDate()) & _
                    " ON DUPLICATE KEY UPDATE" & _
                        "  sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd")) & _
                        ", sDetlInfo = " & "'" & (JSONObjDetail()) & "'"
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

    'Public Function OpenTransaction(String)
    Public Function OpenTransaction(ByVal fsTransNox As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(fsTransNox))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        jsonObjDet = New GOCAS_Param
        Call populateJSONObject(jsonObjDet, IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), ""))
        Debug.Print(JsonConvert.SerializeObject(jsonObjDet))

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
            lsSQL = AddCondition(getSQ_Browse, "cTranStat IN (" & strDissect(p_nTranStat) & ")")
        Else
            lsSQL = AddCondition(getSQ_Browse, "NOT cTranStat IS NULL")
        End If
        If IsDBNull(p_oDTMstr(0).Item("cEvaluatr")) Then
            'lsSQL = AddCondition(lsSQL, IsDBNull("cEvaluatr"))
        Else
            lsSQL = AddCondition(lsSQL, "cEvaluatr = " & strParm(IIf(fbEvaluate, 1, 0)))
        End If


        'If p_sBranchCd <> "" Then
        '    lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
        'End If

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "sTransNox LIKE " & strParm("%" & fsValue)
        Else
            lsFilter = "sClientNm LIKE " & strParm(fsValue & "%")
        End If
        Debug.Print(lsSQL)
        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , "Trans No»Client»Date", _
                                        , "sTransNox»sClientNm»dTransact" _
                                        , IIf(fbByCode, 1, 2))
        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            Return False
        Else
            If IsDBNull(loDta.Item("sLoadedBy")) Then GoTo moveNext
            If loDta.Item("sLoadedBy") <> p_oApp.UserID Then
                Dim lnRep As Integer

                lnRep = MsgBox("Transaction was already loaded by other evaluator..." & vbCrLf & _
                                "Do you want to open transaction", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "CONFIRMATION")

                If lnRep = vbNo Then Return False
            End If
MoveNext:

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

    Function TruncateDecimal(ByVal value As Decimal, ByVal precision As Integer) As Decimal
        Dim stepper As Decimal = Math.Pow(10, precision)
        Dim tmp As Decimal = Math.Truncate(stepper * value)
        Return tmp / stepper
    End Function

    Public Function getNextCustomer() As String
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = AddCondition(getSQ_Master, " sCatInfox IS NULL" & _
                                            " AND sLoadedBy IS NULL" & _
                                            " AND sSourceCd = 'CAR'" & _
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
        Return "SELECT sTransNox" & _
                    ", sClientNm" & _
                    ", dTransact" & _
                    ", sLoadedBy" & _
                  " FROM " & p_sMasTable & _
                  " WHERE sSourceCD = 'CAR'"
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
                    "  sDetlInfo = " & "'" & (JSONObjDetail()) & "'" & _
                    ", cTranStat = " & strParm(1) & _
                    ", sConfirmd = " & strParm(p_oApp.UserID) & _
                    ", dConfirmd = " & dateParm(p_oApp.getSysDate) & _
                " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox"))

        If p_oApp.Execute(lsSQL, "Credit_Online_Application", p_sBranchCd) = 0 Then
            MsgBox("Unable to Confirm Transaction!!!", vbCritical, "Warning")
            Return False
        End If

        Return OpenTransaction(p_oDTMstr(0)("sTransNox"))
    End Function

    Function evaluateTransaction() As Boolean
        Dim lsSQL As String

        lsSQL = "UPDATE Credit_Online_Application SET" & _
                    "  sDetlInfo = " & "'" & (JSONObjDetail()) & "'" & _
                    ", cTranStat = " & strParm(2) & _
                    ", sVerified = " & strParm(p_oApp.UserID) & _
                    ", dVerified = " & dateParm(p_oApp.getSysDate) & _
                " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox"))

        If p_oApp.Execute(lsSQL, "Credit_Online_Application", p_sBranchCd) = 0 Then
            MsgBox("Unable to Confirm Transaction!!!", vbCritical, "Warning")
            Return False
        End If

        Return OpenTransaction(p_oDTMstr(0)("sTransNox"))
    End Function

    Function GenerateQM() As String
        Dim loQMResult As QMResult
        Dim loFrm As frmQuickMatch

        loQMResult = New QMResult
        loFrm = New frmQuickMatch

        With loQMResult
            .AppDriver = p_oApp
            .Branch = p_sBranchCd
            .ApplicationNo = IIf(p_oApp.ProductID <> "LRTrackr", p_oDTMstr.Rows(0)("sTransNox"), "")

            .InitTransaction()
            'Set the Applicant info
            .Applicant("sClientID") = ""

            .Applicant("sLastName") = Detail.applicant_info.sLastName & IIf(IFNull(Detail.applicant_info.sSuffixNm) = "", "", " " & Detail.applicant_info.sSuffixNm)
            .Applicant("sFrstName") = Detail.applicant_info.sFrstName
            .Applicant("sMiddName") = Detail.applicant_info.sMiddName
            .Applicant("dBirthDte") = Detail.applicant_info.dBirthDte
            .Applicant("sBirthPlc") = Detail.applicant_info.sBirthPlc
            .Applicant("sTownIDxx") = Detail.residence_info.present_address.sTownIDxx

            'Set the spouse info
            If Not IsNothing(Detail.spouse_info.personal_info.sLastName) Then
                If IFNull(Detail.spouse_info.personal_info.sLastName) <> "" Then
                    .Spouse("sClientID") = ""

                    .Spouse("sLastName") = Detail.spouse_info.personal_info.sLastName
                    .Spouse("sFrstName") = Detail.spouse_info.personal_info.sFrstName & IIf(IFNull(Detail.spouse_info.personal_info.sSuffixNm) = "", "", " " & Detail.spouse_info.personal_info.sSuffixNm)
                    .Spouse("sMiddName") = Detail.spouse_info.personal_info.sMiddName
                    .Spouse("dBirthDte") = Detail.spouse_info.personal_info.dBirthDte
                    .Spouse("sBirthPlc") = Detail.spouse_info.personal_info.sBirthPlc
                    .Spouse("sTownIDxx") = Detail.spouse_info.residence_info.present_address.sTownIDxx
                End If
            End If

            .Term("sModelIDx") = Detail.sModelIDx
            .Term("nDownPaym") = Detail.nDownPaym
            .Term("nAcctTerm") = Detail.nAcctTerm

            'Execute quickmatch here
            GenerateQM = .QuickMatch

            If GenerateQM = "" Then
                Exit Function
            End If

            loFrm = New frmQuickMatch
            loFrm.Appdriver = p_oApp

            loFrm.txtField00.Text = .TransNo
            loFrm.txtField04.Text = Detail.applicant_info.sLastName & _
                         ", " & Detail.applicant_info.sFrstName & IIf(IFNull(Detail.applicant_info.sSuffixNm) = "", "", " " & Detail.applicant_info.sSuffixNm) & _
                          " " & Detail.applicant_info.sMiddName
            loFrm.txtField20.Text = Detail.residence_info.present_address.sAddress1 & _
                         ", " & getTownCity(Detail.residence_info.present_address.sTownIDxx, True, True, "")

            If Not IsNothing(Detail.spouse_info.personal_info.sLastName) Then
                'Display spouse info
                If IFNull(Detail.spouse_info.personal_info.sLastName) = "" Then
                    loFrm.txtField06.Text = "N-O-N-E"
                    loFrm.txtField07.Text = "N-O-N-E"
                Else
                    loFrm.txtField06.Text = Detail.spouse_info.personal_info.sLastName & _
                                 ", " & Detail.spouse_info.personal_info.sFrstName & IIf(IFNull(Detail.spouse_info.personal_info.sSuffixNm) = "", "", " " & Detail.spouse_info.personal_info.sSuffixNm) & _
                                  " " & Detail.spouse_info.personal_info.sMiddName
                    loFrm.txtField07.Text = Detail.spouse_info.residence_info.present_address.sAddress1 & _
                                 ", " & getTownCity(Detail.spouse_info.residence_info.present_address.sTownIDxx, True, True, "")
                End If
            End If

            p_oDTMstr.Rows(0)("sQMatchNo") = GenerateQM
            loFrm.txtField08.Text = p_oDTMstr.Rows(0)("sQMatchNo")
            loFrm.txtField09.Text = p_oDTMstr.Rows(0)("sTransNox")
            loFrm.txtField05.Text = Format(Detail.dAppliedx, "Mmmm DD, YYYY")

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

    Function JSONObjDetail() As String
        Dim loSettings As New JsonSerializerSettings

        loSettings.NullValueHandling = NullValueHandling.Ignore
        loSettings.DefaultValueHandling = DefaultValueHandling.Ignore

        Return JsonConvert.SerializeObject(jsonObjDet, loSettings)
    End Function

    Private Sub populateJSONObject(ByVal foJSONObject As GOCAS_Param, _
                                    ByVal fsJSONValue As String)
        Dim loJSONObject As New GOCAS_Param
        Dim loSettings As New JsonSerializerSettings

        loSettings.DefaultValueHandling = DefaultValueHandling.Populate
        Debug.Print(fsJSONValue)
        loJSONObject = JsonConvert.DeserializeObject(Of GOCAS_Param)(fsJSONValue, loSettings)


        If fsJSONValue = "" Then Exit Sub

        With foJSONObject
            .sBranchCd = loJSONObject.sBranchCd
            .dAppliedx = loJSONObject.dAppliedx
            'added by jovan 11-19-2020
            .nRebatesx = loJSONObject.nRebatesx
            .nPNValuex = loJSONObject.nPNValuex
            .sClientNm = loJSONObject.sClientNm
            .cUnitAppl = loJSONObject.cUnitAppl
            .nDownPaym = loJSONObject.nDownPaym
            .dCreatedx = loJSONObject.dCreatedx
            .cApplType = loJSONObject.cApplType
            .sUnitAppl = loJSONObject.sUnitAppl
            .sModelIDx = loJSONObject.sModelIDx
            .nAcctTerm = loJSONObject.nAcctTerm
            .nMonAmort = loJSONObject.nMonAmort
            .dTargetDt = loJSONObject.dTargetDt

            With .applicant_info
                .sLastName = loJSONObject.applicant_info.sLastName
                .sFrstName = loJSONObject.applicant_info.sFrstName
                .sSuffixNm = loJSONObject.applicant_info.sSuffixNm
                .sMiddName = loJSONObject.applicant_info.sMiddName
                .sNickName = loJSONObject.applicant_info.sNickName
                .dBirthDte = loJSONObject.applicant_info.dBirthDte
                .sBirthPlc = loJSONObject.applicant_info.sBirthPlc
                .sCitizenx = loJSONObject.applicant_info.sCitizenx

                If Not IsNothing(loJSONObject.applicant_info.mobile_number) Then
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.mobile_number.Count - 1
                        .mobile_number.Add(New CARConst.mobileno_param)
                        .mobile_number(nCtr).sMobileNo = loJSONObject.applicant_info.mobile_number(nCtr).sMobileNo
                        .mobile_number(nCtr).cPostPaid = loJSONObject.applicant_info.mobile_number(nCtr).cPostPaid
                        .mobile_number(nCtr).nPostYear = loJSONObject.applicant_info.mobile_number(nCtr).nPostYear
                    Next
                End If
                If Not IsNothing(loJSONObject.applicant_info.landline.Count) Then
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.landline.Count - 1
                        .landline.Add(New CARConst.landline_param)
                        .landline(nCtr).sPhoneNox = loJSONObject.applicant_info.landline(nCtr).sPhoneNox
                    Next
                End If

                .cCvilStat = loJSONObject.applicant_info.cCvilStat
                .cGenderCd = loJSONObject.applicant_info.cGenderCd
                .sMaidenNm = loJSONObject.applicant_info.sMaidenNm

                If Not IsNothing(loJSONObject.applicant_info.email_address.Count) Then
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.email_address.Count - 1
                        .email_address.Add(New CARConst.email_param)
                        .email_address(nCtr).sEmailAdd = loJSONObject.applicant_info.email_address(nCtr).sEmailAdd
                    Next
                End If

                .facebook.sFBAcctxx = loJSONObject.applicant_info.facebook.sFBAcctxx
                .facebook.cAcctStat = loJSONObject.applicant_info.facebook.cAcctStat
                .facebook.nNoFriend = loJSONObject.applicant_info.facebook.nNoFriend
                .facebook.nYearxxxx = loJSONObject.applicant_info.facebook.nYearxxxx
                .sVibeAcct = loJSONObject.applicant_info.sVibeAcct
            End With

            If Not IsNothing(loJSONObject.residence_info) Then
                With .residence_info
                    .cOwnershp = loJSONObject.residence_info.cOwnershp
                    .cOwnOther = loJSONObject.residence_info.cOwnOther

                    If Not IsNothing(loJSONObject.residence_info.rent_others) Then
                        .rent_others.cRntOther = loJSONObject.residence_info.rent_others.cRntOther
                        .rent_others.nLenStayx = loJSONObject.residence_info.rent_others.nLenStayx
                        .rent_others.nRentExps = loJSONObject.residence_info.rent_others.nRentExps
                    End If

                    .sCtkReltn = loJSONObject.residence_info.sCtkReltn
                    .cHouseTyp = loJSONObject.residence_info.cHouseTyp
                    .cGaragexx = loJSONObject.residence_info.cGaragexx

                    .present_address.sLandMark = loJSONObject.residence_info.present_address.sLandMark
                    .present_address.sHouseNox = loJSONObject.residence_info.present_address.sHouseNox
                    .present_address.sAddress1 = loJSONObject.residence_info.present_address.sAddress1
                    .present_address.sAddress2 = loJSONObject.residence_info.present_address.sAddress2
                    .present_address.sTownIDxx = loJSONObject.residence_info.present_address.sTownIDxx
                    .present_address.sBrgyIDxx = loJSONObject.residence_info.present_address.sBrgyIDxx

                    .permanent_address.sLandMark = loJSONObject.residence_info.permanent_address.sLandMark
                    .permanent_address.sHouseNox = loJSONObject.residence_info.permanent_address.sHouseNox
                    .permanent_address.sAddress1 = loJSONObject.residence_info.permanent_address.sAddress1
                    .permanent_address.sAddress2 = loJSONObject.residence_info.permanent_address.sAddress2
                    .permanent_address.sTownIDxx = loJSONObject.residence_info.permanent_address.sTownIDxx
                    .permanent_address.sBrgyIDxx = loJSONObject.residence_info.permanent_address.sBrgyIDxx
                End With
            End If

            If Not IsNothing(loJSONObject.means_info) Then
                With .means_info
                    .cIncmeSrc = loJSONObject.means_info.cIncmeSrc
                    .employed.cEmpSectr = loJSONObject.means_info.employed.cEmpSectr
                    .employed.cUniforme = loJSONObject.means_info.employed.cUniforme
                    .employed.cMilitary = loJSONObject.means_info.employed.cMilitary
                    .employed.cGovtLevl = loJSONObject.means_info.employed.cGovtLevl
                    .employed.cCompLevl = loJSONObject.means_info.employed.cCompLevl
                    .employed.cEmpLevlx = loJSONObject.means_info.employed.cEmpLevlx
                    .employed.cOcCatgry = loJSONObject.means_info.employed.cOcCatgry
                    .employed.cOFWRegnx = loJSONObject.means_info.employed.cOFWRegnx
                    .employed.sOFWNatnx = loJSONObject.means_info.employed.sOFWNatnx
                    .employed.sIndstWrk = loJSONObject.means_info.employed.sIndstWrk
                    .employed.sEmployer = loJSONObject.means_info.employed.sEmployer
                    .employed.sWrkAddrx = loJSONObject.means_info.employed.sWrkAddrx
                    .employed.sWrkTownx = loJSONObject.means_info.employed.sWrkTownx
                    .employed.sPosition = loJSONObject.means_info.employed.sPosition
                    .employed.sFunction = loJSONObject.means_info.employed.sFunction
                    .employed.cEmpStatx = loJSONObject.means_info.employed.cEmpStatx
                    .employed.nLenServc = loJSONObject.means_info.employed.nLenServc
                    .employed.nSalaryxx = loJSONObject.means_info.employed.nSalaryxx
                    .employed.sWrkTelno = loJSONObject.means_info.employed.sWrkTelno
                    .self_employed.sIndstBus = loJSONObject.means_info.self_employed.sIndstBus
                    .self_employed.sBusiness = loJSONObject.means_info.self_employed.sBusiness
                    .self_employed.sBusAddrx = loJSONObject.means_info.self_employed.sBusAddrx
                    .self_employed.sBusTownx = loJSONObject.means_info.self_employed.sBusTownx
                    .self_employed.cBusTypex = loJSONObject.means_info.self_employed.cBusTypex
                    .self_employed.nBusLenxx = loJSONObject.means_info.self_employed.nBusLenxx
                    .self_employed.nBusIncom = loJSONObject.means_info.self_employed.nBusIncom
                    .self_employed.nMonExpns = loJSONObject.means_info.self_employed.nMonExpns
                    .self_employed.cOwnTypex = loJSONObject.means_info.self_employed.cOwnTypex
                    .self_employed.cOwnSizex = loJSONObject.means_info.self_employed.cOwnSizex
                    .pensioner.cPenTypex = loJSONObject.means_info.pensioner.cPenTypex
                    .pensioner.nPensionx = loJSONObject.means_info.pensioner.nPensionx
                    .pensioner.nRetrYear = loJSONObject.means_info.pensioner.nRetrYear
                    .financed.sReltnCde = loJSONObject.means_info.financed.sReltnCde
                    .financed.sFinancer = loJSONObject.means_info.financed.sFinancer
                    .financed.nEstIncme = loJSONObject.means_info.financed.nEstIncme
                    .financed.sNatnCode = loJSONObject.means_info.financed.sNatnCode
                    .financed.sMobileNo = loJSONObject.means_info.financed.sMobileNo
                    .financed.sFBAcctxx = loJSONObject.means_info.financed.sFBAcctxx
                    .financed.sEmailAdd = loJSONObject.means_info.financed.sEmailAdd
                    If Not IsNothing(loJSONObject.means_info.other_income) Then
                        .other_income.sOthrIncm = loJSONObject.means_info.other_income.sOthrIncm
                        .other_income.nOthrIncm = loJSONObject.means_info.other_income.nOthrIncm
                    End If
                End With
            End If
            If Not IsNothing(loJSONObject.other_info) Then
                With .other_info
                    .sUnitUser = loJSONObject.other_info.sUnitUser
                    .sUsr2Buyr = loJSONObject.other_info.sUsr2Buyr
                    .sPurposex = loJSONObject.other_info.sPurposex
                    .sUnitPayr = loJSONObject.other_info.sUnitPayr
                    .sPyr2Buyr = loJSONObject.other_info.sPyr2Buyr
                    .sSrceInfo = loJSONObject.other_info.sSrceInfo
                    For nCtr As Integer = 0 To loJSONObject.other_info.personal_reference.Count - 1
                        .personal_reference.Add(New CARConst.personal_reference_param)
                        .personal_reference(nCtr).sRefrNmex = loJSONObject.other_info.personal_reference(nCtr).sRefrNmex
                        .personal_reference(nCtr).sRefrMPNx = loJSONObject.other_info.personal_reference(nCtr).sRefrMPNx
                        .personal_reference(nCtr).sRefrAddx = loJSONObject.other_info.personal_reference(nCtr).sRefrAddx
                        .personal_reference(nCtr).sRefrTown = loJSONObject.other_info.personal_reference(nCtr).sRefrTown
                    Next
                End With
            End If

            If Not IsNothing(loJSONObject.comaker_info) Then
                With .comaker_info
                    .sLastName = loJSONObject.comaker_info.sLastName
                    .sFrstName = loJSONObject.comaker_info.sFrstName
                    .sSuffixNm = loJSONObject.comaker_info.sSuffixNm
                    .sMiddName = loJSONObject.comaker_info.sMiddName
                    .sNickName = loJSONObject.comaker_info.sNickName
                    .dBirthDte = loJSONObject.comaker_info.dBirthDte
                    .sBirthPlc = loJSONObject.comaker_info.sBirthPlc
                    .cIncmeSrc = loJSONObject.comaker_info.cIncmeSrc
                    .sReltnCde = loJSONObject.comaker_info.sReltnCde
                    If Not IsNothing(loJSONObject.comaker_info.mobile_number) Then
                        For nCtr As Integer = 0 To loJSONObject.comaker_info.mobile_number.Count - 1
                            .mobile_number.Add(New CARConst.mobileno_param)
                            .mobile_number(nCtr).sMobileNo = loJSONObject.comaker_info.mobile_number(nCtr).sMobileNo
                            .mobile_number(nCtr).cPostPaid = loJSONObject.comaker_info.mobile_number(nCtr).cPostPaid
                            .mobile_number(nCtr).nPostYear = loJSONObject.comaker_info.mobile_number(nCtr).nPostYear
                        Next
                    End If
                    .sFBAcctxx = loJSONObject.comaker_info.sFBAcctxx
                End With
            End If

            If Not IsNothing(loJSONObject.spouse_info) Then
                With .spouse_info
                    .personal_info.sLastName = loJSONObject.spouse_info.personal_info.sLastName
                    .personal_info.sFrstName = loJSONObject.spouse_info.personal_info.sFrstName
                    .personal_info.sSuffixNm = loJSONObject.spouse_info.personal_info.sSuffixNm
                    .personal_info.sMiddName = loJSONObject.spouse_info.personal_info.sMiddName
                    .personal_info.sNickName = loJSONObject.spouse_info.personal_info.sNickName
                    .personal_info.dBirthDte = loJSONObject.spouse_info.personal_info.dBirthDte
                    .personal_info.sBirthPlc = loJSONObject.spouse_info.personal_info.sBirthPlc
                    .personal_info.sCitizenx = loJSONObject.spouse_info.personal_info.sCitizenx

                    For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.mobile_number.Count - 1
                        .personal_info.mobile_number.Add(New CARConst.mobileno_param)
                        .personal_info.mobile_number(nCtr).sMobileNo = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).sMobileNo
                        .personal_info.mobile_number(nCtr).cPostPaid = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).cPostPaid
                        .personal_info.mobile_number(nCtr).nPostYear = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).nPostYear
                    Next
                    If Not IsNothing(loJSONObject.spouse_info.personal_info.landline) Then
                        For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.landline.Count - 1
                            .personal_info.landline.Add(New CARConst.landline_param)
                            .personal_info.landline(nCtr).sPhoneNox = loJSONObject.spouse_info.personal_info.landline(nCtr).sPhoneNox
                        Next
                    End If

                    .personal_info.cCvilStat = loJSONObject.spouse_info.personal_info.cCvilStat
                    .personal_info.cGenderCd = loJSONObject.spouse_info.personal_info.cGenderCd
                    .personal_info.sMaidenNm = loJSONObject.spouse_info.personal_info.sMaidenNm
                    If Not IsNothing(loJSONObject.spouse_info.personal_info.email_address) Then
                        For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.email_address.Count - 1
                            .personal_info.email_address.Add(New CARConst.email_param)
                            .personal_info.email_address(nCtr).sEmailAdd = loJSONObject.spouse_info.personal_info.email_address(nCtr).sEmailAdd
                        Next
                    End If
                    If Not IsNothing(loJSONObject.spouse_info.personal_info.facebook) Then
                        .personal_info.facebook.sFBAcctxx = loJSONObject.spouse_info.personal_info.facebook.sFBAcctxx
                        .personal_info.facebook.cAcctStat = loJSONObject.spouse_info.personal_info.facebook.cAcctStat
                        .personal_info.facebook.nNoFriend = loJSONObject.spouse_info.personal_info.facebook.nNoFriend
                        .personal_info.facebook.nYearxxxx = loJSONObject.spouse_info.personal_info.facebook.nYearxxxx
                    End If
                    .personal_info.sVibeAcct = loJSONObject.spouse_info.personal_info.sVibeAcct
                    If Not IsNothing(loJSONObject.spouse_info.residence_info) Then
                        .residence_info.cOwnershp = loJSONObject.spouse_info.residence_info.cOwnershp
                        .residence_info.cOwnOther = loJSONObject.spouse_info.residence_info.cOwnOther
                        If Not IsNothing(loJSONObject.spouse_info.residence_info.rent_others) Then
                            .residence_info.rent_others.cRntOther = loJSONObject.spouse_info.residence_info.rent_others.cRntOther
                            .residence_info.rent_others.nLenStayx = loJSONObject.spouse_info.residence_info.rent_others.nLenStayx
                            .residence_info.rent_others.nRentExps = loJSONObject.spouse_info.residence_info.rent_others.nRentExps
                        End If
                        .residence_info.sCtkReltn = loJSONObject.spouse_info.residence_info.sCtkReltn
                        .residence_info.cHouseTyp = loJSONObject.spouse_info.residence_info.cHouseTyp
                        .residence_info.cGaragexx = loJSONObject.spouse_info.residence_info.cGaragexx
                        If Not IsNothing(loJSONObject.spouse_info.residence_info.present_address) Then
                            .residence_info.present_address.sLandMark = loJSONObject.spouse_info.residence_info.present_address.sLandMark
                            .residence_info.present_address.sHouseNox = loJSONObject.spouse_info.residence_info.present_address.sHouseNox
                            .residence_info.present_address.sAddress1 = loJSONObject.spouse_info.residence_info.present_address.sAddress1
                            .residence_info.present_address.sAddress2 = loJSONObject.spouse_info.residence_info.present_address.sAddress2
                            .residence_info.present_address.sTownIDxx = loJSONObject.spouse_info.residence_info.present_address.sTownIDxx
                            .residence_info.present_address.sBrgyIDxx = loJSONObject.spouse_info.residence_info.present_address.sBrgyIDxx
                        End If
                        If Not IsNothing(loJSONObject.spouse_info.residence_info.permanent_address) Then
                            .residence_info.permanent_address.sLandMark = loJSONObject.spouse_info.residence_info.permanent_address.sLandMark
                            .residence_info.permanent_address.sHouseNox = loJSONObject.spouse_info.residence_info.permanent_address.sHouseNox
                            .residence_info.permanent_address.sAddress1 = loJSONObject.spouse_info.residence_info.permanent_address.sAddress1
                            .residence_info.permanent_address.sAddress2 = loJSONObject.spouse_info.residence_info.permanent_address.sAddress2
                            .residence_info.permanent_address.sTownIDxx = loJSONObject.spouse_info.residence_info.permanent_address.sTownIDxx
                            .residence_info.permanent_address.sBrgyIDxx = loJSONObject.spouse_info.residence_info.permanent_address.sBrgyIDxx
                        End If
                    End If
                End With
                If Not IsNothing(loJSONObject.spouse_means) Then
                    With .spouse_means
                        .cIncmeSrc = loJSONObject.spouse_means.cIncmeSrc
                        If Not IsNothing(loJSONObject.spouse_means.employed) Then
                            .employed.cEmpSectr = loJSONObject.spouse_means.employed.cEmpSectr
                            .employed.cUniforme = loJSONObject.spouse_means.employed.cUniforme
                            .employed.cMilitary = loJSONObject.spouse_means.employed.cMilitary
                            .employed.cGovtLevl = loJSONObject.spouse_means.employed.cGovtLevl
                            .employed.cCompLevl = loJSONObject.spouse_means.employed.cCompLevl
                            .employed.cEmpLevlx = loJSONObject.spouse_means.employed.cEmpLevlx
                            .employed.cOcCatgry = loJSONObject.spouse_means.employed.cOcCatgry
                            .employed.cOFWRegnx = loJSONObject.spouse_means.employed.cOFWRegnx
                            .employed.sOFWNatnx = loJSONObject.spouse_means.employed.sOFWNatnx
                            .employed.sIndstWrk = loJSONObject.spouse_means.employed.sIndstWrk
                            .employed.sEmployer = loJSONObject.spouse_means.employed.sEmployer
                            .employed.sWrkAddrx = loJSONObject.spouse_means.employed.sWrkAddrx
                            .employed.sWrkTownx = loJSONObject.spouse_means.employed.sWrkTownx
                            .employed.sPosition = loJSONObject.spouse_means.employed.sPosition
                            .employed.sFunction = loJSONObject.spouse_means.employed.sFunction
                            .employed.cEmpStatx = loJSONObject.spouse_means.employed.cEmpStatx
                            .employed.nLenServc = loJSONObject.spouse_means.employed.nLenServc
                            .employed.nSalaryxx = loJSONObject.spouse_means.employed.nSalaryxx
                            .employed.sWrkTelno = loJSONObject.spouse_means.employed.sWrkTelno
                        End If
                        If Not IsNothing(loJSONObject.spouse_means.self_employed) Then
                            .self_employed.sIndstBus = loJSONObject.spouse_means.self_employed.sIndstBus
                            .self_employed.sBusiness = loJSONObject.spouse_means.self_employed.sBusiness
                            .self_employed.sBusAddrx = loJSONObject.spouse_means.self_employed.sBusAddrx
                            .self_employed.sBusTownx = loJSONObject.spouse_means.self_employed.sBusTownx
                            .self_employed.cBusTypex = loJSONObject.spouse_means.self_employed.cBusTypex
                            .self_employed.nBusLenxx = loJSONObject.spouse_means.self_employed.nBusLenxx
                            .self_employed.nBusIncom = loJSONObject.spouse_means.self_employed.nBusIncom
                            .self_employed.nMonExpns = loJSONObject.spouse_means.self_employed.nMonExpns
                            .self_employed.cOwnTypex = loJSONObject.spouse_means.self_employed.cOwnTypex
                            .self_employed.cOwnSizex = loJSONObject.spouse_means.self_employed.cOwnSizex
                        End If
                        If Not IsNothing(loJSONObject.spouse_means.pensioner) Then
                            .pensioner.cPenTypex = loJSONObject.spouse_means.pensioner.cPenTypex
                            .pensioner.nPensionx = loJSONObject.spouse_means.pensioner.nPensionx
                            .pensioner.nRetrYear = loJSONObject.spouse_means.pensioner.nRetrYear
                        End If
                        If Not IsNothing(loJSONObject.spouse_means.financed) Then
                            .financed.sReltnCde = loJSONObject.spouse_means.financed.sReltnCde
                            .financed.sFinancer = loJSONObject.spouse_means.financed.sFinancer
                            .financed.nEstIncme = loJSONObject.spouse_means.financed.nEstIncme
                            .financed.sNatnCode = loJSONObject.spouse_means.financed.sNatnCode
                            .financed.sMobileNo = loJSONObject.spouse_means.financed.sMobileNo
                            .financed.sFBAcctxx = loJSONObject.spouse_means.financed.sFBAcctxx
                            .financed.sEmailAdd = loJSONObject.spouse_means.financed.sEmailAdd
                        End If
                        If Not IsNothing(loJSONObject.spouse_means.other_income) Then
                            .other_income.sOthrIncm = loJSONObject.spouse_means.other_income.sOthrIncm
                            .other_income.nOthrIncm = loJSONObject.spouse_means.other_income.nOthrIncm
                        End If
                    End With
                End If
            End If

            If Not IsNothing(loJSONObject.disbursement_info) Then
                With .disbursement_info
                    .dependent_info.nHouseHld = loJSONObject.disbursement_info.dependent_info.nHouseHld
                    If Not IsNothing(loJSONObject.disbursement_info.dependent_info.children) Then
                        For nCtr As Integer = 0 To loJSONObject.disbursement_info.dependent_info.children.Count - 1
                            .dependent_info.children.Add(New CARConst.children_param)
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

                    If Not IsNothing(loJSONObject.disbursement_info.properties) Then
                        If Not IsNothing(loJSONObject.disbursement_info.properties.with4Whls_info) Then
                            .properties.with4Whls_info.cWithWhls = loJSONObject.disbursement_info.properties.with4Whls_info.cWithWhls
                            .properties.with4Whls_info.cOwnerShp = loJSONObject.disbursement_info.properties.with4Whls_info.cOwnerShp
                            .properties.with4Whls_info.cStatusxx = loJSONObject.disbursement_info.properties.with4Whls_info.cStatusxx
                            .properties.with4Whls_info.cTermxxxx = loJSONObject.disbursement_info.properties.with4Whls_info.cTermxxxx
                        End If

                        If Not IsNothing(loJSONObject.disbursement_info.properties.with3Whls_info) Then
                            .properties.with3Whls_info.cWithWhls = loJSONObject.disbursement_info.properties.with3Whls_info.cWithWhls
                            .properties.with3Whls_info.cOwnerShp = loJSONObject.disbursement_info.properties.with3Whls_info.cOwnerShp
                            .properties.with3Whls_info.cStatusxx = loJSONObject.disbursement_info.properties.with3Whls_info.cStatusxx
                            .properties.with3Whls_info.cTermxxxx = loJSONObject.disbursement_info.properties.with3Whls_info.cTermxxxx
                        End If

                        If Not IsNothing(loJSONObject.disbursement_info.properties.with2Whls_info) Then
                            .properties.with2Whls_info.cWithWhls = loJSONObject.disbursement_info.properties.with2Whls_info.cWithWhls
                            .properties.with2Whls_info.cOwnerShp = loJSONObject.disbursement_info.properties.with2Whls_info.cOwnerShp
                            .properties.with2Whls_info.cStatusxx = loJSONObject.disbursement_info.properties.with2Whls_info.cStatusxx
                            .properties.with2Whls_info.cTermxxxx = loJSONObject.disbursement_info.properties.with2Whls_info.cTermxxxx
                        End If

                        .monthly_expenses.nElctrcBl = loJSONObject.disbursement_info.monthly_expenses.nElctrcBl
                        .monthly_expenses.nWaterBil = loJSONObject.disbursement_info.monthly_expenses.nWaterBil
                        .monthly_expenses.nFoodAllw = loJSONObject.disbursement_info.monthly_expenses.nFoodAllw
                        .monthly_expenses.nLoanAmtx = loJSONObject.disbursement_info.monthly_expenses.nLoanAmtx
                        .monthly_expenses.nEductnxx = loJSONObject.disbursement_info.monthly_expenses.nEductnxx

                        .bank_account.sBankName = loJSONObject.disbursement_info.bank_account.sBankName
                        .bank_account.sAcctType = loJSONObject.disbursement_info.bank_account.sAcctType
                        .credit_card.sBankName = loJSONObject.disbursement_info.credit_card.sBankName
                        .credit_card.nCrdLimit = loJSONObject.disbursement_info.credit_card.nCrdLimit
                        .credit_card.nSinceYrx = loJSONObject.disbursement_info.credit_card.nSinceYrx
                    End If
                End With
            End If
        End With
    End Sub

    Class GOCAS_Param
        Property sBranchCd As String
        Property dAppliedx As String
        Property sClientNm As String
        Property cUnitAppl As String
        Property sUnitAppl As String
        Property nDownPaym As String
        'added by jovan 11-19-2020
        Property nRebatesx As String
        Property nPNValuex As String

        Property dCreatedx As String
        Property cApplType As String
        Property sModelIDx As String
        Property nAcctTerm As String
        Property nMonAmort As String
        Property dTargetDt As String
        Property applicant_info As New CARConst.applicant_param
        Property residence_info As New CARConst.client_param
        Property means_info As New CARConst.means_param
        Property other_info As New CARConst.other_param
        Property comaker_info As New CARConst.comaker_param
        Property spouse_info As New CARConst.spouse_param
        Property spouse_means As New CARConst.spouse__means_param
        Property disbursement_info As New CARConst.disbursement_param
    End Class

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