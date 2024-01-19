'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     GOCAS Object
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
'       Started creating this object.
'  mac [ 06/15/2021 08:00 ]
'       - modified the procedure on how to approve transaction/compute credit score
'       - implements Try/Catch and Begin/Commit/Rollback database transaction
'       - added DisapproveTransaction ->> computed credit score and approved to be disapproved by CSS Head/Supervisor
'       - change the notification sending on branches
'           initial notification -> message informing the branch that the application is on process (from QMProcessor Utility).
'           final notification -> CSS will do the notification thru messenger GC for the final result.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Strict Off

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports rmjGOCAS
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ggcGOCAS.GOCASConst

Public Class GOCASApplication
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nEditMode As xeEditMode
    Private p_sBranchCd As String 'Branch code of the transaction to retrieve
    Private p_sBranchNm As String 'Branch Name of the transaction to retrieve
    Private p_nTranStat As Int32  'Transaction Status of the transaction to retrieve
    Private p_sParent As String

    Private Const p_sMasTable As String = "Credit_Online_Application"
    Private Const p_sPointsDetail As String = "Credit_Online_Application_Points_Detail"
    Private Const p_sMsgHeadr As String = "Credit_Online_Application"
    Private Const pxeCSSNumber As String = "09158683181" 'default
    Private Const pxeMaxReferx As Integer = 3

    Private jsonDet As String
    Private jsonObjDet As New GOCAS_Param
    Private jsonObjCat As New GOCAS_Param
    Private jsonObjDes As New GOCAS_Param

    Private p_sCSSNumbr As String

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

    Public Property Category() As GOCAS_Param
        Get
            Return jsonObjCat
        End Get

        Set(ByVal Value As GOCAS_Param)
            jsonObjCat = Value
        End Set
    End Property

    Public Property Description() As GOCAS_Param
        Get
            Return jsonObjDes
        End Get

        Set(ByVal Value As GOCAS_Param)
            jsonObjDes = Value
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

        jsonObjDet = New GOCAS_Param
        jsonObjCat = New GOCAS_Param
        jsonObjDes = New GOCAS_Param

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

        jsonObjDet = New GOCAS_Param
        'new mac 2022.05.19
        If IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), "") = "" Then
            Call populateJSONObject(jsonObjDet, IFNull(p_oDTMstr.Rows(0)("sCatInfox"), ""))
        Else
            Call populateJSONObject(jsonObjDet, IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), ""))
        End If

        'old
        'Call populateJSONObject(jsonObjDet, IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), ""))

        jsonObjCat = New GOCAS_Param
        If IFNull(p_oDTMstr.Rows(0)("sCatInfox"), "") = "" Then
            Call populateJSONObject(jsonObjCat, IFNull(p_oDTMstr.Rows(0)("sDetlInfo"), ""))
        Else
            Call populateJSONObject(jsonObjCat, IFNull(p_oDTMstr.Rows(0)("sCatInfox"), ""))
        End If

        jsonObjDes = New GOCAS_Param
        Call populateJSONObject(jsonObjDes, IFNull(p_oDTMstr.Rows(0)("sDesInfox"), ""))

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

        lsSQL = AddCondition(lsSQL, "a.cEvaluatr = " & strParm(IIf(fbEvaluate, 1, 0)))


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
                                        , "sTransNox»sClientNm»dTransact»sBranchNm" _
                                        , "Trans No»Client»Date»Branch", _
                                        , "a.sTransNox»a.sClientNm»a.dTransact»b.sBranchNm" _
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

    'mac 2021-06-15
    '   the form calling this function is not used/working
    Function DisApproved() As Boolean
        Dim instance As New GOCASCalculator
        Dim lnUnitTpye As String
        Dim lnDownPaym As Double

        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Try
            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            instance.setAppDriver = p_oApp
            instance.setJSON = IIf(IFNull(p_oDTMstr(0)("sCatInfox"), "") = "", p_oDTMstr(0)("sDetlInfo"), p_oDTMstr(0)("sCatInfox"))

            lnDownPaym = getDownpayment(Detail.cUnitAppl, _
                                        lnUnitTpye, _
                                        Detail.sModelIDx, _
                                        instance.Compute(), _
                                        Detail.dAppliedx)

            p_oDTMstr(0).Item("cTranStat") = "3"
            p_oDTMstr(0).Item("cEvaluatr") = "0"
            p_oDTMstr(0).Item("sGOCASNox") = createGOCAS(True, lnDownPaym)
            p_oDTMstr(0).Item("dModified") = p_oApp.SysDate
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            'mac 2021.10.06
            '   issue disapprove script of MC_Credit_Application
            lsSQL = "UPDATE MC_Credit_Application SET" & _
                        "  cTranStat = '3'" & _
                        ", sApproved = " & strParm(p_oApp.UserID) & _
                    " WHERE sReferNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            Call p_oApp.Execute(lsSQL, "xxxTableBranch", p_oApp.BranchCode)

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try
endwithroll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
    End Function

    'mac 2021-06-15
    '   disapprove already computed and approved application due to some circumstances
    '   only allow users with supervisor account and up
    Public Function DisapproveTransaction() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then
            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'If Not isUserHighRank() Then
        '    MsgBox("User is not allowed to DISAPPROVE an application.", vbCritical, p_sMsgHeadr)
        '    Return False
        'End If

        If p_oDTMstr(0).Item("cTranStat") = "3" Then
            MsgBox("Application was currently disapproved.", vbCritical, "Notice")
            Return False
        ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
            MsgBox("Application was currently voided.", vbCritical, "Notice")
            Return False
        End If

        Try
            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cTranStat") = "3"
            p_oDTMstr(0).Item("dModified") = p_oApp.SysDate
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

            If p_oApp.Execute(lsSQL, p_sMasTable) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", MsgBoxStyle.Critical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            'mac 2021.10.06
            '   issue disapprove script of MC_Credit_Application
            lsSQL = "UPDATE MC_Credit_Application SET" & _
                        "  cTranStat = '3'" & _
                        ", sApproved = " & strParm(p_oApp.UserID) & _
                    " WHERE sReferNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            Call p_oApp.Execute(lsSQL, "xxxTableBranch", p_oApp.BranchCode)

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithroll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
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
        If (Not isUserHighRank()) Then
            If p_oDTMstr(0).Item("cTranStat") = "1" Or p_oDTMstr(0).Item("cTranStat") = "2" Then
                If MsgBox("Request was already approved! Do you continue?", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr) = MsgBoxResult.Cancel Then
                    Return False
                End If
            ElseIf p_oDTMstr(0).Item("cTranStat") = "3" Then
                If MsgBox("Request was already cancelled! Do you continue?", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, p_sMsgHeadr) = MsgBoxResult.Cancel Then
                    Return False
                End If
            ElseIf p_oDTMstr(0).Item("cTranStat") = "4" Then
                MsgBox("Request was already voided!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
                Return False
            End If
        End If

        Try
            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cTranStat") = "4"
            p_oDTMstr(0).Item("cEvaluatr") = "0"
            p_oDTMstr(0).Item("dModified") = p_oApp.SysDate
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")))

            If p_oApp.Execute(lsSQL, p_sMasTable) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", MsgBoxStyle.Critical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            'mac 2023.12.14
            lsSQL = "UPDATE MC_Credit_Application SET cTranStat = '5' WHERE sReferNox = " & strParm(p_oDTMstr(0).Item("sTransNox"))
            p_oApp.Execute(lsSQL, "MC_Credit_Application")

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithroll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
    End Function

    Function PostQuickMatch() As Boolean
        Dim lnUnitTpye As String
        Dim lnDownPaym As Double
        Dim instance As New GOCASCalculator

        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Try
            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            instance.setAppDriver = p_oApp
            instance.setJSON = p_oDTMstr(0)("sCatInfox")

            lnDownPaym = getDownpayment(Detail.cUnitAppl, _
                                        lnUnitTpye, _
                                        Detail.sModelIDx, _
                                        instance.Compute(), _
                                        Detail.dAppliedx)


            p_oDTMstr(0).Item("sGOCASNox") = createGOCAS(True, lnDownPaym)

            Select Case Trim(Left(p_oDTMstr.Rows(0)("sQMatchNo"), 2))
                Case "DA", "BA"
                    Call getModel(Detail.sModelIDx, True, True, "", lnUnitTpye)

                    p_oDTMstr(0).Item("cTranStat") = "1"
                    p_oDTMstr(0).Item("nDownPaym") = 100
                    p_oDTMstr(0).Item("cWithCIxx") = xeLogical.YES
                Case Else
                    p_oDTMstr(0).Item("cTranStat") = "0"
                    p_oDTMstr(0).Item("nDownPaym") = lnDownPaym
            End Select

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), p_oApp.UserID, p_oApp.SysDate.ToString)

            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithroll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
    End Function

    'mac 2021-06-15
    '   fixed logic on computing credit score and approving COA
    Function Approved(ByVal fbComputeScore As Boolean) As Boolean
        Try
            If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

                MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
                Return False
            End If

            If p_sParent = "" Then p_oApp.BeginTransaction()

            'is the user wants to compute the credit score?
            If fbComputeScore Then
                If Not computeCreditScore() Then
                    If p_sParent = "" Then p_oApp.RollBackTransaction()
                    Return False
                End If
            End If

            'process approval
            If Not Approved() Then GoTo endwithroll

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithRoll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
    End Function

    'mac 2021-06-15
    '   callee will do the try/catch and the begin/commit transaction
    Private Function computeCreditScore() As Boolean
        Dim instance As New GOCASCalculator
        Dim lsSQL As String

        instance.setAppDriver = p_oApp
        Debug.Print(p_oDTMstr(0)("sCatInfox"))
        instance.setJSON = p_oDTMstr(0)("sCatInfox")
        p_oDTMstr(0).Item("nCrdtScrx") = instance.Compute()

        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)

        If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
            MsgBox("Unable to update Credit Score.", vbCritical, p_sMsgHeadr)
            Return False
        End If

        If CDbl(p_oApp.getConfiguration("CrdtScrSve")) = 1 Then
            lsSQL = "INSERT INTO " + p_sPointsDetail + " SET" & _
                    "  sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")) & _
                    ", nContactx = " & instance.getContactInfoRate & _
                    ", nResidnce = " & instance.getResidenceInfoRate & _
                    ", nDsposble = " & instance.getDisposableIncomeRate & _
                    ", nMobilePt = " & instance.getMobilePoints & _
                    ", nCvilStPt = " & instance.getCivilStatPoints & _
                    ", nFBPoints = " & instance.getFBPoints & _
                    ", nSelfEmpx = " & instance.getSelfEmployedPoints & _
                    ", nEmployed = " & instance.getEmployedPoints & _
                    ", nFinancer = " & instance.getFinancedPoints & _
                    ", nPensionr = " & instance.getPensionerPoints & _
                    ", nDpndntPt = " & instance.getDependentsPoints & _
                    ", dModified = " & dateParm(p_oApp.SysDate)

            If p_oApp.Execute(lsSQL, p_sPointsDetail, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update Credit Score Detail.", vbCritical, p_sMsgHeadr)
                Return False
            End If
        End If

        Return True
    End Function

    'mac 2021-06-15
    '   callee will do the try/catch and the begin/commit transaction
    Private Function Approved() As Boolean
        Dim lnUnitTpye As String
        Dim lnDownPaym As Double
        Dim instance As New GOCASCalculator
        Dim lsSQLBranch As String
        Dim loDTBranch As DataTable
        Dim loDT As DataTable

        instance.setAppDriver = p_oApp
        If IFNull(p_oDTMstr(0)("sCatInfox"), "") = "" Then
            instance.setJSON = p_oDTMstr(0)("sDetlInfo")
        Else
            instance.setJSON = p_oDTMstr(0)("sCatInfox")
        End If

        lnDownPaym = getDownpayment(Detail.cUnitAppl, _
                                    lnUnitTpye, _
                                    Detail.sModelIDx, _
                                    instance.Compute(), _
                                    Detail.dAppliedx)

        p_oDTMstr(0).Item("cTranStat") = "1"
        p_oDTMstr(0).Item("cEvaluatr") = "0"

        'jovan 01-28-2020' per sic mac need to populate this 2 field
        p_oDTMstr(0).Item("sVerified") = p_oApp.UserID
        p_oDTMstr(0).Item("dVerified") = p_oApp.getSysDate.ToString

        If IFNull(p_oDTMstr(0)("nCrdtScrx"), 0) = 0 Then
            instance.setJSON = IIf(IFNull(p_oDTMstr(0)("sCatInfox"), "") = "", p_oDTMstr(0)("sDetlInfo"), p_oDTMstr(0)("sCatInfox"))
            p_oDTMstr(0).Item("nCrdtScrx") = instance.Compute()
        End If

        p_oDTMstr(0).Item("sGOCASNox") = createGOCAS(IIf(p_oDTMstr(0)("nCrdtScrx") >= 50, False, True), IIf(lnDownPaym = 0, 200, lnDownPaym))

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery("SELECT" & _
                                        "  nSelPrice" & _
                                        ", nMinDownx" & _
                                    " FROM MC_Model_Price" & _
                                    " WHERE sModelIDx = " & strParm(Detail.sModelIDx))

        Dim instanceGen As GOCASCodeGen
        instanceGen = New GOCASCodeGen

        If Not instanceGen.Decode(p_oDTMstr(0).Item("sGOCASNox")) Then
            MsgBox("Unable to decode GOCAS Number.", MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        p_oDTMstr(0).Item("nDownPaym") = instanceGen.DownPayment
        p_oDTMstr(0).Item("cWithCIxx") = IIf(p_oDTMstr(0)("nCrdtScrx") >= 50, 0, 1)

        'mac 2021.03.22
        '   assign CI in-charge for this account
        'If Not reassignCI() Then
        '    MsgBox("Unable to assign Credit Specialist.", vbCritical, p_sMsgHeadr)
        '    Return False
        'End If

        Dim lsSQL As String

        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)

        If (p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0) Then
            MsgBox("Unable to update " + p_sMasTable + ".", vbCritical, p_sMsgHeadr)
            Return False
        End If

        lsSQLBranch = "SELECT *" & _
                            " FROM Branch_Mobile" & _
                            " WHERE sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd"))

        loDTBranch = New DataTable
        loDTBranch = ExecuteQuery(lsSQLBranch, p_oApp.Connection)

        lsSQL = "GOCAS #: " & p_oDTMstr(0).Item("sGOCASNox") & vbCrLf & _
                "Application of Mr./Ms. " & p_oDTMstr(0).Item("sClientNm") & " is on Process." & vbCrLf & _
                "Valid Until 60 days upon application." & vbCrLf & _
                "REF. #: " & p_oDTMstr(0).Item("sTransNox") & vbCrLf & _
                "-GUANZON Group"

        'For Branch
        For lnCtr As Integer = 0 To loDTBranch.Rows.Count - 1
            If Not createReply(lsSQL, loDTBranch(lnCtr)("sMobileNo"), p_oDTMstr(0).Item("sTransNox")) Then
                MsgBox("Unable to create BRANCH notification for this transaction.", vbCritical, p_sMsgHeadr)
                Return False
            End If
        Next

        'For CSS
        If Not createReply(lsSQL, p_sCSSNumbr, p_oDTMstr(0).Item("sTransNox")) Then
            MsgBox("Unable to create CSS notification for this transaction.", vbCritical, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    Function overrideResult(ByVal fnDownPayF As Double _
                            , ByRef fnMonPaymx As Double) As Boolean
        Dim loDt As DataTable
        Dim lsSQLBranch As String
        Dim loDTBranch As DataTable
        loDt = New DataTable

        If Not (p_nEditMode = xeEditMode.MODE_READY Or
                p_nEditMode = xeEditMode.MODE_UPDATE) Then
            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If Not isUserHighRank() Then
            MsgBox("User is not allowed to OVERRIDE an application.", vbCritical, p_sMsgHeadr)
            Return False
        End If

        Try
            lsSQLBranch = "SELECT" & _
                                "  a.nSelPrice" & _
                                ", b.nMiscChrg" & _
                                ", b.nRebatesx" & _
                                ", b.nEndMrtgg" & _
                                ", c.nFactorRt" & _
                            " FROM MC_Model_Price a" & _
                                ", MC_Category b" & _
                                ", MC_Term_Category c" & _
                            " WHERE a.sModelIDx = " & strParm(Detail.sModelIDx) & _
                                " AND a.sMCCatIDx = b.sMCCatIDx" & _
                                " AND b.sMCCatIDx = c.sMCCatIDx" & _
                                " AND " & strParm(Detail.nAcctTerm) & " BETWEEN c.nAcctTerm AND c.nAcctThru"
            loDt = p_oApp.ExecuteQuery(lsSQLBranch)

            If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                    p_nEditMode = xeEditMode.MODE_UPDATE) Then
                MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
                Return False
            End If

            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("nDownPayF") = TruncateDecimal(fnDownPayF / loDt(0)("nSelPrice"), 4) * 100
            p_oDTMstr(0).Item("sGOCASNoF") = createGOCAS(True, p_oDTMstr(0)("nDownPayF"))
            fnMonPaymx = Math.Round(((loDt(0)("nSelPrice") - fnDownPayF + loDt(0)("nMiscChrg")) _
                         * loDt(0)("nFactorRt") / Detail.nAcctTerm) + loDt(0)("nRebatesx") + (loDt(0)("nEndMrtgg") / Detail.nAcctTerm), 0)

            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)

            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            Dim instanceGen As GOCASCodeGen
            instanceGen = New GOCASCodeGen
            If Not instanceGen.Decode(p_oDTMstr(0).Item("sGOCASNoF")) Then
                MsgBox("Unable to decode GOCAS Number.", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            lsSQLBranch = "SELECT *" & _
                               " FROM Branch_Mobile" & _
                               " WHERE sBranchCd = " & strParm(p_oDTMstr(0).Item("sBranchCd"))

            loDTBranch = New DataTable
            loDTBranch = ExecuteQuery(lsSQLBranch, p_oApp.Connection)

            lsSQL = "FINAL GOCAS #: " & p_oDTMstr(0).Item("sGOCASNoF") & vbCrLf & _
                    "Application of Mr./Ms. " & p_oDTMstr(0).Item("sClientNm") & " is on Process." & vbCrLf & _
                    "Valid Until 60 days upon application." & vbCrLf & _
                    "REF. #: " & p_oDTMstr(0).Item("sTransNox") & vbCrLf & _
                    "-GUANZON Group"

            'For Branch
            For lnCtr As Integer = 0 To loDTBranch.Rows.Count - 1
                If Not createReply(lsSQL, loDTBranch(lnCtr)("sMobileNo"), p_oDTMstr(0).Item("sTransNox")) Then
                    MsgBox("Unable to create BRANCH notification for this transaction.", vbCritical, p_sMsgHeadr)
                    GoTo endwithroll
                End If
            Next

            'For collection
            If Not createReply(lsSQL, p_sCSSNumbr, p_oDTMstr(0).Item("sTransNox")) Then
                MsgBox("Unable to create CSSS notification for this transaction.", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithroll:
        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return False
    End Function

    Function TruncateDecimal(ByVal value As Decimal, ByVal precision As Integer) As Decimal
        Dim stepper As Decimal = Math.Pow(10, precision)
        Dim tmp As Decimal = Math.Truncate(stepper * value)
        Return tmp / stepper
    End Function

    Function Evaluate() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Try
            Dim lsSQL As String

            If p_sParent = "" Then p_oApp.BeginTransaction()

            p_oDTMstr(0).Item("cEvaluatr") = "1"
            p_oDTMstr(0).Item("cTranStat") = "0" 'mac 2022.10.05
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sTransNox = " & strParm(p_oDTMstr(0).Item("sTransNox")), , p_oApp.SysDate.ToString)

            If p_oApp.Execute(lsSQL, p_sMasTable, Left(p_oDTMstr.Rows(0).Item("sTransNox"), 4)) <= 0 Then
                MsgBox("Unable to update " + p_sMasTable + ".", vbCritical, p_sMsgHeadr)
                GoTo endwithroll
            End If

            If p_sParent = "" Then p_oApp.CommitTransaction()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
        End Try

endwithroll:
        If p_sParent = "" Then p_oApp.RollBackTransaction()

        Return False
    End Function

    Public Function getNextCustomer() As String
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = AddCondition(getSQ_Master, " sCatInfox IS NULL" & _
                                            " AND sLoadedBy IS NULL" & _
                                            " AND sSourceCd = 'APP'" & _
                                            " AND cTranStat IN ('0', '3')") & _
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
        For nCtr As Integer = 0 To Detail.other_info.personal_reference.Count - 1
            If Detail.other_info.personal_reference(nCtr).sRefrMPNx = fsMobileNo Then
                If nCtr + 1 = Detail.other_info.personal_reference.Count Then
                    fsRefNamex = Detail.other_info.personal_reference(0).sRefrNmex
                    Return Detail.other_info.personal_reference(0).sRefrMPNx
                Else
                    fsRefNamex = Detail.other_info.personal_reference(nCtr + 1).sRefrNmex
                    Return Detail.other_info.personal_reference(nCtr + 1).sRefrMPNx
                End If
            End If
        Next
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
                    ", sCredInvx" & _
                    ", sCoMkrRs1" & _
                    ", sCoMkrRs2" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sTransNox" & _
                    ", a.sClientNm" & _
                    ", a.dTransact" & _
                    ", b.sBranchNm" & _
                    ", a.sLoadedBy" & _
                  " FROM " & p_sMasTable & " a" & _
                    ", Branch b" & _
                  " WHERE a.sSourceCD = 'APP'" & _
                    " AND a.sBranchCd = b.sBranchCd"
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

    'mac 2021.03.22
    Private Function getSQL_CI(ByVal fbByCode As Boolean, Optional ByVal fsValue As String = "") As String
        Dim lsSQL As String

        lsSQL = "SELECT" & _
                    "  a.sCredInvx" & _
                    ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) sFullName" & _
                    ", e.sBranchNm" & _
                " FROM Route_Area a" & _
                        " LEFT JOIN Route_Area_Town b ON a.sRouteIDx = b.sRouteIDx" & _
                        " LEFT JOIN Branch e ON a.sBranchCd = e.sBranchCd" & _
                    ", Employee_Master001 c" & _
                        " LEFT JOIN Client_Master d ON c.sEmployID = d.sClientID" & _
                " WHERE a.sCredInvx = c.sEmployID" & _
                    " AND a.cTranStat = '1'" & _
                    " AND c.cRecdStat = '1'" & _
                " GROUP BY a.sCredInvx" & _
                " ORDER BY sFullName"

        If fbByCode Then
            lsSQL = AddCondition(lsSQL, "b.sTownIDxx =  " & strParm(fsValue))
        End If

        Return lsSQL
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
                " FROM MC_Model"
    End Function

    Private Function getSQL_Occupation() As String
        Return "SELECT" & _
                    "  sOccptnID" & _
                    ", sOccptnNm" & _
                " FROM Occupation" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    'jovan 04-22-2021 added to filter brangay depends on town idx selected
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

        p_sCSSNumbr = p_oApp.getConfiguration("CSSNmbr") 'SMS receiving mobile number of CSS Department
        If p_sCSSNumbr = "" Then p_sCSSNumbr = pxeCSSNumber 'assign the pre-defined number if configuration was empty

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

    'jovan 04-22-21 - add this code to search baragy via selected townid
    Function getBarangay(ByVal sValue As String _
                                , ByVal bSearch As Boolean _
                                , ByVal bByCode As Boolean _
                                , ByRef sBrgyIDxx As String _
                                , Optional ByVal sTownIDxx As String = "") As String
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

        Try
            loDT = New DataTable
            loDT = ExecuteQuery("SELECT * FROM Credit_Online_Application_Verification_History" & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")), p_oApp.Connection)

            lsSQL = "INSERT INTO Credit_Online_Application_Verification_History SET" & _
                        "  sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                        ", nEntryNox = " & CDbl(loDT.Rows.Count + 1) & _
                        ", sModified = " & strParm(p_oApp.UserID) & _
                        ", dModified = " & dateParm(p_oApp.SysDate)

            If p_oApp.Execute(lsSQL, "Credit_Online_Application_Verification_History", p_sBranchCd) = 0 Then
                MsgBox("Unable to Save History Info!!!", vbCritical, p_sMsgHeadr)
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
            Return False
        End Try

        Return True
    End Function

    Function confirmTransaction() As Boolean
        Dim lsSQL As String

        Try
            lsSQL = "UPDATE Credit_Online_Application SET" & _
                    "  sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
                    ", sDesInfox = " & "'" & (JSONObjDescription()) & "'" & _
                " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox"))

            If p_oApp.ExecuteActionQuery(lsSQL) = 0 Then
                MsgBox("Unable to Confirm Transaction!!!", vbCritical, p_sMsgHeadr)
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
            Return False
        End Try

        Return OpenTransaction(p_oDTMstr(0)("sTransNox"))
    End Function

    Function saveReference(ByVal fsMobileNo As String) As Boolean
        Dim lsSQL As String

        Try
            lsSQL = "INSERT INTO Credit_Online_Application_Reference SET" & _
                    "  sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                    ", sMobileNo = " & strParm(fsMobileNo) & _
                    ", sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
                    ", sDesInfox = " & "'" & (JSONObjDescription()) & "'" & _
                " ON DUPLICATE KEY UPDATE" & _
                    "  sCatInfox = " & "'" & (JSONObjCategory()) & "'" & _
                    ", sDesInfox = " & "'" & (JSONObjDescription()) & "'"

            If p_oApp.ExecuteActionQuery(lsSQL) = 0 Then
                MsgBox("Unable to Save Reference Transaction!!!", vbCritical, p_sMsgHeadr)
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, p_sMsgHeadr)
            Return False
        End Try

        Return True
    End Function

    Function callApplicant() As String()
        Dim lsMobile(Detail.applicant_info.mobile_number.Count - 1) As String

        For nCtr As Integer = 0 To Detail.applicant_info.mobile_number.Count - 1
            lsMobile(nCtr) = Detail.applicant_info.mobile_number(nCtr).sMobileNo
        Next nCtr

        p_oApp.Execute("UPDATE Credit_Online_Application SET" & _
                            " sLoadedBY = " & strParm(p_oApp.UserID) & _
                        " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")), "Credit_Online_Application", p_sBranchCd)

        Return lsMobile
    End Function

    'mac 2022.07.21
    '   check if all references was called.
    Function isReferenceOK() As Boolean
        Dim loDT As DataTable

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery("SELECT *" & _
                                        " FROM Credit_Online_Application_Reference" & _
                                        " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")))

        isReferenceOK = pxeMaxReferx <= loDT.Rows.Count
        loDT = Nothing
    End Function

    'mac 2022.07.21
    '   check if For-CI entry was approved or not needed.
    Function isForCIOkay() As Boolean
        Dim loDT As DataTable

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery("SELECT sTransNox, cTranStat" & _
                                    " FROM Credit_Online_Application_CI" & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")) & _
                                        " AND cTranStat <> '3'")
        isForCIOkay = False
        If loDT.Rows.Count = 0 Then
            isForCIOkay = MsgBox("No data is set for For-CI Reco. Do you want to continue?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes

        Else
            isForCIOkay = loDT(0)("cTranStat") = "2"
        End If

        loDT = Nothing
    End Function

    Function callReference(ByRef fsRefNamex As String) As String
        Dim loDT As DataTable
        Dim lbIsEqual As Boolean

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery("SELECT *" & _
                                    " FROM Credit_Online_Application_Reference" & _
                                    " WHERE sTransNox = " & strParm(p_oDTMstr(0)("sTransNox")))
        If loDT.Rows.Count = 0 Then
            If Detail.other_info.personal_reference.Count > 0 Then
                fsRefNamex = Detail.other_info.personal_reference(0).sRefrNmex
                Return Detail.other_info.personal_reference(0).sRefrMPNx
            End If
        Else
            For nCtr As Integer = 0 To Detail.other_info.personal_reference.Count - 1
                lbIsEqual = False
                For nCtr1 As Integer = 0 To loDT.Rows.Count - 1
                    If Not lbIsEqual Then
                        If Detail.other_info.personal_reference(nCtr).sRefrMPNx = loDT(nCtr1)("sMobileNo") Then
                            lbIsEqual = True
                            Exit For
                        End If
                    End If
                Next nCtr1

                If Not lbIsEqual Then
                    fsRefNamex = Detail.other_info.personal_reference(nCtr).sRefrNmex
                    Return Detail.other_info.personal_reference(nCtr).sRefrMPNx
                Else
                    GoTo movenext
                End If
movenext:
            Next nCtr

            Return ""
        End If

    End Function

    Function GenerateQM() As String
        Dim loQMResult As QMResult
        Dim loFrm As frmQuickMatch

        Dim lsCoMkrQM1 As String
        Dim lsCoMkrQM2 As String

        loQMResult = New QMResult
        loFrm = New frmQuickMatch

        With loQMResult
            .AppDriver = p_oApp
            .Branch = p_sBranchCd
            .ApplicationNo = IIf(p_oApp.ProductID <> "LRTrackr", p_oDTMstr.Rows(0)("sTransNox"), "")

            .InitTransaction()
            'Set the Applicant info
            .Applicant("sClientID") = ""

            .Applicant("sLastName") = Detail.applicant_info.sLastName
            .Applicant("sFrstName") = Detail.applicant_info.sFrstName & IIf(IFNull(Detail.applicant_info.sSuffixNm) = "", "", " " & Detail.applicant_info.sSuffixNm)
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

            'mac 2021.02.24
            '   added co-maker info for validation on QM
            If Not IsNothing(Detail.comaker_info.sLastName) Then
                If IFNull(Detail.comaker_info.sLastName) <> "" Then
                    .CoMaker("sClientID") = ""

                    .CoMaker("sLastName") = Detail.comaker_info.sLastName
                    .CoMaker("sFrstName") = Detail.comaker_info.sFrstName & IIf(IFNull(Detail.comaker_info.sSuffixNm) = "", "", " " & Detail.comaker_info.sSuffixNm)
                    .CoMaker("sMiddName") = Detail.comaker_info.sMiddName
                    .CoMaker("dBirthDte") = Detail.comaker_info.dBirthDte
                    .CoMaker("sBirthPlc") = Detail.comaker_info.sBirthPlc

                    If IFNull(Detail.comaker_info.residence_info.present_address.sAddress1) <> "" Then
                        .CoMaker("sAddressx") = Detail.comaker_info.residence_info.present_address.sAddress1
                    End If

                    If IFNull(Detail.comaker_info.residence_info.present_address.sAddress2) <> "" Then
                        .CoMaker("sAddressx") = .CoMaker("sAddressx") & " " & Detail.comaker_info.residence_info.present_address.sAddress2
                    End If

                    .CoMaker("sAddressx") = Trim(.CoMaker("sAddressx"))
                    .CoMaker("sBrgyIDxx") = Detail.comaker_info.residence_info.present_address.sBrgyIDxx
                    .CoMaker("sTownIDxx") = Detail.comaker_info.residence_info.present_address.sTownIDxx
                End If
            End If
            'end - 'mac 2021.02.24

            .Term("sModelIDx") = Detail.sModelIDx
            .Term("nDownPaym") = Detail.nDownPaym
            .Term("nAcctTerm") = Detail.nAcctTerm

            'Execute quickmatch here
            GenerateQM = .QuickMatch
            lsCoMkrQM1 = .QuickMatchResult("comaker1")
            lsCoMkrQM2 = .QuickMatchResult("comaker2")

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

            'mac 2021.05.15
            If Not IsNothing(Detail.comaker_info.sLastName) Then
                'Display spouse info
                If IFNull(Detail.comaker_info.sLastName) <> "" Then
                    loFrm.txtField10.Text = "N-O-N-E"
                    loFrm.txtField11.Text = "N-O-N-E"
                    loFrm.txtField12.Text = "N-O-N-E"
                Else
                    loFrm.txtField10.Text = .CoMaker("sLastName") & ", " & .CoMaker("sFrstName") & " " & .CoMaker("sMiddName")
                    loFrm.txtField11.Text = .CoMaker("sAddressx") & _
                                 ", " & getTownCity(.CoMaker("sTownIDxx"), True, True, "")
                    loFrm.txtField12.Text = lsCoMkrQM1
                End If
            End If

            loFrm.txtField13.Text = "N-O-N-E"
            loFrm.txtField14.Text = "N-O-N-E"
            loFrm.txtField15.Text = "N-O-N-E"
            'end - mac 2021.05.15

            p_oDTMstr.Rows(0)("sQMatchNo") = GenerateQM
            p_oDTMstr.Rows(0)("sCoMkrRs1") = lsCoMkrQM1
            p_oDTMstr.Rows(0)("sCoMkrRs2") = lsCoMkrQM2
            loFrm.txtField08.Text = p_oDTMstr.Rows(0)("sQMatchNo")
            loFrm.txtField09.Text = p_oDTMstr.Rows(0)("sTransNox")
            loFrm.txtField05.Text = Format(Detail.dAppliedx, "Mmmm DD, YYYY")

            loFrm.Result = .Result
            p_oFrmResult = loFrm
            loFrm.ShowDialog()
        End With
    End Function

    Private Function createGOCAS(ByVal fbIsCINeeded As Boolean, _
                         ByVal fnDownPaymnt As Long) As String
        Dim instance As GOCASCodeGen

        instance = New GOCASCodeGen

        With instance
            .UserID = p_oDTMstr.Rows(0)("sCreatedx") 'created
            .TransactionNo = p_oDTMstr.Rows(0)("sTransNox") 'table transaction number
            .LastName = Detail.applicant_info.sLastName
            .FirstName = Detail.applicant_info.sFrstName
            .MiddleName = Detail.applicant_info.sMiddName
            .SuffixName = Detail.applicant_info.sSuffixNm
            .IsCINeeded = fbIsCINeeded 'is CI needed
            .DownPayment = fnDownPaymnt 'approved downpayment
            .Encode() 'generate code
        End With

        Return instance.GOCASApprvl
    End Function

    Function getDownpayment(ByVal fcLoanType As String, _
                            ByVal fcUnitType As String, _
                            ByVal fsModelIDx As String, _
                            ByVal fnCredtScr As Double, _
                            ByVal fdTransact As Date) As Long
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = "SELECT" & _
                    " IFNULL(b.nDownPaym, a.nDownPaym) nDownPaym" & _
                " FROM Credit_Score_By_Model a" & _
                    " LEFT JOIN Credit_Score_By_Model_History b" & _
                        " ON a.sCSBMIDxx = b.sCSBMIDxx" & _
                        " AND " & dateParm(fdTransact) & " BETWEEN b.dDateFrom AND b.dDateThru" & _
                " WHERE a.sModelIDx = " & strParm(fsModelIDx) & _
                    " AND a.cLoanType = " & strParm(fcLoanType) & _
                    " AND " & fnCredtScr & " BETWEEN a.nScoreFrm AND a.nScoreThr"

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count > 0 Then
            Return loDT(0)("nDownPaym")
        End If

        lsSQL = "SELECT" & _
                    " IFNULL(b.nDownPaym, a.nDownPaym) nDownPaym" & _
                " FROM Credit_Score_By_Type a" & _
                    " LEFT JOIN Credit_Score_By_Type_History b" & _
                        " ON a.sCSBTIDxx = b.sCSBTIDxx" & _
                        " AND " & dateParm(fdTransact) & " BETWEEN b.dDateFrom AND b.dDateThru" & _
                " WHERE a.cUnitType = " & strParm(fcUnitType) & _
                    " AND a.cLoanType = " & strParm(fcLoanType) & _
                    " AND " & fnCredtScr & " BETWEEN a.nScoreFrm AND a.nScoreThr"

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count > 0 Then
            Return loDT(0)("nDownPaym")
        End If

        Return 0
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
                    ", a.sCLastNme" & _
                    ", a.sCFrstNme" & _
                    ", a.sCMiddNme" & _
                    ", a.dCBrthDte" & _
                    ", a.sCBrthPlc" & _
                    ", d.sTownName sCTownNme" & _
                    ", a.sCoMkrRs1" & _
                    ", a.sCLastNm2" & _
                    ", a.sCFrstNm2" & _
                    ", a.sCMiddNm2" & _
                    ", a.dCBrthDt2" & _
                    ", a.sCBrthPl2" & _
                    ", e.sTownName sCTownNm2" & _
                    ", a.sCoMkrRs2" & _
                " FROM MC_LR_QuickMatch a" & _
                    " LEFT JOIN TownCity b" & _
                        " ON a.sTownIDxx = b.sTownIDxx" & _
                    " LEFT JOIN TownCity c" & _
                        " ON a.sSTownIDx = c.sTownIDxx" & _
                    " LEFT JOIN TownCity d" & _
                        " ON a.sCTownIDx = d.sTownIDxx" & _
                    " LEFT JOIN TownCity e" & _
                        " ON a.sCTownID2 = e.sTownIDxx" & _
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

            .txtField10.Text = loDT(0)("sCLastNme") & ", " & loDT(0)("sCFrstNme") & " " & loDT(0)("sCMiddNme")
            .txtField11.Text = IFNull(loDT(0)("sCTownNme"), "")
            .txtField12.Text = IFNull(loDT(0)("sCoMkrRs1"), "N-O-N-E")

            .txtField13.Text = loDT(0)("sCLastNm2") & ", " & loDT(0)("sCFrstNm2") & " " & loDT(0)("sCMiddNm2")
            .txtField14.Text = IFNull(loDT(0)("sCTownNm2"), "")
            .txtField15.Text = IFNull(loDT(0)("sCoMkrRs2"), "N-O-N-E")

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

    Function JSONObjCategory() As String
        Dim loSettings As New JsonSerializerSettings

        loSettings.NullValueHandling = NullValueHandling.Ignore
        loSettings.DefaultValueHandling = DefaultValueHandling.Ignore

        Return JsonConvert.SerializeObject(jsonObjCat, loSettings)
    End Function

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

    Function JSONObjDescription() As String
        Dim loSettings As New JsonSerializerSettings

        loSettings.NullValueHandling = NullValueHandling.Ignore
        loSettings.DefaultValueHandling = DefaultValueHandling.Ignore

        Return JsonConvert.SerializeObject(jsonObjDes)
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
                        .mobile_number.Add(New GOCASConst.mobileno_param)
                        .mobile_number(nCtr).sMobileNo = loJSONObject.applicant_info.mobile_number(nCtr).sMobileNo
                        .mobile_number(nCtr).cPostPaid = loJSONObject.applicant_info.mobile_number(nCtr).cPostPaid
                        .mobile_number(nCtr).nPostYear = loJSONObject.applicant_info.mobile_number(nCtr).nPostYear
                    Next
                End If
                If Not IsNothing(loJSONObject.applicant_info.landline.Count) Then
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.landline.Count - 1
                        .landline.Add(New GOCASConst.landline_param)
                        .landline(nCtr).sPhoneNox = loJSONObject.applicant_info.landline(nCtr).sPhoneNox
                    Next
                End If

                .cCvilStat = loJSONObject.applicant_info.cCvilStat
                .cGenderCd = loJSONObject.applicant_info.cGenderCd
                .sMaidenNm = loJSONObject.applicant_info.sMaidenNm

                If Not IsNothing(loJSONObject.applicant_info.email_address.Count) Then
                    For nCtr As Integer = 0 To loJSONObject.applicant_info.email_address.Count - 1
                        .email_address.Add(New GOCASConst.email_param)
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
                        .personal_reference.Add(New GOCASConst.personal_reference_param)
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
                            .mobile_number.Add(New GOCASConst.mobileno_param)
                            .mobile_number(nCtr).sMobileNo = loJSONObject.comaker_info.mobile_number(nCtr).sMobileNo
                            .mobile_number(nCtr).cPostPaid = loJSONObject.comaker_info.mobile_number(nCtr).cPostPaid
                            .mobile_number(nCtr).nPostYear = loJSONObject.comaker_info.mobile_number(nCtr).nPostYear
                        Next
                    End If
                    .sFBAcctxx = loJSONObject.comaker_info.sFBAcctxx

                    'mac 2021.02.21
                    '   added address on comaker info
                    If Not IsNothing(loJSONObject.comaker_info.residence_info) Then
                        If Not IsNothing(loJSONObject.comaker_info.residence_info.present_address) Then
                            .residence_info.present_address.sLandMark = loJSONObject.comaker_info.residence_info.present_address.sLandMark
                            .residence_info.present_address.sHouseNox = loJSONObject.comaker_info.residence_info.present_address.sHouseNox
                            .residence_info.present_address.sAddress1 = loJSONObject.comaker_info.residence_info.present_address.sAddress1
                            .residence_info.present_address.sAddress2 = loJSONObject.comaker_info.residence_info.present_address.sAddress2
                            .residence_info.present_address.sTownIDxx = loJSONObject.comaker_info.residence_info.present_address.sTownIDxx
                            .residence_info.present_address.sBrgyIDxx = loJSONObject.comaker_info.residence_info.present_address.sBrgyIDxx
                        End If
                    End If
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
                        .personal_info.mobile_number.Add(New GOCASConst.mobileno_param)
                        .personal_info.mobile_number(nCtr).sMobileNo = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).sMobileNo
                        .personal_info.mobile_number(nCtr).cPostPaid = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).cPostPaid
                        .personal_info.mobile_number(nCtr).nPostYear = loJSONObject.spouse_info.personal_info.mobile_number(nCtr).nPostYear
                    Next
                    If Not IsNothing(loJSONObject.spouse_info.personal_info.landline) Then
                        For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.landline.Count - 1
                            .personal_info.landline.Add(New GOCASConst.landline_param)
                            .personal_info.landline(nCtr).sPhoneNox = loJSONObject.spouse_info.personal_info.landline(nCtr).sPhoneNox
                        Next
                    End If

                    .personal_info.cCvilStat = loJSONObject.spouse_info.personal_info.cCvilStat
                    .personal_info.cGenderCd = loJSONObject.spouse_info.personal_info.cGenderCd
                    .personal_info.sMaidenNm = loJSONObject.spouse_info.personal_info.sMaidenNm
                    If Not IsNothing(loJSONObject.spouse_info.personal_info.email_address) Then
                        For nCtr As Integer = 0 To loJSONObject.spouse_info.personal_info.email_address.Count - 1
                            .personal_info.email_address.Add(New GOCASConst.email_param)
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

                    If Not IsNothing(loJSONObject.disbursement_info.properties) Then
                        .properties.sProprty1 = loJSONObject.disbursement_info.properties.sProprty1
                        .properties.sProprty2 = loJSONObject.disbursement_info.properties.sProprty2
                        .properties.sProprty3 = loJSONObject.disbursement_info.properties.sProprty3
                        .properties.cWith4Whl = loJSONObject.disbursement_info.properties.cWith4Whl
                        .properties.cWith3Whl = loJSONObject.disbursement_info.properties.cWith3Whl
                        .properties.cWith2Whl = loJSONObject.disbursement_info.properties.cWith2Whl
                        .properties.cWithRefx = loJSONObject.disbursement_info.properties.cWithRefx
                        .properties.cWithTVxx = loJSONObject.disbursement_info.properties.cWithTVxx
                        .properties.cWithACxx = loJSONObject.disbursement_info.properties.cWithACxx
                        .monthly_expenses.nElctrcBl = loJSONObject.disbursement_info.monthly_expenses.nElctrcBl
                        .monthly_expenses.nWaterBil = loJSONObject.disbursement_info.monthly_expenses.nWaterBil
                        .monthly_expenses.nFoodAllw = loJSONObject.disbursement_info.monthly_expenses.nFoodAllw
                        .monthly_expenses.nLoanAmtx = loJSONObject.disbursement_info.monthly_expenses.nLoanAmtx
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

    Private Function createReply(ByVal fsMessages As String, _
                                  ByVal fsMobileNo As String, _
                                  ByVal fsTransNox As String) As Boolean
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
        Return True
    End Function

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
        Return (token.Type = JTokenType.Array And Not token.HasValues) Or
               (token.Type = JTokenType.Object And Not token.HasValues) Or
               (token.Type = JTokenType.Null)
    End Function

    'mac 2021.03.22
    Private Function getCreditInvestigator(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sCredInvx As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getCreditInvestigator"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) LIKE " & strParm(sValue & "%")
                Else
                    lsCondition = "CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) = " & strParm(sValue)
                End If
            Else
                lsCondition = "a.sCredInvx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_CI(False), lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getCreditInvestigator = loDT(0)("sFullName")
            sCredInvx = loDT(0)("sCredInvx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sCredInvx»sFullName»sBranchNm", _
                                "ID»Name»Branch", _
                                "", _
                                "sCredInvx»CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName)»sBranchNm", _
                                1)

            If Not IsNothing(loDataRow) Then
                getCreditInvestigator = loDataRow("sFullName")
                sCredInvx = loDataRow("sCredInvx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getCreditInvestigator = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    'mac 2022.03.22
    Private Function createSMSCI(ByVal fsValue As String) As Boolean
        Dim lsSQL As String
        Dim loRS As DataTable
        Dim loBranch As DataTable

        If fsValue = "" Then Return False

        lsSQL = "SELECT" & _
                    "  IF(IFNULL(sMobileNo, ''), IFNULL(sPhoneNox, ''), '') sMobileNo" & _
                " FROM Client_Master" & _
                " WHERE sClientID = " & strParm(fsValue)

        loRS = p_oApp.ExecuteQuery(lsSQL)

        If loRS.Rows.Count = 0 Then Return True

        If loRS(0)("sMobileNo") <> "" Then
            lsSQL = "SELECT sBranchNm FROM Branch WHERE sBranchCd = " & Str(p_oDTMstr(0).Item("sBranchCd"))
            loBranch = p_oApp.ExecuteQuery(lsSQL)

            lsSQL = "Good day. You have 1 new for CI." & vbCrLf & _
                    "Name: " & p_oDTMstr(0).Item("sClientNm") & vbCrLf & _
                    "Ref. #: " & p_oDTMstr(0).Item("sTransNox") & vbCrLf & _
                    "Branch: " & loBranch(0).Item("sBranchNm") & vbCrLf & _
                    "Thank you." & vbCrLf & _
                    "-Guanzon Group"

            'mac 2021.05.27
            'comment muna, ibalik pag iiimplement na yung evaluation android
            'createReply(lsSQL, loRS(0)("sMobileNo"), p_oDTMstr(0)("sTransNox"))
            'createReply(lsSQL, p_sCSSNumbr, p_oDTMstr(0)("sTransNox"))
        End If

        Return True
    End Function

    'mac 2021.03.22
    '   add CI assignment confirmation
    '   note: sCredInvx field was auto filled up by the utility QMProcessor if cWithCIxx is equal was set to 1
    Private Function reassignCI() As Boolean
        Dim lsSQL As String = ""

        p_oDTMstr(0).Item("sCredInvx") = IIf(p_oDTMstr(0).Item("cWithCIxx") = xeLogical.YES, p_oDTMstr(0).Item("sCredInvx"), "")

        If p_oDTMstr(0).Item("cWithCIxx") = xeLogical.YES Then
            If IFNull(p_oDTMstr(0).Item("sCredInvx"), "") <> "" Then
                'ask user if he wants to change the auto assigned CI.
getCI01:
                If MsgBox("Credit specialist " & getCreditInvestigator(p_oDTMstr(0).Item("sCredInvx"), True, True, "") &
                        " was assigned as customer CI." & vbCrLf & vbCrLf & "Do you want to change?", vbQuestion + vbYesNo, "Confirm") = vbYes Then

                    Call getCreditInvestigator("%", True, False, lsSQL)
                    If lsSQL <> "" Then
                        p_oDTMstr(0).Item("sCredInvx") = lsSQL
                    Else
                        GoTo getCI01
                    End If
                End If
            Else
getCI02:
                MsgBox("Please assign a credit specialist for this application.", vbInformation, "Select")

                Call getCreditInvestigator("%", True, False, lsSQL)
                If lsSQL <> "" Then
                    p_oDTMstr(0).Item("sCredInvx") = lsSQL
                Else
                    GoTo getCI02
                End If
            End If
        End If

        Return createSMSCI(p_oDTMstr(0)("sCredInvx"))
    End Function

    'mac 2021.06.16
    '   check if user is a CSS supervisor of higher
    Public Function isUserHighRank() As Boolean
        Dim lsSQL As String

        lsSQL = "SELECT" +
                    "  sDeptIDxx" +
                    ", sPositnID" +
                    ", sEmployID" +
                " FROM Employee_Master001" +
                " WHERE sEmployID = " + strParm(p_oApp.EmployNo) +
                    " AND cRecdStat = '1'"

        Dim loDT As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 1 Then
            If loDT(0)("sEmployID") = "M00122000390" Or
                loDT(0)("sEmployID") = "M00112001921" Or
                loDT(0)("sEmployID") = "M00122001049" Then 'ericka gawid, city valdez, erika abulencia
                Return True
            ElseIf loDT(0)("sDeptIDxx") = "022" Then
                Select Case loDT(0)("sPositnID")
                    Case "002", "053", "126", "098"
                        Return True
                End Select
            End If
        End If

        'If p_oApp.UserLevel >= xeUserRights.ENGINEER Then Return True

        Return False
    End Function
End Class