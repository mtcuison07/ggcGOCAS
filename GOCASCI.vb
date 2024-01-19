Option Strict Off

Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports rmjGOCAS
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ggcGOCAS.GOCASConst

Public Class GOCASCI
    Private Const p_sMasTable As String = "Credit_Online_Application"
    Private Const p_sCITable As String = "Credit_Online_Application_CI"

    Private p_oApp As GRider
    Private p_nEditMode As xeEditMode

    Private p_sTransNox As String
    Private p_oDTMaster As DataTable
    Private p_oDTSource As DataTable

    Private p_sSQL As String
    Private p_sCredInvx As String

    Private p_oResidence As residence_info
    Private p_oPropertyx As properties_info
    Private p_oMeansInfo As means_info

    Private p_xResidence As residence_info
    Private p_xPropertyx As properties_info
    Private p_xMeansInfo As means_info

    Sub New(ByVal foApp As GRider)
        p_oApp = foApp

        p_oDTMaster = Nothing
        p_oDTSource = Nothing

        p_oResidence = Nothing
        p_oPropertyx = Nothing
        p_oMeansInfo = Nothing

        p_xResidence = Nothing
        p_xPropertyx = Nothing
        p_xMeansInfo = Nothing

        p_sCredInvx = ""

        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Protected Overrides Sub Finalize()
        p_oApp = Nothing

        p_oDTMaster = Nothing
        p_oDTSource = Nothing

        p_oResidence = Nothing
        p_oPropertyx = Nothing
        p_oMeansInfo = Nothing

        p_xResidence = Nothing
        p_xPropertyx = Nothing
        p_xMeansInfo = Nothing

        MyBase.Finalize()
    End Sub

    Public WriteOnly Property TransNo() As String
        Set(ByVal fsValue As String)
            p_sTransNox = fsValue
        End Set
    End Property

    Public ReadOnly Property Master(ByVal Index As String) As Object
        Get
            If Not IsNothing(p_oDTSource) Then
                Return p_oDTSource(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property

    Public ReadOnly Property Others(ByVal Index As String) As Object
        Get
            If Not IsNothing(p_oDTMaster) Then
                Return p_oDTMaster(0).Item(Index)
            Else
                Return vbEmpty
            End If
        End Get
    End Property


    Public Property CI_Residence As residence_info
        Get
            Return p_oResidence
        End Get

        Set(ByVal foValue As residence_info)
            p_oResidence = foValue
        End Set
    End Property

    Public Property Result_Residence As residence_info
        Get
            Return p_xResidence
        End Get

        Set(ByVal foValue As residence_info)
            p_xResidence = foValue
        End Set
    End Property

    Public Property CI_Property As properties_info
        Get
            Return p_oPropertyx
        End Get

        Set(ByVal foValue As properties_info)
            p_oPropertyx = foValue
        End Set
    End Property

    Public Property Result_Property As properties_info
        Get
            Return p_xPropertyx
        End Get

        Set(ByVal foValue As properties_info)
            p_xPropertyx = foValue
        End Set
    End Property

    Public Property CI_Means_Info As means_info
        Get
            Return p_oMeansInfo
        End Get

        Set(ByVal foValue As means_info)
            p_oMeansInfo = foValue
        End Set
    End Property

    Public Property Result_Means_Info As means_info
        Get
            Return p_xMeansInfo
        End Get

        Set(ByVal foValue As means_info)
            p_xMeansInfo = foValue
        End Set
    End Property

    Public Function NewRecord() As Boolean
        If TypeName(p_oApp) = "Nothing" Then
            MsgBox("Application driver is not set.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        If TypeName(p_oDTMaster) = "Nothing" Then isRecordExist()

        If p_oDTMaster.Rows.Count > 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN

            MsgBox("This application's CI record was already processed.", MsgBoxStyle.Information, "Notice")
            Return False
        Else
            'If p_oDTSource(0)("cEvaluatr") = "1" Then
            '    MsgBox("This application is still For Evaluation.", MsgBoxStyle.Information, "Notice")
            '    Return False
            'End If

            'If p_oDTSource(0)("cTranStat") = "0" Then
            '    MsgBox("This application is still For Evaluation.", MsgBoxStyle.Information, "Notice")
            '    Return False
            'End If

            If p_oDTSource(0)("cEvaluatr") = "0" Then
                MsgBox("This application is already evaluated.", MsgBoxStyle.Information, "Notice")
                Return False
            End If

            If p_oDTSource(0)("cTranStat") = "1" Then
                MsgBox("This application is already evaluated.", MsgBoxStyle.Information, "Notice")
                Return False
            End If

            If p_oDTSource(0)("cTranStat") = "3" Or p_oDTSource(0)("cTranStat") = "4" Then
                MsgBox("This application is Disapproved/Void.", MsgBoxStyle.Information, "Notice")
                Return False
            End If

            loadJSON(p_oDTSource(0)("sCatInfox"))
            initQuestion()

            p_nEditMode = xeEditMode.MODE_ADDNEW
        End If

        Return True
    End Function

    Public Function LoadRecord() As Boolean
        If TypeName(p_oApp) = "Nothing" Then
            MsgBox("Application driver is not set.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        If TypeName(p_oDTMaster) = "Nothing" Then
            MsgBox("No record loaded.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        loadResult()

        p_nEditMode = xeEditMode.MODE_READY

        Return True
    End Function

    Public Function SaveRecord() As Boolean
        If TypeName(p_oApp) = "Nothing" Then
            MsgBox("Application driver is not set.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        If p_nEditMode <> xeEditMode.MODE_ADDNEW Then
            MsgBox("Invalid edit mode detected.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        If p_sCredInvx.Trim = "" Then
            MsgBox("No Credit Investigator was assigned.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If

        p_sSQL = "INSERT INTO " & p_sCITable & " SET" & _
                    "  sTransNox = " & strParm(p_sTransNox) & _
                    ", sCredInvx = " & strParm(p_sCredInvx) & _
                    ", sAddressx = '" & JsonConvert.SerializeObject(p_oResidence) & "'" & _
                    ", sAddrFndg = '" & JsonConvert.SerializeObject(p_xResidence) & "'" & _
                    ", sAssetsxx = '" & JsonConvert.SerializeObject(p_oPropertyx) & "'" & _
                    ", sAsstFndg = '" & JsonConvert.SerializeObject(p_xPropertyx) & "'" & _
                    ", sIncomexx = '" & JsonConvert.SerializeObject(p_oMeansInfo) & "'" & _
                    ", sIncmFndg = '" & JsonConvert.SerializeObject(p_xMeansInfo) & "'" & _
                    ", cTranStat = '0'" & _
                    ", dModified = " & datetimeParm(p_oApp.SysDate)
        p_oApp.BeginTransaction()
        If p_oApp.ExecuteActionQuery(p_sSQL) <= 0 Then 'no replication log
            p_oApp.RollBackTransaction()
            MsgBox("Unable to save request.", MsgBoxStyle.Critical, "Warning")
            Return False
        End If
        p_oApp.CommitTransaction()

        Return True
    End Function

    '0 - no record; 1 - has record; 2 - failed;
    Public Function isRecordExist() As Integer
        p_sSQL = AddCondition(getSQ_Master, "a.sTransNox = " & strParm(p_sTransNox))

        p_oDTSource = p_oApp.ExecuteQuery(p_sSQL)

        If p_oDTSource.Rows.Count <= 0 Then
            p_oDTSource = Nothing
            Return 2
        End If

        'If IsDBNull(p_oDTSource(0)("nCrdtScrx")) Then
        '    MsgBox("Application not yet computed Credit Score.", MsgBoxStyle.Critical, "Warning")
        '    p_oDTSource = Nothing
        '    Return 2
        'End If

        p_sSQL = AddCondition(getSQ_Master_CI, "a.sTransNox = " & strParm(p_sTransNox))

        p_oDTMaster = p_oApp.ExecuteQuery(p_sSQL)

        If p_oDTMaster.Rows.Count = 0 Then
            Return 0
        Else
            Return 1
        End If
    End Function

    Public Function getCreditInvestigator(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getCreditInvestigator"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "CONCAT(d.sFrstName, ' ', d.sLastName) LIKE " & strParm(sValue & "%")
                Else
                    lsCondition = "CONCAT(d.sFrstName, ' ', d.sLastName) = " & strParm(sValue)
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
            p_sCredInvx = loDT(0)("sCredInvx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sCredInvx»sFullName»sBranchNm", _
                                "ID»Name»Branch", _
                                "", _
                                "sCredInvx»CONCAT(d.sFrstName, ' ', d.sLastName)»sBranchNm", _
                                1)

            If Not IsNothing(loDataRow) Then
                getCreditInvestigator = loDataRow("sFullName")
                p_sCredInvx = loDataRow("sCredInvx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        p_sCredInvx = ""
        getCreditInvestigator = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getSQL_CI(ByVal fbByCode As Boolean, Optional ByVal fsValue As String = "") As String
        Dim lsSQL As String

        lsSQL = "SELECT" & _
                    "  a.sCredInvx" & _
                    ", CONCAT(d.sLastName, ', ', d.sFrstName) sFullName" & _
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

    Private Sub initQuestion()
        p_xResidence = New residence_info
        p_xResidence.present_address = New address_info
        p_xResidence.primary_address = New address_info

        p_xPropertyx = New properties_info

        p_xMeansInfo = New means_info
        p_xMeansInfo.employed = New employed
        p_xMeansInfo.self_employed = New self_employed
        p_xMeansInfo.financed = New financed
        p_xMeansInfo.pensioner = New pensioner
    End Sub

    Private Sub loadResult()
        Dim loSettings As New JsonSerializerSettings

        loSettings.DefaultValueHandling = DefaultValueHandling.Populate
        p_oResidence = JsonConvert.DeserializeObject(Of residence_info)(p_oDTMaster(0)("sAddressx"), loSettings)
        p_xResidence = JsonConvert.DeserializeObject(Of residence_info)(p_oDTMaster(0)("sAddrFndg"), loSettings)

        p_oPropertyx = JsonConvert.DeserializeObject(Of properties_info)(p_oDTMaster(0)("sAssetsxx"), loSettings)
        p_xPropertyx = JsonConvert.DeserializeObject(Of properties_info)(p_oDTMaster(0)("sAsstFndg"), loSettings)

        p_oMeansInfo = JsonConvert.DeserializeObject(Of means_info)(p_oDTMaster(0)("sIncomexx"), loSettings)
        p_xMeansInfo = JsonConvert.DeserializeObject(Of means_info)(p_oDTMaster(0)("sIncmFndg"), loSettings)
    End Sub

    Private Sub loadJSON(ByVal fsValue As String)
        Dim loJSONObject As New GOCAS_Param
        Dim loSettings As New JsonSerializerSettings

        loSettings.DefaultValueHandling = DefaultValueHandling.Populate
        loJSONObject = JsonConvert.DeserializeObject(Of GOCAS_Param)(fsValue, loSettings)

        If Not IsNothing(loJSONObject.residence_info) Then
            p_oResidence = New residence_info

            Dim loPresent As address_info
            loPresent = New address_info

            loPresent.cAddrType = "0"
            p_sSQL = loJSONObject.residence_info.present_address.sHouseNox
            p_sSQL += " " & loJSONObject.residence_info.present_address.sAddress1
            p_sSQL += " " & loJSONObject.residence_info.present_address.sAddress2

            If loJSONObject.residence_info.present_address.sBrgyIDxx <> "" Then
                p_sSQL += ", " & getBarangay(loJSONObject.residence_info.present_address.sBrgyIDxx)
            End If
            If loJSONObject.residence_info.present_address.sTownIDxx <> "" Then
                p_sSQL += ", " & getTown(loJSONObject.residence_info.present_address.sTownIDxx)
            End If

            loPresent.sAddressx = Trim(p_sSQL)
            loPresent.sAddrImge = ""

            Dim loPrimary As address_info
            loPrimary = New address_info

            loPrimary.cAddrType = "1"
            p_sSQL = loJSONObject.residence_info.permanent_address.sHouseNox
            p_sSQL += " " & loJSONObject.residence_info.permanent_address.sAddress1
            p_sSQL += " " & loJSONObject.residence_info.permanent_address.sAddress2

            If loJSONObject.residence_info.permanent_address.sBrgyIDxx <> "" Then
                p_sSQL += ", " & getBarangay(loJSONObject.residence_info.permanent_address.sBrgyIDxx)
            End If

            If loJSONObject.residence_info.permanent_address.sTownIDxx <> "" Then
                p_sSQL += ", " & getTown(loJSONObject.residence_info.permanent_address.sTownIDxx)
            End If

            loPrimary.sAddressx = Trim(p_sSQL)
            loPrimary.sAddrImge = ""

            p_oResidence.present_address = loPresent
            p_oResidence.primary_address = loPrimary
        End If

        If Not IsNothing(loJSONObject.disbursement_info) Then
            If Not IsNothing(loJSONObject.disbursement_info.properties) Then
                p_oPropertyx = New properties_info
                p_oPropertyx.sProprty1 = loJSONObject.disbursement_info.properties.sProprty1
                p_oPropertyx.sProprty2 = loJSONObject.disbursement_info.properties.sProprty2
                p_oPropertyx.sProprty3 = loJSONObject.disbursement_info.properties.sProprty3
                p_oPropertyx.cWith4Whl = loJSONObject.disbursement_info.properties.cWith4Whl
                p_oPropertyx.cWith3Whl = loJSONObject.disbursement_info.properties.cWith3Whl
                p_oPropertyx.cWith2Whl = loJSONObject.disbursement_info.properties.cWith2Whl
                p_oPropertyx.cWithRefx = loJSONObject.disbursement_info.properties.cWithRefx
                p_oPropertyx.cWithTVxx = loJSONObject.disbursement_info.properties.cWithTVxx
                p_oPropertyx.cWithACxx = loJSONObject.disbursement_info.properties.cWithACxx
            End If
        End If

        If Not IsNothing(loJSONObject.means_info) Then
            p_oMeansInfo = New means_info

            If Not IsNothing(loJSONObject.means_info.employed) Then
                If Trim(loJSONObject.means_info.employed.sEmployer) <> "" Then
                    Dim loEmployed As employed
                    loEmployed = New employed

                    p_sSQL = loJSONObject.means_info.employed.sWrkAddrx
                    If loJSONObject.means_info.employed.sWrkTownx <> "" Then
                        p_sSQL += ", " & getTown(loJSONObject.means_info.employed.sWrkTownx)
                    End If

                    loEmployed.sEmployer = loJSONObject.means_info.employed.sEmployer
                    loEmployed.sWrkAddrx = Trim(p_sSQL)
                    loEmployed.sPosition = loJSONObject.means_info.employed.sPosition
                    loEmployed.nLenServc = loJSONObject.means_info.employed.nLenServc
                    loEmployed.nSalaryxx = loJSONObject.means_info.employed.nSalaryxx

                    p_oMeansInfo.employed = loEmployed
                End If
            End If

            If Not IsNothing(loJSONObject.means_info.self_employed) Then
                If Trim(loJSONObject.means_info.self_employed.sBusiness) <> "" Then
                    Dim loSelfEmployed As self_employed
                    loSelfEmployed = New self_employed

                    p_sSQL = loJSONObject.means_info.self_employed.sBusAddrx
                    If loJSONObject.means_info.self_employed.sBusTownx <> "" Then
                        p_sSQL += ", " & getTown(loJSONObject.means_info.self_employed.sBusTownx)
                    End If

                    loSelfEmployed.sBusiness = loJSONObject.means_info.self_employed.sBusiness
                    loSelfEmployed.sBusAddrx = Trim(p_sSQL)
                    loSelfEmployed.nBusIncom = loJSONObject.means_info.self_employed.nBusIncom
                    loSelfEmployed.nMonExpns = loJSONObject.means_info.self_employed.nMonExpns
                    loSelfEmployed.nBusLenxx = loJSONObject.means_info.self_employed.nBusLenxx

                    p_oMeansInfo.self_employed = loSelfEmployed
                End If
            End If

            If Not IsNothing(loJSONObject.means_info.financed) Then
                If loJSONObject.means_info.financed.nEstIncme > 0.0# Then
                    Dim loFinanced As financed
                    loFinanced = New financed

                    If loJSONObject.means_info.financed.sNatnCode <> "" Then
                        loFinanced.sCntryNme = getCountry(loJSONObject.means_info.financed.sNatnCode)
                    End If

                    loFinanced.sFinancer = loJSONObject.means_info.financed.sFinancer

                    If loJSONObject.means_info.financed.sReltnCde <> "" Then
                        loFinanced.sReltnDsc = getRelationship(loJSONObject.means_info.financed.sReltnCde)
                    End If

                    loFinanced.nEstIncme = loJSONObject.means_info.financed.nEstIncme

                    p_oMeansInfo.financed = loFinanced
                End If
            End If

            If Not IsNothing(loJSONObject.means_info.pensioner) Then
                If loJSONObject.means_info.pensioner.nPensionx > 0.0# Then
                    Dim loPensioner As pensioner
                    loPensioner = New pensioner

                    loPensioner.nPensionx = loJSONObject.means_info.pensioner.nPensionx

                    If loJSONObject.means_info.pensioner.cPenTypex <> "" Then
                        loPensioner.sPensionx = getPension(loJSONObject.means_info.pensioner.cPenTypex)
                    End If

                    p_oMeansInfo.pensioner = loPensioner
                End If
            End If
        End If
    End Sub

#Region "Parameter"
    Private Function getTown(ByVal fsValue As String) As String
        p_sSQL = "SELECT" & _
                    " TRIM(CONCAT(a.sTownName, ' ', b.sProvName)) sTownName" & _
                " FROM TownCity a" & _
                    ", Province b" & _
                " WHERE a.sProvIDxx = b.sProvIDxx" & _
                    " AND a.sTownIDxx = " & strParm(fsValue)

        Dim loRS As DataTable = p_oApp.ExecuteQuery(p_sSQL)

        p_sSQL = ""
        If loRS.Rows.Count = 1 Then
            p_sSQL = loRS(0)("sTownName")
        End If

        Return p_sSQL
    End Function

    Private Function getBarangay(ByVal fsValue As String) As String
        p_sSQL = "SELECT sBrgyName FROM Barangay WHERE sBrgyIDxx = " & strParm(fsValue)

        Dim loRS As DataTable = p_oApp.ExecuteQuery(p_sSQL)

        p_sSQL = ""
        If loRS.Rows.Count = 1 Then
            p_sSQL = loRS(0)("sBrgyName")
        End If

        Return p_sSQL
    End Function

    Public Function getPosition(ByVal fsValue As String) As String
        p_sSQL = "SELECT sOccptnNm FROM Occupation WHERE sOccptnID = " & strParm(fsValue)

        Dim loRS As DataTable = p_oApp.ExecuteQuery(p_sSQL)

        p_sSQL = ""
        If loRS.Rows.Count = 1 Then
            p_sSQL = loRS(0)("sOccptnNm")
        End If

        Return p_sSQL
    End Function

    Private Function getCountry(ByVal fsValue As String) As String
        p_sSQL = "SELECT sCntryNme FROM Country WHERE sCntryCde = " & strParm(fsValue)

        Dim loRS As DataTable = p_oApp.ExecuteQuery(p_sSQL)

        p_sSQL = ""
        If loRS.Rows.Count = 1 Then
            p_sSQL = loRS(0)("sCntryNme")
        End If

        Return p_sSQL
    End Function

    Private Function getRelationship(ByVal fsValue As String) As String
        Select Case fsValue
            Case "0"
                Return "Children"
            Case "1"
                Return "Parents"
            Case "2"
                Return "Sibllings"
            Case "3"
                Return "Relatives"
            Case Else
                Return "Others"
        End Select
    End Function

    Private Function getPension(ByVal fsValue As String) As String
        Return IIf(fsValue = "0", "Public", "Private")
    End Function
#End Region

#Region "SQL"
    Private Function getSQ_Master() As String
        Return "SELECT" & _
                    "  a.sTransNox" & _
                    ", a.sBranchCd" & _
                    ", a.dTransact" & _
                    ", a.dTargetDt" & _
                    ", a.sClientNm" & _
                    ", a.sGOCASNox" & _
                    ", a.sGOCASNoF" & _
                    ", a.cUnitAppl" & _
                    ", a.sSourceCD" & _
                    ", IFNULL(a.sCatInfox, a.sDetlInfo) sCatInfox" & _
                    ", a.sDesInfox" & _
                    ", a.sQMatchNo" & _
                    ", a.sQMAppCde" & _
                    ", a.nCrdtScrx" & _
                    ", a.nDownPaym" & _
                    ", a.nDownPayF" & _
                    ", a.sRemarksx" & _
                    ", a.sCreatedx" & _
                    ", a.dReceived" & _
                    ", a.sVerified" & _
                    ", a.dVerified" & _
                    ", a.cWithCIxx" & _
                    ", a.cTranStat" & _
                    ", a.cDivision" & _
                    ", a.cEvaluatr" & _
                    ", a.sLoadedBy" & _
                    ", a.dModified" & _
                    ", a.sCredInvx" & _
                    ", a.sCoMkrRs1" & _
                    ", a.sCoMkrRs2" & _
                    ", b.sBranchNm" & _
                " FROM " & p_sMasTable & " a" & _
                    " LEFT JOIN Branch b ON a.sBranchCd = b.sBranchCd"
    End Function

    Private Function getSQ_Master_CI() As String
        Return "SELECT" +
                    "  a.sTransNox" +
                    ", a.sCredInvx" +
                    ", a.sAddressx" +
                    ", a.sAddrFndg" +
                    ", a.sAssetsxx" +
                    ", a.sAsstFndg" +
                    ", a.sIncomexx" +
                    ", a.sIncmFndg" +
                    ", a.cHasRecrd" +
                    ", a.sRecrdRem" +
                    ", a.sPrsnBrgy" +
                    ", a.sPrsnPstn" +
                    ", a.sPrsnNmbr" +
                    ", a.sNeighbr1" +
                    ", a.sNeighbr2" +
                    ", a.sNeighbr3" +
                    ", a.dRcmdRcv1" +
                    ", a.dRcmdtnx1" +
                    ", a.cRcmdtnx1" +
                    ", a.sRcmdtnx1" +
                    ", a.dRcmdRcv2" +
                    ", a.dRcmdtnx2" +
                    ", a.cRcmdtnx2" +
                    ", a.sRcmdtnx2" +
                    ", a.cTranStat" +
                    ", a.sApproved" +
                    ", a.dApproved" +
                    ", a.dModified" +
                    ", c.sBranchNm" +
                    ", CONCAT(d.sFrstName, ' ', d.sLastName) xCredInvx" +
                " FROM " + p_sCITable + " a" +
                    " LEFT JOIN " + p_sMasTable + " b ON a.sTransNox = b.sTransNox" +
                    " LEFT JOIN Branch c ON b.sBranchCd = c.sBranchCd" +
                    " LEFT JOIN Client_Master d ON a.sCredInvx = d.sClientID" +
                " WHERE a.cTranStat <> '3'"
    End Function
#End Region

#Region "JSON Properties"
    Class residence_info
        Property present_address As address_info
        Property primary_address As address_info
    End Class

    Class means_info
        Property employed As employed
        Property self_employed As self_employed
        Property financed As financed
        Property pensioner As pensioner
    End Class

    Class address_info
        Property cAddrType As String
        Property sAddressx As String
        Property sAddrImge As String
        Property nLatitude As Decimal
        Property nLongitud As Decimal
    End Class

    Class properties_info
        Property sProprty1 As String
        Property sProprty2 As String
        Property sProprty3 As String
        Property cWith4Whl As String
        Property cWith3Whl As String
        Property cWith2Whl As String
        Property cWithRefx As String
        Property cWithTVxx As String
        Property cWithACxx As String
    End Class

    Class employed
        Property sEmployer As String
        Property sWrkAddrx As String
        Property sPosition As String
        Property nLenServc As Decimal
        Property nSalaryxx As Decimal
    End Class

    Class self_employed
        Property sBusiness As String
        Property sBusAddrx As String
        Property nBusLenxx As Decimal
        Property nBusIncom As Decimal
        Property nMonExpns As Decimal
    End Class

    Class financed
        Property sFinancer As String
        Property sReltnDsc As String
        Property sCntryNme As String
        Property nEstIncme As Decimal
    End Class

    Class pensioner
        Property sPensionx As String
        Property nPensionx As Decimal
    End Class
#End Region
End Class
