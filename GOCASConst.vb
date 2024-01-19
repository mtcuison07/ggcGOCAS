'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     GOCAS JSON Constant Parameter
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

Public Class GOCASConst
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

    Class client_param
        Property cOwnershp As String
        Property cOwnOther As String
        Property rent_others As New rntOther_param
        Property sCtkReltn As String
        Property cHouseTyp As String
        Property cGaragexx As String
        Property present_address As New addres_param
        Property permanent_address As New addres_param
    End Class

    Class other_param
        Property sUnitUser As String
        Property sUsr2Buyr As String
        Property sPurposex As String
        Property sUnitPayr As String
        Property sPyr2Buyr As String
        Property sSrceInfo As String
        Property personal_reference As New List(Of personal_reference_param)
    End Class

    Class personal_reference_param
        Property sRefrNmex As String
        Property sRefrMPNx As String
        Property sRefrAddx As String
        Property sRefrTown As String
    End Class

    Class addres_param
        Property sLandMark As String
        Property sHouseNox As String
        Property sAddress1 As String
        Property sAddress2 As String
        Property sTownIDxx As String
        Property sBrgyIDxx As String
    End Class

    Class means_param
        Property cIncmeSrc As String
        Property employed As New employed_param
        Property self_employed As New self_employed_param
        Property pensioner As New pensioner_param
        Property financed As New financed_param
        Property other_income As New othrincm_param
    End Class

    Class othrincm_param
        Property nOthrIncm As String
        Property sOthrIncm As String
    End Class

    Class rntOther_param
        Property cRntOther As String
        Property nLenStayx As String
        Property nRentExps As String
    End Class

    Class employed_param
        Property cEmpSectr As String
        Property cUniforme As String
        Property cMilitary As String
        Property cGovtLevl As String
        Property cCompLevl As String
        Property cEmpLevlx As String
        Property cOcCatgry As String
        Property cOFWRegnx As String
        Property sOFWNatnx As String
        Property sIndstWrk As String
        Property sEmployer As String
        Property sWrkAddrx As String
        Property sWrkTownx As String
        Property sPosition As String
        Property sFunction As String
        Property cEmpStatx As String
        Property nLenServc As String
        Property nSalaryxx As String
        Property sWrkTelno As String
    End Class

    Class self_employed_param
        Property sIndstBus As String
        Property sBusiness As String
        Property sBusAddrx As String
        Property sBusTownx As String
        Property cBusTypex As String
        Property nBusLenxx As String
        Property nBusIncom As String
        Property nMonExpns As String
        Property cOwnTypex As String
        Property cOwnSizex As String
    End Class

    Class pensioner_param
        Property cPenTypex As String
        Property nPensionx As String
        Property nRetrYear As String
    End Class

    Class financed_param
        Property sReltnCde As String
        Property sFinancer As String
        Property nEstIncme As String
        Property sNatnCode As String
        Property sMobileNo As String
        Property sFBAcctxx As String
        Property sEmailAdd As String
    End Class

    Class mobileno_param
        Property sMobileNo As String
        Property cPostPaid As String
        Property nPostYear As String
    End Class

    Class applicant_param
        Property sLastName As String
        Property sFrstName As String
        Property sSuffixNm As String
        Property sMiddName As String
        Property sNickName As String
        Property dBirthDte As String
        Property sBirthPlc As String
        Property sCitizenx As String
        Property mobile_number As New List(Of mobileno_param)
        Property landline As New List(Of landline_param)
        Property cCvilStat As String
        Property cGenderCd As String
        Property sMaidenNm As String
        Property email_address As New List(Of email_param)
        Property facebook As New facebook_param
        Property sVibeAcct As String
    End Class

    Class facebook_param
        Property sFBAcctxx As String
        Property cAcctStat As String
        Property nNoFriend As String
        Property nYearxxxx As String
    End Class

    Class spouse_param
        Property personal_info As New applicant_param
        Property residence_info As New client_param
    End Class

    Class email_param
        Property sEmailAdd As String
    End Class

    Class landline_param
        Property sPhoneNox As String
    End Class

    Class spouse__means_param
        Property cIncmeSrc As String
        Property employed As New employed_param
        Property self_employed As New self_employed_param
        Property pensioner As New pensioner_param
        Property financed As New financed_param
        Property other_income As New othrincm_param
    End Class

    Class comaker_param
        Property sLastName As String
        Property sFrstName As String
        Property sSuffixNm As String
        Property sMiddName As String
        Property sNickName As String
        Property dBirthDte As String
        Property sBirthPlc As String
        Property cIncmeSrc As String
        Property sReltnCde As String
        Property mobile_number As New List(Of mobileno_param)
        Property sFBAcctxx As String
        Property residence_info As New client_param
    End Class

    Class dependent_param
        Property nHouseHld As String
        Property children As New List(Of children_param)
    End Class

    Class children_param
        Property sFullName As String
        Property sRelatnCD As String
        Property nDepdAgex As String
        Property cIsPupilx As String
        Property sSchlName As String
        Property sSchlAddr As String
        Property sSchlTown As String
        Property cIsPrivte As String
        Property sEducLevl As String
        Property cIsSchlrx As String
        Property cHasWorkx As String
        Property cWorkType As String
        Property sCompanyx As String
        Property cHouseHld As String
        Property cDependnt As String
        Property cIsChildx As String
        Property cIsMarrdx As String
    End Class

    Class disbursement_param
        Property dependent_info As New dependent_param
        Property properties As New properties_param
        Property monthly_expenses As New monthly_expense_param
        Property bank_account As New bank_account_param
        Property credit_card As New credit_card_param
    End Class

    Class properties_param
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

    Class monthly_expense_param
        Property nElctrcBl As String
        Property nWaterBil As String
        Property nFoodAllw As String
        Property nLoanAmtx As String
    End Class

    Class bank_account_param
        Property sBankName As String
        Property sAcctType As String
    End Class

    Class credit_card_param
        Property sBankName As String
        Property nCrdLimit As String
        Property nSinceYrx As String
    End Class
End Class