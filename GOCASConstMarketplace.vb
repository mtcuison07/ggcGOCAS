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

Public Class GOCASConstMarketplace
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
        Property nUnitPrce As String
        Property nAcctTerm As String
        Property nMonAmort As String
        Property dTargetDt As String
        Property applicant_info As New GOCASConstMarketplace.applicant_param
        Property means_info As New GOCASConstMarketplace.means_param
        Property disbursement_info As New GOCASConstMarketplace.disbursement_param
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

    Class employed_param
        Property sIndstWrk As String
        Property sPosition As String
        Property nSalaryxx As String
    End Class

    Class self_employed_param
        Property sIndstBus As String
        Property nBusIncom As String
    End Class

    Class pensioner_param
        Property cPenTypex As String
        Property nPensionx As String
    End Class

    Class financed_param
        Property sReltnCde As String
        Property nEstIncme As String
    End Class
    Class applicant_param
        Property sLastName As String
        Property sFrstName As String
        Property sSuffixNm As String
        Property sMiddName As String
        Property dBirthDte As String
        Property sBirthPlc As String
        Property cCvilStat As String
        Property cGenderCd As String
        Property sMaidenNm As String
        Property facebook As New facebook_param
        Property sLandMark As String
        Property sHouseNox As String
        Property sAddress1 As String
        Property sAddress2 As String
        Property sTownIDxx As String
        Property sBrgyIDxx As String
    End Class
    Class facebook_param
        Property sFBAcctxx As String
    End Class
    Class disbursement_param
        Property bank_account As New bank_account_param
    End Class
    Class bank_account_param
        Property sBankName As String
        Property sAcctType As String
    End Class
End Class