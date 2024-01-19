Option Strict Off

Imports ggcAppDriver
Imports ggcGOCAS.GOCASCI
Imports Newtonsoft.Json

Module modMain
    Public p_oAppDriver As GRider
    Public p_oTrans As GOCASCI

    Private p_oResidence As residence_info
    Private p_oPropertyx As properties_info
    Private p_oMeansInfo As means_info

    Private p_xResidence As residence_info
    Private p_xPropertyx As properties_info
    Private p_xMeansInfo As means_info

    Sub Main()
        p_oAppDriver = New GRider("LRTrackr")

        If Not p_oAppDriver.LoadEnv() Then
            MsgBox("Unable to load configuration file!")
            Exit Sub
        End If

        If Not p_oAppDriver.LogUser("M001111122") Then
            MsgBox("User unable to log!")
            Exit Sub
        End If

        p_oTrans = New GOCASCI(p_oAppDriver)

        With p_oTrans
            p_oTrans.TransNo = "CI5UV2200018"

            If p_oTrans.isRecordExist() Then
                'use history form here
                Debug.Print(p_oTrans.Master("sTransNox"))
                Debug.Print(p_oTrans.Master("sClientNm"))

                'sample usage of the class
                If p_oTrans.LoadRecord Then
                    'get field values
                    p_oResidence = p_oTrans.CI_Residence
                    p_oPropertyx = p_oTrans.CI_Property
                    p_oMeansInfo = p_oTrans.CI_Means_Info
                    'get result field values
                    p_xResidence = p_oTrans.Result_Residence
                    p_xPropertyx = p_oTrans.Result_Property
                    p_xMeansInfo = p_oTrans.Result_Means_Info

                    'displaying field values
                    If Not IsNothing(p_oResidence) Then 'residence info
                        'present address
                        Debug.Print(p_oResidence.present_address.sAddressx)
                        Debug.Print(p_oResidence.present_address.sAddrImge)
                        Debug.Print(p_oResidence.present_address.nLatitude)
                        Debug.Print(p_oResidence.present_address.nLongitud)

                        'primary address
                        Debug.Print(p_oResidence.primary_address.sAddressx)
                        Debug.Print(p_oResidence.primary_address.sAddrImge)
                        Debug.Print(p_oResidence.primary_address.nLatitude)
                        Debug.Print(p_oResidence.primary_address.nLongitud)
                    End If

                    If Not IsNothing(p_oPropertyx) Then 'property info
                        Debug.Print("Property 1: " & p_oPropertyx.sProprty1)
                        Debug.Print("Property 2: " & p_oPropertyx.sProprty2)
                        Debug.Print("Property 3: " & p_oPropertyx.sProprty3)
                        Debug.Print("Property 3: " & p_oPropertyx.sProprty3)

                        Debug.Print("W/ 4 Wheels: " & p_oPropertyx.cWith4Whl)
                        Debug.Print("W/ 3 Wheels: " & p_oPropertyx.cWith3Whl)
                        Debug.Print("W/ 2 Wheels: " & p_oPropertyx.cWith2Whl)
                        Debug.Print("W/ TV: " & p_oPropertyx.cWithRefx)
                        Debug.Print("W/ Ref: " & p_oPropertyx.cWithTVxx)
                        Debug.Print("W/ AC: " & p_oPropertyx.cWithACxx)
                    End If

                    If Not IsNothing(p_oMeansInfo) Then 'means info
                        'employed
                        If JsonConvert.SerializeObject(p_oMeansInfo.employed) <> "null" Then
                            Debug.Print("Employer: " & p_oMeansInfo.employed.sEmployer)
                            Debug.Print("Address: " & p_oMeansInfo.employed.sWrkAddrx)
                            Debug.Print("Service: " & p_oMeansInfo.employed.nLenServc)
                            Debug.Print("Position: " & p_oMeansInfo.employed.sPosition)
                            Debug.Print("Salary: " & p_oMeansInfo.employed.nSalaryxx)
                        End If

                        'self employed
                        If JsonConvert.SerializeObject(p_oMeansInfo.self_employed) <> "null" Then
                            Debug.Print("Business: " & p_oMeansInfo.self_employed.sBusiness)
                            Debug.Print("Address: " & p_oMeansInfo.self_employed.sBusAddrx)
                            Debug.Print("Bus. Len: " & p_oMeansInfo.self_employed.nBusLenxx)
                            Debug.Print("Income: " & p_oMeansInfo.self_employed.nBusIncom)
                            Debug.Print("Expense: " & p_oMeansInfo.self_employed.nMonExpns)
                        End If

                        'financed
                        If JsonConvert.SerializeObject(p_oMeansInfo.financed) <> "null" Then
                            Debug.Print("Financer: " & p_oMeansInfo.financed.sFinancer)
                            Debug.Print("Relationship: " & p_oMeansInfo.financed.sReltnDsc)
                            Debug.Print("Country: " & p_oMeansInfo.financed.sCntryNme)
                            Debug.Print("Est. Amount: " & p_oMeansInfo.financed.nEstIncme)
                        End If

                        'pensioner
                        If JsonConvert.SerializeObject(p_oMeansInfo.pensioner) <> "null" Then
                            Debug.Print("Pen. Source: " & p_oMeansInfo.pensioner.sPensionx)
                            Debug.Print("Est. Amount: " & p_oMeansInfo.pensioner.nPensionx)
                        End If
                    End If

                    'result of evaluation
                    '-1 - no answer yet
                    '0 - wrong info
                    '1 - correct info

                    'i want to check the result of CI on present and primary address of the applicant
                    Debug.Print(p_xResidence.present_address.sAddressx)
                    Debug.Print(p_xResidence.primary_address.sAddressx)

                    'get the long lat of the addresses
                    Debug.Print(p_xResidence.present_address.nLongitud)
                    Debug.Print(p_xResidence.present_address.nLatitude)

                    Debug.Print(p_xResidence.primary_address.nLongitud)
                    Debug.Print(p_xResidence.primary_address.nLatitude)

                    'get the images source of the addresses
                    Debug.Print(p_xResidence.present_address.sAddrImge)
                    Debug.Print(p_xResidence.primary_address.sAddrImge)

                    'i want to check the result of CI if the applicant has TV and Ref at home
                    Debug.Print(p_xPropertyx.cWithTVxx)
                    Debug.Print(p_xPropertyx.cWithRefx)

                    'i want to check the result of CI on the employment address and the salary of the applicant
                    Debug.Print(p_xMeansInfo.employed.sWrkAddrx)
                    Debug.Print(p_xMeansInfo.employed.nSalaryxx)
                End If
            Else
                'use entry form here

                'sample usage of the class
                If p_oTrans.NewRecord Then 'create new record
                    'get field values
                    p_oResidence = p_oTrans.CI_Residence
                    p_oPropertyx = p_oTrans.CI_Property
                    p_oMeansInfo = p_oTrans.CI_Means_Info
                    'get result field values
                    p_xResidence = p_oTrans.Result_Residence
                    p_xPropertyx = p_oTrans.Result_Property
                    p_xMeansInfo = p_oTrans.Result_Means_Info

                    'displaying field values
                    If Not IsNothing(p_oResidence) Then 'residence info
                        'present address
                        Debug.Print(p_oResidence.present_address.sAddressx)
                        Debug.Print(p_oResidence.present_address.sAddrImge)
                        Debug.Print(p_oResidence.present_address.nLatitude)
                        Debug.Print(p_oResidence.present_address.nLongitud)

                        'primary address
                        Debug.Print(p_oResidence.primary_address.sAddressx)
                        Debug.Print(p_oResidence.primary_address.sAddrImge)
                        Debug.Print(p_oResidence.primary_address.nLatitude)
                        Debug.Print(p_oResidence.primary_address.nLongitud)
                    End If

                    If Not IsNothing(p_oPropertyx) Then 'property info
                        Debug.Print("Property 1: " & p_oPropertyx.sProprty1)
                        Debug.Print("Property 2: " & p_oPropertyx.sProprty2)
                        Debug.Print("Property 3: " & p_oPropertyx.sProprty3)
                        Debug.Print("Property 3: " & p_oPropertyx.sProprty3)

                        Debug.Print("W/ 4 Wheels: " & p_oPropertyx.cWith4Whl)
                        Debug.Print("W/ 3 Wheels: " & p_oPropertyx.cWith3Whl)
                        Debug.Print("W/ 2 Wheels: " & p_oPropertyx.cWith2Whl)
                        Debug.Print("W/ TV: " & p_oPropertyx.cWithRefx)
                        Debug.Print("W/ Ref: " & p_oPropertyx.cWithTVxx)
                        Debug.Print("W/ AC: " & p_oPropertyx.cWithACxx)
                    End If

                    If Not IsNothing(p_oMeansInfo) Then 'means info
                        'employed
                        If JsonConvert.SerializeObject(p_oMeansInfo.employed) <> "null" Then
                            Debug.Print("Employer: " & p_oMeansInfo.employed.sEmployer)
                            Debug.Print("Address: " & p_oMeansInfo.employed.sWrkAddrx)
                            Debug.Print("Service: " & p_oMeansInfo.employed.nLenServc)
                            Debug.Print("Position: " & p_oMeansInfo.employed.sPosition)
                            Debug.Print("Salary: " & p_oMeansInfo.employed.nSalaryxx)
                        End If

                        'self employed
                        If JsonConvert.SerializeObject(p_oMeansInfo.self_employed) <> "null" Then
                            Debug.Print("Business: " & p_oMeansInfo.self_employed.sBusiness)
                            Debug.Print("Address: " & p_oMeansInfo.self_employed.sBusAddrx)
                            Debug.Print("Bus. Len: " & p_oMeansInfo.self_employed.nBusLenxx)
                            Debug.Print("Income: " & p_oMeansInfo.self_employed.nBusIncom)
                            Debug.Print("Expense: " & p_oMeansInfo.self_employed.nMonExpns)
                        End If

                        'financed
                        If JsonConvert.SerializeObject(p_oMeansInfo.financed) <> "null" Then
                            Debug.Print("Financer: " & p_oMeansInfo.financed.sFinancer)
                            Debug.Print("Relationship: " & p_oMeansInfo.financed.sReltnDsc)
                            Debug.Print("Country: " & p_oMeansInfo.financed.sCntryNme)
                            Debug.Print("Est. Amount: " & p_oMeansInfo.financed.nEstIncme)
                        End If

                        'pensioner
                        If JsonConvert.SerializeObject(p_oMeansInfo.pensioner) <> "null" Then
                            Debug.Print("Pen. Source: " & p_oMeansInfo.pensioner.sPensionx)
                            Debug.Print("Est. Amount: " & p_oMeansInfo.pensioner.nPensionx)
                        End If
                    End If

                    'by default values of string is null and for numeric field is zero
                    'assign "-1" to string and -1 to numeric fields if selected
                    'how to assign if selected

                    'i want the CI to check the present and primary address of the applicant
                    p_xResidence.present_address.sAddressx = "-1"
                    p_xResidence.primary_address.sAddressx = "-1"
                    .Result_Residence = p_xResidence 'submit the selection to the class

                    'i want the CI to check if the applicant has TV and Ref at home
                    p_xPropertyx.cWithTVxx = "-1"
                    p_xPropertyx.cWithRefx = "-1"
                    .Result_Property = p_xPropertyx

                    'i want the CI to check the employment address and the salary of the applicant
                    p_xMeansInfo.employed.sWrkAddrx = "-1"
                    p_xMeansInfo.employed.nSalaryxx = -1
                    .Result_Means_Info = p_xMeansInfo

                    'save the request
                    If .SaveRecord Then
                        MsgBox("Request saved successfully.", MsgBoxStyle.Information, "Notice")
                    End If
                End If
            End If
        End With
    End Sub
End Module
