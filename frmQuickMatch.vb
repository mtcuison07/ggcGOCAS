Imports ggcAppDriver
Imports System.Windows.Forms

Public Class frmQuickMatch
    Private p_oApp As GRider
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_oResult As DataTable

    Public Property Appdriver() As ggcAppDriver.GRider
        Get
            Appdriver = p_oApp
        End Get
        Set(ByVal value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
    End Property

    Public WriteOnly Property Result As DataTable
        Set(ByVal value As DataTable)
            p_oResult = value
        End Set
    End Property

    Private Sub frmQuickMatch_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmQuickMatch_Activated")
        If pnLoadx = 1 Then
            Call ShowResult()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmQuickMatch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Debug.Print("frmQuickMatch_Load")
        If pnLoadx = 0 Then
            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            pnLoadx = 1
        End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 ' Exit
                Me.Hide()
        End Select
    End Sub

    Private Sub ShowResult()
        Dim lnCtr As Integer
        With dgv1
            For lnCtr = 0 To p_oResult.Rows.Count - 1
                .Rows.Add()
                .Rows(lnCtr).Cells(0).Value = Format(lnCtr + 1, "00")
                .Rows(lnCtr).Cells(1).Value = p_oResult(lnCtr)("sFullName")
                .Rows(lnCtr).Cells(2).Value = p_oResult(lnCtr)("sResltCde")
                .Rows(lnCtr).Cells(3).Value = p_oResult(lnCtr)("sAcctNmbr")
                .Rows(lnCtr).Cells(4).Value = p_oResult(lnCtr)("sMCSONmbr")
                .Rows(lnCtr).Cells(5).Value = p_oResult(lnCtr)("sApplNmbr")
            Next lnCtr
        End With
    End Sub
End Class