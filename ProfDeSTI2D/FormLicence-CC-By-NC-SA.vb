Imports System.Windows.Forms

Public Class FormLicence_CC_By_NC_SA
    Private Sub LinkLabelLicenceSTI2D_LinkClicked(sender As Object, e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabelLicenceSTI2D.LinkClicked
        Me.LinkLabelLicenceSTI2D.LinkVisited = True
        System.Diagnostics.Process.Start("https://creativecommons.org/licenses/by-nc-sa/4.0/deed.fr")
    End Sub

    Private Sub FormLicence_CC_By_NC_SA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = "Complément " & Me.ProductName & " Version " & Me.ProductVersion
        Label9.Text = RepModel
        Label10.Text = RepComplément
        Label11.Text = ModeDebugage
        Label12.Text = RepProgramFiles
        Label13.Text = NomAppli
        Label14.Text = CompanyAppli
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As EventArgs) Handles ButtonOk.Click
        Me.Close()
    End Sub

    Private Sub FormLicence_CC_By_NC_SA_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        ButtonOk.Left = Me.Width - ButtonOk.Width - 20

    End Sub
End Class