Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
    Private Sub ButtonDivision_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonDivision.Click
        Call Globals.ThisAddIn.Division()
    End Sub

    Private Sub ButtonComplementer_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonComplementer.Click
        Call Globals.ThisAddIn.Complementer()
    End Sub

    Private Sub ButtonSansBord_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonSansBord.Click
        Call Globals.ThisAddIn.ZoneTexteSansBord()
    End Sub

    Private Sub ButtonSTI2D_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonSTI2D.Click
        Call Globals.ThisAddIn.ChargerSTI2D()
    End Sub

    Private Sub ButtonNvxSTI2D_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Call Globals.ThisAddIn.NouveauSTI2D()
    End Sub

    Private Sub ButtonSujet_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonSujet.Click
        Call Globals.ThisAddIn.CreeSujet()
    End Sub

    Private Sub ButtonMarque_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonMarque.Click
        Call Globals.ThisAddIn.MarquerCorrection()
    End Sub

    Private Sub ButtonDemarque_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonDemarque.Click
        Call Globals.ThisAddIn.DemarquerCorrection()
    End Sub

    Private Sub ButtonLicenceCC_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonLicenceCC.Click
        System.Diagnostics.Process.Start("http://creativecommons.org/choose/?lang=fr")
    End Sub

    Private Sub ButtonAPropos_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonAPropos.Click
        Dim formLicence_CC_By_NC_SA As New FormLicence_CC_By_NC_SA
        formLicence_CC_By_NC_SA.ShowDialog()
    End Sub

    Private Sub ButtonAide_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonAide.Click
        Call Globals.ThisAddIn.MontrerAide()
    End Sub

    Private Sub ButtonRelance_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonRelance.Click
        'On Error Resume Next
        If Globals.ThisAddIn.TestLeModèle Then
            Globals.ThisAddIn.Application.Run("ChargerLaBoitePrincipale")
        Else
            Call Globals.ThisAddIn.NouveauSTI2D()
        End If
    End Sub

    Private Sub ButtonAidePDF_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonAidePDF.Click
        Call Globals.ThisAddIn.MontrerAidePDF()
    End Sub

    Private Sub ButtonZoneDeCorrection_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonZoneDeCorrection.Click
        'If Globals.ThisAddIn.TestLeModèle Then
        '    Globals.ThisAddIn.Application.Run("InsèreZoneDeTexteArrondi")
        'End If
        Call Globals.ThisAddIn.InsèreZoneDeTexteArrondi()
    End Sub

    Private Sub ButtonAjouterModèlePerso_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonAjouterModèlePerso.Click
        System.Diagnostics.Process.Start(RepModel & "\STI2D\STI2D.mp4")

    End Sub

End Class
