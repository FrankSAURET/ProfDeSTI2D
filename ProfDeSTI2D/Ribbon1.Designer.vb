Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Requis pour la prise en charge du Concepteur de composition de classes Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Cet appel est requis par le Concepteur de composants.
        InitializeComponent()

    End Sub

    'Component remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur de composants
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur de composants
    'Elle peut être modifiée à l'aide du Concepteur de composants.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabProf = Me.Factory.CreateRibbonTab
        Me.GroupMath = Me.Factory.CreateRibbonGroup
        Me.ButtonDivision = Me.Factory.CreateRibbonButton
        Me.ButtonComplementer = Me.Factory.CreateRibbonButton
        Me.GroupMiseEnForme = Me.Factory.CreateRibbonGroup
        Me.ButtonSansBord = Me.Factory.CreateRibbonButton
        Me.GroupProf = Me.Factory.CreateRibbonGroup
        Me.ButtonSTI2D = Me.Factory.CreateRibbonButton
        Me.ButtonRelance = Me.Factory.CreateRibbonButton
        Me.ButtonSujet = Me.Factory.CreateRibbonButton
        Me.ButtonZoneDeCorrection = Me.Factory.CreateRibbonButton
        Me.ButtonMarque = Me.Factory.CreateRibbonButton
        Me.ButtonDemarque = Me.Factory.CreateRibbonButton
        Me.GroupAPropos = Me.Factory.CreateRibbonGroup
        Me.SplitButtonAide = Me.Factory.CreateRibbonSplitButton
        Me.ButtonAPropos = Me.Factory.CreateRibbonButton
        Me.ButtonAide = Me.Factory.CreateRibbonButton
        Me.ButtonAidePDF = Me.Factory.CreateRibbonButton
        Me.ButtonAjouterModèlePerso = Me.Factory.CreateRibbonButton
        Me.ButtonLicenceCC = Me.Factory.CreateRibbonButton
        Me.TabProf.SuspendLayout()
        Me.GroupMath.SuspendLayout()
        Me.GroupMiseEnForme.SuspendLayout()
        Me.GroupProf.SuspendLayout()
        Me.GroupAPropos.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabProf
        '
        Me.TabProf.Groups.Add(Me.GroupMath)
        Me.TabProf.Groups.Add(Me.GroupMiseEnForme)
        Me.TabProf.Groups.Add(Me.GroupProf)
        Me.TabProf.Groups.Add(Me.GroupAPropos)
        Me.TabProf.Label = "Prof"
        Me.TabProf.Name = "TabProf"
        '
        'GroupMath
        '
        Me.GroupMath.Items.Add(Me.ButtonDivision)
        Me.GroupMath.Items.Add(Me.ButtonComplementer)
        Me.GroupMath.Label = "Math"
        Me.GroupMath.Name = "GroupMath"
        '
        'ButtonDivision
        '
        Me.ButtonDivision.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonDivision.Image = Global.ProfDeSTI2D.My.Resources.Resource1.asurb
        Me.ButtonDivision.Label = "Division"
        Me.ButtonDivision.Name = "ButtonDivision"
        Me.ButtonDivision.ShowImage = True
        '
        'ButtonComplementer
        '
        Me.ButtonComplementer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonComplementer.Image = Global.ProfDeSTI2D.My.Resources.Resource1.abarre
        Me.ButtonComplementer.Label = "Complémenter"
        Me.ButtonComplementer.Name = "ButtonComplementer"
        Me.ButtonComplementer.ShowImage = True
        '
        'GroupMiseEnForme
        '
        Me.GroupMiseEnForme.Items.Add(Me.ButtonSansBord)
        Me.GroupMiseEnForme.Label = "Mise en forme"
        Me.GroupMiseEnForme.Name = "GroupMiseEnForme"
        '
        'ButtonSansBord
        '
        Me.ButtonSansBord.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSansBord.Image = Global.ProfDeSTI2D.My.Resources.Resource1.ZoneSansBord
        Me.ButtonSansBord.Label = "Sans Bord"
        Me.ButtonSansBord.Name = "ButtonSansBord"
        Me.ButtonSansBord.ShowImage = True
        '
        'GroupProf
        '
        Me.GroupProf.Items.Add(Me.ButtonSTI2D)
        Me.GroupProf.Items.Add(Me.ButtonRelance)
        Me.GroupProf.Items.Add(Me.ButtonSujet)
        Me.GroupProf.Items.Add(Me.ButtonZoneDeCorrection)
        Me.GroupProf.Items.Add(Me.ButtonMarque)
        Me.GroupProf.Items.Add(Me.ButtonDemarque)
        Me.GroupProf.Label = "Prof de STI2D"
        Me.GroupProf.Name = "GroupProf"
        '
        'ButtonSTI2D
        '
        Me.ButtonSTI2D.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSTI2D.Image = Global.ProfDeSTI2D.My.Resources.Resource1.ModeleStidd
        Me.ButtonSTI2D.Label = "Modèle STI2D"
        Me.ButtonSTI2D.Name = "ButtonSTI2D"
        Me.ButtonSTI2D.ShowImage = True
        '
        'ButtonRelance
        '
        Me.ButtonRelance.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonRelance.Image = Global.ProfDeSTI2D.My.Resources.Resource1.BoitePrincipale_Petite
        Me.ButtonRelance.Label = "Nouveau ou Relance"
        Me.ButtonRelance.Name = "ButtonRelance"
        Me.ButtonRelance.ShowImage = True
        '
        'ButtonSujet
        '
        Me.ButtonSujet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonSujet.Image = Global.ProfDeSTI2D.My.Resources.Resource1.doceleve
        Me.ButtonSujet.Label = "Sujet"
        Me.ButtonSujet.Name = "ButtonSujet"
        Me.ButtonSujet.ShowImage = True
        '
        'ButtonZoneDeCorrection
        '
        Me.ButtonZoneDeCorrection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonZoneDeCorrection.Image = Global.ProfDeSTI2D.My.Resources.Resource1.Carré_Vert
        Me.ButtonZoneDeCorrection.Label = "Zone de correction"
        Me.ButtonZoneDeCorrection.Name = "ButtonZoneDeCorrection"
        Me.ButtonZoneDeCorrection.ShowImage = True
        '
        'ButtonMarque
        '
        Me.ButtonMarque.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonMarque.Image = Global.ProfDeSTI2D.My.Resources.Resource1.marquer
        Me.ButtonMarque.Label = "Elément de Correction"
        Me.ButtonMarque.Name = "ButtonMarque"
        Me.ButtonMarque.ShowImage = True
        '
        'ButtonDemarque
        '
        Me.ButtonDemarque.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonDemarque.Image = Global.ProfDeSTI2D.My.Resources.Resource1.demarquer
        Me.ButtonDemarque.Label = "Elément de sujet"
        Me.ButtonDemarque.Name = "ButtonDemarque"
        Me.ButtonDemarque.ShowImage = True
        '
        'GroupAPropos
        '
        Me.GroupAPropos.Items.Add(Me.SplitButtonAide)
        Me.GroupAPropos.Label = "À Propos"
        Me.GroupAPropos.Name = "GroupAPropos"
        '
        'SplitButtonAide
        '
        Me.SplitButtonAide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButtonAide.Image = Global.ProfDeSTI2D.My.Resources.Resource1.aide
        Me.SplitButtonAide.Items.Add(Me.ButtonAPropos)
        Me.SplitButtonAide.Items.Add(Me.ButtonAide)
        Me.SplitButtonAide.Items.Add(Me.ButtonAidePDF)
        Me.SplitButtonAide.Items.Add(Me.ButtonAjouterModèlePerso)
        Me.SplitButtonAide.Items.Add(Me.ButtonLicenceCC)
        Me.SplitButtonAide.Label = "Aide"
        Me.SplitButtonAide.Name = "SplitButtonAide"
        '
        'ButtonAPropos
        '
        Me.ButtonAPropos.Image = Global.ProfDeSTI2D.My.Resources.Resource1.apropos
        Me.ButtonAPropos.Label = "A propos ..."
        Me.ButtonAPropos.Name = "ButtonAPropos"
        Me.ButtonAPropos.ShowImage = True
        '
        'ButtonAide
        '
        Me.ButtonAide.Image = Global.ProfDeSTI2D.My.Resources.Resource1.aide
        Me.ButtonAide.Label = "Aide ..."
        Me.ButtonAide.Name = "ButtonAide"
        Me.ButtonAide.ShowImage = True
        '
        'ButtonAidePDF
        '
        Me.ButtonAidePDF.Image = Global.ProfDeSTI2D.My.Resources.Resource1.iconePDF
        Me.ButtonAidePDF.Label = "Afficher l'aide au format PDF ..."
        Me.ButtonAidePDF.Name = "ButtonAidePDF"
        Me.ButtonAidePDF.ShowImage = True
        '
        'ButtonAjouterModèlePerso
        '
        Me.ButtonAjouterModèlePerso.Image = Global.ProfDeSTI2D.My.Resources.Resource1.Caméra
        Me.ButtonAjouterModèlePerso.Label = "Ajouter un modèle perso dans « nouveau »."
        Me.ButtonAjouterModèlePerso.Name = "ButtonAjouterModèlePerso"
        Me.ButtonAjouterModèlePerso.ShowImage = True
        '
        'ButtonLicenceCC
        '
        Me.ButtonLicenceCC.Image = Global.ProfDeSTI2D.My.Resources.Resource1.cc1
        Me.ButtonLicenceCC.Label = "Licence Creative Common"
        Me.ButtonLicenceCC.Name = "ButtonLicenceCC"
        Me.ButtonLicenceCC.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.TabProf)
        Me.TabProf.ResumeLayout(False)
        Me.TabProf.PerformLayout()
        Me.GroupMath.ResumeLayout(False)
        Me.GroupMath.PerformLayout()
        Me.GroupMiseEnForme.ResumeLayout(False)
        Me.GroupMiseEnForme.PerformLayout()
        Me.GroupProf.ResumeLayout(False)
        Me.GroupProf.PerformLayout()
        Me.GroupAPropos.ResumeLayout(False)
        Me.GroupAPropos.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabProf As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupMath As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GroupMiseEnForme As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GroupProf As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GroupAPropos As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonDivision As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonComplementer As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSansBord As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSTI2D As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonRelance As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSujet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonZoneDeCorrection As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonMarque As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonDemarque As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButtonAide As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents ButtonAPropos As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonAide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonAidePDF As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonAjouterModèlePerso As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLicenceCC As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
