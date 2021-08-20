Imports Microsoft.Win32
Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Call Initialisation()
    End Sub
    Public Sub Complementer()
        'Ok le 4/12/98 par F.S.
        ' réécris le 3/05/2011 par F.S.
        Dim LaSelection As Object
        Dim Test As String = ""
        Dim Cpt As Integer

        LaSelection = Application.Selection.Text
        'Teste pour voir s'il y a des saut de ligne dans la sélection
        For Cpt = 1 To Len(LaSelection)
            If Mid(LaSelection, Cpt, 1) = Chr(13) Then
                MsgBox("La sélection ne doit pas inclure de marques de paragraphe. Sélectionnez à nouveau votre texte.", 16)
                GoTo Sortir
            End If
        Next Cpt

        Application.ScreenUpdating = False
        Application.Selection.Cut()
        Application.Selection.Fields.Add(Range:=Application.Selection.Range, Type:=Word.WdFieldType.wdFieldEmpty, Text:=
          "EQ \X \to() ", PreserveFormatting:=False)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)

        Application.Selection.Fields(1).ShowCodes = True
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        '  Supprime les espaces chiantes dans le champ
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=11)
        Application.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=2)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.Paste()
        Application.Selection.Fields.ToggleShowCodes()
        Application.ScreenUpdating = True
Sortir:
    End Sub
    Public Sub Division()
        ' Ok le 4/12/98 par F.S.
        ' modifié le 2/4/99 par F.S.
        ' modifié le 3/05/2011 par F.S.

        Dim LaSelection As Object
        Dim Test As String = ""
        Dim Cpt As Integer = 0
        Dim Position As Integer = 0

        LaSelection = Application.Selection.Text

        'Recherche du /
        For Cpt = 1 To Len(LaSelection)
            If Mid(LaSelection, Cpt, 1) = "/" Then Position = Cpt
        Next Cpt
        If Position = 0 Then
            MsgBox("La sélection doit contenir une barre de division [/].", 16)
            GoTo Sortir
        End If

        'Test pour voir s'il y a des saut de ligne dans la sélection
        For Cpt = 1 To Len(LaSelection)
            If Mid(LaSelection, Cpt, 1) = Chr(13) Then
                MsgBox("La sélection ne doit pas inclure de marques de paragraphe. Sélectionnez à nouveau votre texte.", 16)
                GoTo Sortir
            End If
        Next Cpt

        Application.ScreenUpdating = False
        Application.Selection.Cut()
        Application.Selection.TypeText(Text:="££")
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.Paste()
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.Fields.Add(Range:=Application.Selection.Range, Type:=Word.WdFieldType.wdFieldEmpty, Text:=
          "EQ \F(;) ", PreserveFormatting:=False)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)
        Application.Selection.Fields(1).ShowCodes = True
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)

        '  Supprime les espace chiant dans le champ
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=8)
        Application.Selection.Delete(Unit:=Word.WdUnits.wdCharacter, Count:=2)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=9)

        For Cpt = 1 To 2
            With Application.Selection.Find
                .Forward = False
                .ClearFormatting()
                .MatchWholeWord = False
                .Execute(FindText:="£")
            End With
        Next
        With Application.Selection
            .MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
            .Extend(Character:="/")
            .MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)
        End With
        Application.Selection.Cut()
        With Application.Selection.Find
            .ClearFormatting()
            .Text = "^d"
            .Forward = True
            .Execute()
        End With
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=3)
        Application.Selection.Paste()
        For Cpt = 1 To 2
            With Application.Selection.Find
                .Forward = False
                .ClearFormatting()
                .MatchWholeWord = False
                .Execute(FindText:="£")
            End With
        Next
        With Application.Selection
            .Extend(Character:="/")
            .MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
            .Extend(Character:="£")
            .MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1, Extend:=Word.WdMovementType.wdExtend)
            .Cut()
        End With
        With Application.Selection.Find
            .ClearFormatting()
            .Text = "^d"
            .Forward = True
            .Execute()
        End With
        Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCharacter, Count:=1)
        Application.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=2)
        Application.Selection.Paste()
        For Cpt = 1 To 2
            With Application.Selection.Find
                .Forward = False
                .ClearFormatting()
                .MatchWholeWord = False
                .Execute(FindText:="£")
            End With
        Next
        With Application.Selection
            .Extend(Character:="£")
        End With
        Application.Selection.Cut()

        With Application.Selection.Find
            .ClearFormatting()
            .Text = "^d"
            .Forward = True
            .Execute()
        End With
        Application.Selection.Fields.ToggleShowCodes()
        Application.ScreenUpdating = True
Sortir:
    End Sub
    Sub ZoneTexteSansBord()
        On Error GoTo ZoneDeTexte
        Application.Selection.ShapeRange.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        Application.Selection.ShapeRange.TextFrame.MarginLeft = 0.0#
        Application.Selection.ShapeRange.TextFrame.MarginRight = 0.0#
        Application.Selection.ShapeRange.TextFrame.MarginTop = 0.0#
        Application.Selection.ShapeRange.TextFrame.MarginBottom = 0.0#
        Application.Selection.ShapeRange.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
        Application.Selection.ShapeRange.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        Application.Selection.ShapeRange.RelativeHorizontalSize = Word.WdRelativeHorizontalSize.wdRelativeHorizontalSizePage
        Application.Selection.ShapeRange.RelativeVerticalSize = Word.WdRelativeVerticalSize.wdRelativeVerticalSizePage
        Application.Selection.ShapeRange.TextFrame.AutoSize = True
        Application.Selection.ShapeRange.TextFrame.WordWrap = True
        Exit Sub
ZoneDeTexte:
        MsgBox("Cette fonction ne marche que sur un objet graphique !")
    End Sub
    Sub ChargerSTI2D()
        ' CrééUnTP 
        '
        Application.AddIns.Add(FileName:=RepModel & "\STI2D.dotm", Install:=True)
        With Application.ActiveDocument
            .UpdateStylesOnOpen = True
            .AttachedTemplate = RepModel & "\STI2D.dotm"
            .XMLSchemaReferences.AutomaticValidation = True
            .XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
        End With
    End Sub
    Function TestLeModèle() As Boolean
        Dim returnValue As Object
        Dim NomDuModèle As String
        returnValue = Application.ActiveDocument.AttachedTemplate
        NomDuModèle = returnValue.FullName

        If NomDuModèle = RepModel & "\STI2D.dotm" Then TestLeModèle = True Else TestLeModèle = False
    End Function
    Sub NouveauSTI2D()
        Application.Documents.Add(Template:=RepModel & "\STI2D.dotm", NewTemplate:=False, DocumentType:=0)
    End Sub
    Sub CreeSujet()
        '
        'Macro de création du sujet à partir du document contenant le corrigé
        '
        Dim NombreDeForme As Integer
        Dim NomCourant As String
        Dim NomDeLaSousForme As String
        Dim Position As Integer = 0
        Dim Cpt As Integer = 0
        Dim Cpt2 As Integer = 0
        Dim IndexForme As Integer = 1
        Dim IndexSousForme As Integer = 1
        Dim NombreDeSousForme As Integer
        Dim RepertoireCourant As String
        Dim NomDuFichier As String
        Dim NomDuFichierDOrigine As String
        Dim NomDuFichierMin As String
        Dim Position1 As Integer
        Dim NbCar As Integer
        Dim FormPatiente As New FormPatientez
        Dim CoordonnéeX As Integer
        Dim CoordonnéeY As Integer
        Dim Largeur As Integer
        Dim Hauteur As Integer
        Dim LargeurFille As Integer
        Dim HauteurFille As Integer
        Dim CoordonnéeXFille As Integer
        Dim CoordonnéeYFille As Integer
        Dim Fond As String
        Dim VersionWord As String
        Dim ConserverTableau As String
        Dim NombreDeTableau As Integer
        Dim CouleurDuFond As Integer

        On Error Resume Next
        Application.ScreenUpdating = False
        'D'abord on sauvegarde
        Application.ActiveDocument.Save()
        'Définition de la position de la fenêtre parente
        Largeur = Application.ActiveDocument.ActiveWindow.Width
        Hauteur = Application.ActiveDocument.ActiveWindow.Height
        CoordonnéeX = Application.ActiveDocument.ActiveWindow.Left
        CoordonnéeY = Application.ActiveDocument.ActiveWindow.Top
        LargeurFille = FormPatiente.Width
        HauteurFille = FormPatiente.Height
        CoordonnéeXFille = (Largeur - LargeurFille) + CoordonnéeX
        CoordonnéeYFille = (Hauteur - HauteurFille) + CoordonnéeY

        FormPatiente.Show()
        FormPatiente.Top = CoordonnéeYFille
        FormPatiente.Left = CoordonnéeXFille
        FormPatiente.ProgressBar1.Visible = True
        FormPatiente.Label1.Visible = True
        FormPatiente.Refresh()

        VersionWord = Application.Version
        '*************************************************************************************************************************
        'Supprime le tableau pédagogique si demandé
        ConserverTableau = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\" & VersionWord & "\Word\STI2D", False).GetValue("Tableau élève")
        If ConserverTableau = "Faux" Or ConserverTableau = "False" Then

            NombreDeTableau = Application.ActiveDocument.Tables.Count
            IndexForme = 1
            For Cpt = 1 To NombreDeTableau
                NomCourant = Application.ActiveDocument.Tables.Item(IndexForme).Title

                If NomCourant = "TableauPédagogique" Then
                    Application.ActiveDocument.Tables.Item(IndexForme).Delete()
                End If
                IndexForme = IndexForme + 1
                NomCourant = ""
            Next
        End If

        '*************************************************************************************************************************
        'Recherche et supprime tous les caractères Verts foncés (La correction qui ne laisse pas de trou)
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = Word.WdColor.wdColorGreen
        Application.Selection.Find.Replacement.ClearFormatting()
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        Application.ScreenUpdating = True
        'Recherche et supprime ce qui est du style "Réponse Verte"
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Style = Application.ActiveDocument.Styles("Réponse Verte")
        Application.Selection.Find.Replacement.ClearFormatting()
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        FormPatiente.ProgressBar1.Value = 20
        FormPatiente.Refresh()
        '*************************************************************************************************************************
        'Recherche et supprime tous les caractères bleu (Les commentaires pédagogiques)
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = Word.WdColor.wdColorBlue
        Application.Selection.Find.Replacement.ClearFormatting()
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        'Recherche et supprime ce qui est du style "Commentaire Pédagogique"
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Style = Application.ActiveDocument.Styles("Commentaire Pédagogique")
        Application.Selection.Find.Replacement.ClearFormatting()
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        FormPatiente.ProgressBar1.Value = 40
        FormPatiente.Refresh()
        '*************************************************************************************************************************
        'Recherche et remplace tous les caractères rouges par des x de la couleur du fond en respectant les marques de paragraphe (La correction qui laisse des trous)
        Fond = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\" & VersionWord & "\Word\STI2D", False).GetValue("Fond Doc élève")

        If Fond = "Vrai" Or Fond = "True" Then
            CouleurDuFond = RGB(255, 255, 204)
        Else
            CouleurDuFond = RGB(255, 255, 255)
        End If

        'avant sauvegarde des marques de paragraphe
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = Word.WdColor.wdColorRed
        Application.Selection.Find.Replacement.ClearFormatting()
        Application.Selection.Find.Replacement.Font.Color = CouleurDuFond
        With Application.Selection.Find
            .Text = "^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = Word.WdColor.wdColorRed
        Application.Selection.Find.Replacement.ClearFormatting()
        Application.Selection.Find.Replacement.Font.Size = 16
        Application.Selection.Find.Replacement.Font.Color = CouleurDuFond
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = "x"
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        FormPatiente.ProgressBar1.Value = 60
        FormPatiente.Refresh()

        '*************************************************************************************************************************
        'Recherche et remplace tous les caractères violets par des ..
        'avant sauvegarde des marques de paragraphe
        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = RGB(204, 0, 204)
        Application.Selection.Find.Replacement.ClearFormatting()
        Application.Selection.Find.Replacement.Font.Color = Word.WdColor.wdColorAutomatic
        With Application.Selection.Find
            .Text = "^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)

        Application.Selection.Find.ClearFormatting()
        Application.Selection.Find.Font.Color = RGB(204, 0, 204)
        Application.Selection.Find.Replacement.ClearFormatting()
        Application.Selection.Find.Replacement.Style = Application.ActiveDocument.Styles("Pointillé Car")
        With Application.Selection.Find
            .Text = "^?"
            .Replacement.Text = ".."
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Application.Selection.Find.Execute(Replace:=Word.WdReplace.wdReplaceAll)
        Application.ScreenUpdating = True
        FormPatiente.ProgressBar1.Value = 70
        FormPatiente.Refresh()
        '*************************************************************************************************************************
        'Supprime les éléments de dessin marquée

        'supprime les formes et les formes en ligne marquées avec correction
        NombreDeForme = Application.ActiveDocument.Shapes.Count
        IndexForme = 1
        For Cpt = 1 To NombreDeForme
            'Attention, ne regarde que la forme IndexForme car elles sont automatiquement renumérotés à la suite de l'effacement
            NomCourant = Application.ActiveDocument.Shapes.Item(IndexForme).Title
            If NomCourant = "Correction" Then
                Application.ActiveDocument.Shapes.Item(IndexForme).Delete()
            Else
                'Pareil pour les sous formes
                On Error Resume Next
                NombreDeSousForme = Application.ActiveDocument.Shapes.Item(IndexForme).CanvasItems.Count 'Erreur ici si pas de sous forme
                IndexSousForme = 1
                For Cpt2 = 1 To NombreDeSousForme
                    'Attention, ne regarde que la forme IndexSousForme car elles sont automatiquement renumérotés à la suite de l'effacement
                    NomCourant = Application.ActiveDocument.Shapes.Item(IndexForme).CanvasItems(IndexSousForme).Title
                    If NomCourant = "Correction" Then
                        Application.ActiveDocument.Shapes.Item(IndexForme).CanvasItems.Item(IndexSousForme).Delete()
                    Else : IndexSousForme = IndexSousForme + 1
                    End If
                Next
            End If
            IndexForme = IndexForme + 1
        Next
        FormPatiente.ProgressBar1.Value = 80
        FormPatiente.Refresh()
        NombreDeForme = Application.ActiveDocument.InlineShapes.Count
        IndexForme = 1
        For Cpt = 1 To NombreDeForme
            'Attention, ne regarde que la forme IndexForme car elles sont automatiquement renumérotés à la suite de l'effacement
            NomCourant = Application.ActiveDocument.InlineShapes.Item(IndexForme).Title
            If NomCourant = "Correction" Then
                Application.ActiveDocument.InlineShapes.Item(IndexForme).Delete()
            Else : IndexForme = IndexForme + 1
            End If
        Next

        'supprime tableaux marquées avec correction
        NombreDeForme = Application.ActiveDocument.Tables.Count
        IndexForme = 1
        For Cpt = 1 To NombreDeForme
            'Attention, ne regarde que la forme IndexForme car elles sont automatiquement renumérotés à la suite de l'effacement
            NomCourant = Application.ActiveDocument.Tables.Item(IndexForme).Title()
            If NomCourant = "Correction" Then
                Application.ActiveDocument.Tables.Item(IndexForme).Delete()
            Else
                IndexForme = IndexForme + 1
            End If

        Next
        FormPatiente.ProgressBar1.Value = 90
        FormPatiente.Refresh()

        'Sauvegarde avec ajout de " - Elèves" à la fin - Supprime prof s'il existe dans le nom
        NomDuFichier = Application.ActiveDocument.FullName
        NomDuFichierDOrigine = NomDuFichier
        NomDuFichierMin = LCase(NomDuFichier)
        Position1 = InStr(NomDuFichierMin, ".docx", Microsoft.VisualBasic.vbTextCompare)
        If Position1 = 0 Then
            Position1 = InStr(NomDuFichierMin, ".doc", Microsoft.VisualBasic.vbTextCompare)
        End If
        NomDuFichier = Left(NomDuFichier, Position1 - 1)
        NomDuFichier = NomDuFichier & " - Elèves.Docx"

        NomDuFichierMin = LCase(NomDuFichier)

        NbCar = 0
        Position1 = InStr(NomDuFichierMin, " - prof", Microsoft.VisualBasic.vbTextCompare)
        If Position1 = 0 Then
            Position1 = InStr(NomDuFichierMin, " prof", Microsoft.VisualBasic.vbTextCompare)
        Else : NbCar = 7
        End If
        If Position1 = 0 Then
            Position1 = InStr(NomDuFichierMin, "-prof", Microsoft.VisualBasic.vbTextCompare)
        ElseIf NbCar = 0 Then
            NbCar = 5
        End If
        If Position1 = 0 Then
            Position1 = InStr(NomDuFichierMin, "prof", Microsoft.VisualBasic.vbTextCompare)
        ElseIf NbCar = 0 Then
            NbCar = 5
        End If
        If Position1 > 0 And NbCar = 0 Then NbCar = 4

        If NbCar > 0 Then
            NomDuFichier = NomDuFichier.Remove(Position1 - 1, NbCar)
        End If


        'changement de la couleur de fond de page

        Fond = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\" & VersionWord & "\Word\STI2D", False).GetValue("Fond Doc élève")

        If Fond = "Vrai" Or Fond = "True" Then
            Application.ActiveDocument.Background.Fill.ForeColor.RGB = RGB(255, 255, 204)
            Application.ActiveDocument.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
            Application.ActiveDocument.Background.Fill.Solid()
        Else
            Application.ActiveDocument.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        End If

        Dim DocElève = Application.Documents(NomDuFichier)
        If DocElève Is Nothing Then
            Application.ActiveDocument.SaveAs2(FileName:=
                NomDuFichier, FileFormat:=Word.WdSaveFormat.wdFormatXMLDocument, LockComments:=False, Password:="",
                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False,
                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                :=False, SaveAsAOCELetter:=False, CompatibilityMode:=14)
        Else
            'Application.Windows(NomDuFichier).Activate()
            Application.Documents.Item(NomDuFichier).Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            'Application.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            'Application.Windows(NomDuFichierDOrigine).Activate()
            Application.Documents.Item(NomDuFichierDOrigine).SaveAs2(FileName:=
                NomDuFichier, FileFormat:=Word.WdSaveFormat.wdFormatXMLDocument, LockComments:=False, Password:="",
                AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False,
                EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                :=False, SaveAsAOCELetter:=False, CompatibilityMode:=14)
        End If

        Application.Documents.Open(FileName:=NomDuFichierDOrigine, ConfirmConversions:=True, ReadOnly:=False, AddToRecentFiles:=True)

        Dim LargeurEcran As Integer
        Dim HauteurEcran As Integer
        LargeurEcran = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width * 0.75
        HauteurEcran = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height * 0.75

        Application.WindowState = Word.WdWindowState.wdWindowStateNormal

        Application.Documents.Item(NomDuFichierDOrigine).Select()
        Application.ActiveWindow.Width = LargeurEcran / 2
        Application.ActiveWindow.Height = HauteurEcran
        Application.ActiveWindow.Left = 1
        Application.ActiveWindow.Top = 1
        Application.ActiveDocument.
        Application.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:="Travail")

        Application.Documents.Item(NomDuFichier).Select()
        Application.ActiveWindow.Width = LargeurEcran / 2
        Application.ActiveWindow.Height = HauteurEcran
        Application.ActiveWindow.Left = LargeurEcran / 2
        Application.ActiveWindow.Top = 1
        Application.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:="Travail")

        Application.ScreenUpdating = True
        FormPatiente.Close()
    End Sub
    Sub MarquerCorrection()
        'Marque les formes et les formes en lignes ainsi que les tableaux en leur mettant le titre "Correction"
        Dim NomCourant As String
        Dim Position As Integer
        Dim IndexForme As Integer = 1
        Dim NombreDeForme As Integer = 1
        Dim Cpt As Integer = 0
        Dim NombreDeSousForme As Integer
        Dim NomDeLaSousForme As String
        Dim NomDeLaForme As String

        On Error Resume Next 'évite de gérer les erreurs

        ' une erreur se produit  s'il ne s'agit pas d'un tableau
        Application.Selection.Tables(1).Title = "Correction"

        ' une erreur se produit s'il ne s'agit pas d'une shape
        NombreDeForme = Application.Selection.ShapeRange.Count
        NombreDeSousForme = Application.Selection.ChildShapeRange.Count
        NomDeLaForme = Application.Selection.ShapeRange.Name
        NomDeLaSousForme = Application.Selection.ChildShapeRange.Name

        If NombreDeSousForme = 0 Then
            For Cpt = 1 To NombreDeForme
                Application.Selection.ShapeRange(Cpt).Title = "Correction"
            Next
        Else
            For Cpt = 1 To NombreDeSousForme
                Application.Selection.ChildShapeRange.Item(Cpt).Title = "Correction"
            Next
        End If

        ' une erreur se produit  s'il ne s'agit pas d'une inlineshape
        NombreDeForme = Application.Selection.InlineShapes.Count
        For Cpt = 1 To NombreDeForme
            Application.Selection.InlineShapes(Cpt).Title = "Correction"
        Next Cpt
    End Sub
    Sub DemarquerCorrection()
        'DéMarque les formes et les formes en leur supprimant le titre
        Dim NomCourant As String
        Dim Position As Integer
        Dim IndexForme As Integer = 1
        Dim NombreDeForme As Integer = 1
        Dim Cpt As Integer = 0
        Dim NombreDeSousForme As Integer
        Dim NomDeLaSousForme As String
        Dim NomDeLaForme As String

        On Error Resume Next 'évite de gérer les erreurs

        ' une erreur se produit  s'il ne s'agit pas d'un tableau
        Application.Selection.Tables(1).Title = ""

        ' une erreur se produit s'il ne s'agit pas d'une shape
        NombreDeForme = Application.Selection.ShapeRange.Count
        NombreDeSousForme = Application.Selection.ChildShapeRange.Count
        NomDeLaForme = Application.Selection.ShapeRange.Name
        NomDeLaSousForme = Application.Selection.ChildShapeRange.Name

        If NombreDeSousForme = 0 Then
            For Cpt = 1 To NombreDeForme
                Application.Selection.ShapeRange(Cpt).Title = ""
            Next
        Else
            For Cpt = 1 To NombreDeSousForme
                Application.Selection.ChildShapeRange.Item(Cpt).Title = ""
            Next
        End If

        ' une erreur se produit  s'il ne s'agit pas d'une inlineshape
        NombreDeForme = Application.Selection.InlineShapes.Count
        For Cpt = 1 To NombreDeForme
            Application.Selection.InlineShapes(Cpt).Title = ""
        Next Cpt
    End Sub
    Sub InsèreZoneDeTexteArrondi()
        Dim PosX, PosY, LargeurPage, MargeDroite, Largeur As Integer
        Dim NomZoneDeTexte As String

        PosX = Application.Selection.Information(Word.WdInformation.wdHorizontalPositionRelativeToPage)
        PosY = Application.Selection.Information(Word.WdInformation.wdVerticalPositionRelativeToPage)

        LargeurPage = Application.ActiveDocument.PageSetup.PageWidth
        'MargeGauche = ActiveDocument.PageSetup.LeftMargin
        MargeDroite = Application.ActiveDocument.PageSetup.RightMargin
        Largeur = LargeurPage - MargeDroite - PosX
        'NomZoneDeTexte = ActiveDocument.Shapes.AddShape(msoShapeRoundedRectangle, Posx, Posy, Largeur, 30).Name
        NomZoneDeTexte = Application.ActiveDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, PosX, PosY, Largeur, 30).Name
        With Application.ActiveDocument.Shapes(NomZoneDeTexte)
            .ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset42
            .TextFrame.MarginBottom = Application.CentimetersToPoints(0.15)
            .TextFrame.MarginTop = Application.CentimetersToPoints(0.15)
            .TextFrame.MarginLeft = Application.CentimetersToPoints(0.25)
            .TextFrame.MarginRight = Application.CentimetersToPoints(0.25)
            .TextFrame.AutoSize = True
            .ConvertToInlineShape()
            .AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle
            .Title = "Correction"
            .Select()
        End With
    End Sub
    Sub MontrerAide()
        System.Diagnostics.Process.Start(RepModel & "\STI2D\STI2D.chm")
    End Sub
    Sub MontrerAidePDF()
        System.Diagnostics.Process.Start(RepModel & "\STI2D\STI2D.pdf")
    End Sub
    Private Function assem() As String
        Throw New NotImplementedException
    End Function
    Public Sub Initialisation()

        RepProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        NomAppli = My.Application.Info.ProductName
        CompanyAppli = My.Application.Info.CompanyName
        Dim VersionWord = Application.Version
        'Dim NuméroApplicationWord = Application.a
        'Dim NumUtilisateur = System.Security.Principal.WindowsIdentity.GetCurrent().User.Value

        '**** Test si on est en mode debuggage en cherchant la clef d'installation du complément
        Dim RegWordAddin As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\Word\Addins")

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\" & CompanyAppli & "\" & NomAppli, "Version", Nothing) Is Nothing Then
            '***Répertoire pour debug
            RepComplément = Environment.CurrentDirectory
            ModeDebugage = True
        Else
            '*** Répertoire pour version installée
            RepComplément = RepProgramFiles & "\" & CompanyAppli & "\" & NomAppli
            ModeDebugage = False
        End If

        '*** Détection du dossier de modèle personnalisés
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\" & VersionWord & "\Word\STI2D\Rep", "RepDeBase", Nothing) IsNot Nothing Then
            RepModel = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\" & VersionWord & "\Word\STI2D\Rep").GetValue("RepDeBase")
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\" & VersionWord & "\Word\Options", "PersonalTemplates", Nothing) IsNot Nothing Then
            RepModel = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\" & VersionWord & "\Word\Options").GetValue("PersonalTemplates")
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\" & VersionWord & "\Word\Options", "PersonalTemplates", Nothing) IsNot Nothing Then
            RepModel = Registry.CurrentUser.OpenSubKey("SOFTWARE\Policies\Microsoft\Office\" & VersionWord & "\Word\Options").GetValue("PersonalTemplates")
        Else
            If Val(VersionWord) < 16 Then
                RepModel = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Roaming\Microsoft\Templates"
            Else
                RepModel = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\Modèles Office personnalisés"
            End If
        End If
    End Sub

End Class
