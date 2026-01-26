Private WithEvents App As Application

Private Sub Workbook_Open()
    Set App = Application
    auto_open
End Sub

Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrHandler
    
    ' 1. Limiter la portée : Ne pas exécuter le code si le clic droit a lieu
    '    dans le classeur de l'add-in lui-même (ex: sur la feuille "Setup" visible).
    If Sh.Parent Is ThisWorkbook Then Exit Sub
    
    ' 2. Pertinence : S'assurer que le menu n'apparaît que sur les feuilles de calcul.
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    ' D'abord, on nettoie les anciens menus personnalisés pour éviter les doublons.
    Call modMain.CleanupContextMenu
    
    ' Ensuite, on crée le nouveau menu dynamique en fonction de la cellule cliquée.
    Call modMain.CreateDynamicContextMenu(Target)
    Exit Sub

ErrHandler:
    ' En cas d'erreur critique, on s'assure que le menu est nettoyé et on ne plante pas Excel.
    On Error Resume Next
    Call modMain.CleanupContextMenu
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    '------------------------------------------------------------------------------------
    ' PROCÉDURE   : Workbook_SheetDeactivate
    ' DESCRIPTION : Se déclenche lorsqu'une feuille est désactivée. Si c'est la
    '               feuille "Help", elle est automatiquement masquée.
    '------------------------------------------------------------------------------------
    If Sh.name = "Help" Then
        Sh.Visible = xlSheetVeryHidden
    End If
End Sub
