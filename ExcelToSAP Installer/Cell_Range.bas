Attribute VB_Name = "Cell_Range"
Option Explicit
'====================================================================================
' MODULE      : Cell_Range
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module fournit des fonctions utilitaires robustes pour la
'               manipulation et l'analyse des plages de cellules Excel. La fonction
'               principale, `Last`, permet de déterminer de manière fiable la dernière
'               ligne, colonne ou cellule non vide dans une plage donnée.
'====================================================================================


'====================================================================================
' PROCÉDURE DE DÉMONSTRATION
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : g_LastRow_Example
' DESCRIPTION : Procédure de démonstration montrant comment utiliser la fonction `Last`
'               pour trouver la dernière ligne de la sélection actuelle et stocker
'               le résultat dans la variable globale `g_LastRow`.
'
'               NOTE : Cette procédure est un exemple et n'est pas destinée à être
'               utilisée directement dans le flux de production principal.
'------------------------------------------------------------------------------------
Sub g_LastRow_Example()
    Dim Rng As Range  ' Plage de travail

    ' Utilise la s?lection active comme plage
    Set Rng = Selection

    ' Appel de la fonction Last() pour d?terminer la derni?re ligne
    g_LastRow = Last(1, Rng) ' Le choix 1 correspond à la dernière ligne
End Sub


'====================================================================================
' FONCTION UTILITAIRE PRINCIPALE
'====================================================================================

'------------------------------------------------------------------------------------
' FONCTION    : Last
' DESCRIPTION : Renvoie la dernière ligne, colonne ou l'adresse de la dernière cellule
'               contenant des données dans une plage spécifiée.
'               Cette méthode est robuste car elle utilise `Range.Find` en recherche
'               inversée, ce qui la rend insensible aux cellules vides.
'
' PARAMÈTRES  :
'   - choice (Long) : Spécifie l'information à retourner.
'                       1 = Index de la dernière ligne (Long).
'                       2 = Index de la dernière colonne (Long).
'                       3 = Adresse de la dernière cellule (String, ex: "D25").
'   - Rng (Range)   : La plage de cellules à analyser.
'
' RETOUR      : Variant - Un Long (pour le choix 1 ou 2) ou un String (pour le choix 3).
'               Retourne 0 (pour ligne/colonne) ou l'adresse de la première cellule
'               de la plage si la plage est vide.
'------------------------------------------------------------------------------------
Function Last(choice As Long, Rng As Range) As Variant
    Dim lrw As Long      ' Derni?re ligne
    Dim lcol As Long     ' Derni?re colonne

    Select Case choice
        Case 1
            ' --- Trouve la dernière LIGNE non vide ---
            On Error Resume Next
            Last = Rng.Find(What:="*", _
                            After:=Rng.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row
            On Error GoTo 0 ' Désactive la gestion d'erreur pour ne pas masquer d'autres erreurs

        Case 2
            ' --- Trouve la dernière COLONNE non vide ---
            On Error Resume Next
            Last = Rng.Find(What:="*", _
                            After:=Rng.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0

        Case 3
            ' --- Trouve l'adresse de la dernière CELLULE non vide ---
            On Error Resume Next
            
            ' Trouve d'abord la dernière ligne
            lrw = Rng.Find(What:="*", _
                           After:=Rng.Cells(1), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).row
                           
            ' Puis trouve la dernière colonne
            lcol = Rng.Find(What:="*", _
                            After:=Rng.Cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
                            
            ' Si aucune erreur et que les résultats sont valides, construit l'adresse
            If Err.Number = 0 And lrw > 0 And lcol > 0 Then
                Last = Rng.Parent.Cells(lrw, lcol).Address(False, False)
            Else ' Si la plage est vide, retourne l'adresse de la première cellule
                Last = Rng.Cells(1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0
    End Select
End Function

