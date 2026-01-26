VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' Classe : Node
' Objectif :
'   Repr?sente un n?ud dans la structure hi?rarchique du graphique Krona.
'   Chaque Node peut contenir :
'     - un nom (cat?gorie)
'     - une magnitude (quantit?)
'     - un score (valeur associ?e)
'     - des enfants (sous-n?uds)
'------------------------------------------------------------

Option Explicit    ' Oblige la d?claration explicite de toutes les variables

'--- Variables publiques et priv?es ---
Public name As String             ' Nom du n?ud courant

Private magnitude() As Double     ' Tableau des magnitudes (une par dataset)
Private score() As Double         ' Tableau des scores (m?me dimension que magnitude)
Private html As String            ' (non utilis? ici ? peut servir ? stocker une portion HTML)
Private children() As Node        ' Tableau dynamique de sous-n?uds (objets Node)

'------------------------------------------------------------
' Sub : add
' Objectif :
'   Ajoute une ligne de donn?es dans la hi?rarchie (dans ce n?ud ou un sous-n?ud).
'   Construit r?cursivement la structure ? partir des donn?es Excel.
'------------------------------------------------------------
Public Sub add(row As Integer, col As Integer, dataset As Integer)
    ' Ajoute une ligne ? la hi?rarchie, selon la structure :
    ' Col1 = Dataset | Col2...ColFin-2 = hi?rarchie | ColFin-1 = Score | ColFin = Magnitude
    
    If dataset > UBound(magnitude) Then                    ' Si le tableau magnitude n?a pas assez d??l?ments pour ce dataset
        ReDim Preserve magnitude(0 To dataset)             ' ?tend le tableau magnitude pour inclure cet index
        ReDim Preserve score(0 To dataset)                 ' ?tend le tableau score de la m?me mani?re
    End If
    
    Dim newMagnitude As Double                             ' Variable pour stocker la magnitude de la ligne
    '--- Lecture de la magnitude (derni?re colonne) ---'
    If Not IsEmpty(Cells(Lgn_1 + 1 + row, ColFin)) _
       And IsNumeric(Cells(Lgn_1 + 1 + row, ColFin).value) Then
           
        ' Si la cellule n?est pas vide et contient un nombre ? on l?utilise
        newMagnitude = Cells(Lgn_1 + 1 + row, ColFin).value        ' Lit la valeur num?rique de magnitude
    Else
    ' Sinon (vide ou non num?rique) ? on garde la valeur par d?faut
    newMagnitude = 1                                           ' Valeur par d?faut si pas de nombre valide
    End If

    ' Ajoute la magnitude au total du dataset correspondant
    magnitude(dataset) = magnitude(dataset) + newMagnitude
    
    ' Ajoute le score si pr?sent
    If Not IsEmpty(Cells(Lgn_1 + 1 + row, ColFin - 1)) _
       And IsNumeric(Cells(Lgn_1 + 1 + row, ColFin - 1).value) Then
                                                                  ' V?rifie si une valeur de score existe (avant-derni?re colonne)
        score(dataset) = score(dataset) + _
            Cells(Lgn_1 + 1 + row, ColFin - 1).value * _
            newMagnitude * -1                                     ' Multiplie le score par la magnitude (pond?ration) et l?ajoute (ici invers?)
    End If
    
    '--- Si on n?a pas encore atteint la derni?re colonne hi?rarchique ---'
    If col <= ColFin - 2 Then                                     ' Tant qu?il reste des niveaux de hi?rarchie ? explorer
        Dim newName As String
        newName = Trim(Cells(Lgn_1 + 1 + row, col).value)         ' Lit le nom du n?ud ? ce niveau hi?rarchique
        'If newName = "" Then Exit Sub                            ' Si la cellule est vide, on arr?te la r?cursion
        
        Dim targetChild As Long
        targetChild = -1                                          ' Initialise l?indice du sous-n?ud cible ? ?non trouv?
        
        ' Recherche d?un enfant existant avec le m?me nom
        Dim i As Long
        For i = 0 To UBound(children)                             ' Parcourt tous les sous-n?uds existants
            If Not children(i) Is Nothing Then
                If children(i).name = newName Then                 ' Si un sous-n?ud porte d?j? ce nom
                    targetChild = i                                ' On le r?utilise
                    Exit For
                End If
            End If
        Next
        
        ' Si aucun sous-n?ud existant ne correspond, on en cr?e un nouveau
        If targetChild = -1 Then
            If Not children(0) Is Nothing Then                     ' Si le tableau des enfants contient d?j? au moins un ?l?ment
                ReDim Preserve children(0 To UBound(children) + 1) ' On agrandit le tableau pour ajouter un nouvel enfant
            End If
            
            Set children(UBound(children)) = New Node              ' On cr?e un nouveau n?ud enfant
            targetChild = UBound(children)                         ' On r?cup?re son index
            children(targetChild).name = newName                   ' On lui attribue le nom du niveau hi?rarchique lu dans la cellule
        End If
        
        ' Appel r?cursif : ajoute la ligne dans le sous-n?ud suivant
        children(targetChild).add row, col + 1, dataset            ' Passe ? la colonne suivante (niveau inf?rieur de hi?rarchie)
    End If
End Sub



'------------------------------------------------------------
' Fonction : scoreMax
' Objectif :
'   Calcule le score moyen maximum parmi tous les sous-n?uds.
'   Utilis? pour d?finir les couleurs dans Krona.
'------------------------------------------------------------
Public Function scoreMax() As Double
    Dim i As Integer
    
    '--- Recherche du score maximum dans ce n?ud ---
    For i = 0 To UBound(score)
        If magnitude(i) > 0 Then
            If Not IsNumeric(scoreMax) Or score(i) / magnitude(i) > scoreMax Then
                scoreMax = score(i) / magnitude(i)
            End If
        End If
    Next
    
    '--- Recherche r?cursive dans les enfants ---
    Dim child As Variant
    For Each child In children
        If Not child Is Nothing Then
            Dim maxChild As Double
            maxChild = child.scoreMax()
            
            If maxChild > scoreMax Then
                scoreMax = maxChild
            End If
        End If
    Next
End Function

'------------------------------------------------------------
' Fonction : scoreMin
' Objectif :
'   Calcule le score moyen minimum parmi tous les sous-n?uds.
'   (M?me logique que scoreMax, mais cherche la plus petite valeur)
'------------------------------------------------------------
Public Function scoreMin() As Double
    
    Dim i As Integer
    
    '--- Recherche du score minimum pour ce n?ud ---
    For i = 0 To UBound(score)
        If magnitude(i) > 0 Then
            If Not IsNumeric(scoreMin) Or score(i) / magnitude(i) < scoreMin Then
                scoreMin = score(i) / magnitude(i)
            End If
        End If
    Next
    
    '--- Recherche r?cursive dans les enfants ---
    Dim child As Variant
    For Each child In children
        If Not child Is Nothing Then
            Dim minChild As Double
            minChild = child.scoreMin()
            
            If minChild < scoreMin Then
                scoreMin = minChild
            End If
        End If
    Next
End Function

'------------------------------------------------------------
' Sub : writeChart
' Objectif :
'   D?marre la g?n?ration XML du graphique Krona.
'   Appelle writeXML (r?cursive) pour ?crire chaque n?ud.
'------------------------------------------------------------
Public Sub writeChart(magName As String, useScore As Boolean, scoreName As String)
    writeXML useScore    ' Appelle la m?thode interne writeXML
End Sub

'------------------------------------------------------------
' Sub : writeXML
' Objectif :
'   ?crit la repr?sentation XML du n?ud et de ses sous-n?uds
'   dans le fichier HTML via la fonction printOut.
'   Cette proc?dure est r?cursive.
'------------------------------------------------------------
Public Sub writeXML(useScore As Boolean)
    
    Dim i As Integer
    
    '--- Ouvre une balise <node> avec le nom du n?ud courant ---
    printOut "<node name=""" & name & """><magnitude>"
    
    '--- ?crit la magnitude de chaque dataset ---
    For i = 0 To UBound(magnitude)
        printOut "<val>" & str$(magnitude(i)) & "</val>"
    Next
    
    printOut "</magnitude>"
    
    '--- Si les scores sont utilis?s, on ?crit aussi les valeurs ---
    If useScore Then
        printOut "<score>"
        For i = 0 To UBound(score)
            Dim scoreAve As Double
            
            ' Calcul du score moyen (score total / magnitude)
            If magnitude(i) > 0 Then
                scoreAve = score(i) / magnitude(i)
            End If
            
            printOut "<val>" & str$(scoreAve) & "</val>"
        Next
        
        printOut "</score>"
    End If
    
    '--- Parcours r?cursif des enfants ---
    Dim child As Variant
    For Each child In children
        If Not child Is Nothing Then
            child.writeXML useScore
        End If
    Next
    
    '--- Ferme la balise du n?ud ---
    printOut "</node>"
    
End Sub

'------------------------------------------------------------
' Sub : class_Initialize
' Objectif :
'   Initialise un nouvel objet Node avec des tableaux pr?ts ? l?emploi.
'   Appel? automatiquement ? la cr?ation (Set x = New Node).
'------------------------------------------------------------
Private Sub class_Initialize()
    ' Initialise les tableaux avec une seule valeur (index 0)
    ReDim magnitude(0 To 0) As Double
    ReDim score(0 To 0) As Double
    ReDim children(0 To 0) As Node
    
    ' Valeurs initiales = 0
    magnitude(0) = 0
    score(0) = 0
End Sub

