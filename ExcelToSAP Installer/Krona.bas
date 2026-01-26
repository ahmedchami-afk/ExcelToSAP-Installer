Attribute VB_Name = "Krona"
Option Explicit

Dim buffer() As String
Dim flen As Long
Dim unicode As Boolean
Dim fnum As Integer

Public Lgn_1, Col_1, LgnFin, ColFin As Integer
Public CharType, title, Series, DataX, DataY As String

Sub createChart()
    '=== Cr?e un graphique Krona ? partir du tableau Excel ===
    
    Dim datasetsByName As New Collection                  ' Collection associant chaque nom de dataset ? un identifiant unique
    Dim datasetNames As New Collection                    ' Collection stockant la liste ordonn?e des noms de datasets
    
    '// Variables Krona //
    Dim magName As String: magName = "Amount"             ' Nom par d?faut de la colonne Magnitude (valeur quantitative)
    Dim scoreName As String: scoreName = "Score"          ' Nom par d?faut de la colonne Score (valeur qualitative)
    Dim useScore As Boolean: useScore = False             ' Indique si une colonne de score est utilis?e dans les donn?es
    unicode = False                                       ' Indique si un encodage Unicode (UTF-16) est n?cessaire pour le fichier HTML
    
    '--- D?termination automatique de la zone de donn?es autour de la cellule active ---'
    If WorksheetFunction.CountA(ActiveCell.CurrentRegion) = 0 Then
        MsgBox "Aucune donn?e d?tect?e autour de la cellule active." & vbCrLf & _
               "Veuillez placer le curseur ? l?int?rieur d?un tableau de donn?es avant de lancer la macro.", _
               vbCritical, "Erreur : aucun dataset trouv?"
        Exit Sub
    End If
    Lgn_1 = ActiveCell.CurrentRegion.row                    ' Premi?re ligne du bloc de donn?es
    Col_1 = ActiveCell.CurrentRegion.Column                 ' Premi?re colonne du bloc de donn?es
    LgnFin = Lgn_1 + ActiveCell.CurrentRegion.Rows.count - 1 ' Derni?re ligne du bloc
    ColFin = Col_1 + ActiveCell.CurrentRegion.Columns.count - 1 ' Derni?re colonne du bloc
                                                         ' Trouve la derni?re colonne de donn?es (jusqu?? la droite)
    title = ActiveSheet.name                                      ' R?cup?re le titre du graphique dans la cellule B5
    
    '// Cr?ation du n?ud racine //
    Dim head As Node                                     ' D?clare un objet Node qui sera le n?ud racine de la hi?rarchie
    Set head = New Node                                  ' Cr?e une nouvelle instance de Node
    head.name = title                                    ' Donne au n?ud racine le titre du graphique comme nom
    
    Dim row As Integer                                   ' Variable compteur pour parcourir les lignes de donn?es
    Application.ActiveSheet.Cells(Lgn_1 + 1, Col_1).Select   ' S?lectionne la premi?re cellule de la zone de donn?es (juste sous l?en-t?te)
    row = 0                                              ' Initialise le compteur de ligne
    
    '=== Boucle principale sur les lignes ===
    Do Until IsEmpty(Cells(Lgn_1 + 1 + row, Col_1))      ' Boucle jusqu?? rencontrer une ligne vide dans la colonne du dataset
        
        Dim datasetName As String                        ' Variable pour stocker le nom du dataset de la ligne courante
        datasetName = Trim(Cells(Lgn_1 + 1 + row, Col_1).value)   ' Lit le nom du dataset (dans la premi?re colonne)
        
        If datasetName <> "" Then checkUnicode datasetName   ' V?rifie si le nom contient des caract?res non ASCII
        
        ' V?rifie coh?rence des noms de dataset
        If row > 0 Then                                       ' ? partir de la deuxi?me ligne, on v?rifie la coh?rence
            If (datasetName = "" And datasetNames(1) <> "") Or (datasetName <> "" And datasetNames(1) = "") Then
                Cells(Lgn_1 + 1 + row, 1).Select              ' S?lectionne la cellule fautive
                MsgBox "Dataset name must be provided for all rows if provided for any", , "Error"   ' Alerte si certains noms sont manquants
                Exit Sub                                      ' Arr?te l?ex?cution si incoh?rence
            End If
        End If
        
        ' Ajout du dataset dans la collection si nouveau
        If Not Exists(datasetsByName, datasetName) Then       ' V?rifie si le dataset n?est pas d?j? pr?sent dans la collection
            datasetsByName.add key:=datasetName, Item:=datasetsByName.count   ' L?ajoute avec un identifiant unique
            datasetNames.add datasetName                      ' Ajoute le nom ? la liste ordonn?e
        End If
        
        ' Ajoute les donn?es ? la hi?rarchie
        head.add row, Col_1 + 1, datasetsByName.Item(datasetName)   ' Ajoute la ligne au n?ud racine en commen?ant ? la colonne suivante (hi?rarchie)
        
    ' V?rifie s?il existe un score valide dans la colonne avant-derni?re
    If Not IsEmpty(Cells(Lgn_1 + 1 + row, ColFin - 1)) _
       And IsNumeric(Cells(Lgn_1 + 1 + row, ColFin - 1).value) Then
        useScore = True                                             ' Active le mode "Score" uniquement si la cellule contient une valeur num?rique
    End If
        
        row = row + 1                                               ' Passe ? la ligne suivante
    Loop
    
    '// Cr?ation du fichier HTML Krona //
    Dim fileName As String                                          ' Nom complet du fichier de sortie
    
    ' D?finit automatiquement le nom du fichier HTML dans le dossier du classeur actif
    fileName = ThisWorkbook.path & Application.PathSeparator & "index.html"
    
    ' R?cup?re les noms des colonnes pour magnitude et score
    If Not IsEmpty(Cells(Lgn_1 + row, ColFin - 1)) _
       And IsNumeric(Cells(Lgn_1 + row, ColFin - 1).value) Then                  ' Si l?avant-derni?re colonne a un en-t?te
        scoreName = Cells(Lgn_1, ColFin - 1).value                  ' Utilise ce texte comme nom du champ Score
    End If
    
    If Not IsEmpty(Cells(Lgn_1 + row, ColFin)) _
       And IsNumeric(Cells(Lgn_1 + row, ColFin).value) Then                       ' Si la derni?re colonne a un en-t?te
        magName = Cells(Lgn_1, ColFin).value                          ' Utilise ce texte comme nom du champ Magnitude
    End If
    
    checkUnicode magName                                            ' V?rifie la compatibilit? Unicode pour Magnitude
    checkUnicode scoreName                                          ' V?rifie la compatibilit? Unicode pour Score
    
    If fileName <> "" Then                                          ' Si un nom de fichier valide a ?t? s?lectionn?
        DeleteFile fileName                                         ' Supprime le fichier existant s?il existe d?j?
        fnum = FreeFile()                                           ' R?cup?re un num?ro de fichier libre
        
        If unicode Then                                             ' Si du texte Unicode est n?cessaire
            Open fileName For Binary Access Write As fnum           ' Ouvre le fichier en mode binaire
            Put #fnum, 1, &HFF                                      ' ?crit l?indicateur BOM UTF-16 (premier octet)
            Put #fnum, 2, &HFE                                      ' ?crit l?indicateur BOM UTF-16 (deuxi?me octet)
            Put #fnum, 3, 32                                        ' ?crit un espace pour initialiser le flux
        Else
            Open fileName For Output As fnum                        ' Sinon, ouvre le fichier en mode texte normal (UTF-8 / ANSI)
        End If
        
        writeHeader head, fnum, magName, useScore, scoreName, datasetNames   ' ?crit l?en-t?te HTML et les m?tadonn?es Krona
        head.writeChart magName, useScore, scoreName               ' ?crit le contenu du graphique sous forme XML (arborescence des n?uds)
        WriteFooter                                                ' ?crit le pied de page HTML (balises de fermeture)
        Close #fnum                                                ' Ferme le fichier
    End If
    
    '--- Ouvre automatiquement le fichier cr?? ---'
    OuvrirFichier fileName                                         ' Ouvre le fichier HTML dans le navigateur par d?faut
End Sub

Sub clearChart()
    ' Efface automatiquement le tableau (classique, structur? ou crois? dynamique)
    ' trouv? ? la cellule active (ActiveCell)

    Dim ws As Worksheet
    Dim cell As Range
    Dim lo As ListObject
    Dim pt As PivotTable
    Dim msg As String
    Dim title As String

    Set ws = ActiveSheet
    Set cell = ActiveCell
    msg = "Effacer le tableau d?tect? ? la cellule active ? (Op?ration irr?versible)"
    title = "Krona - Effacement"

    ' Demande confirmation ? l?utilisateur
    If MsgBox(msg, vbOKCancel + vbExclamation, title) = vbCancel Then Exit Sub

    On Error Resume Next

    ' --- Cas 1 : la cellule appartient ? un tableau structur? (ListObject)
    For Each lo In ws.ListObjects
        If Not Intersect(cell, lo.Range) Is Nothing Then
            lo.Unlist                        ' Convertit le tableau structur? en plage
            lo.Range.Clear                   ' Efface tout le contenu et la mise en forme
            MsgBox "Tableau structur? effac?.", vbInformation, title
            Exit Sub
        End If
    Next lo

    ' --- Cas 2 : la cellule appartient ? un tableau crois? dynamique
    For Each pt In ws.PivotTables
        If Not Intersect(cell, pt.TableRange2) Is Nothing Then
            pt.TableRange2.Clear             ' Efface tout le tableau crois?
            MsgBox "Tableau crois? dynamique effac?.", vbInformation, title
            Exit Sub
        End If
    Next pt

    ' --- Cas 3 : sinon, efface la plage courante
    If Not cell.CurrentRegion Is Nothing Then
        cell.CurrentRegion.Clear             ' Efface le contenu et la mise en forme du bloc
        MsgBox "Plage de donn?es effac?e.", vbInformation, title
    Else
        MsgBox "Aucune table d?tect?e ? la cellule active.", vbInformation, title
    End If

    On Error GoTo 0
End Sub



'------------------------------------------------------------
' Fonction : getFileName
' Objectif : Demande ? l'utilisateur de choisir o? enregistrer
'             le fichier HTML g?n?r? (fichier Krona)
'------------------------------------------------------------
Function getFileName()
   
    #If Mac Then
        '--- Partie sp?cifique ? MacOS ---
        Dim fileName
        
        ' Ouvre une bo?te de dialogue Mac pour choisir le nom de fichier
        ' et renvoie le chemin complet du fichier s?lectionn?.
        fileName = MacScript("try" & Chr(13) & _
            "(choose file name default name ""krona.html"") as text" & Chr(13) & _
            "on error" & Chr(13) & """""" & Chr(13) & "end try")
        
        ' Si le nom de fichier est vide, on sort ; sinon, on ajoute ".html" s?il manque.
        If fileName <> "" And Right(fileName, 5) <> ".html" Then
            fileName = fileName & ".html"
        End If
        
        ' Renvoie le nom complet choisi par l?utilisateur.
        getFileName = fileName
        
    #Else
        '--- Partie sp?cifique ? Windows ---
        Dim chemin As String
        
        ' Ouvre une bo?te de dialogue standard "Enregistrer sous" pour un fichier HTML.
        getFileName = Application.GetSaveAsFilename( _
            FileFilter:="HTML files (*.html), *.html")

    #End If

End Function

'------------------------------------------------------------
' Fonction : Exists
' Objectif : V?rifie si une cl? existe dans une collection VBA
' Entr?e : coll (Collection), key (String)
' Sortie : True si la cl? existe, False sinon
'------------------------------------------------------------
Function Exists(coll As Collection, key As String) As Boolean
    On Error GoTo EH                  ' Si erreur (cl? inexistante), aller ? EH
    coll.Item key                     ' Essaie d'acc?der ? l'?l?ment
    Exists = True                     ' Si pas d'erreur ? la cl? existe
EH:                                   ' En cas d?erreur, on sort silencieusement
End Function

'------------------------------------------------------------
' Sub : DeleteFile
' Objectif : Supprime un fichier s?il existe
' Entr?e : fileName (String)
'------------------------------------------------------------
Sub DeleteFile(fileName As String)
    On Error GoTo EH                  ' Ignore les erreurs (par ex. si le fichier n?existe pas)
    Kill fileName                     ' Supprime le fichier sp?cifi?
EH:
End Sub

'------------------------------------------------------------
' Sub : WriteFooter
' Objectif : ?crit la fin du fichier HTML Krona (balises fermantes)
'------------------------------------------------------------
Private Sub WriteFooter()
    printOut "    </krona></div>"     ' Ferme les balises principales du contenu Krona
    printOut "    </body>"            ' Ferme la balise <body>
    printOut "</html>"                ' Ferme la balise <html>
End Sub

'------------------------------------------------------------
' Sub : writeHeader
' Objectif : ?crit le d?but du fichier HTML Krona
'             (balises <html>, <head>, scripts, images, etc.)
'------------------------------------------------------------
Private Sub writeHeader( _
    head As Node, fnum As Integer, _
    magName As String, useScore As Boolean, _
    scoreName As String, datasetNames As Collection)
    
    '--- D?but du document HTML (d?claration et ent?te) ---
    printOut "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
    printOut "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">"
    printOut "    <head>"
    
    ' D?finition de l?encodage en fonction de la variable globale unicode
    If unicode Then
        printOut "        <meta charset=""UTF-16""/>"
    Else
        printOut "        <meta charset=""UTF-8""/>"
    End If
    
    printOut "        "
    
    '--- Acc?s ? la feuille "images" contenant les ressources visuelles ---
    Sheets("images").Visible = True
    Sheets("images").Select
    
    ' Ajoute la favicon (petite ic?ne de l?onglet du navigateur)
    printOut "        <link rel=""shortcut icon"" href=""" & Range("B1").value & """/>"
    
    ' Message d?erreur si le script principal ne peut pas ?tre charg?
    printOut "        <script id=""notfound"">window.onload=function(){document.body.innerHTML=""Could not get resources from \""http://krona.sourceforge.net\"".""}</script>"
    
    ' D?but du script JavaScript principal
    printOut "        <script>"
    
    '--- Lecture du contenu de la feuille "krona-2.0.js" ---
    Sheets("krona-2.0.js").Visible = True
    Sheets("krona-2.0.js").Select
    
    ' Recherche de la derni?re ligne utilis?e dans la colonne A
    Dim g_LastRow As Long
    With ActiveSheet
        g_LastRow = .Cells(.Rows.count, "A").End(xlUp).row
    End With
    
    ' S?lection de tout le script JavaScript dans la colonne A
    Range("A1", "A" & g_LastRow).Select
    
    ' ?crit chaque ligne du script dans le fichier HTML
    Dim cell As Range
    For Each cell In Application.Selection.Cells
        printOut cell.value
    Next
    
    ' Cache la feuille du script apr?s utilisation
    Sheets("krona-2.0.js").Visible = False
    
    ' Termine la balise <script> et la section <head>
    printOut "</script>"
    printOut "    </head>"
    printOut "    <body>"
    
    '--- Section <body> : insertion d?images cach?es et structure du graphique ---
    Sheets("images").Select
    
    ' Images cach?es (utilis?es par Krona pour le rendu visuel)
    printOut "        <img id=""hiddenImage"" src=""" & Range("B2").value & """  style=""display:none""/>"
    printOut "        <img id=""loadingImage"" src=""" & Range("B3").value & """  style=""display:none""/>"
    printOut "        <img id=""logo"" src=""" & Range("B4").value & Range("B5").value & Range("B6").value & """  style=""display:none""/>"
    
    ' Message affich? si JavaScript est d?sactiv?
    printOut "        <noscript>Javascript must be enabled to view this page.</noscript>"
    
    ' Balises ouvrantes du conteneur Krona
    printOut "        <div style=""display:none""><krona>"
    
    ' D?finition des attributs du graphique Krona (magnitude, score)
    printOut "            <attributes magnitude=""magnitude"">"
    printOut "                <attribute display=""" & magName & """>magnitude</attribute>"
    
    ' Si l?utilisation du score est activ?e, on ajoute cet attribut
    If useScore Then
        printOut "               <attribute display=""" & scoreName & """>score</attribute>"
    End If
    
    printOut "            </attributes>"
    
    '--- D?finition de la coloration si score utilis? ---
    If useScore Then
        printOut "            <color attribute=""score"" valueStart=""" & head.scoreMin() & _
                  """ valueEnd=""" & head.scoreMax() & """ hueStart=""0"" hueEnd=""120""></color>"
    End If
    
    ' Retour ? la feuille principale apr?s inclusion des ressources
    Sheets("images").Visible = False
    Sheets(title).Select
    Range("A1").Select
    
    '--- Si plusieurs datasets existent, les ?crire dans la section <datasets> ---
    If datasetNames.count > 0 And datasetNames(1) <> "" Then
        printOut "            <datasets>"
        Dim i As Integer
        
        For i = 1 To datasetNames.count
            printOut "               <dataset>" & datasetNames(i) & "</dataset>"
        Next
        
        printOut "            </datasets>"
    End If
    
End Sub

'------------------------------------------------------------
' Fonction : printOut
' Objectif : ?crit une ligne dans le fichier HTML, en tenant compte
'             du mode d?encodage (UTF-8 ou UTF-16)
' Entr?e : str (String)
'------------------------------------------------------------
Function printOut(str As String)
    If unicode Then
        str = str & ChrW(10)                  ' Ajoute un saut de ligne Unicode
        Dim data() As Byte
        data = str                            ' Convertit la cha?ne en tableau de bytes
        Put #fnum, LOF(fnum) + 1, data        ' ?crit les donn?es binaires dans le fichier
    Else
        Print #fnum, str                      ' ?crit la ligne en ASCII standard
    End If
End Function

'------------------------------------------------------------
' Sub : checkUnicode
' Objectif : V?rifie si une cha?ne contient des caract?res non ASCII
'             pour d?cider si le fichier doit ?tre ?crit en Unicode
' Entr?e : str (String)
'------------------------------------------------------------
Sub checkUnicode(str As String)
    If unicode = True Then Exit Sub           ' Si d?j? Unicode, inutile de rev?rifier
    
    Dim j As Integer
    Dim x() As Byte
    
    x = str                                   ' Convertit la cha?ne en tableau d?octets
    For j = 0 To UBound(x)
        If x(j) > 127 Then                    ' Si un caract?re d?passe ASCII (127)...
            unicode = True                    ' ...on bascule en mode Unicode
            Exit Sub
        End If
    Next
End Sub

'------------------------------------------------------------
' Fonction : OuvrirFichier
' Objectif : Ouvre le fichier HTML g?n?r? dans son application par d?faut
' Source : https://excel-malin.com
'------------------------------------------------------------
Public Function OuvrirFichier(MonFichier As String)
   
On Error GoTo OuvertureFichierErreur
   
   ' V?rifie si le fichier existe
   If Len(Dir(MonFichier)) = 0 Then
      OuvrirFichier = False
      Exit Function
   End If
   
   ' Cr?e un objet Shell pour ouvrir le fichier avec l'application associ?e
   Dim MonApplication As Object
   Set MonApplication = CreateObject("Shell.Application")
   
   MonApplication.Open (MonFichier)           ' Ouvre le fichier
   OuvrirFichier = True                       ' Retourne True si succ?s
   Set MonApplication = Nothing               ' Lib?re l?objet Shell
   
Exit Function

' Gestion des erreurs d?ouverture
OuvertureFichierErreur:
   Set MonApplication = Nothing
   OuvrirFichier = False
End Function

