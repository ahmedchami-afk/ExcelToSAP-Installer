Attribute VB_Name = "SAP"
'================================================================================
' MODULE      : SAP.bas (Gestion des interactions SAP)
' AUTEUR      : Votre Nom / Gemini Code Assist
' DATE        : 26/10/2025 (Dernière modification : 02/12/2025)
' DESCRIPTION : Ce module gère l'ensemble des interactions avec l'application SAP GUI.
'               Module principal de gestion de la session SAP.
'               Contient les routines de connexion/deconnexion ("onSAP", "offSAP")
'               et les wrappers de lancement des transactions SAP (Z_IW33, Z_IE03, etc.).
'
' DEPENDANCES :
'   - module1.bas (pour les variables globales g_Session, g_DataType, doNotRun)
'   - Feuille "Setup" (pour lire les configurations "Exe_Mode", "LAY_...", "ppf", "mpf")
'================================================================================

'================================================================================
' SECTION : VARIABLES PUBLIQUES
'================================================================================
Option Explicit

Public StatusBar As Object ' Objet pour interagir avec la barre de statut SAP

'================================================================================
' SECTION : DÉCLARATIONS API (Presse-papiers Windows)
'================================================================================

#If VBA7 Then
    ' Pour les versions 32 bits et 64 bits de VBA
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowTextLengthA Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowTextA Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
#Else
    ' Pour les versions 32 bits de VBA
    'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    'Private Declare Function EmptyClipboard Lib "user32" () As Long
    'Private Declare Function CloseClipboard Lib "user32" () As Long
#End If

'================================================================================
' SECTION : UTILITAIRES (Fonctions liées au Presse-papiers)
'================================================================================

Sub ViderPressePapier()
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : Vide le contenu du presse-papiers Windows.
    ' APPELÉ PAR  : offSAP()
    '--------------------------------------------------------------------------------
    ' Ouvre le presse-papiers. Le handle de la fenêtre est 0 pour l'application courante.
    OpenClipboard 0&

    ' Vider le presse-papiers
    EmptyClipboard

    ' Fermer le presse-papiers
    CloseClipboard
End Sub

'--------------------------------------------------------------------------------
' PROCÉDURE   : WaitForSAP
' DESCRIPTION : Met en pause l'exécution tant que la session SAP est occupée (Busy).
'               Remplace Application.Wait pour une attente dynamique.
'               Intègre un timeout de sécurité pour éviter les boucles infinies.
'--------------------------------------------------------------------------------
Public Sub WaitForSAP(Optional ByVal timeOut As Double = 10)
    Dim startTime As Double
    startTime = Timer
    
    Do
        DoEvents ' Garde Excel réactif
        
        On Error Resume Next
        ' Si g_Session est perdu ou vide, ou si SAP n'est pas occupé, on sort
        If g_Session Is Nothing Then Exit Do
        If g_Session.Busy = False Then Exit Do
        If Err.Number <> 0 Then Exit Do ' Erreur d'accès à l'objet
        On Error GoTo 0
        
        ' Timeout de sécurité (gestion basique du passage à minuit incluse via Abs)
        If Abs(Timer - startTime) > timeOut Then Exit Do
    Loop
End Sub

'--------------------------------------------------------------------------------
' PROCÉDURE   : FillSAPSelectionList
' DESCRIPTION : Remplit une liste de sélection multiple SAP (Select-Options)
'               directement depuis une plage Excel, sans utiliser le presse-papiers.
'               Gère la pagination du Table Control.
'               Permet l'exclusion de valeurs et un contrôle fin sur l'ouverture/fermeture de la popup.
'--------------------------------------------------------------------------------
Public Sub FillSAPSelectionList(ByVal buttonId As String, ByVal dataRange As Range, Optional ByVal Exclude As Boolean = False, Optional ByVal openPopup As Boolean = True, Optional ByVal clearList As Boolean = True, Optional ByVal closePopup As Boolean = True)
    On Error GoTo ErrHandler
    
    ' 1. Ouvrir la fenêtre de sélection multiple
    If openPopup Then
        g_Session.findById(buttonId).press
        WaitForSAP
    End If
    
    ' Vérifier qu'on est bien sur une popup (wnd[1])
    If g_Session.ActiveWindow.name <> "wnd[1]" Then Exit Sub
    
    ' --- OPTIMISATION 1 : Sélection du bon onglet (Inclusion vs Exclusion) ---
    ' Cela évite de devoir gérer le signe "E" ligne par ligne dans la boucle
    Dim tabID As String
    If Exclude Then
        tabID = "NOSV" ' Onglet Exclusion Valeurs Simples
    Else
        tabID = "SIVA" ' Onglet Inclusion Valeurs Simples
    End If
    
    On Error Resume Next
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabp" & tabID).Select
    If Err.Number <> 0 Then
        tabID = "SIVA" ' Fallback sur l'onglet par défaut si l'onglet cible n'existe pas
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' 2. Identifier le Table Control standard
    ' Note : L'ID est standard pour la plupart des rapports (Logical Database)
    Dim tableID As String
    tableID = "wnd[1]/usr/tabsTAB_STRIP/tabp" & tabID & "/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
    
    Dim sapTable As Object
    On Error Resume Next
    Set sapTable = g_Session.findById(tableID)
    
    ' 3. Vider la liste existante
    If clearList Then
        g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
        If Not sapTable Is Nothing Then Set sapTable = g_Session.findById(tableID) ' Rafraîchir l'objet après le roundtrip serveur
    End If

    If Not dataRange Is Nothing Then
        ' Optimisation : Restreindre la plage aux cellules utilisées pour éviter de boucler sur des colonnes entières
        Set dataRange = Intersect(dataRange, dataRange.Parent.UsedRange)
    End If

    If Not dataRange Is Nothing Then
    On Error GoTo ErrHandler
    
    If sapTable Is Nothing Then
        ' Fallback : Si la table n'est pas trouvée, on tente le presse-papiers standard
        dataRange.Copy
        g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
        GoTo Finalize
    End If
    
    ' --- OPTIMISATION 2 : Utilisation du presse-papiers (Méthode rapide) ---
    ' Cette méthode réduit la charge sur SAP de N appels à 1 seul appel.
    On Error Resume Next
    dataRange.Copy
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
    If Err.Number = 0 Then
        ' Si le collage a réussi, on valide et on sort immédiatement
        On Error GoTo ErrHandler
        GoTo Finalize
    End If
    
    ' --- OPTIMISATION 3 : Boucle de repli (Méthode robuste) ---
    ' Si le presse-papiers échoue, on utilise la boucle mais optimisée par le choix de l'onglet
    Err.Clear
    On Error GoTo ErrHandler
    
    ' Trouver la fin si on n'a pas vidé (si on a vidé, sapRowIndex est 0)
    Dim sapRowIndex As Long ' Index de la ligne visible dans la grille SAP (0-based)
    sapRowIndex = 0

    If Not clearList Then
        ' Mode ajout : trouve la première ligne vide sur la page visible.
        ' NOTE : Pour la stabilité, cette version ne scrolle pas pour trouver une ligne vide.
        ' Si la première page est pleine, l'ajout commencera à la première ligne.
        Dim i As Long
        For i = 0 To sapTable.VisibleRowCount - 1
            If g_Session.findById(tableID & "/ctxtRSCSEL_255-SLOW_I[1," & i & "]").text = "" Then
                sapRowIndex = i
                Exit For
            End If
        Next i
    End If
    
    ' 4. Boucler sur les cellules
    Dim cell As Range
    Dim cellValue As String
    Dim visibleRows As Long
    visibleRows = sapTable.VisibleRowCount
    
    For Each cell In dataRange
        cellValue = CStr(cell.value)
        If cellValue <> "" Then
            ' Si la page visible est pleine, on passe à la page suivante
            If sapRowIndex >= visibleRows Then
                g_Session.findById("wnd[1]").sendVKey 82 ' VKey for Page Down
                WaitForSAP
                sapRowIndex = 0 ' Réinitialise l'index pour la nouvelle page
            End If
            
            ' Injection de la valeur (colonne 1 = Low value)
            ' On utilise l'ID direct de la cellule visible
            g_Session.findById(tableID & "/ctxtRSCSEL_255-SLOW_I[1," & sapRowIndex & "]").text = cellValue
            
            ' Gestion de l'exclusion : Uniquement nécessaire si on n'a pas pu changer d'onglet
            If Exclude And tabID = "SIVA" Then
                g_Session.findById(tableID & "/ctxtRSCSEL_255-SIGN_I[0," & sapRowIndex & "]").text = "E"
            End If
            
            sapRowIndex = sapRowIndex + 1
        End If
    Next cell
    End If ' End If Not dataRange Is Nothing
    
Finalize:
    ' 5. Valider (F8)
    If closePopup Then
        WaitForSAP
        ' Pause de stabilisation (0.5s) pour éviter le crash du GUI après un remplissage intensif
        Dim t As Double: t = Timer: Do While Timer < t + 0.5: DoEvents: Loop
        g_Session.findById("wnd[1]").sendVKey 8
    End If
    Exit Sub
    
ErrHandler:
    DisplayAndLogError "FillSAPSelectionList", Err
    On Error Resume Next
    If closePopup Then g_Session.findById("wnd[1]").Close ' Tente de fermer la popup en cas d'erreur
End Sub

'================================================================================
' SECTION : GESTION DE SESSION SAP (Fonctions Cœur)
'================================================================================

Function onSAP() As Boolean
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : Établit ou réutilise une connexion à une instance SAP GUI déjà ouverte.
    '               Récupère l'objet "g_Session" (session SAP active) et le rend disponible
    '               pour les interactions. Gère les cas où SAP n'est pas ouvert ou est occupé.
    ' RETOUR      : True si la connexion est établie/réutilisée avec succès, False sinon.
    ' APPELÉ PAR  : Toutes les procédures d'action nécessitant une interaction SAP.
    '--------------------------------------------------------------------------------
    Dim SapGuiAuto As Object
    Dim Aplicacion As Object
    Dim Connection As Object
    'Dim g_Session As Object ' Deplace vers modConfig.bas
    Dim busySAP As VbMsgBoxResult

    ' ÉTAPE 0 : Vérification du flag "Do Not Run"
    ' Si l'option "Do Not Run" est activée, la macro s'arrête immédiatement.
    If g_DoNotRun Then
        onSAP = False ' Indique que la connexion n'a pas été établie.
    End If

    ' --- INITIALISATION ---
    ' --- ÉTAPE 1 : Vérifier si une session a déjà été sélectionnée et est valide ---
    If IsSAPConnectionAlive() Then
        ' Une session valide existe déjà (probablement sélectionnée via la ComboBox).
        ' On l'utilise directement sans la réinitialiser.
        ' LogMessage "Réutilisation de la session SAP dé
    Else
        ' ÉTAPE 2 : Si aucune session n'est active ou sélectionnée, tenter de se connecter à la première disponible.
        On Error Resume Next ' Active la gestion d'erreurs pour la connexion SAP.
        Set SapGuiAuto = GetObject("SAPGUI")
        If Err.Number <> 0 Then
            MsgBox "L'application SAP Logon n'est pas ouverte. Veuillez lancer SAP et réessayer.", vbCritical, "Erreur de Connexion SAP"
            onSAP = False
            Exit Function
        End If

        ' Accède au moteur de script et à la première connexion/session.
        Set Aplicacion = SapGuiAuto.GetScriptingEngine
        Set Connection = Aplicacion.children(0)

        If Err.Number <> 0 Or Connection Is Nothing Then
            MsgBox "Aucune connexion SAP n'est ouverte. La macro va s'arrêter.", vbCritical, "Erreur de Connexion SAP"
            offSAP
            Exit Function
        End If

        Set g_Session = Connection.children(0) ' Récupère la première session.
        On Error GoTo 0 ' Désactive la gestion d'erreurs.

        ' Vérifie si l'objet session a été correctement initialisé.
        If g_Session Is Nothing Then
            MsgBox "Impossible de récupérer la session SAP. La macro va s'arrêter.", vbCritical, "Erreur de Session SAP"
            Exit Function
        End If
    End If

    ' --- ÉTAPE 4 : Vérifications finales et activation de la session ---
    LoadConfiguration ' S'assure que les paramètres de configuration sont chargés.

    ' Vérifie à nouveau la validité de la connexion avant de continuer.
    If Not IsSAPConnectionAlive() Then
        MsgBox "La connexion ? SAP a ?t? perdue ou est instable. La macro va s'arr?ter.", vbCritical, "Erreur de Connexion SAP"
        onSAP = False
        Exit Function
    End If
    ' Vérifie si la session SAP est occupée par un autre processus.
    If g_Session.Busy = True Then
        busySAP = MsgBox("La session SAP est occupee par un autre processus." & vbCrLf & vbCrLf & "Voulez-vous annuler l'execution de cette macro ?", vbYesNo + vbQuestion, "SAP Occup?")
        If busySAP = vbYes Then
            onSAP = False
            Exit Function
        End If
    End If

    Set StatusBar = g_Session.findById("wnd[0]/sbar") ' Initialise l'objet barre de statut.

    ' Met la fenêtre SAP au premier plan et lui donne le focus.
    If IsSAPConnectionAlive() Then
        ' Met la fenêtre SAP au premier plan et lui donne le focus.
        g_Session.findById("wnd[0]").iconify
        g_Session.findById("wnd[0]").Restore
        g_Session.findById("wnd[0]").JumpForward
        g_Session.findById("wnd[0]").SetFocus
    End If

    onSAP = True ' Indique que la connexion est prête.
End Function

Sub offSAP()
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : Libère tous les objets SAP et réinitialise les variables globales liées à la session.
    ' APPELÉ PAR  : Toutes les procédures d'action (généralement en fin d'exécution ou en cas d'erreur).
    '--------------------------------------------------------------------------------
    Dim Connection As Object, Aplicacion As Object, SapGuiAuto As Object ' Déclaration des objets locaux.
    
    'pour optimisé la largeur des cologne si elle existe
    OptimizeGridColumns
    
    ' Libère les objets SAP en les mettant à Nothing.
    Set StatusBar = Nothing
    Set g_Session = Nothing
    Set Connection = Nothing
    Set Aplicacion = Nothing
    Set SapGuiAuto = Nothing

    ViderPressePapier
    g_DataType = "" ' Réinitialise le type de données global.


    ' La ligne ci-dessous était commentée car elle accédait directement à la feuille "Setup".
    ' Si cette fonctionnalité est nécessaire, elle devrait être gérée via une fonction
    ' dans un module de configuration pour une meilleure abstraction.
    ' Application.Workbooks("ExcelToSAP Installer.xlam").Sheets("Setup").Range("verifCell").Value = ""
End Sub

Public Function IsSAPConnectionAlive() As Boolean
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : V?rifie si la connexion ? la session SAP est toujours active.
    '               A appeler avant chaque interaction avec g_Session pour ?viter
    '               les erreurs si la connexion est perdue.
    ' RETOUR      : True si la connexion est active, False sinon.
    '--------------------------------------------------------------------------------
    On Error Resume Next

    ' Si g_Session n'est pas initialisé, la connexion est inactive.
    If g_Session Is Nothing Then
        IsSAPConnectionAlive = False
        Exit Function
    End If

    ' La simple vérification que l'objet g_Session n'est pas Nothing est suffisante.
    ' Tenter de lire une propriété comme .Info.systemName peut échouer si la fenêtre
    ' n'est pas active, menant à une reconnexion inutile.
    IsSAPConnectionAlive = Not (g_Session Is Nothing)
    On Error GoTo 0
End Function

Public Function CheckSAPError(ByVal transactionName As String) As Boolean
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : Vérifie si un message d'erreur est présent dans la barre de statut SAP après une action.
    '               Typiquement utilise apres avoir lance une transaction pour verifier les autorisations.
    ' RETOUR      : True si une erreur est trouvee, False sinon.
    '--------------------------------------------------------------------------------
    On Error Resume Next
    If Not IsSAPConnectionAlive() Then Exit Function

    ' Vérifie simplement si un message d'erreur de type 'E' est présent.
    Dim errorType As String
    errorType = g_Session.findById("wnd[0]/sbar").MessageType

    If errorType = "E" Or errorType = "A" Then
        ' Si on n'est PAS en mode test, on affiche la MsgBox à l'utilisateur.
        If Not g_IsTestMode And g_Session.findById("wnd[0]/sbar").text <> "" Then
            MsgBox "Erreur SAP : " & g_Session.findById("wnd[0]/sbar").text, vbCritical, "Erreur SAP"
        ElseIf Not g_IsTestMode Then
            ' Cas où l'erreur est détectée avant l'action SAP (ex: config manquante)
            MsgBox "Erreur de prérequis pour la transaction '" & transactionName & "'. " & _
                   "Veuillez vérifier la configuration (ex: Division manquante dans Setup).", vbCritical, "Erreur de Configuration"
        End If
        CheckSAPError = True ' Erreur trouvee
        Exit Function
    End If
    
    On Error GoTo 0
    CheckSAPError = False ' Aucune erreur détectée.
    
End Function

'--------------------------------------------------------------------------------
' PROCÉDURE   : CloseSecondaryWindows
' DESCRIPTION : Tente de fermer toutes les fenêtres SAP secondaires (wnd[1], wnd[2], etc.)
'               pour s'assurer que seule la fenêtre principale (wnd[0]) reste ouverte.
'               Utilise "On Error Resume Next" pour ignorer les erreurs si une fenêtre
'               secondaire n'existe pas.
' APPELÉ PAR  : AutomatedTestWrapper (à la fin de chaque test).
'--------------------------------------------------------------------------------
Public Sub CloseSecondaryWindows()
    ' Si g_Session est vide, on tente de récupérer une session active pour le nettoyage.
    If g_Session Is Nothing Then
        On Error Resume Next
        Dim SapGuiAuto As Object, Aplicacion As Object
        Set SapGuiAuto = GetObject("SAPGUI")
        If SapGuiAuto Is Nothing Then Exit Sub ' Si SAP n'est pas ouvert, on ne peut rien faire.

        Set Aplicacion = SapGuiAuto.GetScriptingEngine
        If Aplicacion.Connections.count > 0 Then
            Set g_Session = Aplicacion.Connections(0).children(0)
        End If
    End If

    On Error Resume Next ' Ignore les erreurs si une fenêtre n'existe pas

    ' Boucle pour fermer les fenêtres modales potentielles (wnd[1] à wnd[5])
    Dim i As Integer
    For i = 5 To 1 Step -1 ' Boucle inversée pour fermer les fenêtres les plus récentes en premier
        g_Session.findById("wnd[" & i & "]").Close
    Next i

    ' Ajoute une courte pause pour laisser le temps à l'interface SAP de se stabiliser
    ' après la fermeture des fenêtres.
    'Application.Wait Now + TimeValue("00:00:02")

    On Error GoTo 0 ' Réactive la gestion d'erreurs standard
End Sub

'================================================================================
' SECTION : UTILITAIRES DE VÉRIFICATION SAP
'================================================================================

'--------------------------------------------------------------------------------
' FONCTION    : IsOnTransaction
' DESCRIPTION : Vérifie si la session SAP se trouve sur l'écran de la transaction attendue.
'               Affiche un message d'erreur et retourne False si ce n'est pas le cas.
' PARAMÈTRE   : expectedTCode (String) - Le code de transaction attendu (ex: "IW33").
' RETOUR      : Boolean - True si la transaction est correcte, False sinon.
'--------------------------------------------------------------------------------
Public Function IsOnTransaction(ByVal expectedTCode As String) As Boolean
    On Error Resume Next
    IsOnTransaction = False
    If g_Session Is Nothing Then Exit Function
    
    Dim currentTCode As String
    currentTCode = CStr(g_Session.Info.Transaction)
    
    If LCase(currentTCode) = LCase(expectedTCode) Then
        IsOnTransaction = True
    Else
        ' Si la transaction est incorrecte, on journalise et on informe l'utilisateur.
        LogMessage "Vérification de transaction échouée. Attendu: " & expectedTCode & ", Actuel: " & currentTCode
        If Not g_IsTestMode Then ' N'affiche pas de MsgBox en mode test pour ne pas bloquer l'exécution
            MsgBox "La transaction attendue (" & expectedTCode & ") n'est pas active." & vbCrLf & vbCrLf & "L'écran actuel est : '" & currentTCode & "'." & vbCrLf & vbCrLf & "L'opération a été annulée.", vbExclamation, "Erreur de Transaction"
        End If
    End If
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' PROCÉDURE   : RunSAPTransaction
' DESCRIPTION : Fonction générique pour lancer une transaction SAP et remplir un champ.
'               Factorise le code répétitif des wrappers de transaction.
' PARAMÈTRES  :
'   - tCode (String)       : Code de la transaction (ex: "IW33").
'   - fieldId (String)     : ID du champ à remplir (ex: "ctxtCAUFVD-AUFNR"). Optionnel.
'   - fieldValue (Variant) : Valeur à mettre dans le champ. Si omis, utilise la cellule active.
'--------------------------------------------------------------------------------
Public Sub RunSAPTransaction(ByVal tCode As String, Optional ByVal fieldId As String = "", Optional ByVal fieldValue As Variant = Empty)
    On Error GoTo SapErrorHandler
    
    If Not IsSAPConnectionAlive() Then Exit Sub

    ' Lancement de la transaction
    g_Session.findById("wnd[0]/tbar[0]/okcd").text = "/n " & tCode
    g_Session.findById("wnd[0]").sendVKey 0
    
    ' Vérifications
    If CheckSAPError(tCode) Then Exit Sub
    If Not IsOnTransaction(tCode) Then Exit Sub

    ' Remplissage du champ
    If fieldId <> "" Then
        If IsEmpty(fieldValue) Then fieldValue = Cells(ActiveCell.row, ActiveCell.Column).value
        
        ' Gestion des ID complets ou relatifs
        Dim fullFieldId As String
        If Left(fieldId, 3) = "wnd" Then fullFieldId = fieldId Else fullFieldId = "wnd[0]/usr/" & fieldId
        
        g_Session.findById(fullFieldId).text = fieldValue
    End If
    
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "RunSAPTransaction (" & tCode & ")", Err
    offSAP
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (CC - Product Structure)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CC04 (Product Structure Browser).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme article.
'--------------------------------------------------------------------------------
Sub Z_CC04()
    RunSAPTransaction "CC04", "tabsBROWSER_TAB_STRIP/tabpFCMAT/ssubINPUT_MAT:SAPLCPDMOBJECTBROWSER:0300/ctxtMARA-MATNR"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (CN - Project System)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN42N (Liste des projets, vue d'ensemble).
' CONFIG      : Utilise les settings "PS_PROF" et "LAY_CN42N".
'--------------------------------------------------------------------------------
Sub Z_CN42N()
    RunSAPTransaction "CN42N"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0


    If GetSetting("LAY_CN42N") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_DISVAR").text = GetSetting("LAY_CN42N")
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN43N (Liste des documents projet).
' CONFIG      : Utilise les settings "PS_PROF" et "LAY_CN43N".
'--------------------------------------------------------------------------------
Sub Z_CN43N()
    RunSAPTransaction "CN43N"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0

    If GetSetting("LAY_CN43N") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_DISVAR").text = GetSetting("LAY_CN43N")
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKVR").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN46N (Liste des réseaux, vue d'ensemble).
' CONFIG      : Utilise les settings "PS_PROF" et "LAY_CN46N".
'--------------------------------------------------------------------------------
Sub Z_CN46N()
    RunSAPTransaction "CN46N"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0

    If GetSetting("LAY_CN46N") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_DISVAR").text = GetSetting("LAY_CN46N")
    g_Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (CS - Nomenclatures)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CS15 (Utilisation article dans nomenclatures).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme article.
' CONFIG      : Utilise le setting "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_CS15()
    RunSAPTransaction "CS15", "ctxtRC29L-MATNR"
    
    g_Session.findById("wnd[0]/usr/chkRC29L-EQUTP").Selected = True
    g_Session.findById("wnd[0]/usr/chkRC29L-EQUTP").SetFocus
    g_Session.findById("wnd[0]/usr/chkRC29L-DIRKT").Selected = True
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    
    CheckSAPError ("CS15")
    
    If GetSetting("SAP_PLANT_PF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = GetSetting("SAP_PLANT_PF")
    Else
        MsgBox ("This transaction requires the Planning Plant, it's not defined in the setup. Execution Canceled")
        Exit Sub
    End If
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN47 (Liste des activités, vue d'ensemble).
' CONFIG      : Utilise le setting "PS_PROF".
'--------------------------------------------------------------------------------
Sub Z_CN47()
    RunSAPTransaction "CN47"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0

    g_Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN47N (Liste des activités, vue d'ensemble).
' CONFIG      : Utilise les settings "PS_PROF" et "LAY_CN47N".
'--------------------------------------------------------------------------------
Sub Z_CN47N()
    RunSAPTransaction "CN47N"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0

    If GetSetting("LAY_CN47N") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_DISVAR").text = GetSetting("LAY_CN47N")
    g_Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CNS41 (Liste des composants).
' CONFIG      : Utilise le setting "PS_PROF".
'--------------------------------------------------------------------------------
Sub Z_CNS41()
    RunSAPTransaction "CNS41"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0

    g_Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_MATNR-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP CN48N (Liste des confirmations, vue d'ensemble).
' CONFIG      : Utilise les settings "PS_PROF" et "LAY_CN48N".
'--------------------------------------------------------------------------------
Sub Z_CN48N()
    RunSAPTransaction "CN48N"

    On Error Resume Next
    g_Session.findById("wnd[1]/usr/ctxtTCNTT-PROFID").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").text = GetSetting("PS_PROF")
    g_Session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0

    If GetSetting("LAY_CN48N") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_DISVAR").text = GetSetting("LAY_CN48N")
    g_Session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (CV - Documents)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction CV02N (Modifier Document).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de document.
' CONFIG      : Le type de document est fixé à "ZPM".
'--------------------------------------------------------------------------------
Sub Z_CV02N()
    RunSAPTransaction "CV02N", "ctxtDRAW-DOKNR"
    ' Champs spécifiques supplémentaires
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKAR").text = "ZPM" ' Renseigne le type de document.
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").text = "" ' Efface la partie du document.
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKVR").text = "" ' Efface la version du document.
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction CV03N (Afficher Document).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de document.
' CONFIG      : Le type de document est fixé à "ZPM".
'--------------------------------------------------------------------------------
Sub Z_CV03N()
    RunSAPTransaction "CV03N", "ctxtDRAW-DOKNR"
    ' Champs spécifiques supplémentaires
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKAR").text = "" '"ZPM" ' Renseigne le type de document.
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKTL").text = "" ' Efface la partie du document.
    g_Session.findById("wnd[0]/usr/ctxtDRAW-DOKVR").text = "" ' Efface la version du document.
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction CV04N (Rechercher Document).
'--------------------------------------------------------------------------------
Sub Z_CV04N()
    RunSAPTransaction "CV04N"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IA - Gammes)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IA01 (Créer gamme opératoire).
'--------------------------------------------------------------------------------
Sub Z_IA01()
    RunSAPTransaction "IA01"
End Sub
'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IA02 (Modifier gamme opératoire).
'--------------------------------------------------------------------------------
Sub Z_IA02()
    RunSAPTransaction "IA02", "ctxtRC27E-EQUNR"
    
    g_Session.findById("wnd[0]/usr/ctxtRC271-VERWE").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC271-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IA03 (Afficher gamme opératoire).
'--------------------------------------------------------------------------------
Sub Z_IA03()
    RunSAPTransaction "IA03", "ctxtRC27E-EQUNR"
    
    g_Session.findById("wnd[0]/usr/ctxtRC271-VERWE").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC271-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IA11 (Creer Liste de taches poste tech.).
'               Passe la valeur de la cellule active en parametre.
'--------------------------------------------------------------------------------
Sub Z_IA11()
    RunSAPTransaction "IA11", "ctxtRC27E-TPLNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IA12 (Modifier Liste de taches poste tech.).
'               Passe la valeur de la cellule active en parametre.
'--------------------------------------------------------------------------------
Sub Z_IA12()
    RunSAPTransaction "IA12", "ctxtRC27E-TPLNR"
    
    g_Session.findById("wnd[0]/usr/txtRC271-PLNAL").text = "" ' Efface le groupe de fiches de tâches.
    g_Session.findById("wnd[0]/usr/ctxtRC271-VERWE").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC271-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IA13 (Afficher Liste de taches poste tech.).
'               Passe la valeur de la cellule active en parametre.
'--------------------------------------------------------------------------------
Sub Z_IA13()
    RunSAPTransaction "IA13", "ctxtRC27E-TPLNR"
    
    g_Session.findById("wnd[0]/usr/txtRC271-PLNAL").text = "" ' Efface le groupe de fiches de tâches.
    g_Session.findById("wnd[0]/usr/ctxtRC271-VERWE").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC271-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IA17 (Afficher listes de tâches, multi-niveaux).
'--------------------------------------------------------------------------------
Sub Z_IA17()
    RunSAPTransaction "IA17"
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[0]/usr/ctxtPN_STATU-LOW").text = "4"
    g_Session.findById("wnd[0]/usr/ctxtPN_VERWE-LOW").text = "4"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IB - Nomenclatures Article)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IB01 (Créer nomenclature article).
'--------------------------------------------------------------------------------
Sub Z_IB01()
    RunSAPTransaction "IB01"
    
    g_Session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IB02 (Modifier nomenclature article).
'--------------------------------------------------------------------------------
Sub Z_IB02()
    RunSAPTransaction "IB02"
    
    g_Session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "4"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IB03 (Afficher nomenclature article).
'--------------------------------------------------------------------------------
Sub Z_IB03()
    RunSAPTransaction "IB03"
    
    g_Session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "4"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = GetSetting("SAP_PLANT_PF")
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IE - Equipements)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IE03 (Afficher un équipement).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'équipement.
'--------------------------------------------------------------------------------
Sub Z_IE03()
    RunSAPTransaction "IE03", "ctxtRM63E-EQUNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IE07 (Liste d'équipements, multi-niveaux).
'--------------------------------------------------------------------------------
Sub Z_IE07()
    RunSAPTransaction "IE07"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IH - Postes Techniques)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IH01 (arboressence Plant).
' CONFIG      : Utilise le setting "LAY_IH06".
'--------------------------------------------------------------------------------
Sub Z_IH01()
    RunSAPTransaction "IH01"

    If GetSetting("SAP_PLANT_AF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtDY_TPLNR").text = GetSetting("SAP_PLANT_AF")
    End If

    g_Session.findById("wnd[0]/usr/chkDY_FLHIE").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_EQUBI").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_EQHIE").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IHBTY").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_BOMEX").Selected = True
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IH06 (Liste de nomenclatures d'équipement).
' CONFIG      : Utilise le setting "LAY_IH06".
'--------------------------------------------------------------------------------
Sub Z_IH06()
    RunSAPTransaction "IH06"
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IH06") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IH06")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IH08 (Liste Postes techniques).
'               Configure le layout ("LAY_IH08").
'--------------------------------------------------------------------------------
Sub Z_IH08()
    RunSAPTransaction "IH08"

    ' Efface les champs de date pour une recherche sans restriction temporelle.
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""

    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IH08") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IH08")
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IL - Postes Techniques)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IL03 (Afficher Poste technique).
'               Passe la valeur de la cellule active en param?tre.
'--------------------------------------------------------------------------------
Sub Z_IL03()
    RunSAPTransaction "IL03", "ctxtIFLO-TPLNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IL07 (Liste de nomenclatures de poste technique).
'--------------------------------------------------------------------------------
Sub Z_IL07()
    RunSAPTransaction "IL07"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IP - Plans d'entretien)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IP02 (Modifier Plan d'entretien).
' UTILISE     : La variable globale "mplan" pour pré-remplir le champ du plan.
'--------------------------------------------------------------------------------
Sub Z_IP02()
    RunSAPTransaction "IP02", "ctxtRMIPM-WARPL"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IP03 (Afficher Plan d'entretien).
' UTILISE     : La variable globale "mplan" pour pré-remplir le champ du plan.
'--------------------------------------------------------------------------------
Sub Z_IP03()
    RunSAPTransaction "IP03", "ctxtRMIPM-WARPL"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IP06 (Afficher un poste de plan d'entretien).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme poste.
'--------------------------------------------------------------------------------
Sub Z_IP06()
    RunSAPTransaction "IP06", "ctxtRMIPM-WAPOS"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IP16 (Liste des postes d'entretien planifiés).
' CONFIG      : Utilise le setting "LAY_IP16".
'--------------------------------------------------------------------------------
Sub Z_IP16()
    RunSAPTransaction "IP16"
    
    If GetSetting("LAY_IP16") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IP16")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IP18 (Liste Plans d'entretien).
'               Configure les statuts, le layout et la division ("LAY_IP18", "ppf").
'--------------------------------------------------------------------------------
Sub Z_IP18()
    RunSAPTransaction "IP18"
    
    g_Session.findById("wnd[0]/usr/chkSPERRE").Selected = True ' Sélectionne les plans bloqués.
    g_Session.findById("wnd[0]/usr/chkOBLIS").Selected = True ' Sélectionne les plans obligatoires.
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IP18") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IP18")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IP18 (Liste Plans d'entretien).
'               Version sans selection de statut (tous).
'--------------------------------------------------------------------------------
Sub Z_IP18_All()
    RunSAPTransaction "IP18"
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IP18") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IP18")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IP24 (Liste Postes d'entretien).
'               Configure statut, layout et division ("LAY_IP24", "ppf").
'--------------------------------------------------------------------------------
Sub Z_IP24()
    RunSAPTransaction "IP24"
    
    g_Session.findById("wnd[0]/usr/chkOBLIS").Selected = True
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IP24") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IP24")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Wrapper pour T-Code IP24 (Liste Postes d'entretien).
'               Version sans selection de statut.
'--------------------------------------------------------------------------------
Sub Z_IP24_NoOL()
    RunSAPTransaction "IP24"
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IP24") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IP24")
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IW - Ordres, Avis, Confirmations)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW22 (Modifier un avis PM).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'avis.
'--------------------------------------------------------------------------------
Sub Z_IW22()
    RunSAPTransaction "IW22", "ctxtRIWO00-QMNUM"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW23 (Afficher un avis PM).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'avis.
'--------------------------------------------------------------------------------
Sub Z_IW23()
    RunSAPTransaction "IW23", "ctxtRIWO00-QMNUM"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW28 (Liste des avis PM, variante).
' CONFIG      : Utilise les settings "LAY_IW28" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW28()
    RunSAPTransaction "IW28"
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IW28") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW28")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW29 (Liste des avis PM).
' CONFIG      : Utilise les settings "LAY_IW29" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW29()
    RunSAPTransaction "IW29"
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_IW29") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW28")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW29 (Liste des avis PM, historique).
' CONFIG      : Utilise les settings "LAY_IW29H" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW29Hist()
    RunSAPTransaction "IW29"
    
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = False
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = False
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("LAY_IW29H") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29H")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW30 (Affichage multi-niveaux).
'--------------------------------------------------------------------------------
Sub Z_IW30()
    RunSAPTransaction "IW30"
End Sub

'================================================================================
' SECTION : WRAPPERS DE TRANSACTIONS SAP (IWBK - Composants d'ordre)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IWBK (Liste des composants d'ordre).
' CONFIG      : Pré-remplit la division et le magasin.
'--------------------------------------------------------------------------------
Sub Z_IWBK()
    RunSAPTransaction "IWBK"
    
    ' Renseigne la division (Planning Plant) depuis les paramètres de configuration
    If GetSetting("SAP_PLANT_MF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")

    End If
    
    ' Renseigne le magasin (Storage Location) depuis les paramètres de configuration
    If GetSetting("Storage_Location") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = GetSetting("Storage_Location")

    End If
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW32 (Modifier un ordre de maintenance).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'ordre.
'--------------------------------------------------------------------------------
Sub Z_IW32()
    RunSAPTransaction "IW32", "ctxtCAUFVD-AUFNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW33 (Afficher un ordre de maintenance).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'ordre.
'--------------------------------------------------------------------------------
Sub Z_IW33()
    RunSAPTransaction "IW33", "ctxtCAUFVD-AUFNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW37 (Liste des opérations de maintenance).
' CONFIG      : Utilise les settings "LAY_IW38" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW37()
    RunSAPTransaction "IW37"
    
    If GetSetting("LAY_IW38") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW38")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW38 (Liste des ordres de maintenance, variante).
' CONFIG      : Utilise les settings "LAY_IW38" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW38()
    RunSAPTransaction "IW38"
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("LAY_IW38") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW38")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW39 (Liste des ordres de maintenance).
' CONFIG      : Utilise les settings "LAY_IW39" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW39()
    RunSAPTransaction "IW39"
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("LAY_IW39") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW38")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW39 (Liste des ordres de maintenance, historique).
' CONFIG      : Utilise le setting "LAY_IW39H".
'--------------------------------------------------------------------------------
Sub Z_IW39H()
    RunSAPTransaction "IW39"
    
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = False
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = False
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_HIS").Selected = True
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("LAY_IW39H") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW39H")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW40 (Affichage multi-niveaux des ordres).
' CONFIG      : Coche tous les statuts et renseigne la division de planification.
'--------------------------------------------------------------------------------
Sub Z_IW40()
    RunSAPTransaction "IW40"

    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_HIS").Selected = True
    
    g_Session.findById("wnd[0]/usr/btnISEL_ALL").press
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW41 (Saisie de confirmation).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'ordre.
'--------------------------------------------------------------------------------
Sub Z_IW41()
    RunSAPTransaction "IW41", "ctxtCORUF-AUFNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW44 (Afficher confirmation d'ordre PM).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de confirmation.
'--------------------------------------------------------------------------------
Sub Z_IW44()
    RunSAPTransaction "IW44", "ctxtAFRUD-RUECK"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW45 (Annuler confirmation).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro d'ordre.
'--------------------------------------------------------------------------------
Sub Z_IW45()
    RunSAPTransaction "IW45", "ctxtCORUF-AUFNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW47 (Afficher confirmations d'ordres).
' CONFIG      : Coche les statuts et renseigne la division.
'--------------------------------------------------------------------------------
Sub Z_IW47()
    RunSAPTransaction "IW47"
    
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_ABG").Selected = True
    g_Session.findById("wnd[0]/usr/chkNO_CANC").Selected = True
    
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtWERKS_C-LOW").text = GetSetting("SAP_PLANT_PF")
    End If
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW49 (Afficher confirmations d'ordres).
'--------------------------------------------------------------------------------
Sub Z_IW49()
    RunSAPTransaction "IW49"
    
    If GetSetting("LAY_IW49") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW49")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW49N (Afficher confirmations d'ordres, nouvelle version).
'--------------------------------------------------------------------------------
Sub Z_IW49N()
    RunSAPTransaction "IW49N"

    g_Session.findById("wnd[0]/usr/chkSP_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    If GetSetting("SAP_PLANT_PF") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_IWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    End If

    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    If GetSetting("SAP_PLANT_MF") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/ctxtS_SWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    End If

    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49n") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49n")
    End If
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP IW69 (Liste des activités pour avis PM).
' CONFIG      : Utilise les settings "LAY_IW29" et "SAP_PLANT_PF".
'--------------------------------------------------------------------------------
Sub Z_IW69()
    RunSAPTransaction "IW69"
    
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    If GetSetting("LAY_IW29") <> "" Then g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
End Sub

'================================================================================
' SECTION     : WRAPPERS DE TRANSACTIONS SAP (MB - Mouvements de stock)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MB24 (Liste des réservations).
' CONFIG      : Utilise les settings "SAP_PLANT_MF" et "LAY_MB24".
'--------------------------------------------------------------------------------
Sub Z_MB24()
    RunSAPTransaction "MB24", "ctxtMATNR-LOW"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_MB24") <> "" Then g_Session.findById("wnd[0]/usr/ctxtALV_DEF").text = GetSetting("LAY_MB24")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MB25 (Liste des réservations).
' CONFIG      : Utilise les settings "SAP_PLANT_MF" et "LAY_MB24".
'--------------------------------------------------------------------------------
Sub Z_MB25()
    RunSAPTransaction "MB25"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_MB25") <> "" Then g_Session.findById("wnd[0]/usr/ctxtALV_DEF").text = GetSetting("LAY_MB24")
End Sub
'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MB51 (Liste des mouvements de stock).
' CONFIG      : Utilise les settings "SAP_PLANT_MF" et "LAY_MB51".
'--------------------------------------------------------------------------------
Sub Z_MB51()
    RunSAPTransaction "MB51", "ctxtMATNR-LOW"

    g_Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = ""
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("LAY_MB51") <> "" Then g_Session.findById("wnd[0]/usr/ctxtALV_DEF").text = GetSetting("LAY_MB51")
    
    On Error Resume Next
    g_Session.findById("wnd[0]/usr/radRFLAT_L").Select
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MB52 (Liste des stocks magasin).
' CONFIG      : Utilise les settings "SAP_PLANT_MF" et "LAY_MB52".
'--------------------------------------------------------------------------------
Sub Z_MB52()
    RunSAPTransaction "MB52"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("Storage_Location") <> "" Then g_Session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = GetSetting("Storage_Location")
    If GetSetting("LAY_MB52") <> "" Then g_Session.findById("wnd[0]/usr/ctxtP_VARI").text = GetSetting("LAY_MB52")
    
    g_Session.findById("wnd[0]/usr/radPA_FLT").Select
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MB53 (Disponibilité des articles).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme article.
' CONFIG      : Utilise le setting "SAP_PLANT_MF".
'--------------------------------------------------------------------------------
Sub Z_MB53()
    RunSAPTransaction "MB53", "ctxtMATNR"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/chkAUTOSEL").Selected = True
End Sub

'================================================================================
' SECTION     : WRAPPERS DE TRANSACTIONS SAP (ME - Achats)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME23 (Afficher commande d'achat).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de commande.
'--------------------------------------------------------------------------------
Sub Z_ME23()
    RunSAPTransaction "ME23", "ctxtRM06E-BSTNR"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME23N (Afficher commande d'achat, Enjoy).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de commande.
'--------------------------------------------------------------------------------
Sub Z_ME23N()
    RunSAPTransaction "ME23N"
    
    On Error Resume Next
    g_Session.findById("wnd[0]/mbar/menu[0]/menu[0]").Select
    g_Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_F").Select
    g_Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = Cells(ActiveCell.row, ActiveCell.Column).value
    g_Session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME2N (Liste des commandes d'achat).
' CONFIGURATION : Utilise la configuration "SAP_PLANT_MF" pour la division.
'               Configure la division ("mpf").
'--------------------------------------------------------------------------------
Sub Z_ME2N()
    RunSAPTransaction "ME2N", "ctxtS_MATNR-LOW"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/ctxtEN_EBELN-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME53 (Afficher demande d'achat).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de demande.
'--------------------------------------------------------------------------------
Sub Z_ME53()
    RunSAPTransaction "ME53", "ctxtEBAN-BANFN"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME53N (Afficher demande d'achat, Enjoy).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme numéro de demande.
'--------------------------------------------------------------------------------
Sub Z_ME53N()
    RunSAPTransaction "ME53N"
    
    On Error Resume Next
    g_Session.findById("wnd[0]/mbar/menu[0]/menu[0]").Select
    g_Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_B").Select
    g_Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").text = Cells(ActiveCell.row, ActiveCell.Column).value
    g_Session.findById("wnd[1]").sendVKey 0
    
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME5A (Liste des demandes d'achat).
' CONFIG      : Utilise le setting "SAP_PLANT_MF".
'--------------------------------------------------------------------------------
Sub Z_ME5A()
    RunSAPTransaction "ME5A"
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/chkP_ERLBA").Selected = True
    
    g_Session.findById("wnd[0]/usr/ctxtBA_BANFN-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").text = ""
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME80FN (Reporting général achats).
'--------------------------------------------------------------------------------
Sub Z_ME80FN()
    RunSAPTransaction "ME80FN"

    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSP$00011-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("Storage_Location") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSP$00009-LOW").text = GetSetting("Storage_Location")
    
    g_Session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = "F"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP ME91F (Messages pour commandes d'achat).
'--------------------------------------------------------------------------------
Sub Z_ME91F()
    RunSAPTransaction "ME91F"
    
    If GetSetting("Purchasing_Org") <> "" Then g_Session.findById("wnd[0]/usr/ctxtEN_EKORG-LOW").text = GetSetting("Purchasing_Org")
    g_Session.findById("wnd[0]/usr/ctxtEN_EBELN-LOW").text = ""
End Sub

'================================================================================
' SECTION     : WRAPPERS DE TRANSACTIONS SAP (MM - Gestion des articles)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction SAP MMBE (Synthèse des stocks).
' PARAMÈTRE   : La valeur de la cellule active est utilisée comme article.
' CONFIG      : Utilise le setting "SAP_PLANT_MF".
'--------------------------------------------------------------------------------
Sub Z_MMBE()
    RunSAPTransaction "MMBE", "ctxtMS_MATNR-LOW"
    
    ' Paramètres supplémentaires spécifiques à MMBE
    g_Session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtMS_WERKS-HIGH").text = ""
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtMS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    If GetSetting("Storage_Location") <> "" Then g_Session.findById("wnd[0]/usr/ctxtMS_LGORT-LOW").text = GetSetting("Storage_Location")
    
    g_Session.findById("wnd[0]/usr/chkKZNUL").Selected = False
End Sub

'================================================================================
' SECTION     : ACTIONS SAP (F8 / ENTREE)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Simule l'appui sur la touche F8 (Executer).
'               G?re deux modes : "SAP Control" (scripting) ou "SendKeys".
'               Gere l'etat de la touche "NumLock".
'--------------------------------------------------------------------------------
Sub Z_F8()
    On Error GoTo SapErrorHandler
    
    If g_DoNotRun = False Then
        If Not IsSAPConnectionAlive() Then Exit Sub
        
        If GetSetting("Exe_Mode") = "SAP Control" Then
            ' --- Mode "SAP Control" ---
            g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
        Else
            g_Session.findById("wnd[0]").sendVKey 8 ' F8
        End If
        
    End If

    If Not CheckSAPError("F8") Then Exit Sub
    
SapErrorHandler:
    'DisplayAndLogError "Z_F8", Err
    offSAP
    Exit Sub
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Simule l'appui sur la touche "Entree".
'               G?re deux modes : "SAP Control" (scripting) ou "SendKeys".
'               Gere l'etat de la touche "NumLock".
'--------------------------------------------------------------------------------
Sub Z_Enter()
    Dim titleBefore As String, titleAfter As String
    On Error GoTo SapErrorHandler
    
    If g_DoNotRun = False Then
        If Not IsSAPConnectionAlive() Then Exit Sub
        
        If GetSetting("Exe_Mode") = "SAP Control" Then
            ' --- Mode "SAP Control" ---
            g_Session.findById("wnd[0]").sendVKey 0
        Else
            g_Session.findById("wnd[0]").sendVKey 0 ' Enter
        End If
        
    End If
    
    If Not CheckSAPError("Enter") Then Exit Sub
    
SapErrorHandler:
    'DisplayAndLogError "Z_Enter", Err
    offSAP
    Exit Sub
End Sub

'================================================================================
' SECTION     : KPI (T-CODES SP?CIFIQUES)
'================================================================================

'---------- New code for KPIs ----------------
'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ProGroup KPIP_SM.
'--------------------------------------------------------------------------------
Sub Z_KPIP_SM()
    RunSAPTransaction "/PROGROUP/KPIP_SM"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ProGroup /PGPNL/SHIFT pour la capacité.
'--------------------------------------------------------------------------------
Public Sub Z_SHIFT()
    RunSAPTransaction "/PGPNL/SHIFT"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ZOA2C_GRANT_GUI pour l'authentification Google.
'--------------------------------------------------------------------------------
Sub Z_ZGOOGLE()
    RunSAPTransaction "ZOA2C_GRANT_GUI"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ProGroup /PROGROUP/PPM pour l'impression des ordres.
'--------------------------------------------------------------------------------
Sub Z_PPM()
    RunSAPTransaction "/PROGROUP/PPM"
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ProGroup /PGP/SCHEDULER (Gantt Scheduler).
'--------------------------------------------------------------------------------
Sub Z_Gannt_Scheduler()
    RunSAPTransaction "/PGP/SCHEDULER"
    
    ' Vérifie si la transaction a bien été lancée avant de continuer
    If Not IsSAPConnectionAlive() Then Exit Sub

    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_SWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ProGroup /PGPNL/GS_D (Gantt Display).
'--------------------------------------------------------------------------------
Sub Z_Gannt_Display()
    RunSAPTransaction "/PGPNL/GS_D"
    
    ' Vérifie si la transaction a bien été lancée avant de continuer
    If Not IsSAPConnectionAlive() Then Exit Sub

    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_SWERK-LOW").text = GetSetting("SAP_PLANT_MF")
End Sub

'--------------------------------------------------------------------------------
' DESCRIPTION : Lance la transaction ZPM004.
'--------------------------------------------------------------------------------
Sub Z_ZPM004()
    RunSAPTransaction "ZPM004"
End Sub

'================================================================================
' SECTION     : ACTIONS SAP (GRILLES)
'================================================================================

'--------------------------------------------------------------------------------
' DESCRIPTION : Optimise la largeur des colonnes de la première grille (GRID1)
'               trouvée dans la fenêtre active.
'               Gère les erreurs si aucune grille n'est présente.
'--------------------------------------------------------------------------------
Public Sub OptimizeGridColumns()
    On Error Resume Next ' Active la gestion d'erreurs pour éviter un arrêt si la grille n'existe pas

    If IsSAPConnectionAlive() Then
        g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 0
        g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&OPTIMIZE"
    End If

    On Error GoTo 0 ' Désactive la gestion d'erreurs
End Sub

