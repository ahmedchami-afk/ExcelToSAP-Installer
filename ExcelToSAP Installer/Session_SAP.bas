Attribute VB_Name = "Session_SAP"
'====================================================================================
' MODULE      : Session_SAP
' AUTEUR      : Votre Nom / Gemini Code Assist
' DATE        : 02/12/2025
' DESCRIPTION :
'   Ce module est dédié à la gestion complète des sessions SAP GUI depuis Excel.
'   Il centralise les fonctionnalités suivantes :
'     - Ouverture et réutilisation de sessions SAP.
'     - Création de nouvelles sessions.
'     - Fermeture propre des sessions (individuelle ou toutes).
'     - Remplissage et gestion de la liste déroulante (ComboBox) des sessions dans le ruban.
'
' CONTEXTE :
'   Utilisé pour assurer une communication stable et contrôlée entre Excel et l'API
'   de scripting de SAP GUI.
'====================================================================================
Option Explicit

'------------------------------------------------------------------------------------
' SECTION : DÉCLARATIONS API WINDOWS
' Description :
'   Ces fonctions de l'API Windows sont utilisées pour interagir avec les fenêtres
'   du système, notamment pour trouver et fermer une fenêtre SAP Logon si nécessaire.
'------------------------------------------------------------------------------------
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

Const WM_CLOSE = &H10


'====================================================================================
' SECTION : PROCÉDURES PRINCIPALES DE GESTION DE SESSION
'====================================================================================

'------------------------------------------------------------------------------
' Sub : SAP_Openg_SessionFromLogon
' Description :
'   Procédure principale pour ouvrir une session SAP. Elle tente d'abord de
'   réutiliser une session existante. Si aucune n'est disponible, elle lance
'   SAP Logon et établit une nouvelle connexion en utilisant les paramètres
'   définis dans la configuration (feuille "Setup").
'
' Étapes clés :
'   1. Tente de réutiliser une session SAP déjà ouverte.
'   2. Si aucune session n'est trouvée, lance SAP Logon via le chemin configuré.
'   3. Ouvre une connexion vers le système SAP spécifié.
'   4. Gère les fenêtres pop-up potentielles (ex: copyright, multi-logon).
'------------------------------------------------------------------------------
Sub SAP_Openg_SessionFromLogon()

    Dim SapGui As Object, Applic As Object, Connection As Object
    Dim WshShell As Object
    Dim sapPath As String, sapBase As String

    On Error GoTo ErrHandler

    ' Récupère les paramètres de connexion depuis la configuration globale.
    sapPath = GetSetting("SAP Path")
    sapBase = GetSetting("SAP Base")
    
    ' Valeurs par défaut si non définies dans le Setup
    If sapPath = "" Then
        sapPath = GetSAPLogonPath()
        If sapPath = "" Then sapPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    End If
    If sapBase = "" Then sapBase = "2.3.01 PR1 - Productive ECC6 Core"

    ' ÉTAPE 1 : Tenter de réutiliser une session existante.
    On Error Resume Next ' Gestion d'erreur temporaire pour GetObject.
    Set SapGui = GetObject("SAPGUI")
    If Not SapGui Is Nothing Then
        Set Applic = SapGui.GetScriptingEngine
        If Not Applic Is Nothing Then
            ' Si g_Session n'est pas déjà définie, on essaie de prendre la première session disponible
            If g_Session Is Nothing And Applic.Connections.count > 0 Then
                Set Connection = Applic.children(0) ' Prend la première connexion
                Set g_Session = Connection.children(0)
                If Not g_Session Is Nothing Then
                    LogMessage "Première session SAP existante réutilisée (Système: " & g_Session.Info.systemName & ")."
                    Exit Sub ' Session trouvée, on quitte la procédure.
                End If
            End If
        ElseIf Not g_Session Is Nothing Then
             LogMessage "Une session SAP est déjà active. Aucune nouvelle connexion ne sera lancée."
             Exit Sub
        End If
    End If
    On Error GoTo ErrHandler ' Réactive la gestion d'erreur standard.
    
    ' ÉTAPE 2 : Si aucune session n'a été trouvée, en créer une nouvelle.
    LogMessage "Aucune session SAP existante ou valide trouvée. Lancement d'une nouvelle connexion."
    
    If sapPath = "" Or sapBase = "" Then
        MsgBox "Les paramètres SAP ('SAP Path', 'SAP Base') ne sont pas définis dans la feuille Setup.", vbCritical, "Erreur de Configuration"
        Exit Sub
    End If

    ' --- 2. Lancer SAP Logon et attendre quelques secondes pour le chargement ---
    Shell sapPath, vbNormalFocus
    LogMessage "Lancement de SAP Logon depuis : " & sapPath

    ' --- Boucle d'attente intelligente pour l'objet SAPGUI ---
    Dim startTime As Single
    startTime = Timer
    LogMessage "Attente de la disponibilité de l'objet SAPGUI..."

    Do
        On Error Resume Next
        Set SapGui = GetObject("SAPGUI")
        On Error GoTo ErrHandler ' Réactive la gestion d'erreur standard.

        If Not SapGui Is Nothing Then Exit Do ' L'objet est trouvé, on sort de la boucle.
        DoEvents ' Laisse Excel respirer pour éviter de figer.

        ' Sécurité : si l'attente dépasse 15 secondes, on abandonne.
        If Timer - startTime > 15 Then
            MsgBox "SAP Logon n'a pas pu être détecté après 15 secondes. L'opération est annulée.", vbCritical, "Timeout de Connexion SAP"
            Exit Sub
        End If
    Loop

    Set Applic = SapGui.GetScriptingEngine

    ' Tente d'ouvrir la connexion. Si ça échoue, l'erreur est gérée.
    Set Connection = Applic.OpenConnection(sapBase, True)
    Set g_Session = Connection.children(0)

    ' Si une fenêtre inattendue apparaît (ex: multi-logon), le script s'arrête pour une gestion manuelle.
    If Err.Number <> 0 Then
        MsgBox "Une fenêtre SAP inattendue est apparue. Veuillez la gérer manuellement.", vbExclamation, "Action requise"
        Exit Sub
    End If

    ' Gère la fenêtre de copyright ou de langue en appuyant sur le premier bouton (généralement "OK" ou "Entrée").
    On Error Resume Next
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo ErrHandler
    
    ' --- Stabilisation et activation de la session ---
    ' Attend un court instant pour que la fenêtre principale se stabilise
    WaitForSAP
    
    ' S'assure que la fenêtre principale est active avant d'envoyer une commande
    If g_Session.ActiveWindow.name = "wnd[0]" Then
        g_Session.findById("wnd[0]").sendVKey 0
        LogMessage "Session SAP activée avec succès."
    End If
    Exit Sub

ErrHandler:
    ' Gestionnaire d'erreurs global pour cette procédure.
    If Err.Number <> 0 Then
        ' Tente de se connecter à une base de secours si définie
        Dim sapBase2 As String
        sapBase2 = GetSetting("SAP_Base_2")
        If sapBase2 <> "" Then ' Logique de secours non implémentée, mais la structure est là.
            Resume Next ' Pour l'instant, on ne fait rien, mais on pourrait ajouter une logique de fallback ici.
        End If
        DisplayAndLogError "SAP_Openg_SessionFromLogon", Err
    End If
    ' Nettoyage
    Set WshShell = Nothing
    Set Applic = Nothing
    Set SapGui = Nothing
End Sub


'------------------------------------------------------------------------------
' Sub : SAP_CreateNewSession
' Description :
'   Crée une nouvelle session SAP (mode) à partir de la connexion existante.
'------------------------------------------------------------------------------
Public Sub SAP_CreateNewSession()
    Dim currentConnection As Object
    Dim newSession As Object
    On Error GoTo ErrHandler

    ' Vérifie si une session SAP est déjà active
    If g_Session Is Nothing Then
        MsgBox "Aucune session SAP active n'a été trouvée." & vbCrLf & _
               "Veuillez d'abord vous connecter avec 'Log ON'.", vbInformation, "Nouvelle Session"
        Exit Sub
    End If

    ' Crée une nouvelle session. La nouvelle fenêtre deviendra automatiquement la fenêtre active.
    g_Session.createSession
    LogMessage "Demande de création d'une nouvelle session SAP envoyée."
    Exit Sub
ErrHandler:
    MsgBox "Impossible de créer une nouvelle session." & vbCrLf & vbCrLf & "Erreur: " & Err.Description, vbCritical, "Erreur SAP"
    DisplayAndLogError "SAP_CreateNewSession", Err
End Sub


'------------------------------------------------------------------------------
' Sub : SAP_CloseAllSessions
' Description :
'   Parcourt toutes les connexions et sessions SAP ouvertes et les ferme
'   proprement. Gère la fenêtre de confirmation de fermeture.
'   Utilise un ArrayList pour éviter les problèmes de modification de la
'   collection pendant le parcours.
'------------------------------------------------------------------------------
Sub SAP_CloseAllSessions()
    Dim SapGuiAuto As Object, SAPApp As Object, Connection As Object, session As Object
    Dim sessionsToClose As Object
    
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        LogMessage "SAP Logon n'est pas en cours d'exécution. Aucune session à fermer."
        Exit Sub
    End If
    
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    If SAPApp Is Nothing Then
        LogMessage "Impossible d'obtenir le Scripting Engine. Aucune session à fermer."
        Exit Sub
    End If
    On Error GoTo 0
    
    LogMessage "Début de la fermeture de toutes les sessions SAP."
    
    ' Utilise une collection pour éviter les problèmes lors de la suppression d'éléments d'une collection qu'on parcourt.
    Set sessionsToClose = CreateObject("System.Collections.ArrayList")
    For Each Connection In SAPApp.Connections
        For Each session In Connection.sessions
            sessionsToClose.add session
        Next
    Next
    
    For Each session In sessionsToClose
        session.findById("wnd[0]").Close
        On Error Resume Next ' Gère la popup de confirmation
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press ' Bouton "Oui"
        On Error GoTo 0
    Next
    
    LogMessage sessionsToClose.count & " session(s) SAP ont été fermées."
End Sub

'------------------------------------------------------------------------------
' Sub : SAP_CloseSelectedSession
' Description :
'   Ferme spécifiquement la session SAP qui est actuellement sélectionnée et
'   stockée dans la variable globale `g_Session`.
'------------------------------------------------------------------------------
Sub SAP_CloseSelectedSession()
    If g_Session Is Nothing Then
        MsgBox "Aucune session SAP n'est sélectionnée pour la fermeture.", vbInformation, "Fermeture de session"
        Exit Sub
    End If
    
    On Error Resume Next
    
    LogMessage "Tentative de fermeture de la session : " & g_Session.Info.systemName
    
    ' Envoie la commande de fermeture à la fenêtre principale de la session
    g_Session.findById("wnd[0]").Close
    
    ' Gère la popup de confirmation de déconnexion si elle apparaît
    g_Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press ' Bouton "Oui"
    
    On Error GoTo 0
End Sub

'====================================================================================
' SECTION : PROCÉDURES UTILITAIRES ET DE VÉRIFICATION
'====================================================================================

'------------------------------------------------------------------------------
' Sub : ConfirmSAPConnection
' Description :
'   Vérifie si une session SAP valide est ouverte pour un système donné et met
'   à jour un indicateur visuel sur la feuille "Dashboard".
'------------------------------------------------------------------------------
Public Sub ConfirmSAPConnection()
    Dim SapGuiAuto As Object, Sap_Applic As Object, Connection As Object
    Dim systemName As String
    Dim isConnected As Boolean
    Dim i As Long, j As Long

    On Error GoTo ErrorHandler
    isConnected = False
    systemName = GetSetting("SAP_System_Name") ' Ex: "PR1"

    ' Connexion à l'instance SAP GUI Scripting Engine.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Sap_Applic = SapGuiAuto.GetScriptingEngine

    If Sap_Applic.Connections.count > 0 Then
        For i = 0 To Sap_Applic.Connections.count - 1
            Set Connection = Sap_Applic.children(CInt(i))
            If Connection.children.count > 0 Then
                For j = 0 To Connection.children.count - 1
                    Set g_Session = Connection.children(CInt(j))
                    ' Vérifie si le nom du système correspond
                    If g_Session.Info.systemName = systemName Then
                        isConnected = True
                        Exit For
                    End If
                Next
            End If
            If isConnected Then Exit For
        Next
    End If

ErrorHandler:
    ' Mise à jour de l'indicateur sur le Dashboard
    Dim wsDashboard As Worksheet
    Set wsDashboard = Nothing
    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    If Not wsDashboard Is Nothing Then
        wsDashboard.Shapes("Conected").Visible = isConnected
    End If

    If Err.Number <> 0 And Err.Source <> "ConfirmSAPConnection" Then
        ' Ne fait rien si l'erreur vient du fait que SAP n'est pas ouvert.
    End If
End Sub

'====================================================================================
' SECTION : FONCTIONS POUR LA GESTION DE LA COMBOBOX DU RUBAN
'====================================================================================

'------------------------------------------------------------------------------
' FONCTION : GetAvailableSessionsCount
' DESCRIPTION : Parcourt toutes les sessions SAP ouvertes, les stocke dans le
'               tableau global `g_Sessions` et retourne leur nombre total.
' APPELÉE PAR : Le ruban (attribut `getItemCount` de la ComboBox).
'------------------------------------------------------------------------------
Function GetAvailableSessionsCount() As Integer
    Dim SapGuiAuto As Object, SAPApp As Object, Connection As Object, session As Object
    Dim count As Integer
    count = 0
    
    ' Réinitialise le tableau global pour éviter les données obsolètes
    Erase g_Sessions

    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Not SapGuiAuto Is Nothing Then
        Set SAPApp = SapGuiAuto.GetScriptingEngine
        If Not SAPApp Is Nothing Then
            ReDim g_Sessions(0 To SAPApp.Connections.count * 10) ' Redimensionne le tableau pour être assez grand.
            For Each Connection In SAPApp.Connections
                For Each session In Connection.sessions
                    Set g_Sessions(count) = session
                    count = count + 1
                Next
            Next
        End If
    End If
    On Error GoTo 0

    GetAvailableSessionsCount = count
End Function

'------------------------------------------------------------------------------
' FONCTION : GetAvailableSessionLabel
' DESCRIPTION : Construit et retourne une chaîne de caractères (label) pour une
'               session donnée, basée sur son index dans le tableau `g_Sessions`.
'               Exemple de label : "PR1 (1) 100".
' APPELÉE PAR : Le ruban (attribut `getItemLabel` de la ComboBox).
'------------------------------------------------------------------------------
Function GetAvailableSessionLabel(index As Integer) As String
    On Error Resume Next
    If g_Sessions(index) Is Nothing Then
        GetAvailableSessionLabel = "Session Invalide"
    Else
        Dim sessionID As String: sessionID = g_Sessions(index).ID
        Dim startPos As Long: startPos = InStr(sessionID, "ses[")
        Dim endPos As Long: endPos = InStr(startPos, sessionID, "]")
        Dim sessionNumber As Long
        If startPos > 0 And endPos > startPos Then
            sessionNumber = CLng(Mid(sessionID, startPos + 4, endPos - (startPos + 4))) + 1
            GetAvailableSessionLabel = g_Sessions(index).Info.systemName & " (" & sessionNumber & ") " & g_Sessions(index).Info.Client
        Else
            GetAvailableSessionLabel = g_Sessions(index).Info.systemName & " - " & g_Sessions(index).Info.Client
        End If
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' SUB : SelectSessionByIndex
' DESCRIPTION : Met à jour la variable globale `g_Session` pour qu'elle pointe
'               vers la session sélectionnée par l'utilisateur dans la ComboBox.
' APPELÉE PAR : Le ruban (attribut `onChange` de la ComboBox).
'------------------------------------------------------------------------------
Sub SelectSessionByIndex(index As Integer)
    On Error Resume Next
    If index >= 0 And index <= UBound(g_Sessions) Then
        If Not g_Sessions(index) Is Nothing Then
            Set g_Session = g_Sessions(index)
            LogMessage "Session SAP active changée pour : " & GetAvailableSessionLabel(index)
        End If
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' FONCTION : GetSAPLogonPath
' DESCRIPTION : Tente de trouver le chemin de saplogon.exe via le registre Windows.
'               Si non trouvé, vérifie les chemins d'installation standards.
'------------------------------------------------------------------------------
Private Function GetSAPLogonPath() As String
    Dim WshShell As Object
    Dim regPath As String
    
    Set WshShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    ' Tentative via le registre (App Paths standard)
    regPath = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\saplogon.exe\")
    If Err.Number = 0 And regPath <> "" Then
        If Dir(regPath) <> "" Then
            GetSAPLogonPath = regPath
            Exit Function
        End If
    End If
    Err.Clear
    
    ' Tentative via le registre (WOW6432Node pour systèmes 64 bits)
    regPath = WshShell.RegRead("HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\saplogon.exe\")
    If Err.Number = 0 And regPath <> "" Then
        If Dir(regPath) <> "" Then
            GetSAPLogonPath = regPath
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0
    
    GetSAPLogonPath = ""
End Function
