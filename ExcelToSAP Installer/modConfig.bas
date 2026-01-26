Attribute VB_Name = "modConfig"
'====================================================================================
' MODULE      : modConfig
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module est le centre névralgique de l'application ExcelToSAP.
'               Il gère :
'                 1. Le chargement de tous les paramètres de configuration depuis la feuille "Setup".
'                 2. Le stockage de ces paramètres dans un dictionnaire global (`ConfigSettings`).
'                 3. La gestion des variables d'état globales qui suivent l'exécution du programme.
'                 4. La vérification et la gestion de la session SAP.
'====================================================================================

Option Explicit

'====================================================================================
' SECTION 1 : DÉCLARATIONS
'====================================================================================

' --- 1.1 Déclarations API Windows ---
'====================================================================================
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetKeyState Lib "user32" _
        (ByVal nVirtKey As Long) As Integer
#Else
    Public Declare Function GetKeyState Lib "user32" _
        (ByVal nVirtKey As Long) As Integer
#End If

' --- 1.2 Variables Globales ---
'====================================================================================

' Dictionnaire de configuration (chargé au démarrage)
Public ConfigSettings As Object

' Variables de session
Public g_Session As Object          ' La session SAP GUI Scripting active.
Public g_Sessions() As Object       ' Tableau dynamique pour stocker les instances de sessions SAP.

' Variables d'état (Runtime)
Public g_DataType As String         ' Type de données en cours de traitement.
Public g_FirstRow As Long           ' Première ligne de la sélection.
Public g_LastRow As Long            ' Dernière ligne de la sélection.
Public g_IsList As Boolean          ' Indicateur si la sélection est une liste.
Public g_DoNotRun As Boolean        ' Flag pour empêcher l'exécution (contrôlé par le ruban).
Public g_IsTestMode As Boolean      ' NOUVEAU: Flag pour indiquer si la suite de tests est en cours.
Public g_StopTests As Boolean       ' Flag pour arrêter la suite de tests (contrôlé par l'utilisateur).
Public g_FormTracker As Integer     ' Suivi de l'état d'un formulaire.
Public g_MaintPlanID As String      ' ID du plan de maintenance en cours.
Public CB As Object                 ' Référence à un objet CommandButton (MSForms).
Public mplan As String              ' ID du plan de maintenance (potentiellement redondant avec g_MaintPlanID).

' --- 1.3 Constantes du Projet ---
'====================================================================================
Public Const c_Key_NumLock As Long = 144                        ' Code virtuel de la touche NumLock.
Public Const c_SetupSheetName As String = "Setup"               ' Nom de la feuille de configuration.
Public Const c_DefaultSAPPath As String = "C:\00-Box_Flow"      ' Chemin SAP par défaut (utilisé si non trouvé dans Setup).
Public Const c_DefaultLogPath As String = "C:\TEMP\ExcelToSAP_Log.txt" ' Chemin du fichier de log par défaut.

'====================================================================================
' SECTION 2 : GESTION DE LA CONFIGURATION
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : LoadConfiguration
' DESCRIPTION : Lit les paramètres depuis la feuille "Setup" du classeur et les charge
'               dans le dictionnaire global `ConfigSettings`. La lecture se fait
'               par blocs de colonnes (Clé/Valeur).
' NOTES       : C'est une procédure critique. Si elle échoue, l'application s'arrête.
'------------------------------------------------------------------------------------
Public Sub LoadConfiguration()
    Dim wsSetup As Worksheet
    Dim wb As Workbook
    Dim settingName As String
    Dim settingValue As Variant
    Dim blockColumns As Variant
    Dim block As Variant
    Dim keyCol As Long, valCol As Long, lastrow As Long, i As Long
    
    On Error GoTo ErrHandler
    
    ' 1. Initialisation du dictionnaire
    If ConfigSettings Is Nothing Then
        Set ConfigSettings = CreateObject("Scripting.Dictionary")
    Else
        ConfigSettings.RemoveAll ' Vide le dictionnaire avant de recharger
    End If
    
    ' 2. Validation des objets Excel (Classeur et Feuille)
    On Error Resume Next
    Set wb = ThisWorkbook
    
    Set wsSetup = wb.Sheets(c_SetupSheetName)
    If wsSetup Is Nothing Then
        MsgBox "Erreur critique : La feuille '" & c_SetupSheetName & "' est introuvable dans le classeur '" & wb.name & "'.", vbCritical, "Erreur de Configuration"
        Exit Sub
    End If
    On Error GoTo ErrHandler ' Réactive la gestion d'erreur standard
    
    ' 3. Lecture des paramètres depuis plusieurs blocs de colonnes
    '    Définit les colonnes de début pour chaque bloc (A=1, D=4, etc.)
    blockColumns = Array(1, 4, 7, 10) ' Correspond aux colonnes A, D, G
    
    For Each block In blockColumns
        keyCol = CLng(block)
        valCol = keyCol + 1
        
        ' Trouve la dernière ligne non vide pour ce bloc spécifique
        lastrow = wsSetup.Cells(wsSetup.Rows.count, keyCol).End(xlUp).row
        
        ' Boucle sur chaque ligne du bloc pour lire les paires Clé/Valeur
        For i = 1 To lastrow
            settingName = CStr(wsSetup.Cells(i, keyCol).value)
            settingValue = wsSetup.Cells(i, valCol).value

            ' Ajoute le paramètre au dictionnaire si le nom n'est pas vide
            If Len(settingName) > 0 Then
                If Not ConfigSettings.Exists(settingName) Then
                    ConfigSettings.add settingName, settingValue
                    ' LogMessage "Config chargée: " & settingName & " = " & settingValue ' Optionnel: décommenter pour le débogage
                End If
            End If
        Next i
    Next block
    
    ' S'assure que "Planners" est bien supprimé du dictionnaire (nettoyage)
    If ConfigSettings.Exists("Planners") Then ConfigSettings.Remove "Planners"
    
    ' --- HARDCODED PATHS OVERRIDE ---
    Dim commonPath As String
    commonPath = ConfigSettings("PATH")
    
    UpdateSetting "Path file for Analyse", commonPath
    UpdateSetting "KPI_PATH", commonPath
    UpdateSetting "PATH_EXPORT", commonPath
    UpdateSetting "PATH_LOGS", commonPath
    
    UpdateSetting "DASHBOARD_PATH", "C:\Dashboard-KPIs\"

    Exit Sub
    
ErrHandler:
    MsgBox "Une erreur critique est survenue lors de la lecture de la configuration." & vbCrLf & vbCrLf & _
           "Module: modConfig.LoadConfiguration" & vbCrLf & _
           "Détails: " & Err.Description, vbCritical, "Erreur de Configuration"
    Exit Sub ' Arrête l'exécution en cas d'échec de chargement de la config
End Sub

'------------------------------------------------------------------------------------
' FONCTION    : GetSetting
' DESCRIPTION : Récupère une valeur de configuration depuis le dictionnaire global.
' PARAMÈTRES  : key (String) - Le nom du paramètre à récupérer.
' RETOUR      : Variant - La valeur du paramètre, ou une chaîne vide si non trouvé.
'------------------------------------------------------------------------------------
Public Function GetSetting(ByVal key As String) As Variant
    If ConfigSettings.Exists(key) Then
        GetSetting = ConfigSettings(key)
    Else
        GetSetting = "" ' Retourne une valeur par défaut (chaîne vide)
        Debug.Print "AVERTISSEMENT: Le paramètre de configuration '" & key & "' n'a pas été trouvé."
    End If
End Function


'------------------------------------------------------------------------------------
' PROCÉDURE   : UpdateSetting
' DESCRIPTION : Met à jour une valeur de configuration dans le dictionnaire en mémoire.
' NOTE        : Cette procédure ne sauvegarde PAS la modification dans la feuille Excel.
'               Utiliser `SaveSettings` pour rendre le changement permanent.
'------------------------------------------------------------------------------------
Public Sub UpdateSetting(ByVal key As String, ByVal newValue As Variant)
    If ConfigSettings Is Nothing Then
        Set ConfigSettings = CreateObject("Scripting.Dictionary")
    End If

    If ConfigSettings.Exists(key) Then
        ConfigSettings(key) = newValue
    Else
        ConfigSettings.add key, newValue
    End If
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : SaveSettings
' DESCRIPTION : Sauvegarde toutes les valeurs du dictionnaire `ConfigSettings`
'               dans la feuille "Setup", en retrouvant chaque clé dans les
'               colonnes de configuration et en mettant à jour la valeur adjacente.
'------------------------------------------------------------------------------------
Public Sub SaveSettings()
    Dim wsSetup As Worksheet
    Dim wb As Workbook
    Dim key As Variant
    Dim foundCell As Range
    Dim blockColumns As Variant
    Dim block As Variant
    Dim searchRange As Range

    On Error GoTo ErrHandler

    ' 1. Validation des objets Excel
    On Error Resume Next
    Set wb = ThisWorkbook
    
    Set wsSetup = wb.Sheets(c_SetupSheetName)
    If wsSetup Is Nothing Then GoTo ErrHandler
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' 2. Définit les colonnes où chercher les clés (doit correspondre à LoadConfiguration)
    blockColumns = Array(1, 4, 7, 10) ' Colonnes A, D, G, J

    ' 3. Boucle sur chaque paramètre du dictionnaire pour le sauvegarder
    For Each key In ConfigSettings.Keys
        ' Cherche la clé dans chaque colonne de bloc définie
        For Each block In blockColumns
            Set searchRange = wsSetup.Columns(CLng(block))
            Set foundCell = searchRange.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            
            ' Si la clé est trouvée, met à jour la valeur dans la cellule adjacente
            If Not foundCell Is Nothing Then
                foundCell.Offset(0, 1).value = ConfigSettings(key)
                Exit For ' Clé trouvée, passe au paramètre suivant
            End If
        Next block
    Next key

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Une erreur est survenue lors de la sauvegarde des paramètres dans la feuille 'Setup'." & vbCrLf & vbCrLf & _
           "Module: modConfig.SaveSettings" & vbCrLf & _
           "Détails: " & Err.Description, vbCritical, "Erreur de Sauvegarde"
End Sub


'====================================================================================
' SECTION 3 : GESTION DE L'ÉTAT GLOBAL ET DE LA SESSION
'====================================================================================
'------------------------------------------------------------------------------------
' PROCÉDURE   : Init_GlobalState
' DESCRIPTION : Réinitialise toutes les variables d'état globales à leur valeur par défaut.
'               Appelée généralement au début d'une nouvelle opération.
'------------------------------------------------------------------------------------
Public Sub Init_GlobalState()
    On Error Resume Next

    Set g_Session = Nothing
    g_DataType = ""
    g_FirstRow = 0
    g_LastRow = 0
    g_IsList = False

    g_DoNotRun = False
    g_FormTracker = 0
    g_MaintPlanID = ""

    Set CB = Nothing

    On Error GoTo 0
End Sub


'------------------------------------------------------------------------------------
' FONCTION    : EnsureSAPSession
' DESCRIPTION : Vérifie la présence d'une session SAP active. Si `g_Session` est vide,
'               tente de se connecter à la première session SAP disponible.
' PARAMÈTRES  : ShowMsg (Boolean, Optional) - Si True, affiche un message si aucune
'               session n'est trouvée.
' RETOUR      : Boolean - True si une session SAP est active, sinon False.
'------------------------------------------------------------------------------------
Public Function EnsureSAPSession(Optional ShowMsg As Boolean = True) As Boolean
    On Error Resume Next

    ' Si aucune session n?est enregistr?e, tentative de connexion automatique
    If g_Session Is Nothing Then
        Dim SapGuiAuto As Object, App As Object
        Set SapGuiAuto = GetObject("SAPGUI")
        If Not SapGuiAuto Is Nothing Then
            Set App = SapGuiAuto.GetScriptingEngine
            If Not App Is Nothing Then
                Set g_Session = App.children(0).children(0)
            End If
        End If
    End If

    On Error GoTo 0

    ' V?rification finale
    If g_Session Is Nothing Then
        EnsureSAPSession = False
        If ShowMsg Then MsgBox "Aucune session SAP active n'a été détectée." & vbCrLf & _
                               "Veuillez ouvrir SAP avant de relancer ExcelToSAP.", _
                               vbExclamation, "ExcelToSAP - Session SAP manquante"
    Else
        EnsureSAPSession = True
    End If
End Function

