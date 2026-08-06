Attribute VB_Name = "modInit"
'====================================================================================
' MODULE      : modInit
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module gère toutes les opérations d'initialisation et de gestion
'               du cycle de vie de l'add-in ExcelToSAP. Il contient les procédures
'               critiques exécutées à l'ouverture d'Excel, ainsi que les routines
'               pour l'installation et la gestion de l'état du classeur.
'
' PROCÉDURES CLÉS :
'   - Auto_Open()     : Exécutée automatiquement à l'ouverture de l'add-in pour
'                     charger la configuration.
'   - Install_AddIn() : Gère l'installation et la mise à jour de l'add-in (.xlam).
'   - Hide_Setup()    : Masque le classeur de configuration en le traitant comme un add-in.
'====================================================================================

Option Explicit

'====================================================================================
' SECTION 1 : PROCÉDURES DE CYCLE DE VIE DE L'ADD-IN
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Auto_Open
' DESCRIPTION : Exécutée automatiquement à l'ouverture de l'add-in. Son rôle principal
'               est de charger la configuration initiale nécessaire au bon
'               fonctionnement de l'application.
' DÉPENDANCES : LoadConfiguration, LogMessage, HandleError.
'------------------------------------------------------------------------------------
Public Sub auto_open()
    On Error GoTo ErrHandler
    LogMessage "Initialisation du module : Auto_Open"

    ' 1. Chargement de la configuration depuis la feuille "Setup"
    Call LoadConfiguration

    ' 1.1 Vérification et création des dossiers requis au démarrage.
    Call EnsureRequiredFoldersExist

    ' 2. Vérification de la présence du classeur add-in
    Dim wbAddin As Workbook
    On Error Resume Next
    Set wbAddin = ThisWorkbook
    On Error GoTo ErrHandler

    LogMessage "Auto_Open terminé avec succès."
    Exit Sub

ErrHandler:
    DisplayAndLogError "Auto_Open", Err
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : EnsureRequiredFoldersExist (Privée)
' DESCRIPTION : Vérifie que les dossiers principaux (TEMP et Dashboard) existent.
'               Si ce n'est pas le cas, tente de les créer.
'               Appelée une seule fois au démarrage par Auto_Open.
' DÉPENDANCES : GetSetting, LogMessage, DisplayAndLogError.
'------------------------------------------------------------------------------------
Private Sub EnsureRequiredFoldersExist()
    On Error GoTo ErrHandler
    
    Dim tempPath As String
    Dim dashboardPath As String
    
    ' Récupère les chemins depuis la configuration.
    ' Ces chemins sont surchargés en dur dans modConfig.LoadConfiguration.
    tempPath = GetSetting("PATH_LOGS") ' Représente C:\TEMP\
    dashboardPath = GetSetting("DASHBOARD_PATH") ' Représente C:\Dashboard-KPIs\
    
    ' Crée le dossier TEMP s'il n'existe pas
    If tempPath <> "" And Dir(tempPath, vbDirectory) = "" Then
        MkDir tempPath
        LogMessage "Dossier requis créé : " & tempPath
    End If
    
    ' Crée le dossier Dashboard s'il n'existe pas
    If dashboardPath <> "" And Dir(dashboardPath, vbDirectory) = "" Then
        MkDir dashboardPath
        LogMessage "Dossier requis créé : " & dashboardPath
    End If
    Exit Sub

ErrHandler:
    DisplayAndLogError "EnsureRequiredFoldersExist", Err
End Sub


'====================================================================================
' SECTION 2 : ROUTINES D'INSTALLATION ET DE CONFIGURATION
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Install_AddIn
' DESCRIPTION : Gère l'installation ou la mise à jour de l'add-in ExcelToSAP.
'               Le processus inclut la désinstallation de l'ancienne version, la
'               sauvegarde du classeur actuel au format .xlam, et l'installation
'               de la nouvelle version.
' DÉPENDANCES : LogMessage, HandleError.
'------------------------------------------------------------------------------------
Public Sub Install_AddIn()
    On Error GoTo ErrHandler
    LogMessage "Démarrage de l'installation de l'add-in ExcelToSAP."

    Dim macroWBD As Workbook
    Dim tempWBD As Workbook
    Dim oAddin As AddIn
    Dim sFile As String
    Dim newPath As String
    Dim boxOptions As VbMsgBoxResult
    Dim addOnFileName As String

    addOnFileName = "ExcelToSAP Installer"
    newPath = Application.UserLibraryPath

    ' 1. Demande de confirmation à l'utilisateur
    boxOptions = MsgBox("Voulez-vous installer ou mettre à jour l'add-in ExcelToSAP ?" & vbCrLf & _
                        "(L'ancienne version sera remplacée)", vbYesNo + vbQuestion)
    If boxOptions = vbNo Then Exit Sub

    ' 2. Désinstallation et suppression de l'ancienne version (si elle existe)
    On Error Resume Next
    AddIns(addOnFileName).Installed = False
    Kill newPath & addOnFileName & ".xlam"
    On Error GoTo ErrHandler

    ' 3. Sauvegarde du classeur actuel au format Add-in (.xlam)
    Set macroWBD = ThisWorkbook
    sFile = newPath & addOnFileName & ".xlam"

    Application.DisplayAlerts = False
    macroWBD.SaveAs fileName:=sFile, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True

    ' 4. Ajout et installation du nouvel add-in
    Set tempWBD = Workbooks.add
    Set oAddin = AddIns.add(fileName:=sFile, CopyFile:=False)
    oAddin.Installed = True
    tempWBD.Close False

    MsgBox "Installation réussie !" & vbCrLf & _
           "Redémarrez Excel pour activer l'add-in.", vbInformation
    LogMessage "Installation " & addOnFileName & " terminée avec succès."

    macroWBD.Close SaveChanges:=False
    Exit Sub

ErrHandler:
    DisplayAndLogError "Install_AddIn", Err
End Sub



'====================================================================================
' SECTION 3 : PROCÉDURES UTILITAIRES DE CONFIGURATION
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Hide_Setup
' DESCRIPTION : Marque le classeur de configuration comme un add-in, ce qui le masque
'               de l'interface utilisateur d'Excel, et sauvegarde les modifications.
'               C'est l'action finale après avoir modifié la feuille "Setup".
' DÉPENDANCES : LogMessage, HandleError.
'------------------------------------------------------------------------------------
Public Sub Hide_Setup()
    On Error GoTo ErrHandler
    LogMessage "Masquage du classeur Setup ExcelToSAP."

    Dim wbAddin As Workbook
    
    Set wbAddin = ThisWorkbook

    If Right(LCase(wbAddin.name), 5) = ".xlam" Then
        wbAddin.IsAddin = True
    End If
    wbAddin.Save

    LogMessage "Setup masqué et sauvegardé avec succès."
    Exit Sub

ErrHandler:
    DisplayAndLogError "Hide_Setup", Err
End Sub

