Attribute VB_Name = "B_Load_Files"
Option Explicit
'====================================================================================
' MODULE      : B_Load_Files
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module centralise et orchestre l'exportation de diverses données
'               SAP vers des fichiers Excel. Il gère la connexion SAP, la configuration
'               des rapports spécifiques et la sauvegarde des données extraites.
'
' DÉPENDANCES EXTERNES (Variables Globales) :
'   - g_Session          : Objet de session SAP GUI (initialisé par 'onSAP').
'   - g_DebutAnnee       : Date de début de période pour les filtres SAP (format string).
'   - g_FinAnnee         : Date de fin de période pour les filtres SAP (format string).
'   - SemaineKPIs        : Semaine KPI courante (définie dans 'KPI_Manager.bas', format string).
'   - g_DebutMois        : Date de début de mois pour les rapports SES (format string).
'   - g_FinMois          : Date de fin de mois pour les rapports SES (format string).
'   - g_ribbonUI         : Objet IRibbonUI pour la gestion du ruban (dans 'RubanX.bas').
'
' DÉPENDANCES EXTERNES (Procédures/Fonctions) :
'   - modUtils.InitLogs  : Initialise le système de journalisation.
'   - modUtils.LogMessage: Enregistre un message dans le journal.
'   - modUtils.SaveFile  : Sauvegarde un fichier Excel.
'   - modUtils.FermerFichierExcel : Ferme un fichier Excel.
'   - modUtils.ExportToExcel : Exporte les données SAP vers Excel.
'   - modConfig.LoadConfiguration : Charge les paramètres de configuration.
'   - SAP.IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'   - onSAP              : Établit une connexion SAP.
'   - offSAP             : Ferme la connexion SAP.
'   - GetSetting         : Récupère un paramètre de configuration.
'   - DateSemain         : Met à jour les variables de date globales (non fournie dans ce module).
'   - RestoreExcel       : Restaure les paramètres d'Excel après l'exécution.
'====================================================================================

'====================================================================================
' ROUTINE PRINCIPALE D'EXPORTATION
'====================================================================================

Private Sub Clear_Export_Files()
    Dim exportPath As String
    Dim fileList As Variant
    Dim vFile As Variant
    
    On Error GoTo ErrorHandler
    exportPath = GetSetting("PATH_EXPORT")
    
    If exportPath = "" Then
        LogMessage "Avertissement: PATH_EXPORT n'est pas défini. Le nettoyage des fichiers d'export est ignoré."
        Exit Sub
    End If
    
    fileList = Array( _
        "MR.xlsx", "WO.xlsx", "OP.xlsx", "GM.xlsx", "KPI.xlsx", _
        "PR.xlsx", "PO.xlsx", "SES.xlsx", "RSV.XLSX", "CNF.XLSX", "PMR.xlsx" _
    )
    
    LogMessage "Début du nettoyage des fichiers d'export..."
    
    For Each vFile In fileList
        ' Appelle la procédure utilitaire qui gère la vérification et le nettoyage
        ClearFileIfExists exportPath, CStr(vFile)
    Next vFile
    
    LogMessage "Nettoyage des fichiers d'export terminé."
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    DisplayAndLogError "Clear_Export_Files", Err
End Sub

Public Sub Export_AllReports()
    On Error GoTo SapErrorHandler

    OptimizeExcel

    LoadConfiguration ' Charge les chemins et parametres depuis la feuille "Setup"

    modUtils.InitLogs
    modUtils.LogMessage "===== D?BUT EXPORTS SAP ====="

    ' Vider les fichiers de destination avant de lancer les extractions
    Clear_Export_Files
    
    ' --- Exécution séquentielle des exports principaux ---
    Z_Load_MR
    Z_Load_WO
    Z_Load_Op
    Z_Load_GM
    Z_Load_KPI
    Z_Load_PR
    Z_Load_PO
    Z_Load_SES
    Z_Load_RSV
    Z_Load_CNF
    Z_Load_PMR
    
    modUtils.LogMessage "===== EXPORTS TERMINÉS ====="
    Application.StatusBar = "Export terminé. Prêt."
    MsgBox "Tous les exports SAP sont termin?s.", vbInformation, "Export termin?"

CleanExit:
    ' --- Restauration des parametres Excel ---
    RestoreExcel
    Exit Sub ' Sortie normale de la proc?dure

SapErrorHandler:
    ' Gestionnaire d'erreurs pour l'ensemble du processus d'exportation
    DisplayAndLogError "Export_AllReports", Err
    Application.StatusBar = "Une erreur est survenue. L'exportation a ?t? interrompue."
End Sub

'====================================================================================
' SECTION : EXPORTS SAP SPÉCIFIQUES (Appels à ExecuteAndExportReport)
' Ces procédures sont des wrappers pour la fonction générique d'exportation.
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_MR
' DESCRIPTION : Charge et exporte le rapport des Demandes de Maintenance (MR).
' DÉPENDANCES : ExecuteAndExportReport, Configure_MR_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_MR()
    ExecuteAndExportReport "MR", "Z_IW69", "Configure_MR_Report", "MR.xlsx"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_WO
' DESCRIPTION : Charge et exporte le rapport des Ordres de Travail (WO).
' DÉPENDANCES : ExecuteAndExportReport, Configure_WO_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_WO()
    ExecuteAndExportReport "WO", "Z_IW39", "Configure_WO_Report", "WO.xlsx"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_Op
' DESCRIPTION : Charge et exporte le rapport des Opérations.
' DÉPENDANCES : ExecuteAndExportReport, Configure_Op_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_Op()
    ExecuteAndExportReport "Op", "Z_IW49", "Configure_Op_Report", "OP.xlsx"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_GM
' DESCRIPTION : Charge et exporte le rapport des Mouvements de Marchandises (GM).
' DÉPENDANCES : ExecuteAndExportReport, Configure_GM_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_GM()
    ExecuteAndExportReport "GM", "Z_IW39", "Configure_GM_Report", "GM.xlsx"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_PMR
' DESCRIPTION : Charge et exporte le rapport des Routines de Maintenance Planifiée (PMR).
' DÉPENDANCES : ExecuteAndExportReport, Configure_PMR_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_PMR()
    ExecuteAndExportReport "PMR", "Z_ZPM004", "Configure_PMR_Report", "PMR.xlsx"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_RSV
' DESCRIPTION : Charge et exporte le rapport des Réservations (RSV).
' DÉPENDANCES : ExecuteAndExportReport, Configure_Reservation_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_RSV()
    ExecuteAndExportReport "RSV", "Z_MB25", "Configure_Reservation_Report", "RSV.XLSX"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_CNF
' DESCRIPTION : Charge et exporte le rapport des Confirmations (CNF).
' DÉPENDANCES : ExecuteAndExportReport, Configure_Confirmations_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_CNF()
    ExecuteAndExportReport "CNF", "Z_IW47", "Configure_Confirmations_Report", "CNF.XLSX"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_SES
' DESCRIPTION : Charge et exporte le rapport des Feuilles de Saisie des Services (SES).
' DÉPENDANCES : ExecuteAndExportReport, Configure_SES_Report.
'------------------------------------------------------------------------------------
Public Sub Z_Load_SES()
    ExecuteAndExportReport "SES", "", "Configure_SES_Report", "SES.xlsx"
End Sub

'====================================================================================
' SECTION : EXPORTS SAP SPÉCIFIQUES (Logique personnalisée)
' Ces procédures ont un flux d'exécution unique qui ne peut pas être généralisé.
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_KPI
' DESCRIPTION : Charge et exporte le rapport des Indicateurs Clés de Performance (KPI)
'               depuis SAP. Cette procédure a un flux spécifique et ne peut pas
'               être factorisée par `ExecuteAndExportReport`.
' DÉPENDANCES : DateSemain, onSAP, Z_KPIP_SM, GetSetting, modUtils.SaveFile,
'               offSAP, modUtils.FermerFichierExcel, SAP.IsSAPConnectionAlive, RestoreExcel.
'------------------------------------------------------------------------------------
Public Sub Z_Load_KPI()
    On Error GoTo SapErrorHandler
    DateSemain
    Run "onSAP"
    ' La connexion SAP est vérifiée par la procédure 'onSAP'.
    
    ' Lance la transaction SAP spécifique pour les KPI
    Run "Z_KPIP_SM"
    
    ' --- Paramétrage des critères de sélection dans SAP ---
    With g_Session
        ' Définit la semaine KPI (basse et haute)
        .findById("wnd[0]/usr/ctxtR_WEEK-LOW").text = SemaineKPIs
        .findById("wnd[0]/usr/ctxtR_WEEK-HIGH").text = SemaineKPIs
        ' Définit la version KPI
        .findById("wnd[0]/usr/ctxtR_VERSI-LOW").text = "3"
        ' Exécute la sélection (bouton "Exécuter" ou "F8")
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Appuie sur un bouton de retour ou de confirmation
        .findById("wnd[0]/tbar[1]/btn[5]").press

        ' --- Chargement de la variante d'affichage (layout) ---
        ' Ouvre le menu contextuel pour les variantes
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_VARIANT"
        ' Sélectionne l'option "Charger"
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem "&LOAD"
        ' Sélectionne la première variante dans la liste
        g_Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    
        ' --- Filtrage par usines (plants) ---
        ' Positionne le curseur sur la colonne des usines
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell -1, "WERKSH"
        ' Sélectionne la colonne des usines
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectColumn "WERKSH"
        ' Applique un filtre
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").PressToolbarButton "&MB_FILTER"
        ' Ouvre la fenêtre de sélection multiple pour les valeurs du filtre
        g_Session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    
        ' Renseigne les usines à filtrer à partir des paramètres de configuration
        If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_PF")
        If GetSetting("SAP_PLANT_PT") <> "" Then g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = GetSetting("SAP_PLANT_PT")
        If GetSetting("SAP_PLANT_PD") <> "" Then g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = GetSetting("SAP_PLANT_PD")

        ' Valide le filtre et ferme les fenêtres de sélection
        g_Session.findById("wnd[2]/tbar[0]/btn[8]").press
        g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
        ' --- Sélection et extraction des données ---
        ' Sélectionne toutes les cellules du grid
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell -1, ""
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectAll
        
        ' Répète la sélection complète (parfois nécessaire pour s'assurer que tout est sélectionné)
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell -1, ""
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectAll
    
        ' Ouvre le menu contextuel pour l'exportation
        On Error Resume Next
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell 4, "OBJNRH"
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").contextMenu
        g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem "&XXL"
    
        ' Configure les options d'exportation (format, toujours afficher)
        g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
        g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
        g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").SetFocus
        g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").Selected = True
        g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    End With
    
    ' Sauvegarde le fichier exporté
    modUtils.SaveFile GetSetting("PATH_EXPORT"), "KPI.xlsx"

    ' Déconnexion SAP et nettoyage
    Run "offSAP"
    modUtils.FermerFichierExcel "KPI.xlsx"
    modUtils.LogMessage "? Z_Load_KPI terminé"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Z_Load_KPI", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_PR
' DESCRIPTION : Charge et exporte le rapport des Demandes d'Achat (PR) depuis SAP.
'               Cette procédure a un flux spécifique et ne peut pas être factorisée.
' DÉPENDANCES : DateSemain, onSAP, Z_ME5A, GetSetting, ExportToExcel,
'               offSAP, FermerFichierExcel, IsSAPConnectionAlive, RestoreExcel.
'------------------------------------------------------------------------------------
Public Sub Z_Load_PR()
    On Error GoTo SapErrorHandler
    DateSemain
    Run "onSAP"
    ' La connexion SAP est vérifiée par la procédure 'onSAP'.
    
    ' Lance la transaction SAP spécifique pour les Demandes d'Achat
    Run "Z_ME5A"
    
    ' --- Paramétrage des critères de sélection dans SAP ---
    With g_Session
        ' Coche les options "Demandes d'achat déjà traitées" et "Demandes d'achat libérées"
        .findById("wnd[0]/usr/chkP_ERLBA").Selected = True
        .findById("wnd[0]/usr/chkP_FREIG").Selected = True
        ' Définit l'usine à partir des paramètres de configuration
        .findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Définit le type d'imputation (ex: "F" pour Ordre de fabrication)
        .findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = "F"
        ' Exécute la sélection
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne une option de menu pour l'exportation (ex: "Liste -> Exporter -> Tableur")
        .findById("wnd[0]/mbar/menu[5]/menu[0]/menu[1]").Select
        ' Configure les options d'exportation (format, toujours afficher)
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    End With
    
    ' Exporte les données vers un fichier Excel
    ExportToExcel "PR.xlsx"

    ' Déconnexion SAP et nettoyage
    Run "offSAP"
    FermerFichierExcel "PR.xlsx"
    LogMessage "? Z_Load_PR terminé"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Z_Load_PR", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_Load_PO
' DESCRIPTION : Charge et exporte le rapport des Commandes d'Achat (PO) depuis SAP.
'               Cette procédure a un flux spécifique et ne peut pas être factorisée.
' DÉPENDANCES : DateSemain, onSAP, Configure_PO_Report, offSAP,
'               modUtils.FermerFichierExcel, SAP.IsSAPConnectionAlive, RestoreExcel.
'------------------------------------------------------------------------------------
Public Sub Z_Load_PO()
    On Error GoTo SapErrorHandler
    DateSemain
    Run "onSAP"
    ' La connexion SAP est vérifiée par la procédure 'onSAP'.
    
    ' Configure et exécute le rapport des Commandes d'Achat
    Configure_PO_Report
    
    ' Déconnexion SAP et nettoyage
    Run "offSAP"
    modUtils.FermerFichierExcel "PO.xlsx"
    modUtils.LogMessage "? Z_Load_PO terminé"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Z_Load_PO", Err
End Sub

'====================================================================================
' SECTION : CONFIGURATION DES RAPPORTS SAP
' Chaque procédure configure un écran de transaction SAP spécifique avant l'exportation.
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_MR_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Demandes de Maintenance (MR).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_MR_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Définit le type d'avis (MR_QMART)
        .findById("wnd[0]/usr/ctxtQMART-LOW").text = GetSetting("MR_QMART") ' Assuming MR_QMART is a setting
        ' Définit la période de début et de fin
        .findById("wnd[0]/usr/ctxtBEZDT-LOW").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtBEZDT-HIGH").text = g_FinAnnee
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "/LHMACSAPV7"
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne l'option d'exportation vers un tableur
        .findById("wnd[0]/mbar/menu[5]/menu[8]/menu[0]").Select
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_MR_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_WO_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Ordres de Travail (WO).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_WO_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Coche les options pour les ordres ouverts, en cours, terminés, etc.
        .findById("wnd[0]/usr/chkDY_OFN").Selected = True
        .findById("wnd[0]/usr/chkDY_IAR").Selected = True
        .findById("wnd[0]/usr/chkDY_MAB").Selected = True
        .findById("wnd[0]/usr/chkDY_HIS").Selected = True
        ' Définit la période de début et de fin
        .findById("wnd[0]/usr/ctxtDATUV").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtDATUB").text = g_FinAnnee
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "/LHSAPWOV7"
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne l'option d'exportation vers un tableur
        .findById("wnd[0]/mbar/menu[5]/menu[8]/menu[0]").Select
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_WO_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_Op_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Opérations.
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_Op_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Coche l'option "Opérations actives" et décoche "Historique"
        .findById("wnd[0]/usr/chkDY_AKT").Selected = True
        .findById("wnd[0]/usr/chkDY_HIS").Selected = False
        ' Définit la période de début et de fin
        .findById("wnd[0]/usr/ctxtADDAT-LOW").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtADDAT-HIGH").text = g_FinAnnee
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "/LHSAPWOV7"
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne l'option d'exportation vers un tableur
        .findById("wnd[0]/mbar/menu[5]/menu[8]/menu[0]").Select
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_Op_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_GM_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Mouvements de Marchandises (GM).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_GM_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Coche les options pour les ordres ouverts, en cours, terminés, etc. (similaire à WO)
        .findById("wnd[0]/usr/chkDY_OFN").Selected = True
        .findById("wnd[0]/usr/chkDY_IAR").Selected = True
        .findById("wnd[0]/usr/chkDY_MAB").Selected = True
        .findById("wnd[0]/usr/chkDY_HIS").Selected = True
        ' Définit la période de début et de fin
        .findById("wnd[0]/usr/ctxtDATUV").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtDATUB").text = g_FinAnnee
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "/LHSAPWOV7"
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne toutes les lignes du grid
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
        ' Sélectionne l'option d'exportation (via le menu)
        .findById("wnd[0]/mbar/menu[4]/menu[4]").Select
        .findById("wnd[0]/mbar/menu[4]/menu[8]/menu[0]").Select
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_GM_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_PMR_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Routines de Maintenance Planifiée (PMR).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_PMR_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        
        ' Définit la date de début
        .findById("wnd[0]/usr/ctxtP_SDATE").text = g_DebutAnnee
        
        ' Sélectionne le mode d'affichage "Grid" (tableau)
        .findById("wnd[0]/usr/radP_GRID").Select
        '.findById("wnd[0]/usr/radP_TREE").Select
        
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        ' Ouvre le menu contextuel du grid pour l'exportation
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_PMR_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_Reservation_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Réservations (RSV).
' DÉPENDANCES : g_Session, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_Reservation_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtWERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        
        ' Coche les options pour les réservations annulées et clôturées
        g_Session.findById("wnd[0]/usr/chkP_OPEN").Selected = True
        g_Session.findById("wnd[0]/usr/chkP_CANCEL").Selected = True
        g_Session.findById("wnd[0]/usr/chkP_CLOSED").Selected = True
        g_Session.findById("wnd[0]/usr/chkP_ISSUES").Selected = True
        g_Session.findById("wnd[0]/usr/chkP_RECEIP").Selected = True
        g_Session.findById("wnd[0]/usr/ctxtBDTER-LOW").text = g_DebutAnnee
        g_Session.findById("wnd[0]/usr/ctxtBDTER-HIGH").text = g_FinAnnee
        
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    
    Exit Sub
    
SapErrorHandler:
    DisplayAndLogError "Configure_Reservation_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_Confirmations_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Confirmations (CNF).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_Confirmations_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Coche les options pour les confirmations individuelles et les confirmations annulées
        .findById("wnd[0]/usr/chkDY_IAR").Selected = True
        .findById("wnd[0]/usr/chkDY_ABG").Selected = True
        ' Définit l'usine d'origine
        .findById("wnd[0]/usr/ctxtWERKS_O-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CNF_ALL"
        ' Définit la période de début et de fin pour la date de comptabilisation
        .findById("wnd[0]/usr/ctxtBUDAT_C-LOW").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtBUDAT_C-HIGH").text = g_FinAnnee
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_Confirmations_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_DemandeAchat_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Demandes d'Achat.
'               (Note: Cette procédure semble être un doublon partiel de Z_Load_PR
'               et n'est pas directement appelée par ExecuteAndExportReport).
' DÉPENDANCES : g_Session, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_DemandeAchat_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Coche les options pour les demandes d'achat déjà traitées et libérées
        .findById("wnd[0]/usr/chkP_ERLBA").Selected = True
        .findById("wnd[0]/usr/chkP_FREIG").Selected = True
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Définit le type d'imputation
        .findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = "F"
        ' Charge une variante d'affichage spécifique
        .findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CNF_ALL"
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne l'option d'exportation vers un tableur
        .findById("wnd[0]/mbar/menu[5]/menu[8]/menu[0]").Select
        .findById("wnd[0]/mbar/menu[5]/menu[0]/menu[1]").Select
        ' Configure les options d'exportation
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
        ' Ouvre le menu contextuel du grid pour l'exportation
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    End With
SapErrorHandler:
    DisplayAndLogError "Configure_DemandeAchat_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_PO_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Commandes d'Achat (PO)
'               via la transaction ME80FN et exporte les données.
' DÉPENDANCES : g_Session, GetSetting, modUtils.SaveFile, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_PO_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' === TRANSACTION ===
        ' Saisit le code de transaction ME80FN
        .findById("wnd[0]/tbar[0]/okcd").text = "/nme80fn"
        .findById("wnd[0]").sendVKey 0

        ' === PLANT + OPTIONS ===
        ' Définit l'usine et le type d'imputation
        .findById("wnd[0]/usr/ctxtSP$00011-LOW").text = GetSetting("SAP_PLANT_PF")
        .findById("wnd[0]/usr/ctxtSP$00004-LOW").text = "F"

        ' Selection multiple des types de documents d'achat
        ' Ouvre la fenêtre de sélection multiple pour les types de documents
        .findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press
        ' Renseigne les types de documents spécifiques (NB, ZURG, FO)
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "NB"
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ZURG"
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "FO"
        ' Valide la sélection
        .findById("wnd[1]/tbar[0]/btn[8]").press

        ' === EX?CUTION ===
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press

        ' === LAYOUT ===
        ' Ouvre le menu contextuel pour les variantes d'affichage
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
        ' Sélectionne l'option "Charger"
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem "&LOAD"
        ' Configure les options de chargement de la variante
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
        
        ' === EXPORT / SAUVEGARDE ===
        ' Positionne le curseur et ouvre le menu contextuel pour l'exportation
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").setCurrentCell 0, "EKKO-EBELN"
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectedRows = "0"
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem "&XXL"

        ' Configure les options d'exportation (format, toujours afficher)
        .findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
        .findById("wnd[1]/usr/chkCB_ALWAYS").Selected = True
        .findById("wnd[1]/tbar[0]/btn[0]").press
        
        ' Sauvegarde le fichier exporté
        modUtils.SaveFile GetSetting("PATH_EXPORT"), "PO.xlsx"
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_PO_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_BonCommandes_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Bons de Commandes.
'               (Note: Cette procédure n'est pas directement appelée dans le module).
' DÉPENDANCES : g_Session, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_BonCommandes_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Sélectionne l'option d'exportation vers un tableur
        .findById("wnd[0]/mbar/menu[5]/menu[8]/menu[0]").Select
        ' Ouvre le menu contextuel du grid pour l'exportation
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_BonCommandes_Report", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Configure_SES_Report
' DESCRIPTION : Configure l'écran de sélection pour le rapport des Feuilles de Saisie des Services (SES).
' DÉPENDANCES : g_Session, GetSetting, g_DebutAnnee, g_FinAnnee, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub Configure_SES_Report()
    On Error GoTo SapErrorHandler
    With g_Session
        ' Saisit le code de transaction ML84
        .findById("wnd[0]/tbar[0]/okcd").text = "/nml84"
        .findById("wnd[0]").sendVKey 0
        ' Définit l'usine
        .findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
        ' Ouvre la fenêtre de sélection multiple pour les groupes de marchandises
        .findById("wnd[0]/usr/btn%_S_MATKL_%_APP_%-VALU_PUSH").press
        ' Renseigne les groupes de marchandises spécifiques
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "03*"
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "04*"
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "06*"
        ' Valide la sélection
        .findById("wnd[1]/tbar[0]/btn[8]").press
        ' Sélectionne les options pour les feuilles de saisie bloquées et acceptées
        .findById("wnd[0]/usr/radP_LOCK_A").Select
        .findById("wnd[0]/usr/radP_KZAB_A").Select
        ' Définit la période de création
        .findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = g_DebutAnnee
        .findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = g_FinAnnee
        ' Exécute le rapport
        .findById("wnd[0]/tbar[1]/btn[8]").press
        ' Appuie sur un bouton pour l'exportation (souvent "Exporter" ou "Tableur")
        .findById("wnd[0]/tbar[1]/btn[33]").press
        ' Configure les options d'exportation
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    End With
    Exit Sub
SapErrorHandler:
    DisplayAndLogError "Configure_SES_Report", Err
End Sub

'====================================================================================
' SECTION : PROCÉDURES UTILITAIRES
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExecuteAndExportReport
' DESCRIPTION : Procédure générique pour exécuter un rapport SAP et l'exporter.
'               Elle gère la connexion, la configuration, l'exportation, la déconnexion
'               et la journalisation pour les rapports standardisés.
' PARAMÈTRES  :
'   - reportType (String)          : Type de rapport (pour les messages de log).
'   - sapTransactionWrapper (String): Nom de la procédure VBA qui lance la transaction SAP.
'                                     Peut être vide si la transaction est gérée par configActionName.
'   - configActionName (String)    : Nom de la procédure VBA qui configure l'écran de sélection SAP.
'   - outputFileName (String)      : Nom du fichier Excel de sortie.
'   - useSaveFile (Boolean, Optional): Indique si la sauvegarde est gérée par la procédure
'                                     de configuration (actuellement inutilisé, la sauvegarde
'                                     est centralisée via modUtils.ExportToExcel).
' DÉPENDANCES : DateSemain, onSAP, modUtils.ExportToExcel, offSAP,
'               modUtils.FermerFichierExcel, modUtils.LogMessage, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Private Sub ExecuteAndExportReport(ByVal reportType As String, ByVal sapTransactionWrapper As String, ByVal configActionName As String, ByVal outputFileName As String, Optional ByVal useSaveFile As Boolean = False)
    On Error GoTo SapErrorHandler

    ' 1. Initialisation
    DateSemain
    ' Établit la connexion SAP. La procédure 'onSAP' gère la vérification.
    Run "onSAP"
    
    ' 2. Lancement du wrapper de transaction SAP (si fourni)
    If sapTransactionWrapper <> "" Then
        Run sapTransactionWrapper
    End If

    ' 3. Execution de l'action de configuration specifique
    Application.Run configActionName

    ' 4. Exportation du fichier
    If useSaveFile Then
        ' Ce chemin n'est plus utilisé. La sauvegarde est gérée par modUtils.ExportToExcel.
        ' Le paramètre 'useSaveFile' est maintenu pour compatibilité mais n'a pas d'effet.
    Else
        modUtils.ExportToExcel outputFileName
    End If
    
    ' 5. Nettoyage
    Run "offSAP"
    modUtils.FermerFichierExcel outputFileName
    modUtils.LogMessage "? Z_Load_" & reportType & " terminé"
    Exit Sub

SapErrorHandler:
    'DisplayAndLogError "ExecuteAndExportReport (" & reportType & ")", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExportToExcel
' DESCRIPTION : Wrapper pour la fonction d'exportation vers Excel centralisée.
' PARAMÈTRES  :
'   - fileName (String) : Nom du fichier Excel à créer.
' DÉPENDANCES : modUtils.ExportToExcel.
'------------------------------------------------------------------------------------
Private Sub ExportToExcel(ByVal fileName As String)
    ' Cette fonction est maintenant centralisée dans modUtils.bas
    modUtils.ExportToExcel fileName
End Sub

