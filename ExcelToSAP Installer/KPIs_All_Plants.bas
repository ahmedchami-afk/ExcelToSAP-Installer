Attribute VB_Name = "KPIs_All_Plants"
'====================================================================================
' MODULE      : KPIs_All_Plants (Indicateurs de Performance - Toutes les usines)
' AUTEUR      : Votre Nom / Gemini Code Assist
' DATE        : 03/11/2025 (Dernière modification : 02/12/2025)
'====================================================================================
' DESCRIPTION :
'   Ce module regroupe les procédures d'extraction de données KPI (Indicateurs de Performance)
'   destinées à alimenter le tableau de bord principal ("Dashboard").
'   Contrairement au module `KPI_MP`, ce module gère les extractions pour toutes les usines
'   pertinentes et génère des fichiers Excel distincts, optimisés pour être utilisés
'   par des outils de Business Intelligence (BI) tels que Tableau.
'
'   Les procédures de ce module sont nommées avec le suffixe "_Dash" afin de les différencier
'   des procédures similaires destinées à d'autres usages.
'====================================================================================
Option Explicit

' --- Constantes pour les chemins et noms de fichiers ---
Private Const DASHBOARD_PATH_KEY As String = "DASHBOARD_PATH" ' Clé pour le dictionnaire de configuration

' Fichiers principaux
Private Const FILE_SCHEDULED_OP_KPI As String = "01_Schudled_Op_KPI.XLSX"
Private Const FILE_SCHEDULED_CONFIRMATION As String = "02_Schudled_Confirmation.XLSX"
Private Const FILE_CONFIRMED_CONFIRMATION As String = "03_Confirmed_Confirmation.XLSX"
Private Const FILE_SCHEDULED_OPERATIONS As String = "04_Schudled_Operations.XLSX"
Private Const FILE_SCHEDULED_NO_CONFIRMED As String = "05_Schudled_no_Confirmed.XLSX"
Private Const FILE_UNPLANNED_CONFIRMATION As String = "06_Unplanned_Confirmation.XLSX"
Private Const FILE_UNPLANNED As String = "07_Unplanned.XLSX"
Private Const FILE_OVERDUE As String = "08_Overdue.XLSX"
Private Const FILE_PMR As String = "09_PMR.XLSX"
Private Const FILE_PMR_NOT_PERFORMED As String = "10_PMR_not_Performed.XLSX"
Private Const FILE_PMR_MANUAL_CALL As String = "11_PMR_Manual_Call.XLSX"
Private Const FILE_NOTIFICATION_CREATED As String = "12_Notification_Created.XLSX"
Private Const FILE_PM01_WO_MR As String = "13_PM01_Wo_MR.XLSX"
Private Const FILE_AGING_WO As String = "14_Aging_Wo.XLSX"
Private Const FILE_AGING_MR As String = "15_Aging_MR.XLSX"
Private Const FILE_NOTIFICATION_CREATED_CPM As String = "16_Notification_Created_CPM.XLSX"
Private Const FILE_SCHED_RATIO As String = "17_SchRatio.XLSX"
Private Const FILE_OPEN_PM_ORDER As String = "Open_PMOrder.XLSX"
Private Const FILE_OPEN_NOTIFICATIONS As String = "Open_Notifications.XLSX"
Private Const FILE_CNF_ALL_PLANT As String = "CNF_Extraction_ALL_Plant.XLSX"
Private sheetSource As String

'-------------------------------------------------------------------------------
' SUB : Load_data_Dashboard (Chargement complet des données du tableau de bord)
' DESCRIPTION : Proc?dure ma?tresse qui orchestre l'appel de toutes les autres
'               proc?dures d'extraction de ce module pour un rechargement complet
'               des donn?es du tableau de bord.
'-------------------------------------------------------------------------------
Sub Load_data_Dashboard()
    On Error GoTo ErrorHandler

    ' Optimisation des performances Excel pour toute la dur?e du chargement
    DateSemain
    LoadConfiguration
    OptimizeExcel
    
    ' Vider les fichiers existants avant de les recharger
    Clear_Dashboard_Files
    
    Z_KPIs_List_Dash
    
    Z_Schudled_Confirmation_File_Dash
    Z_Confirmed_Confirmation_File_Dash
    Z_unplanned_Confirmation_File_Dash
    
    Z_KPI_Schudled_Op_Dash
    Z_KPI_Schudled_non_Confirmed_Op_Dash
    Z_KPIs_Unplanned_Op_Dash
    Z_KPI_Accuracy_Dash
    
    Z_KPI_SchRatio_Dash
    
    Z_KPI_Overdue_Dash
    Z_KPIs_PMR_Dash
    Z_KPIs_PMR_not_Performed_Dash
    Z_KPIs_PMR_ManualCall_Dash
    Z_KPI_MR_Created_Dash
    Z_KPI_Wo_MR_Dash
    Z_KPI_AgingWo_Dash
    Z_KPI_AgingMR_Dash
    Z_KPI_MR_Created_CPM_Dash
    
    ' Vérifier et fermer les fichiers du tableau de bord s'ils sont restés ouverts
    Close_Dashboard_Files_If_Open
    
    ' Rafraîchir le fichier de statistiques externe
    RefreshKPIStatistics
    
    ' S'assurer que Tableau Public est en cours d'exécution
    EnsureTableauPublicIsRunning
    
    ' R?tablissement des param?tres Excel ? la fin du chargement complet
    RestoreExcel
    Exit Sub

ErrorHandler:
    RestoreExcel
    MsgBox "Une erreur critique est survenue dans Load_data_Dashboard : " & vbCrLf & Err.Description, vbCritical, "Erreur Dashboard"
End Sub

'===============================================================================
'== SECTION : EXTRACTIONS POUR FICHIERS UTILIS?S PAR TABLEAU / DASHBOARD
'===============================================================================

'-------------------------------------------------------------------------------
' SUB : Z_KPIs_List_Dash (Extraction de la liste principale des KPIs)
' DESCRIPTION : Extrait la liste principale des KPIs (similaire ? Z_KPIs_List)
'               mais sauvegarde le r?sultat dans "01_Schudled_Op_KPI.XLSX" pour
'               le tableau de bord.
'-------------------------------------------------------------------------------
Sub Z_KPIs_List_Dash()
    Dim chemin As String
    Dim nomFichier As String
    
    On Error GoTo ErrorHandler
    
'liste des operations dans le fichier KPIs
    DateSemain
    sheetSource = "KPI_Dashboard"
    
    ' Verification si la semaine est deja chargee
    'Dim rngWeek As Range
    'Set rngWeek = ThisWorkbook.Sheets(sheetSource).Rows(1).Find(What:="Week", LookIn:=xlValues, LookAt:=xlPart)
    'If Not rngWeek Is Nothing Then
    '    If CStr(ThisWorkbook.Sheets(sheetSource).Cells(2, rngWeek.Column).Value) = SemaineKPIs Then
    '        MsgBox "schudled work order loaded"
    '        RestoreExcel
    '        Run "offSAP"
    '        Exit Sub
    '    End If
    'End If
    
    With ThisWorkbook.Sheets(sheetSource)
        If .UsedRange.Rows.count > 1 Then
            .Rows("2:" & .Rows.count).Delete
        End If
    End With
    
    '--- ?tape 1 : Initialisation et param?trage SAP ---
    StartSAPTransaction "Z_KPIP_SM"

    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-LOW").text = SemaineKPIs '_Avant
    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-HIGH").text = SemaineKPIs

    'version 3 et 4 des KPIs region MEA
    g_Session.findById("wnd[0]/usr/ctxtR_VERSI-LOW").text = "3"
   
    ' Ex?cution et navigation
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    
    ' S?lection du layout
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton "&MB_VARIANT"
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem "&LOAD"
    g_Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
      
    '--- ÉTAPE 2 : Filtrage sur les usines (Plants) ---
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell -1, "WERKSH"
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").selectColumn "WERKSH"
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").PressToolbarButton "&MB_FILTER"
    g_Session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    
    g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    
    g_Session.findById("wnd[2]/tbar[0]/btn[8]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' S?lection de toutes les donn?es
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell -1, ""
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").SelectAll
    
    '--- ?tape 3 : Exportation et traitement du fichier ---
    nomFichier = FILE_SCHEDULED_OP_KPI
    ExportGridToExcel "wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell", nomFichier

    chemin = GetSetting(DASHBOARD_PATH_KEY)
    RenommerColonnes_KPIs chemin, nomFichier

    ' Charge les donn?es dans la feuille correspondante du classeur principal
    sheetSource = "KPI_Dashboard"
    ChargerDonnees chemin, nomFichier, sheetSource

CleanExit:
    '--- ÉTAPE 4 : Nettoyage ---
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub
    
ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie
    'DisplayAndLogError "Z_KPIs_List_Dash", Err
    ' Le nettoyage est géré par DisplayAndLogError, on peut sortir.
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

'***************************************************************
'*********************                      ********************
'*********************     Confirmations    ********************
'*********************        Files         ********************
'***************************************************************
Sub Z_Schudled_Confirmation_File_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String
    Dim nomFichier As String
    
    sheetSource = "CNF_SCH_Dashboard"
    With ThisWorkbook.Sheets(sheetSource)
        If .UsedRange.Rows.count > 1 Then
            .Rows("2:" & .Rows.count).Delete
        End If
    End With
    
    '*********************************************************
    CopierChampConfirmationSCH_Dash
    StartSAPTransaction "Z_IW47"
    '---------------------------------
    'case a cocher
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_ABG").Selected = True
    g_Session.findById("wnd[0]/usr/chkNO_CANC").Selected = True
    

    'plants
    g_Session.findById("wnd[0]/usr/btn%_WERKS_O_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    

    'Effacer Planing Plant
    g_Session.findById("wnd[0]/usr/btn%_WERKS_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    'Date
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-HIGH").text = FinSemaine
    
    'Paste schdlued Confirmation liste
    g_Session.findById("wnd[0]/usr/btn%_RUECK_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'layer
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CNF_ALL"
    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_SCHEDULED_CONFIRMATION
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

    ' Charger les donn?es et actualiser le TCD
    chemin = GetSetting(DASHBOARD_PATH_KEY)
    sheetSource = "CNF_SCH_Dashboard"
    ChargerDonnees chemin, nomFichier, sheetSource

CleanExit:
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_Schudled_Confirmation_File_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_Confirmed_Confirmation_File_Dash()

    Dim chemin As String
    Dim nomFichier As String
    
    On Error GoTo ErrorHandler
    
    sheetSource = "CNF_Dashboard"
    With ThisWorkbook.Sheets(sheetSource)
        If .UsedRange.Rows.count > 1 Then
            .Rows("2:" & .Rows.count).Delete
        End If
    End With

    StartSAPTransaction "Z_IW47"

    'case a cocher (cocher les cases)
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_ABG").Selected = True
    g_Session.findById("wnd[0]/usr/chkNO_CANC").Selected = True
    
    'plants
    g_Session.findById("wnd[0]/usr/btn%_WERKS_O_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer Planing Plant
    g_Session.findById("wnd[0]/usr/btn%_WERKS_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    'date
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-HIGH").text = FinSemaine
    
   
    'layer
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CNF_ALL"
    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_CONFIRMED_CONFIRMATION
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

    ' Charger les donn?es et actualiser le TCD
    chemin = GetSetting(DASHBOARD_PATH_KEY)
    sheetSource = "CNF_Dashboard"
    ChargerDonnees chemin, nomFichier, sheetSource

CleanExit:
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_Confirmed_Confirmation_File_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

Sub Z_unplanned_Confirmation_File_Dash()

    Dim chemin As String
    Dim nomFichier As String
    
    On Error GoTo ErrorHandler

    StartSAPTransaction "Z_IW47"

    '---------------------------------
    'case a cocher
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_ABG").Selected = True
    g_Session.findById("wnd[0]/usr/chkNO_CANC").Selected = True
    
    'plants
    g_Session.findById("wnd[0]/usr/btn%_WERKS_O_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer Planing Plant
    g_Session.findById("wnd[0]/usr/btn%_WERKS_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    'date
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-HIGH").text = FinSemaine
    
    'paste Include confirmation
    CopierChampConfirmationCNF_Dash
    g_Session.findById("wnd[0]/usr/btn%_RUECK_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
   'paste exclude confirmation
    CopierChampConfirmationSCH_Dash
    g_Session.findById("wnd[0]/usr/btn%_RUECK_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
    'ok
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'layer
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CNF_ALL"
    
    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_UNPLANNED_CONFIRMATION
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_unplanned_Confirmation_File_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier

End Sub
'***************************************************************
'*********************                      ********************
'*********************      Operations      ********************
'*********************                      ********************
'***************************************************************
Public Sub Z_KPI_Schudled_Op_Dash()
    On Error GoTo ErrorHandler
    

    Dim chemin As String, nomFichier As String
    
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select ' Onglet usine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'paste include confirmation
    CopierChampConfirmationSCH_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "*CNF"
    '----------- week date du lundi au samedi ------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_SCHEDULED_OPERATIONS
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_Schudled_Op_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

Public Sub Z_KPI_Schudled_non_Confirmed_Op_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'paste exclude confirmation
    CopierChampConfirmationSCH_CNF_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
    'paste include confirmation
    CopierChampConfirmationSCH_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "*CNF"
    '----------- week date du lundi au samedi ------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_SCHEDULED_NO_CONFIRMED
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_Schudled_non_Confirmed_Op_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

Sub Z_KPI_Accuracy_Dash()

End Sub

Sub Z_KPI_SchRatio_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    'If g_DoNotRun = TRUE Then Exit Sub
    
    StartSAPTransaction "Z_SHIFT"
    
    '-----------------------------------------------------
    'layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHM_CAP_DASH"
    
    'plant
    g_Session.findById("wnd[0]/usr/ctxtS_PLANT-LOW").text = "0P1D"
    
    'workcenter
    g_Session.findById("wnd[0]/usr/ctxtS_WCTR-LOW").text = "*"
    
    'plants
    g_Session.findById("wnd[0]/usr/btn%_S_PLANT_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'date
    g_Session.findById("wnd[0]/usr/ctxtP_LOWDT").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtP_HIGHDT").text = FinSemaine
    
    
    'execute
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    '-----------------------------------------------------

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell 0, "WERKS"
    
    nomFichier = FILE_SCHED_RATIO
    ExportGridToExcel "wnd[0]/usr/cntlCONTAINER/shellcont/shell", nomFichier
    
    chemin = GetSetting(DASHBOARD_PATH_KEY)
    RenommerColonnes chemin, nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_SchRatio_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPIs_Unplanned_Op_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String

    'If g_DoNotRun = TRUE Then Exit Sub
    
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'paste exclude confirmation
    CopierChampConfirmationSCH_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    
    'paste Include confirmation
    CopierChampConfirmationCNF_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for Unplanned ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "*CNF"

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_UNPLANNED
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPIs_Unplanned_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_Overdue_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW49N"
    
    'not TECO, CLSD
    g_Session.findById("wnd[0]/usr/chkSP_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = False
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = False
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM01"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '----------- condition for aging wo ------------
    'aging basic start date
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GSTRP-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRS-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRS-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "2"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRS-LOW").text = AgingDate
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/ctxtS_SWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "REL*"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_VSTAEX_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CNF"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "DLFL"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "DLT"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '---------------------- User Status 4sch -------------------------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").text = "4sch"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").caretPosition = 4
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_OVERDUE
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_Overdue_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPIs_PMR_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    
    StartSAPTransaction "Z_IW49N"
    
    'PM02
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").caretPosition = 4
    g_Session.findById("wnd[0]").sendVKey 0
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for Unplanned ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "*CNF"
    
    'paste Include confirmation
    CopierChampConfirmationCNF_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    'date
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-HIGH").text = FinSemaine
    
    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_PMR
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPIs_PMR_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPIs_PMR_not_Performed_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String

    CopierChampConfirmationSCH_Dash
    'If g_DoNotRun = TRUE Then Exit Sub
    
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'paste include confirmation
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "*CNF"
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_PMR_NOT_PERFORMED
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPIs_PMR_not_Performed_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPIs_PMR_ManualCall_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String

    CopierChampConfirmationSCH_Dash
    'If g_DoNotRun = TRUE Then Exit Sub
    
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'paste include confirmation
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").SetFocus
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").text = "*IP*"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").caretPosition = 4
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "IWERK"
    
    nomFichier = FILE_PMR_MANUAL_CALL
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPIs_PMR_ManualCall_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_MR_Created_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW29"
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '----------- condition for aging wo ------------
    'select status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_RST").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    '************ created in this week **********************
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = FinSemaine
    'Layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "BIS_NTF_ALL_"
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "IWERK"
    
    nomFichier = FILE_NOTIFICATION_CREATED
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_MR_Created_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_MR_Created_CPM_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW29"
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    
    '----------- condition for CPM MR ------------
    'select status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_RST").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/ctxtSTAI1-LOW").text = "*CPM"
    '************ created in this week **********************
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = FinSemaine
    'Layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "BIS_NTF_ALL_"
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "IWERK"
    
    nomFichier = FILE_NOTIFICATION_CREATED_CPM
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_MR_Created_CPM_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_Wo_MR_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW49N"
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'layer
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "TECO"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").caretPosition = 4
    'PM01
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM01"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'no notifications
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_QMNUM-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_QMNUM-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_QMNUM-LOW").text = "0"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_QMNUM-HIGH").text = "999999999"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").caretPosition = 0
    
    'paste include confirmation
    CopierChampConfirmationSCH_CNF_Dash
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "IWERK"
    
    nomFichier = FILE_PM01_WO_MR
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_Wo_MR_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_AgingWo_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW49N"
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '----------- condition for aging wo ------------
    'aging basuc funishe date
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "2"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").text = AgingDate
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").caretPosition = 10
    
    '--- status relased & confirmed ---------------
    g_Session.findById("wnd[0]/usr/chkSP_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = False
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = False
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    
    'layer
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = "CHM_WO_DASH"
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "IWERK"
    
    nomFichier = FILE_AGING_WO
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_AgingWo_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub
Sub Z_KPI_AgingMR_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW29"
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'Layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "BIS_NTF_ALL_"
    
    '----------- condition for aging wo ------------
    'select status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_RST").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = False
    'aging date
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "2"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = AgingDate
    
    '************ exclude Work Order **********************
    g_Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = "0"
    g_Session.findById("wnd[0]/usr/ctxtAUFNR-HIGH").text = "999999999999"
    '************ exclude revision **********************
    g_Session.findById("wnd[0]/usr/ctxtREVNR-LOW").text = "2*"
    g_Session.findById("wnd[0]/usr/ctxtREVNR-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/ctxtREVNR-LOW").caretPosition = 2
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "IWERK"
    
    nomFichier = FILE_AGING_MR
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Le reste de votre code apr?s SaveFile
    ' ...
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_AgingMR_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

'**********************************************
'********** fin procedure pour tableau ********
'**********************************************

Sub Z_Load_Open_PMwo_Dash()
    Dim chemin As String
    Dim nomFichier As String
    
    On Error GoTo ErrorHandler
    StartSAPTransaction "Z_IW39"
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/btn%_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    
    'no date
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    'select layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CHM_WO_RPRT"

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extraction vers Excel
    nomFichier = FILE_OPEN_PM_ORDER
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Fermer SAP
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_Load_Open_PMwo_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

Sub Z_KPI_Open_MR_Dash()
    On Error GoTo ErrorHandler
    
    Dim chemin As String, nomFichier As String
    StartSAPTransaction "Z_IW29"
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0P1D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "011D"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0D1D"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'layout
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = "BIS_NTF_RPR"
    'select status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_RST").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = False
    '-----------------------------------------------------

    'lanche
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' setCurrentCell et Extraction
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 11, "STTXT"
    nomFichier = FILE_OPEN_NOTIFICATIONS
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

CleanExit:
    ' Fermer SAP
    Run ("offSAP")
    FermerFichierExcel nomFichier
    Exit Sub

ErrorHandler:
    'DisplayAndLogError "Z_KPI_Open_MR_Dash", Err
    Run ("offSAP")
    FermerFichierExcel nomFichier
End Sub

'-------------------------------------------------------------------------------
' SUB : Clear_Dashboard_Files (Effacer le contenu des fichiers du tableau de bord)
' DESCRIPTION : Vide le contenu de tous les fichiers Excel générés pour le
'               tableau de bord.
'-------------------------------------------------------------------------------
Sub Clear_Dashboard_Files()
    Dim chemin As String
    Dim fileList As Variant
    Dim vFile As Variant
    Dim fullPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCell As Range
    
    On Error GoTo ErrorHandler
    chemin = GetSetting(DASHBOARD_PATH_KEY)
    
    fileList = Array( _
        FILE_SCHEDULED_CONFIRMATION, _
        FILE_CONFIRMED_CONFIRMATION, _
        FILE_SCHEDULED_OPERATIONS, _
        FILE_SCHEDULED_NO_CONFIRMED, _
        FILE_UNPLANNED_CONFIRMATION, _
        FILE_UNPLANNED, _
        FILE_OVERDUE, _
        FILE_PMR, _
        FILE_PMR_NOT_PERFORMED, _
        FILE_PMR_MANUAL_CALL, _
        FILE_NOTIFICATION_CREATED, _
        FILE_PM01_WO_MR, _
        FILE_AGING_WO, _
        FILE_AGING_MR, _
        FILE_NOTIFICATION_CREATED_CPM, _
        FILE_SCHED_RATIO, _
        FILE_OPEN_PM_ORDER, _
        FILE_OPEN_NOTIFICATIONS, _
        FILE_CNF_ALL_PLANT _
    )
    
    For Each vFile In fileList
        fullPath = chemin & "\" & vFile
        If Dir(fullPath) <> "" Then
            Set wb = Workbooks.Open(fullPath)
            Set ws = wb.Sheets(1)
            
            ' Find the last cell containing data
            Set lastCell = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            
            ' If data exists below the header row, delete those rows
            If Not lastCell Is Nothing Then
                If lastCell.row > 1 Then
                    ws.Rows("2:" & lastCell.row).Delete
                End If
            End If
            
            wb.Close SaveChanges:=True
        End If
    Next vFile
    
CleanExit:
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Erreur dans Clear_Dashboard_Files: " & Err.Description
End Sub

Sub RefreshKPIStatistics()
    Dim wbPath As String
    Dim wbName As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    On Error GoTo ErrorHandler
    
    ' --- Définir les chemins et noms de fichiers ---
    wbPath = "C:\Users\ahmchami\Documents\ExcelToSAP"
    wbName = "Statistique KPI.xlsx"
    fullPath = wbPath & "\" & wbName
    
    ' --- Vérifier si le fichier existe ---
    If Dir(fullPath) = "" Then
        MsgBox "Le fichier '" & wbName & "' n'a pas été trouvé à l'emplacement : " & vbCrLf & wbPath, vbExclamation, "Fichier introuvable"
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(fullPath)
    Set ws = wb.Sheets(1) ' Cible la première feuille
    
    ' Rafraîchissement total du classeur
    wb.RefreshAll
    
    ' Attente de 10 secondes pour laisser le temps aux données de se charger
    Application.Wait (Now + TimeValue("0:00:10"))
    
    ' Second rafraîchissement pour s'assurer que tout est à jour
    wb.RefreshAll
    
    For Each pt In ws.PivotTables
        pt.RefreshTable
    Next pt
    
    'wb.Close SaveChanges:=True
    Exit Sub ' Sortie normale
ErrorHandler:
    MsgBox "Une erreur est survenue lors du rafraîchissement du fichier 'Statistique KPI.xlsx'." & vbCrLf & "Erreur: " & Err.Description, vbCritical, "Erreur de rafraîchissement"
    If Not wb Is Nothing Then wb.Close SaveChanges:=False ' Fermer sans sauvegarder en cas d'erreur
End Sub

Sub RenommerColonnes_KPIs(chemin As String, nomFichier As String)
    Dim classeur As Workbook
    Dim feuille As Worksheet

    ' Sp?cifier le chemin complet du fichier Excel
    Dim cheminFichier As String
    cheminFichier = chemin & "\" & nomFichier

    ' Ouvrir le fichier Excel
    Set classeur = Workbooks.Open(cheminFichier)

    ' R?f?rencer la premi?re feuille du classeur
    Set feuille = classeur.Sheets(1)

    ' Renommer les colonnes
    feuille.Cells(1, 13).value = "Operation WorkCenter"


    ' Enregistrer et fermer le fichier Excel
    classeur.Close SaveChanges:=True
End Sub

Sub Close_Dashboard_Files_If_Open()
    Dim fileList As Variant
    Dim vFile As Variant
    Dim wb As Workbook
    
    On Error Resume Next
    
    ' Liste des fichiers à vérifier
    fileList = Array( _
        FILE_SCHEDULED_OP_KPI, _
        FILE_SCHEDULED_CONFIRMATION, _
        FILE_CONFIRMED_CONFIRMATION, _
        FILE_SCHEDULED_OPERATIONS, _
        FILE_SCHEDULED_NO_CONFIRMED, _
        FILE_UNPLANNED_CONFIRMATION, _
        FILE_UNPLANNED, _
        FILE_OVERDUE, _
        FILE_PMR, _
        FILE_PMR_NOT_PERFORMED, _
        FILE_PMR_MANUAL_CALL, _
        FILE_NOTIFICATION_CREATED, _
        FILE_PM01_WO_MR, _
        FILE_AGING_WO, _
        FILE_AGING_MR, _
        FILE_NOTIFICATION_CREATED_CPM, _
        FILE_SCHED_RATIO, _
        FILE_OPEN_PM_ORDER, _
        FILE_OPEN_NOTIFICATIONS, _
        FILE_CNF_ALL_PLANT, _
        "Statistique KPI.xlsx" _
    )
    
    For Each vFile In fileList
        Set wb = Nothing
        Set wb = Workbooks(vFile)
        If Not wb Is Nothing Then
            wb.Close SaveChanges:=False
        End If
    Next vFile
    
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' SUB : EnsureTableauPublicIsRunning (Vérifier et lancer Tableau Public)
' DESCRIPTION : Vérifie si l'application Tableau Public est déjà en cours
'               d'exécution. Si c'est le cas, elle tente de l'amener au premier
'               plan. Sinon, elle lance l'application.
'-------------------------------------------------------------------------------
Sub EnsureTableauPublicIsRunning()
    Const APP_NAME As String = "tabpublic.exe"
    Const APP_PATH As String = "C:\Program Files\Tableau\Tableau Public 2025.2\bin\tabpublic.exe"
    Dim objWMIService As Object
    Dim colProcesses As Object
    Dim isRunning As Boolean

    On Error GoTo ErrorHandler
    
    isRunning = False
    
    ' Utiliser WMI pour vérifier si le processus est en cours d'exécution
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & APP_NAME & "'")
    
    If colProcesses.count > 0 Then isRunning = True

    If isRunning Then
        ' L'application est en cours d'exécution, on tente de l'activer.
        On Error Resume Next
        AppActivate "Tableau"
        On Error GoTo 0
    Else
        ' L'application n'est pas en cours d'exécution, on la lance.
        If Dir(APP_PATH) <> "" Then
            Shell """" & APP_PATH & """", vbNormalFocus
        Else
            MsgBox "L'exécutable de Tableau Public est introuvable au chemin spécifié : " & vbCrLf & APP_PATH, vbCritical, "Fichier introuvable"
        End If
    End If
Exit Sub
ErrorHandler:
    MsgBox "Une erreur est survenue lors de la vérification/lancement de Tableau Public : " & Err.Description, vbExclamation, "Erreur Processus"
End Sub
'----------------------------------------------------
'------------- fonction support ---------------------
'----------------------------------------------------

Private Sub StartSAPTransaction(ByVal transactionMacro As String)
'-------------------------------------------------------------------------------
' DESCRIPTION : Initialise la connexion SAP et lance la transaction sp?cifi?e.
'               Factorise les appels r?p?titifs ? "onSAP" et ? la macro de transaction.
'-------------------------------------------------------------------------------
    LogMessage "Lancement de la transaction SAP via macro : " & transactionMacro
    Run "onSAP"
    Run transactionMacro
End Sub

Private Sub ExportGridToExcel(ByVal gridID As String, ByVal fileName As String)
'-------------------------------------------------------------------------------
' DESCRIPTION : Gère l'exportation Excel depuis une grille SAP ALV et sauvegarde
'               le fichier via SaveFile.
'-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    ' Context Menu -> Export
    g_Session.findById(gridID).contextMenu
    g_Session.findById(gridID).selectContextMenuItem "&XXL"
    
    ' Vérification si la popup d'exportation s'est bien ouverte (wnd[1])
    ' Si SAP affiche un message (ex: "Pas de données") au lieu d'ouvrir la popup, on arrête.
    If g_Session.ActiveWindow.name <> "wnd[1]" Then
        LogMessage "Avertissement Export : La popup n'est pas apparue pour " & fileName & ". Message SAP : " & g_Session.findById("wnd[0]/sbar").text
        Exit Sub
    End If

    ' Popup Export
    g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
    g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31" ' Excel
    g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").SetFocus
    g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").Selected = True
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ' Sauvegarde
    Dim chemin As String
    chemin = GetSetting(DASHBOARD_PATH_KEY)
    SaveFile chemin, fileName
    Exit Sub

ErrorHandler:
    ' On relance l'erreur pour qu'elle soit gérée par la procédure appelante (qui gère le nettoyage)
    Err.Raise Err.Number, "ExportGridToExcel", "Erreur lors de l'export (" & fileName & ") : " & Err.Description
End Sub

Private Sub CopyConfirmationData(ByVal sourceSheetName As String, ByVal colName As String)
'-------------------------------------------------------------------------------
' DESCRIPTION : Procedure privee et generique pour copier les donnees du champ
'               specifie depuis la feuille.
'-------------------------------------------------------------------------------
    Dim fichierSource As Workbook
    Dim feuilleSource As Worksheet
    Dim plageSource As Range
    Dim rngHeader As Range
    Dim lastrow As Long

    On Error GoTo ErrorHandler
    g_DoNotRun = False

    Set fichierSource = ThisWorkbook
    Set feuilleSource = fichierSource.Sheets(sourceSheetName)
    
    Set rngHeader = feuilleSource.Cells.Find(What:=colName, LookIn:=xlValues, LookAt:=xlPart)
    If rngHeader Is Nothing Then Set rngHeader = feuilleSource.Range(colName)

    With rngHeader
        lastrow = feuilleSource.Cells(feuilleSource.Rows.count, .Column).End(xlUp).row
        If lastrow <= .row Then
            g_DoNotRun = True
            Exit Sub
        End If
        Set plageSource = feuilleSource.Range(.Offset(1, 0), feuilleSource.Cells(lastrow, .Column))
    End With

    plageSource.Copy
    Exit Sub

ErrorHandler:
    MsgBox "Impossible de copier les donnees de " & colName & "." & vbCrLf & _
           "Verifiez que la feuille '" & sourceSheetName & "' contient la colonne.", vbExclamation
    g_DoNotRun = True
End Sub

Private Sub CopierChampConfirmationSCH_Dash()
    CopyConfirmationData "KPI_Dashboard", "Confirmation"
End Sub
Private Sub CopierChampConfirmationSCH_CNF_Dash()
    CopyConfirmationData "CNF_SCH_Dashboard", "Confirmation"
End Sub
Private Sub CopierChampConfirmationCNF_Dash()
    CopyConfirmationData "CNF_Dashboard", "Confirmation"
End Sub



