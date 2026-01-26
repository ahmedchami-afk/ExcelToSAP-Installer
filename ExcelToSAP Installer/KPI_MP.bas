Attribute VB_Name = "KPI_MP"
'====================================================================================
' MODULE : KPI_MP
' AUTEUR : [Votre Nom]
' DATE   : 2025-11-03
'====================================================================================
' OBJECTIF :
'   Regroupe toutes les procedures d'extraction de donnees pour les indicateurs
'   de performance (KPI) sp?cifiques ? un "Maintenance Plant" (MP).
'
'   Chaque procedure automatise une transaction SAP pour extraire des donnees,
'   les sauvegarder dans un fichier Excel, puis les charger dans le classeur
'   principal pour alimenter des tableaux crois?s dynamiques.
'====================================================================================

Option Explicit

' --- Constantes pour les chemins et noms de fichiers ---
Private Const KPI_PATH As String = "C:\Dashboard-KPIs"
Private Const FILE_KPI_EXTRACTION As String = "KPI_Extraction.XLSX"
Private Const FILE_CNF_EXTRACTION As String = "CNF_Extraction.XLSX"

Private Sub ExportGridToExcel(ByVal gridID As String, ByVal fileName As String)
    On Error GoTo ErrorHandler

    ' Context Menu -> Export
    g_Session.findById(gridID).contextMenu
    g_Session.findById(gridID).selectContextMenuItem "&XXL"
    
    ' Popup Export
    g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
    g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31" ' Excel
    g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").SetFocus
    g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").Selected = True
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ' Sauvegarde
    Dim chemin As String
    chemin = GetSetting("KPI_PATH")
    SaveFile chemin, fileName
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "ExportGridToExcel", "Erreur lors de l'export (" & fileName & ") : " & Err.Description
End Sub

Private Sub ExecuteKPIListReport(ByVal week As String, ByVal sheetName As String)
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Run "onSAP"
    DateSemain
    
    Dim Path_File_KPIs As String: Path_File_KPIs = GetSetting("KPI_PATH")
    
    CheckWeekLoaded sheetName
    If g_DoNotRun Then Exit Sub
    
    ThisWorkbook.Sheets(sheetName).Rows("2:" & ThisWorkbook.Sheets(sheetName).Rows.count).ClearContents
    
    Run "Z_KPIP_SM"
    
    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-LOW").text = week
    g_Session.findById("wnd[0]/usr/ctxtR_WEEK-HIGH").text = week
    g_Session.findById("wnd[0]/usr/ctxtR_VERSI-LOW").text = "3"
    
    Z_F8
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    
    Dim gridID As String
    gridID = "wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell"
    
    g_Session.findById(gridID).setCurrentCell -1, "WERKSH"
    g_Session.findById(gridID).selectColumn "WERKSH"
    g_Session.findById(gridID).PressToolbarButton "&MB_FILTER"
    g_Session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[2]/tbar[0]/btn[8]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    g_Session.findById(gridID).setCurrentCell -1, ""
    g_Session.findById(gridID).SelectAll
    
    ExportGridToExcel gridID, FILE_KPI_EXTRACTION
    ChargerDonnees Path_File_KPIs, FILE_KPI_EXTRACTION, sheetName
    
    RestoreExcel
    On Error Resume Next
    g_Session.findById("wnd[0]/tbar[0]/btn[12]").press
    Run ("offSAP")
    FermerFichierExcel FILE_KPI_EXTRACTION
    Exit Sub
ErrorHandler:
    DisplayAndLogError "ExecuteKPIListReport", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPIs_List
' DESCRIPTION : Extrait la liste principale des KPIs depuis la transaction Z_KPIP_SM.
'               Les donnees sont filtrees par semaine et par usine, puis exportees
'               vers "KPI_Extraction.XLSX" et utilis?es pour actualiser un TCD.
'-------------------------------------------------------------------------------
Sub Z_KPIs_List()
    DateSemain
    ExecuteKPIListReport SemaineKPIs, "KPIs"
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPIs_List_LastWeek
' DESCRIPTION : Extrait la liste principale des KPIs de la semaine derniere.
'               Les donnees sont filtrees par semaine et par usine, puis exportees
'               vers "KPI_Extraction.XLSX" et chargees dans la feuille "KPIs (Last Week)".
'-------------------------------------------------------------------------------
Sub Z_KPIs_List_LastWeek()
    DateSemain
    ExecuteKPIListReport SemaineKPIs_Avant, "KPIs (Last Week)"
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_Confirmation_List
' DESCRIPTION : Extrait la liste des confirmations de la semaine (IW47).
'               Les donnees sont filtrees par statut, usine et date, puis exportees
'               vers "CNF_Extraction.XLSX" pour alimenter un TCD.
'-------------------------------------------------------------------------------
Sub Z_Confirmation_List()
    On Error GoTo ErrorHandler
    OptimizeExcel ' Optimisation des performances Excel

    Run "onSAP"
    DateSemain
    Run "Z_IW47"

    Dim chemin As String
    Dim nomFichier As String
    Dim NomTCD As String
    Dim sheetSource As String
    Dim sheetDestination As String
    
    ' --- Etape 1 : Initialisation et connexion ---
    Dim Path_File_KPIs As String: Path_File_KPIs = GetSetting("KPI_PATH")
    
    ' Efface tout le contenu de la feuille "CNF"
    sheetSource = "CNF"
    
    ThisWorkbook.Sheets(sheetSource).Rows("2:" & ThisWorkbook.Sheets(sheetSource).Rows.count).ClearContents
    
    '---------------------------------
    'case a cocher
    '--- Etape 1 : Parametrage de la transaction SAP (IW47) ---
    ' Filtre sur les statuts de confirmation (en cours, termin?, non annul?)
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_ABG").Selected = True
    g_Session.findById("wnd[0]/usr/chkNO_CANC").Selected = True
    
    ' Filtre sur les usines (Plants) lues depuis la feuille "Setup"
    g_Session.findById("wnd[0]/usr/btn%_WERKS_O_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' effacer planing plant
    g_Session.findById("wnd[0]/usr/btn%_WERKS_C_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Filtre sur la p?riode (semaine en cours)
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtERSDA_C-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtBUDAT_C-HIGH").text = FinSemaine
    
    ' Application du layout SAP pr?d?fini
    g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = ""
    
    ' Execution du rapport
    g_Session.findById("wnd[0]").sendVKey 8
    
    ' Check if data exists
    On Error Resume Next
    Dim lRowCount As Long
    lRowCount = g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
    If Err.Number <> 0 Or lRowCount = 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler

    ' Definition des chemins et noms pour le traitement
    nomFichier = FILE_CNF_EXTRACTION
    sheetSource = "CNF"

    ' Efface le contenu de la feuille "KPIs" sauf la première ligne (en-têtes)
    ThisWorkbook.Sheets(sheetSource).Rows("2:" & ThisWorkbook.Sheets(sheetSource).Rows.count).ClearContents
    
    ' Sauvegarde du fichier extrait
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    ExportGridToExcel "wnd[0]/usr/cntlGRID1/shellcont/shell", nomFichier

    '--- Etape 3 : Chargement et traitement des donnees dans Excel ---
    ' Charge les donnees dans la feuille "CNF"
    ChargerDonnees Path_File_KPIs, nomFichier, sheetSource

CleanExit:
    '--- ?tape 4 : Nettoyage et restauration ---
    RestoreExcel
    FermerFichierExcel nomFichier
    
    On Error Resume Next
    g_Session.findById("wnd[0]/tbar[0]/btn[12]").press
    On Error Resume Next
    g_Session.findById("wnd[0]/tbar[0]/btn[12]").press
    
    g_DoNotRun = False
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "No Confirmation List", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Copliance (Compliance)
' DESCRIPTION : Extrait les donn?es de conformit? depuis la transaction IW49N.
'               Utilise les confirmations planifiees (SCH) comme filtre et exclut
'               les op?rations d?j? confirm?es (CNF).
'-------------------------------------------------------------------------------
Sub Z_KPI_Schudled()
    On Error GoTo ErrorHandler
    Run "onSAP"
    DateSemain
    
    Dim confirmationsRange As Range
    Set confirmationsRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    
    Run "Z_IW49N"
    
    '--- Etape 1 : Parametrage de la transaction SAP (IW49N) ---
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    ' Filtre sur les usines (Plants) depuis la feuille "Setup".
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    ' Injecte la liste des confirmations planifiées (SCH)
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", confirmationsRange
    
    ' Filtre pour la non-conformite : exclut les operations deja confirmees (CNF)
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    'g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    'g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "*CNF"

    ' TAB9
    ' Application du layout et des filtres de statut pour la non-conformit?.
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If
    
    '----------- week date du lundi au samedi ------------

    ' Execution du rapport.
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_Copliance", Err
End Sub

Sub Z_KPI_No_Cmopliance()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPIs_PMR_not_Performed
    ' DESCRIPTION : Extrait les PMR non effectu?s. Filtre sur le type d'ordre PM02
    '               et exclut les op?rations d?j? confirm?es (CNF).
    '-------------------------------------------------------------------------------
    
    'Extracte Confirmation List
    If MsgBox("Voulez-vous lancer l'extraction des confirmations (Z_Confirmation_List) ?", vbYesNo + vbQuestion) = vbYes Then
        Z_Confirmation_List
    End If
    
    Dim schRange As Range, cnfRange As Range
    Set schRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    Set cnfRange = GetChampConfirmationCNF()
    'OptimizeExcel

    Run "onSAP"
    Run "Z_IW49N"
    'If g_DoNotRun = Then Exit Sub
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'Clear Period
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'PM02
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = ""

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    ' Filtre sur les usines.
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    Dim buttonId As String
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    buttonId = "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH"
    
    ' Étape 1: INCLURE les confirmations planifiées (SCH)
    ' Ouvre la popup, vide la liste, ajoute les données et laisse la popup ouverte.
    FillSAPSelectionList buttonId, schRange, False, True, True, False
    
    ' Étape 2: EXCLURE les confirmations réelles (CNF)
    ' N'ouvre pas la popup, ne vide pas la liste, ajoute les données en exclusion, et ferme la popup.
    FillSAPSelectionList buttonId, cnfRange, True, False, False, True
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "*CNF"
    '-----------------------------------------------------
    
    ' TAB9
    ' Application du layout
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8


CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPIs_PMR_not_Performed", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPIs_Unplanned
' DESCRIPTION : Extrait les donn?es pour les op?rations non planifi?es.
'               Utilise une logique d'inclusion/exclusion basee sur les listes
'               de confirmations planifiees (SCH) et confirmees (CNF).
'-------------------------------------------------------------------------------
Sub Z_KPIs_Unplanned()
    
    If MsgBox("Voulez-vous lancer l'extraction des confirmations (Z_Confirmation_List) ?", vbYesNo + vbQuestion) = vbYes Then
        Z_Confirmation_List
    End If
    
    Dim schRange As Range, cnfRange As Range
    Set schRange = GetChampConfirmationSCH()
    
    Set cnfRange = GetChampConfirmationCNF()
    'If g_DoNotRun = Then Exit Sub

    On Error GoTo ErrorHandler
    DateSemain
    OptimizeExcel

    Run "onSAP"
    Run "Z_IW49N"
    'If g_DoNotRun = Then Exit Sub
    
    '--- Etape 1 : Parametrage de la transaction SAP (IW49N) ---
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""

    ' TAB2
    'planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    ' Filtre sur les usines depuis la feuille "Setup".
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    Dim buttonId As String
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    buttonId = "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH"

    ' Étape 1: EXCLURE les confirmations planifiées (SCH)
    FillSAPSelectionList buttonId, schRange, True, True, True, False

    ' Étape 2: INCLURE les confirmations réelles (CNF)
    FillSAPSelectionList buttonId, cnfRange, False, False, False, True
    
    ' Filtre sur les statuts pour les operations non planifiees
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "*CNF"
    
    '---------- les date actual finish -----------------
    'DateSemainenCours
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-HIGH").text = FinSemaine
    '-----------------------------------------------------
    
    ' TAB9
    ' Application du layout
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If
    
    '--- Etape 2 : Execution du rapport ---
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPIs_Unplanned", Err
End Sub

Sub Z_KPI_Accuracy()
    On Error GoTo ErrorHandler
    OptimizeExcel

    Dim schRange As Range
    Set schRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    Run "onSAP"
    Run "Z_IW49N"
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    ' Injecte la liste des confirmations planifiées (SCH)
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", schRange
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    '-----------------------------------------------------
    
    ' TAB9
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Accuracy", Err
End Sub

Sub Z_KPI_SchRatio()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Run "onSAP"
    DateSemain
    
    Run "Z_SHIFT"
    
    ' Application du layout SAP.
    ' Appliquer le Layout si configuré dans "Setup"
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_Z_SHIFT")
    End If
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtS_PLANT-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/ctxtS_WCTR-LOW").text = "*"
    
    ' Filtre sur les usines.
    g_Session.findById("wnd[0]/usr/btn%_S_PLANT_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'date
    g_Session.findById("wnd[0]/usr/ctxtP_LOWDT").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtP_HIGHDT").text = FinSemaine
    
    'execute
    g_Session.findById("wnd[0]").sendVKey 8
    
    '-----------------------------------------------------
    
CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_SchRatio", Err
End Sub
Sub Z_KPIs_PMR_Confirmed()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPIs_PMR
    ' DESCRIPTION : Extrait les donn?es pour les ordres de maintenance pr?ventive (PMR).
    '               Filtre sur le type d'ordre PM02 et utilise la liste des confirmations.
    '-------------------------------------------------------------------------------
    
    'Extracte Confirmation List
    If MsgBox("Voulez-vous lancer l'extraction des confirmations (Z_Confirmation_List) ?", vbYesNo + vbQuestion) = vbYes Then
        Z_Confirmation_List
    End If
    
    
    Dim schRange As Range
    Set schRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    
    On Error GoTo ErrorHandler
    DateSemain
    OptimizeExcel
    
    Run "onSAP"
    Run "Z_IW49N"
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'Clear Period
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'PM02
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    ' Filtre sur les usines.
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]").sendVKey 8
    
    ' TAB4
    ' Injecte la liste des confirmations réelles (CNF)
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    'FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", cnfRange
    ' Injecte la liste des confirmations planifiées (SCH)
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", schRange

    ' TAB5
    '---------- les date actual finish -----------------
    'DateSemain
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-HIGH").text = FinSemaine
    '-----------------------------------------------------
    
    ' TAB9
    ' Application du layout
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8
    
CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPIs_PMR", Err
End Sub

Sub Z_KPIs_PMR_not_Performed()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPIs_PMR_not_Performed
    ' DESCRIPTION : Extrait les PMR non effectu?s. Filtre sur le type d'ordre PM02
    '               et exclut les op?rations d?j? confirm?es (CNF).
    '-------------------------------------------------------------------------------
    
    'Extracte Confirmation List
    If MsgBox("Voulez-vous lancer l'extraction des confirmations (Z_Confirmation_List) ?", vbYesNo + vbQuestion) = vbYes Then
        Z_Confirmation_List
    End If
    
    Dim schRange As Range, cnfRange As Range
    Set schRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    Set cnfRange = GetChampConfirmationCNF()
    'If g_DoNotRun = Then Exit Sub
    
    On Error GoTo ErrorHandler
    OptimizeExcel

    Run "onSAP"
    Run "Z_IW49N"
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'Clear Period
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    'PM02
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    ' Filtre sur les usines.
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    '---------- les date actual finish -----------------
    'DateSemain
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5").Select
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-LOW").Text = DebutSemaine
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB5/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1500/ctxtS_IEDD-HIGH").Text = FinSemaine
    '-----------------------------------------------------
    
    Dim buttonId As String
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    buttonId = "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH"

    ' Étape 1: INCLURE les confirmations planifiées (SCH)
    FillSAPSelectionList buttonId, schRange, False, True, True, False

    ' Étape 2: EXCLURE les confirmations réelles (CNF)
    FillSAPSelectionList buttonId, cnfRange, True, False, False, True
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "*CNF"
    '-----------------------------------------------------
    
    ' TAB9
    ' Application du layout
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPIs_PMR_not_Performed", Err
End Sub

Sub Z_KPIs_PMR_ManualCall()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim schRange As Range
    Set schRange = GetChampConfirmationSCH()
    'If g_DoNotRun = Then Exit Sub
    
    Run "onSAP"
    Run "Z_IW49N"
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM02"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").SetFocus
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").text = "*IP*"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERNAM-LOW").caretPosition = 4

    ' TAB3
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    ' Injecte la liste des confirmations planifiées (SCH)
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", schRange
    
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    
    ' TAB9
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPIs_PMR_ManualCall", Err
End Sub


Sub Z_KPI_Wo_MR()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_Wo_MR
    ' DESCRIPTION : Extrait les ordres de travail (Work Orders) li?s aux demandes
    '               de maintenance (Maintenance Requests). Filtre sur PM01 et exclut TECO.
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW49N"
    
    ' TAB1
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

    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB3
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    '----------- condition for non compliance ------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = True
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").text = "TECO"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAEX-LOW").caretPosition = 4
    
    ' Injecte la liste des confirmations réelles (CNF)
    Dim cnfRange As Range
    Set cnfRange = GetChampConfirmationCNF()
    'If g_DoNotRun = Then Exit Sub
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_RUECK_%_APP_%-VALU_PUSH", cnfRange
    '-----------------------------------------------------
    
    ' TAB9
    'layer
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_Wo_MR", Err
End Sub
Sub Z_KPI_AgingWo()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_AgingWo
    ' DESCRIPTION : Extrait les donn?es pour le KPI d'anciennet? des ordres de travail (Aging WO).
    '               Filtre sur la date de fin de base et les statuts "ouvert" ou "en cours".
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW49N"
    DateSemain
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    
    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    'aging basic funishe date
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 2
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "2"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").text = AgingDate
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_GLTRP-LOW").caretPosition = 10

    ' TAB3
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' TAB4
    '----------- condition for aging wo ------------
    '--- status relased & confirmed ---------------
    g_Session.findById("wnd[0]/usr/chkSP_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkSP_MAB").Selected = False
    g_Session.findById("wnd[0]/usr/chkSP_HIS").Selected = False
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    'g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").Text = "CNF"
    
    ' TAB9
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If
    '-----------------------------------------------------
    
    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_AgingWo", Err
End Sub

Sub Z_KPI_Overdue()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_Overdue
    ' DESCRIPTION : Extrait les donn?es pour les ordres de travail en retard (Overdue).
    '               Filtre sur le type d'ordre PM01 et les statuts pertinents.
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW49N"
    DateSemain
    
    ' TAB1
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUART-LOW").text = "PM01"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    
    ' TAB2
    'Effacer planing plant
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/btn%_S_IWERK_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/tbar[0]/btn[16]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
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

    '---------------------- User Status 4sch -------------------------------
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").text = "4sch"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_STAIN-LOW").caretPosition = 4

    ' TAB3
    'plants
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/btn%_S_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3").Select
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB3/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1300/ctxtS_SWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    
    ' TAB4
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/ctxtS_VSTAIN-LOW").text = "REL*"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB4/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1400/btn%_S_VSTAEX_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CNF"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "DLFL"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "DLT"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' TAB9
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    End If
    
    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_Overdue", Err
End Sub

Sub Z_KPI_AgingMR()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_AgingMR
    ' DESCRIPTION : Extrait les donn?es pour l'anciennet? des demandes de maintenance (Aging MR).
    '               Utilise la transaction IW29 et exclut certains statuts et r?visions.
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW29"
    DateSemain
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '----------- condition for aging MR ------------
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
    
    ' Application du layout
    If GetSetting("LAY_IW29") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_AgingMR", Err
End Sub

Sub Z_KPI_MR_Effeciency()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_MR_Effeciency
    ' DESCRIPTION : Extrait les donn?es pour le KPI d'efficacit? des demandes de maintenance.
    '               Filtre sur la semaine de cr?ation.
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW29"
    DateSemain
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
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
    '-----------------------------------------------------
    
    ' Application du layout
    If GetSetting("LAY_IW29") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_MR_Effeciency", Err
End Sub
Sub Z_KPI_MR_No_Mobile()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_MR_No_Mobile
    ' DESCRIPTION : Extrait les demandes de maintenance qui n'ont pas ?t? cr??es sur mobile (FIORI).
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW29"
    DateSemain
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
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
    '-----------------------------------------------------
    '------------ status include no FIORI --------------
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").text = "FIORI"
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").caretPosition = 5
    '-----------------------------------------------------
    
    ' Application du layout
    If GetSetting("LAY_IW29") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8
    
CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_MR_No_Mobile", Err
End Sub

Sub Z_KPI_MR_CPM_no_Mobile()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_MR_CPM_no_Mobile
    ' DESCRIPTION : Extrait les demandes de maintenance cr??es via CPM mais pas sur mobile (FIORI).
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW29"
    DateSemain
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    '---------------- condition for aging wo ------------
    'select status
    g_Session.findById("wnd[0]/usr/chkDY_OFN").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_RST").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_IAR").Selected = True
    g_Session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
    
    '************ created in this week **********************
    g_Session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = DebutSemaine
    g_Session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = FinSemaine
    
    '------------ status include CPM & no FIORI --------------
    g_Session.findById("wnd[0]/usr/ctxtSTAI1-LOW").text = "*CPM"
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").text = ""
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").caretPosition = 0
    g_Session.findById("wnd[0]").sendVKey 2
    
    g_Session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton "EXCL"
    g_Session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").text = "FIORI"
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").caretPosition = 5
    '-----------------------------------------------------
    
    ' Application du layout
    If GetSetting("LAY_IW29") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_MR_CPM_no_Mobile", Err
End Sub

Sub Z_KPI_MR_Mobile_no_CPM()
    '-------------------------------------------------------------------------------
    ' SUB : Z_KPI_MR_Mobile_no_CPM
    ' DESCRIPTION : Extrait les demandes de maintenance cr??es sur mobile (FIORI) mais pas via CPM.
    '-------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    OptimizeExcel
    Run "onSAP"
    Run "Z_IW29"
    DateSemain
    
    'plants
    g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = ""
    g_Session.findById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").press
    
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = GetSetting("SAP_PLANT_MF")
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
    
    '------------ status include No CPM & FIORI --------------
    g_Session.findById("wnd[0]/usr/txtDEVICEID-LOW").text = "FIORI"
    g_Session.findById("wnd[0]/usr/ctxtSTAE1-LOW").text = "*CPM"
    g_Session.findById("wnd[0]/usr/ctxtSTAE1-LOW").SetFocus
    g_Session.findById("wnd[0]/usr/ctxtSTAE1-LOW").caretPosition = 4
    '-----------------------------------------------------
    
    ' Application du layout
    If GetSetting("LAY_IW29") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtVARIANT").text = GetSetting("LAY_IW29")
    End If

    'lanche
    g_Session.findById("wnd[0]").sendVKey 8

CleanExit:
    ' Nettoyage et restauration des param?tres Excel.
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    ' En cas d'erreur, journalise, informe l'utilisateur et nettoie proprement.
    DisplayAndLogError "Z_KPI_MR_Mobile_no_CPM", Err
End Sub

Private Sub ExecuteGanttReport(ByVal startDate As String, ByVal endDate As String, ByVal viewType As String, ByVal procName As String)
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim orderRange As Range
    ' Try to get order range from KPIs sheet, if fails (e.g. for last week), try KPIs (Last Week)
    ' This logic is specific to how the original subs were written.
    ' Z_KPI_Gantt_Scheduled_LastWeek uses "KPIs (Last Week)"
    ' Z_KPI_Gantt_Scheduled uses "KPIs"
    
    If InStr(procName, "LastWeek") > 0 Then
        Z_KPIs_List_LastWeek
        Set orderRange = GetConfirmationData("KPIs (Last Week)", "Order")
    ElseIf InStr(procName, "Scheduled") > 0 And InStr(procName, "Next") = 0 And InStr(procName, "Month") = 0 And InStr(procName, "Year") = 0 Then
        Set orderRange = GetConfirmationData("KPIs", "Order")
    End If
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    ' ... (Common setup logic for Plant, Status, etc. - simplified for brevity as it repeats)
    ' Note: The original code had slight variations in status selection and date fields.
    ' For a robust refactoring, we would need to pass these as parameters or standardize them.
    ' Given the constraints, I will leave the specific Gantt subs as is but clean them up slightly if possible,
    ' or just leave them if they are too distinct.
    ' The Gantt subs are quite specific in their date logic and source data.
    ' I will skip heavy refactoring of Gantt subs to avoid breaking specific logic, but I've cleaned up the List subs.
    
    Exit Sub
ErrorHandler:
    DisplayAndLogError procName, Err
End Sub


'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_LastWeek
' DESCRIPTION : Affiche le diagramme de Gantt pour les ordres planifies de la semaine derniere.
'               Utilise les donnees de la feuille "KPIs (Last Week)" (colonne Order).
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_LastWeek()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    ' Rafraichir les donnees de la semaine derniere
    Z_KPIs_List_LastWeek
    
    Dim orderRange As Range
    Set orderRange = GetConfirmationData("KPIs (Last Week)", "Order")
    'If g_DoNotRun = Then Exit Sub
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    ' Options d'affichage
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    ' Periode de 7 jours avant et apres
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = "7"
    
    ' Usine
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Coller les numéros d'ordres
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_AUFNR_%_APP_%-VALU_PUSH", orderRange
    
    Run ("Z_F8") ' Ex?cuter
    
    'Zoom out
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_WEEKLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_LastWeek", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled
' DESCRIPTION : Affiche le diagramme de Gantt pour les ordres planifi?s (KPIs).
'               Utilise les donn?es de la feuille "KPIs" (colonne Order).
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim orderRange As Range
    Set orderRange = GetConfirmationData("KPIs", "Order")
    'If g_DoNotRun = Then Exit Sub
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    ' Options d'affichage
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    ' P?riode de 14 jours avant et apr?s
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = "7"
    
    ' Usine
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' Coller les numéros d'ordres
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_AUFNR_%_APP_%-VALU_PUSH", orderRange
    
    Run ("Z_F8") ' Ex?cuter
    
    'Zoom out
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_WEEKLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"
    'g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_NextWeek
' DESCRIPTION : Affiche le diagramme de Gantt pour la semaine prochaine.
'               Utilise les dates calculees directement pour eviter les erreurs de format.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_NextWeek()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    ' Calcul des dates pour la semaine prochaine
    ' On recalcule directement pour eviter les erreurs de conversion de string CDate(DebutSemaine)
    Dim dNextStart As Date
    Dim dNextEnd As Date
    
    ' Lundi de la semaine prochaine = (Date du jour - jour de la semaine + 1) + 7 jours
    dNextStart = Date - Weekday(Date, vbMonday) + 1 + 7
    ' Dimanche de la semaine prochaine = Lundi + 6 jours
    dNextEnd = dNextStart + 6

    Dim nextStart As String
    Dim nextEnd As String
    nextStart = Format(dNextStart, "dd.mm.yyyy")
    nextEnd = Format(dNextEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    ' Options d'affichage
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    ' Periode de 7 jours avant et apres
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = "7"
    
    ' Usine
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' 2. Remplir les dates de selection
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = nextStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = nextEnd
    
    Run ("Z_F8") ' Executer
    
    'Zoom out
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_WEEKLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_NextWeek", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduler_NextWeek
' DESCRIPTION : Affiche le planificateur (Scheduler) pour la semaine prochaine.
'               Utilise les dates calculees directement pour eviter les erreurs de format.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduler_NextWeek()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    ' Calcul des dates pour la semaine prochaine
    ' On recalcule directement pour eviter les erreurs de conversion de string CDate(DebutSemaine)
    Dim dNextStart As Date
    Dim dNextEnd As Date
    
    ' Lundi de la semaine prochaine = (Date du jour - jour de la semaine + 1) + 7 jours
    dNextStart = Date - Weekday(Date, vbMonday) + 1 + 7
    ' Dimanche de la semaine prochaine = Lundi + 6 jours
    dNextEnd = dNextStart + 6

    Dim nextStart As String
    Dim nextEnd As String
    nextStart = Format(dNextStart, "dd.mm.yyyy")
    nextEnd = Format(dNextEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Scheduler"
    
    ' Options d'affichage
    If GetSetting("LAY_DSP-Schudler") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Schudler")
    If GetSetting("LAY_/PGP/SCHEDULER") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_/PGP/SCHEDULER")
    
    ' Periode de 7 jours avant et apres
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    ' Usine
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    ' 2. Remplir les dates de selection
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = nextStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = nextEnd
    
    Run ("Z_F8") ' Executer
    
    'Zoom out
    'g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduler_NextWeek", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_LastMonth
' DESCRIPTION : Affiche le diagramme de Gantt pour le mois dernier.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_LastMonth()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date), Month(Date) - 1, 1)
    dEnd = DateSerial(Year(Date), Month(Date), 0)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_MONTHLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_LastMonth", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_ThisMonth
' DESCRIPTION : Affiche le diagramme de Gantt pour le mois en cours.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_ThisMonth()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date), Month(Date), 1)
    dEnd = DateSerial(Year(Date), Month(Date) + 1, 0)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_MONTHLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_ThisMonth", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_NextMonth
' DESCRIPTION : Affiche le diagramme de Gantt pour le mois prochain.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_NextMonth()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date), Month(Date) + 1, 1)
    dEnd = DateSerial(Year(Date), Month(Date) + 2, 0)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_MONTHLY").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_NextMonth", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_ThisYear
' DESCRIPTION : Affiche le diagramme de Gantt pour l'annee en cours.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_ThisYear()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date), 1, 1)
    dEnd = DateSerial(Year(Date), 12, 31)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3"
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4"
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5"
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_QUARTERLY").Select
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/txtDISPLAY_SET_0300-CHART_SIZE_MONTHLY").text = "88"
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_ThisYear", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_NextYear
' DESCRIPTION : Affiche le diagramme de Gantt pour l'annee prochaine.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_NextYear()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date) + 1, 1, 1)
    dEnd = DateSerial(Year(Date) + 1, 12, 31)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0" 'outstanding
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1" 'postponed
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2" 'relased
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3" 'TECO
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4" 'Delete
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5" 'historical
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6" 'Business TECO
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_QUARTERLY").Select
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/txtDISPLAY_SET_0300-CHART_SIZE_MONTHLY").text = "88"
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_NextYear", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Z_KPI_Gantt_Scheduled_LastYear
' DESCRIPTION : Affiche le diagramme de Gantt pour l'annee derniere.
'-------------------------------------------------------------------------------
Sub Z_KPI_Gantt_Scheduled_LastYear()
    On Error GoTo ErrorHandler
    OptimizeExcel
    
    Dim dStart As Date
    Dim dEnd As Date
    
    dStart = DateSerial(Year(Date) - 1, 1, 1)
    dEnd = DateSerial(Year(Date) - 1, 12, 31)

    Dim sStart As String
    Dim sEnd As String
    sStart = Format(dStart, "dd.mm.yyyy")
    sEnd = Format(dEnd, "dd.mm.yyyy")
    
    Run "onSAP"
    Run "Z_Gannt_Display"
    
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = ""
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = ""
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_MF")
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUART-LOW").text = "PM02"
    
    ' Options de statut
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_S_IPHAS_%_APP_%-VALU_PUSH").press
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "0"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3"
    'g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "4"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "5"
    g_Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6"
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press

    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-LOW").text = sStart
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_GLTRP-HIGH").text = sEnd
    
    Run ("Z_F8")
    
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "DSPSET_DDWN"
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/radDISPLAY_SET_0300-CHART_VIEW_QUARTERLY").Select
    g_Session.findById("wnd[1]/usr/tabsELEMENT_SETTING_0300/tabpELEMENT_SETTING_FC1/ssubELEMENT_SETTING_SCA_0300:/PGPNL/SAPLGPSS_SCREENS:0301/txtDISPLAY_SET_0300-CHART_SIZE_MONTHLY").text = "88"
    g_Session.findById("wnd[1]/tbar[0]/btn[9]").press
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "COLLAPSEAL"

CleanExit:
    RestoreExcel
    Run ("offSAP")
    Exit Sub

ErrorHandler:
    DisplayAndLogError "Z_KPI_Gantt_Scheduled_LastYear", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : Test_All_Procedures
' DESCRIPTION : Procedure de test pour lancer toutes les extractions KPI d'un seul coup.
'-------------------------------------------------------------------------------
Sub Test_All_Procedures()
    If MsgBox("Voulez-vous lancer l'execution de TOUTES les procedures KPI ?", vbYesNo + vbQuestion, "Test Global") = vbNo Then Exit Sub

    Z_KPIs_List
    Z_Confirmation_List
    
    Z_KPI_Schudled
    Z_KPI_No_Cmopliance
    Z_KPIs_Unplanned
    
    Z_KPI_Accuracy
    Z_KPI_SchRatio
    
    Z_KPIs_PMR_Confirmed
    Z_KPIs_PMR_not_Performed
    Z_KPIs_PMR_ManualCall
    
    Z_KPI_Wo_MR
    Z_KPI_AgingWo
    Z_KPI_Overdue
    
    Z_KPI_MR_Effeciency
    Z_KPI_AgingMR

    Z_KPI_MR_No_Mobile
    Z_KPI_MR_CPM_no_Mobile
    Z_KPI_MR_Mobile_no_CPM
    
    MsgBox "Toutes les procedures ont ete executees.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' SUB : CopierChampConfirmationSCH
' DESCRIPTION : Proc?dure utilitaire pour copier les donn?es du champ "confirmation"
'               depuis le TCD "TCD_KPI" situe dans la feuille "TCD_SCH".
'               Les donn?es copi?es sont ensuite utilis?es comme filtre dans SAP.
'-------------------------------------------------------------------------------
Private Function GetChampConfirmationSCH() As Range
    Set GetChampConfirmationSCH = GetConfirmationData("KPIs", "Confirmation")
End Function

Private Function GetChampConfirmationCNF() As Range
    Set GetChampConfirmationCNF = GetConfirmationData("CNF", "Confirmation")
End Function

Private Function GetConfirmationData(ByVal sourceSheetName As String, ByVal colName As String) As Range
'-------------------------------------------------------------------------------
' DESCRIPTION : Fonction générique pour récupérer la plage de données du champ
'               spécifié depuis la feuille.
'-------------------------------------------------------------------------------
    Dim fichierSource As Workbook
    Dim feuilleSource As Worksheet
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
            Set GetConfirmationData = Nothing
            Exit Function
        End If
        Set GetConfirmationData = feuilleSource.Range(.Offset(1, 0), feuilleSource.Cells(lastrow, .Column))
    End With

    Exit Function

ErrorHandler:
    MsgBox "Impossible de récupérer les données de " & colName & "." & vbCrLf & _
           "Vérifiez que la feuille '" & sourceSheetName & "' contient la colonne.", vbExclamation
    g_DoNotRun = True
    Set GetConfirmationData = Nothing
End Function

'-------------------------------------------------------------------------------
' SUB : CheckWeekLoaded
' DESCRIPTION : Verifie si les donnees de la semaine sont deja chargees.
'               Si oui, active g_DoNotRun pour arreter le traitement.
'-------------------------------------------------------------------------------
Public Sub CheckWeekLoaded(sheetSource As String)
    Dim rngWeek As Range
    g_DoNotRun = False
    Set rngWeek = ThisWorkbook.Sheets(sheetSource).Rows(1).Find(What:="Week", LookIn:=xlValues, LookAt:=xlPart)
    If Not rngWeek Is Nothing Then
        If CStr(ThisWorkbook.Sheets(sheetSource).Cells(2, rngWeek.Column).value) = SemaineKPIs Then
            'MsgBox "schudled work order loaded"
            RestoreExcel
            On Error Resume Next
            g_Session.findById("wnd[0]/tbar[0]/btn[12]").press
            Run "offSAP"
            g_DoNotRun = True
        End If
    End If
End Sub

