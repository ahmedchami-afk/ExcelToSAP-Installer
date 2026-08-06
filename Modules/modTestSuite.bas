Attribute VB_Name = "modTestSuite"
'====================================================================================
' MODULE      : modTestSuite
' VERSION     : 1.1
' AUTEUR      : [Votre Nom] / Révisé par Gemini Code Assist
' DATE        : 03/12/2025
' DESCRIPTION : Module central de gestion des tests automatisés (Test Suite).
'               Il orchestre l'exécution de scénarios de test pour valider la robustesse
'               des interactions entre Excel et SAP. Le module gère les tests unitaires
'               (par objet) et les tests de masse (listes), génère des échantillons
'               aléatoires et consigne les résultats dans un fichier journal.
'
' UTILISATION :
'   - La procédure principale `Run_All_Project_Tests` peut être lancée pour une
'     vérification complète.
'   - Des tests individuels peuvent être déclenchés depuis le ruban "Tests".
'====================================================================================

Option Explicit

'====================================================================================
' SECTION 1 : CONSTANTES ET VARIABLES GLOBALES DU MODULE
'====================================================================================

Private Const TEST_SHEET_NAME As String = "Tests"   ' Nom de la feuille contenant les données de test
Private Const START_ROW As Long = 2                 ' Première ligne contenant des données (après l'entête)

' Mapping des colonnes de la feuille "Tests" pour chaque type d'objet SAP
Private Const COL_MR As Long = 1
Private Const COL_WO As Long = 2
Private Const COL_MAT As Long = 3
Private Const COL_RSV As Long = 4
Private Const COL_PR As Long = 5
Private Const COL_PO As Long = 6
Private Const COL_PMR As Long = 7
Private Const COL_ITEM As Long = 8
Private Const COL_FL As Long = 9
Private Const COL_EQ As Long = 10
Private Const COL_DOC As Long = 11
Private Const COL_PROJ As Long = 12
Private Const COL_NET As Long = 13
Private Const COL_WBS As Long = 14

Private wsTest As Worksheet ' Référence globale vers la feuille de test pour éviter les appels répétés

'====================================================================================
' SECTION 2 : PROCÉDURE PRINCIPALE DE LA SUITE DE TESTS
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Run_All_Project_Tests
' DESCRIPTION : Orchestrateur principal de la suite de tests.
'               Cette procédure initialise l'environnement, sélectionne un échantillon
'               représentatif de données et lance séquentiellement les tests pour chaque module.
'------------------------------------------------------------------------------------
Public Sub Run_All_Project_Tests()
    Dim lastrow As Long
    Dim testRows As Variant
    Dim rowNum As Variant
    Dim logFilePath As String
    Dim wasHidden As Boolean
    g_StopTests = False ' Réinitialise le flag d'arrêt d'urgence
    g_IsTestMode = True ' Active le mode silencieux (supprime les MsgBox bloquantes)

    ' --- Initialisation ---
    ' 1. Chargement des paramètres globaux
    Call LoadConfiguration
    
    ' 2. Instanciation de la feuille de test
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets(TEST_SHEET_NAME)
    On Error GoTo 0
    If wsTest Is Nothing Then
        MsgBox "La feuille de test '" & TEST_SHEET_NAME & "' est introuvable. Opération annulée.", vbCritical, "Erreur de Test"
        Exit Sub
    End If
    
    If wsTest.Visible <> xlSheetVisible Then
        wasHidden = True
        wsTest.Visible = xlSheetVisible
    End If
    wsTest.Select

    ' 3. Validation de la présence de données
    lastrow = wsTest.Cells(wsTest.Rows.count, 1).End(xlUp).row
    If lastrow < START_ROW Then
        MsgBox "Aucune donnée de test trouvée dans la feuille '" & TEST_SHEET_NAME & "'.", vbExclamation, "Données de Test Manquantes"
        GoTo Cleanup
    End If

    Call ClearLogFile ' Réinitialisation du journal d'exécution
    MsgBox "La suite de tests automatiques va commencer." & vbCrLf & vbCrLf & _
           "Les tests seront effectués sur un échantillon de 10 lignes réparties dans le fichier." & vbCrLf & _
           "Les résultats seront enregistrés et le fichier de log s'ouvrira à la fin.", vbInformation, "Lancement des Tests"

    ' --- Sélection de l'échantillon de test ---
    ' Stratégie : Sélectionner un sous-ensemble de lignes pour éviter une exécution trop longue,
    ' tout en assurant une couverture variée des données.
    Randomize
    
    Const SAMPLE_SIZE As Long = 2
    
    If lastrow < START_ROW + SAMPLE_SIZE - 1 Then
        ' Si le volume de données est faible, on teste toutes les lignes disponibles
        Dim arr() As Variant, c As Long
        ReDim arr(0 To lastrow - START_ROW)
        For c = START_ROW To lastrow
            arr(c - START_ROW) = c
        Next
        testRows = arr
    Else
        ' Sinon, on sélectionne des lignes aléatoires réparties uniformément
        Dim i As Long
        Dim tempRows As New Collection
        Dim stepSize As Long
        stepSize = 10
        
        For i = 0 To SAMPLE_SIZE - 1
            tempRows.add Int((stepSize * Rnd) + (START_ROW + i * stepSize))
        Next i
        testRows = CollectionToArray(tempRows)
    End If

    ' --- Exécution des tests unitaires (par ligne d'échantillon) ---
    For Each rowNum In testRows
        LogMessage vbCrLf & "========================================================" & vbCrLf & "====== DÉBUT DES TESTS POUR LA LIGNE " & rowNum & " ======" & vbCrLf & "========================================================"

        Call Test_WorkOrder_Module(rowNum)
        Call Test_Notification_Module(rowNum)
        Call Test_Equipment_Module(rowNum)
        Call Test_FuncLoc_Module(rowNum)
        Call Test_M_Plan_Module(rowNum)
        Call Test_Materials_Module(rowNum)
        Call Test_PurchOrder_Module(rowNum)
        Call Test_Requisition_Module(rowNum)
        Call Test_Documents_Module(rowNum)
        Call Test_Project_Module(rowNum)
        Call Test_Network_Module(rowNum)
        If g_StopTests Then GoTo Cleanup
    Next rowNum

    ' --- Exécution des tests de listes (Mass Processing) ---
    LogMessage vbCrLf & "========================================================" & vbCrLf & "====== DÉBUT DES TESTS DE LISTE ======" & vbCrLf & "========================================================"
    Call Test_WorkOrder_List_Module(testRows)
    Call Test_Notification_List_Module(testRows)
    Call Test_Equipment_List_Module(testRows)
    Call Test_FuncLoc_List_Module(testRows)
    Call Test_M_Plan_List_Module(testRows)
    Call Test_Materials_List_Module(testRows)
    Call Test_PurchOrder_List_Module(testRows)
    Call Test_Requisition_List_Module(testRows)
    If g_StopTests Then GoTo Cleanup

    ' --- Exécution des tests de transactions générales (sans paramètres) ---
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE TRANSACTIONS GÉNÉRALES ======"
    Call Test_General_List_Transactions
    If g_StopTests Then GoTo Cleanup

    ' --- Exécution des tests de robustesse (Listes vides) ---
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTES VIDES ======"
    Call Test_Empty_List_Transactions

    ' --- Clôture de la session de test ---
    MsgBox "Tous les tests sont terminés ! Le fichier de log va s'afficher.", vbInformation, "Fin de la Suite de Tests"
    
Cleanup:
    If wasHidden Then wsTest.Visible = xlSheetHidden
    If Dir(logFilePath) <> "" Then Shell "notepad.exe " & logFilePath, vbNormalFocus
End Sub

'====================================================================================
' SECTION 3 : TESTS UNITAIRES PAR OBJET MÉTIER
' Ces procédures testent les transactions SAP nécessitant un objet unique (ex: IW33).
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_WorkOrder_Module
' DESCRIPTION : Suite de tests pour un Ordre de Travail (WO).
'               Couvre l'affichage, la modification, et les rapports liés.
'------------------------------------------------------------------------------------
Public Sub Test_WorkOrder_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_WO).value)
    
    wsTest.Cells(rowNum, COL_WO).Select
    Call AutomatedTestWrapper("Z_PMO_IW33_1_Main", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_2_Operations", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_3_Components", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_4_Costs", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_5_Planning", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_6_Enhancement", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW33_7_LOG", "IW33", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_1_Main", "IW32", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_2_Operations", "IW32", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_3_Components", "IW32", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW49N_Operations", "IW49N", "S_AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW40", "IW40", "AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_IWBK", "IWBK", "S_AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW47", "IW47", "AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_MB51_MatChar", "MB51", "AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_ME5A", "ME5A", "S_AUFNR-LOW", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_15_Print_Orders", "IW32", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_16_WO_to_BOOM", "IW32", "CAUFVD-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_17_CNF", "IW41", "CORUF-AUFNR", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW32_18_Cancel_CNF", "IW45", "CORUF-AUFNR", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Notification_Module
' DESCRIPTION : Suite de tests pour un Avis de Maintenance (Notification).
'               Vérifie les transactions IW23, IW22 et les liens vers les ordres.
'------------------------------------------------------------------------------------
Public Sub Test_Notification_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_MR).value)
    
    wsTest.Cells(rowNum, COL_MR).Select
    Call AutomatedTestWrapper("Z_IW23NotM", "IW23", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_IW23NotSum", "IW23", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_IW23NotLoc", "IW23", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_IW23NotMul", "IW23", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_IW23NotLog", "IW23", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_IW22NotM", "IW22", "RIWO00-QMNUM", testValue)
    Call AutomatedTestWrapper("Z_PMO_IW30_Single", "IW30", "QMNUM-LOW", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Equipment_Module
' DESCRIPTION : Suite de tests pour un Équipement.
'               Inclut les données de base, nomenclatures (BOM) et gammes (Task Lists).
'------------------------------------------------------------------------------------
Public Sub Test_Equipment_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_EQ).value)
    
    wsTest.Cells(rowNum, COL_EQ).Select
    Call AutomatedTestWrapper("equipmentInfo", "IE03", "RM63E-EQUNR", testValue)
    Call AutomatedTestWrapper("equipmentBOM", "CC04") ' Test de lancement uniquement (navigation complexe)
    Call AutomatedTestWrapper("eqchars", "IE03", "RM63E-EQUNR", testValue)
    Call AutomatedTestWrapper("equipmentPMO", "IW39", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentPMO_H", "IW39", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentPMOc", "IW38", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentNotif", "IW29", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentNotifH", "IW29", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentNotifc", "IW28", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentplan", "IP18", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentplan24", "IP24", "EQUNR-LOW", testValue)
    Call AutomatedTestWrapper("equipmentBOMIB03", "IB03", "MATNR", testValue)
    Call AutomatedTestWrapper("equipmentBOMIB02", "IB02", "MATNR", testValue)
    Call AutomatedTestWrapper("equipmentBOMIB01", "IB01", "MATNR", testValue)
    Call AutomatedTestWrapper("equipmentIA03", "IA03", "RC27E-EQUNR", testValue)
    Call AutomatedTestWrapper("equipmentIA02", "IA02", "RC27E-EQUNR", testValue)
    Call AutomatedTestWrapper("equipmentIA01", "IA01", "RC27E-EQUNR", testValue)
    Call AutomatedTestWrapper("Z_IA17_E_G_Single", "IA17")
    Call AutomatedTestWrapper("Z_IA17_E_G_Single", "IA17")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_FuncLoc_Module
' DESCRIPTION : Suite de tests pour un Poste Technique (Functional Location).
'               Similaire aux équipements mais adapté à la structure hiérarchique.
'------------------------------------------------------------------------------------
Public Sub Test_FuncLoc_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_FL).value)
    
    wsTest.Cells(rowNum, COL_FL).Select
    Call AutomatedTestWrapper("Z_FLOC_IL03", "IL03", "IFLO-TPLNR", testValue)
    Call AutomatedTestWrapper("Z_FLOC_IL03_c", "CL30N")
    Call AutomatedTestWrapper("Z_FLOC_IL03_B", "CC04")
    Call AutomatedTestWrapper("flocPMO", "IW39", "STRNO-LOW", testValue)
    Call AutomatedTestWrapper("flocPMO_H", "IW39", "STRNO-LOW", testValue)
    Call AutomatedTestWrapper("flocPMOc", "IW38", "STRNO-LOW", testValue)
    Call AutomatedTestWrapper("flocNotif", "IW29", "TPLNR-LOW", testValue)
    Call AutomatedTestWrapper("flocNotifH", "IW29", "TPLNR-LOW", testValue)
    Call AutomatedTestWrapper("flocNotifc", "IW28", "TPLNR-LOW", testValue)
    Call AutomatedTestWrapper("flocplan", "IP18", "STRNO-LOW", testValue)
    Call AutomatedTestWrapper("flocplan24", "IP24", "STRNO-LOW", testValue)
    Call AutomatedTestWrapper("FLOCBOMIB03", "IB03", "TPLNR", testValue)
    Call AutomatedTestWrapper("floCBOMIB02", "IB02", "TPLNR", testValue)
    Call AutomatedTestWrapper("flocOMIB01", "IB01", "TPLNR", testValue)
    Call AutomatedTestWrapper("flocIA13", "IA13", "RC27E-TPLNR", testValue)
    Call AutomatedTestWrapper("flocIA12", "IA12", "RC27E-TPLNR", testValue)
    Call AutomatedTestWrapper("flocIA11", "IA11", "RC27E-TPLNR", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_M_Plan_Module
' DESCRIPTION : Suite de tests pour un Plan de Maintenance.
'               Vérifie la planification (IP03/IP02) et l'historique des appels.
'------------------------------------------------------------------------------------
Public Sub Test_M_Plan_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_PMR).value)
    
    wsTest.Cells(rowNum, COL_PMR).Select
    Call AutomatedTestWrapper("displayPLanM", "IP03", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("ChangePLan", "IP02", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("displayPLanOP", "IP03", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("displayPLanCalls", "IP03", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("IP24plan", "IP24", "WARPL-LOW", testValue)
    Call AutomatedTestWrapper("callH", "IP24", "WARPL-LOW", testValue)
    Call AutomatedTestWrapper("lastcall", "IP24", "WARPL-LOW", testValue)
    Call AutomatedTestWrapper("enh", "IP03", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("changeTLgivenPlan", "IP02", "RMIPM-WARPL", testValue)
    Call AutomatedTestWrapper("pmohistplan", "IW39", "WARPL-LOW", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Materials_Module
' DESCRIPTION : Suite de tests pour un Article (Material).
'               Couvre les stocks (MMBE), mouvements (MB51) et données de base.
'------------------------------------------------------------------------------------
Public Sub Test_Materials_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_MAT).value)
    
    wsTest.Cells(rowNum, COL_MAT).Select
    Call AutomatedTestWrapper("stock", "MMBE", "MS_MATNR-LOW", testValue)
    Call AutomatedTestWrapper("basdata", "MM03", "RMMG1-MATNR", testValue)
    Call AutomatedTestWrapper("matOrders", "ME2N", "S_MATNR-LOW", testValue)
    Call AutomatedTestWrapper("matlt", "ME2N", "S_MATNR-LOW", testValue)
    Call AutomatedTestWrapper("matcons", "MB51", "MATNR-LOW", testValue)
    Call AutomatedTestWrapper("mrpe1", "MM03", "RMMG1-MATNR", testValue)
    Call AutomatedTestWrapper("docdata", "CV04N", "DRAW-OBJKY", testValue)
    Call AutomatedTestWrapper("matmov", "MB51", "MATNR-LOW", testValue)
    Call AutomatedTestWrapper("matmb52", "MB52", "MATNR-LOW", testValue)
    Call AutomatedTestWrapper("maatreservations", "MB25", "MATNR-LOW", testValue)
    Call AutomatedTestWrapper("matnotes", "MM03", "RMMG1-MATNR", testValue)
    Call AutomatedTestWrapper("mattied", "CS15", "RC29L-MATNR", testValue)
    Call AutomatedTestWrapper("IW39Mat", "IW39", "MATNR-LOW", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_PurchOrder_Module
' DESCRIPTION : Suite de tests pour une Commande d'Achat (PO).
'------------------------------------------------------------------------------------
Public Sub Test_PurchOrder_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_PO).value)
    
    wsTest.Cells(rowNum, COL_PO).Select
    Call AutomatedTestWrapper("Z_ME23M", "ME23", "RM06E-BSTNR", testValue)
    Call AutomatedTestWrapper("Z_ME23NM", "ME23N")
    Call AutomatedTestWrapper("Z_ME80FNM", "ME80FN", "SP$00001-LOW", testValue)
    Call AutomatedTestWrapper("Z_ME91FM", "ME91F", "EN_EBELN-LOW", testValue)
    Call AutomatedTestWrapper("Z_ME2NM", "ME2N", "EN_EBELN-LOW", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Requisition_Module
' DESCRIPTION : Suite de tests pour une Demande d'Achat (PR).
'------------------------------------------------------------------------------------
Public Sub Test_Requisition_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_PR).value)
    
    wsTest.Cells(rowNum, COL_PR).Select
    Call AutomatedTestWrapper("Z_ME53M", "ME53", "EBAN-BANFN", testValue)
    Call AutomatedTestWrapper("Z_ME53NM", "ME53N")
    Call AutomatedTestWrapper("Z_ME5AReq", "ME5A", "BA_BANFN-LOW", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Item_Module
' DESCRIPTION : Suite de tests pour un Poste d'Entretien (Maintenance Item).
'------------------------------------------------------------------------------------
Public Sub Test_Item_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_ITEM).value)
    If testValue = "" Then Exit Sub
    
    wsTest.Cells(rowNum, COL_ITEM).Select
    Call AutomatedTestWrapper("Z_IP06I", "IP06", "RMIPM-WAPOS", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Documents_Module
' DESCRIPTION : Suite de tests pour la Gestion Documentaire (DMS).
'------------------------------------------------------------------------------------
Public Sub Test_Documents_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_DOC).value)
    If testValue = "" Then Exit Sub ' Ne pas tester si la cellule est vide
    wsTest.Cells(rowNum, COL_DOC).Select
    Call AutomatedTestWrapper("Z_CV03NDoc")
    Call AutomatedTestWrapper("Z_CV02NDoc")
    Call AutomatedTestWrapper("Z_CV03NDoc", "CV03N", "DRAW-DOKNR", testValue)
    Call AutomatedTestWrapper("Z_CV02NDoc", "CV02N", "DRAW-DOKNR", testValue)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Project_Module
' DESCRIPTION : Placeholder pour les tests unitaires de Projets (PS).
'------------------------------------------------------------------------------------
Public Sub Test_Project_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_PROJ).value)
    If testValue = "" Then Exit Sub ' Ne pas tester si la cellule est vide
    
    ' Note: Les tests PS sont principalement orientés "Liste" (voir Test_Project_List_Module).
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Network_Module
' DESCRIPTION : Placeholder pour les tests unitaires de Réseaux.
'------------------------------------------------------------------------------------
Public Sub Test_Network_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_NET).value)
    If testValue = "" Then Exit Sub ' Ne pas tester si la cellule est vide

    ' Note: Voir Test_Network_List_Module.
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_WBS_Module
' DESCRIPTION : Placeholder pour les tests unitaires d'éléments WBS.
'------------------------------------------------------------------------------------
Public Sub Test_WBS_Module(ByVal rowNum As Long)
    Dim testValue As String
    testValue = CStr(wsTest.Cells(rowNum, COL_WBS).value)
    If testValue = "" Then Exit Sub ' Ne pas tester si la cellule est vide

    ' Note: Voir Test_WBS_List_Module.
End Sub

'------------------------------------------------------------------------------------
' SECTION 4 : TESTS DE LISTES (MASS PROCESSING)
' Ces procédures testent les transactions acceptant des sélections multiples (ex: IW39).
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_WorkOrder_List_Module
' DESCRIPTION : Tests de masse pour les Ordres de Travail.
'------------------------------------------------------------------------------------
Public Sub Test_WorkOrder_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES ORDRES DE TRAVAIL ======"
    ' Sélectionne une plage de 10 cellules réparties pour le test de liste
    Dim testRange As Range, i As Long
    Set testRange = BuildSampleRange(sampleRows, COL_WO)
    testRange.Select
    
    Call AutomatedTestWrapper("Z_IW39_O", "IW39")
    Call AutomatedTestWrapper("Z_IW39_H", "IW39")
    Call AutomatedTestWrapper("Z_IW38O", "IW38")
    Call AutomatedTestWrapper("Z_IW49NL", "IW49N")
    Call AutomatedTestWrapper("Z_IW40_L", "IW40")
    Call AutomatedTestWrapper("Z_IWBK_L", "IWBK")
    Call AutomatedTestWrapper("Z_IW47_L", "IW47")
    Call AutomatedTestWrapper("Z_MB51_L", "MB51")
    Call AutomatedTestWrapper("Z_ME5A_L", "ME5A")
    Call AutomatedTestWrapper("Z_PMO_Gantt", "GS_D")
    'Call AutomatedTestWrapper("Z_PMO_IW38_MassRelease", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_IW38_MassTECO", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_IW38_Change_Date", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_Print_Orders", "PPM")
    'Call AutomatedTestWrapper("Z_PMO_SCHEDULER", "SCHEDULER")
    'Call AutomatedTestWrapper("Z_PMO_IW38_4SCH", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_IW38_5COM", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_IW38_User_Status_mass_Change", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_IW38_MassConfirmation", "IW38")
    'Call AutomatedTestWrapper("Z_PMO_Cancel_Confirmations", "IW47")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Notification_List_Module
' DESCRIPTION : Tests de masse pour les Avis de Maintenance.
'------------------------------------------------------------------------------------
Public Sub Test_Notification_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES AVIS ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_MR)
    testRange.Select
    
    Call AutomatedTestWrapper("Z_IW29Open", "IW29")
    Call AutomatedTestWrapper("Z_IW29Closed", "IW29")
    Call AutomatedTestWrapper("Z_PMO_IW30", "IW30")
    Call AutomatedTestWrapper("Z_IW28Open", "IW28")
    Call AutomatedTestWrapper("Z_IW39NOpen", "IW39")
    Call AutomatedTestWrapper("Z_IW39NClosed", "IW39")
    Call AutomatedTestWrapper("Z_IW38NOpen", "IW38")
    Call AutomatedTestWrapper("Z_IWBKN", "IWBK")
    Call AutomatedTestWrapper("Z_IW49NN", "IW49N")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Equipment_List_Module
' DESCRIPTION : Tests de masse pour les Équipements.
'------------------------------------------------------------------------------------
Public Sub Test_Equipment_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES ÉQUIPEMENTS ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_EQ)
    testRange.Select

    Call AutomatedTestWrapper("Z_IA17_E_G_Single", "IA17")
    Call AutomatedTestWrapper("Z_IH08_E", "IH08")
    Call AutomatedTestWrapper("Z_IE07_E", "IE07")
    Call AutomatedTestWrapper("Z_IW39_E", "IW39")
    Call AutomatedTestWrapper("Z_IW39_EC", "IW39")
    Call AutomatedTestWrapper("Z_IW38_E", "IW38")
    Call AutomatedTestWrapper("Z_IW29_E", "IW29")
    Call AutomatedTestWrapper("Z_IW29_EC", "IW29")
    Call AutomatedTestWrapper("Z_IW28_E", "IW28")
    Call AutomatedTestWrapper("Z_IP18_E", "IP18")
    Call AutomatedTestWrapper("Z_IP24_E", "IP24")
    Call AutomatedTestWrapper("Z_IA17_E_G", "IA17")
    Call AutomatedTestWrapper("Z_IA17_E", "IA17")
    Call AutomatedTestWrapper("Z_IA17_E", "IA17")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_FuncLoc_List_Module
' DESCRIPTION : Tests de masse pour les Postes Techniques.
'------------------------------------------------------------------------------------
Public Sub Test_FuncLoc_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES POSTES TECHNIQUES ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_FL)
    testRange.Select

    Call AutomatedTestWrapper("Z_IH06_F", "IH06")
    Call AutomatedTestWrapper("Z_IH08_F", "IH08")
    Call AutomatedTestWrapper("Z_IL07_F", "IL07")
    Call AutomatedTestWrapper("Z_IW39_F", "IW39")
    Call AutomatedTestWrapper("Z_IW39_FC", "IW39")
    Call AutomatedTestWrapper("Z_IW38_F", "IW38")
    Call AutomatedTestWrapper("Z_IW29_F", "IW29")
    Call AutomatedTestWrapper("Z_IW29_FC", "IW29")
    Call AutomatedTestWrapper("Z_IW28_F", "IW28")
    Call AutomatedTestWrapper("Z_IP18_F", "IP18")
    Call AutomatedTestWrapper("Z_IP24_F", "IP24")
    Call AutomatedTestWrapper("Z_IA17_F_G", "IA17")
    Call AutomatedTestWrapper("Z_IA17_F", "IA17")
    Call AutomatedTestWrapper("Z_IA17_F", "IA17")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_M_Plan_List_Module
' DESCRIPTION : Tests de masse pour les Plans de Maintenance.
'------------------------------------------------------------------------------------
Public Sub Test_M_Plan_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES PLANS DE MAINTENANCE ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_PMR)
    testRange.Select

    Call AutomatedTestWrapper("ip16_L", "IP16")
    Call AutomatedTestWrapper("ip24_L", "IP24")
    Call AutomatedTestWrapper("Z_IW39OL", "IW39")
    Call AutomatedTestWrapper("pmohistplanL", "IW39")
    Call AutomatedTestWrapper("Z_IW38OL", "IW38")
    Call AutomatedTestWrapper("Manual_Call_PMRs", "IP10")
    Call AutomatedTestWrapper("Change_StartCycle_PMRs", "IP10")
    Call AutomatedTestWrapper("Start_PMRs", "IP10")
    Call AutomatedTestWrapper("NewStart_PMRs", "IP10")
    Call AutomatedTestWrapper("Planification_PMRs", "IP30")
    Call AutomatedTestWrapper("Afficher_Arbo_PMR", "IP03")
    ' R_Afficher_Arbo_TL -> Afficher_Arbo_TL -> Z_IA13
    Call AutomatedTestWrapper("Afficher_Arbo_TL", "IA13")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Materials_List_Module
' DESCRIPTION : Tests de masse pour les Articles.
'------------------------------------------------------------------------------------
Public Sub Test_Materials_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES ARTICLES ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_MAT)
    testRange.Select

    Call AutomatedTestWrapper("IW38MatL", "IW38")
    Call AutomatedTestWrapper("IWBKmatL", "IWBK")
    Call AutomatedTestWrapper("matOrdersL", "ME2N")
    Call AutomatedTestWrapper("mattiedcs", "CS15")
    Call AutomatedTestWrapper("IW39MatL", "IW39")
    Call AutomatedTestWrapper("IWBKmat", "IWBK")
    Call AutomatedTestWrapper("maatreservationsL", "MB25")
    Call AutomatedTestWrapper("matmb52L", "MB52")
    Call AutomatedTestWrapper("matmovL", "MB51")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_PurchOrder_List_Module
' DESCRIPTION : Tests de masse pour les Commandes d'Achat.
'------------------------------------------------------------------------------------
Public Sub Test_PurchOrder_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES COMMANDES D'ACHAT ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_PO)
    testRange.Select

    Call AutomatedTestWrapper("Z_ME2NML", "ME2N")
    Call AutomatedTestWrapper("Z_ME80FNML", "ME80FN")
    Call AutomatedTestWrapper("Z_ME91FML", "ME91F")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Requisition_List_Module
' DESCRIPTION : Tests de masse pour les Demandes d'Achat.
'------------------------------------------------------------------------------------
Public Sub Test_Requisition_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES DEMANDES D'ACHAT ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_PR)
    testRange.Select

    Call AutomatedTestWrapper("Z_ME5AReqL", "ME5A")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Item_List_Module
' DESCRIPTION : Tests de masse pour les Postes d'Entretien.
'------------------------------------------------------------------------------------
Public Sub Test_Item_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES POSTES DE MAINTENANCE ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_ITEM)
    testRange.Select

    Call AutomatedTestWrapper("Z_IP18LI", "IP18")
    Call AutomatedTestWrapper("IP24_LI", "IP24")
    Call AutomatedTestWrapper("pmohistplanLI", "IW39")
    Call AutomatedTestWrapper("Z_IW39OLI", "IW39")
    Call AutomatedTestWrapper("Z_IW38OLI", "IW38")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Project_List_Module
' DESCRIPTION : Tests de masse pour les Projets.
'------------------------------------------------------------------------------------
Public Sub Test_Project_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES PROJETS ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_PROJ)
    testRange.Select

    Call AutomatedTestWrapper("Z_CN42NProjL", "CN42N")
    Call AutomatedTestWrapper("Z_CN43NProjL", "CN43N")
    Call AutomatedTestWrapper("Z_CNS41ProjL", "CNS41")
    Call AutomatedTestWrapper("Z_CN46NProjL", "CN46N")
    Call AutomatedTestWrapper("Z_CN47ProjL", "CN47")
    Call AutomatedTestWrapper("Z_CN47NProjL", "CN47N")
    Call AutomatedTestWrapper("Z_CN48NProjL", "CN48N")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Document_List_Module
' DESCRIPTION : Tests de masse pour les Documents.
'------------------------------------------------------------------------------------
Public Sub Test_Document_List_Module(ByVal sampleRows As Variant)
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES DOCUMENTS ======"
    Call AutomatedTestWrapper("Z_CV04NDocL", "CV04N")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Network_List_Module
' DESCRIPTION : Tests de masse pour les Réseaux.
'------------------------------------------------------------------------------------
Public Sub Test_Network_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES RÉSEAUX ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_NET)
    testRange.Select

    Call AutomatedTestWrapper("Z_CN46NNetL", "CN46N")
    Call AutomatedTestWrapper("Z_CNS41NetL", "CNS41")
    Call AutomatedTestWrapper("Z_ME5ANetL", "ME5A")
    Call AutomatedTestWrapper("Z_CN47NetL", "CN47")
    Call AutomatedTestWrapper("Z_CN47NNetL", "CN47N")
    Call AutomatedTestWrapper("Z_CN48NNetL", "CN48N")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_WBS_List_Module
' DESCRIPTION : Tests de masse pour les éléments WBS.
'------------------------------------------------------------------------------------
Public Sub Test_WBS_List_Module(ByVal sampleRows As Variant)
    If IsEmpty(sampleRows) Then Exit Sub
    LogMessage vbCrLf & "====== DÉBUT DES TESTS DE LISTE POUR LES WBS ======"
    Dim testRange As Range
    Set testRange = BuildSampleRange(sampleRows, COL_WBS)
    testRange.Select

    Call AutomatedTestWrapper("Z_CN43N_WBSL", "CN43N")
    Call AutomatedTestWrapper("Z_CN46N_WBSL", "CN46N")
    Call AutomatedTestWrapper("Z_CN47_WBSL", "CN47")
    Call AutomatedTestWrapper("Z_CN47N_WBSL", "CN47N")
    Call AutomatedTestWrapper("Z_CN48N_WBSL", "CN48N")
    Call AutomatedTestWrapper("Z_ME5A_WBSL", "ME5A")
    Call AutomatedTestWrapper("Z_IW38_WBSL", "IW38")
End Sub

'====================================================================================
' SECTION 5 : TESTS SPÉCIAUX ET ROBUSTESSE
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_General_List_Transactions
' DESCRIPTION : Vérifie les macros qui lancent des transactions globales
'               (celles ne nécessitant pas de sélection préalable).
'------------------------------------------------------------------------------------
Public Sub Test_General_List_Transactions()
    Call AutomatedTestWrapper("Z_IW39", "IW39")
    Call AutomatedTestWrapper("Z_ME2N", "ME2N")
    Call AutomatedTestWrapper("Z_CN46N", "CN46N")
    Call AutomatedTestWrapper("Z_CV04NDoc", "CV04N")
    Call AutomatedTestWrapper("IH01", "IH01")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Test_Empty_List_Transactions
' DESCRIPTION : Test de robustesse ("Negative Testing").
'               Vérifie que l'application gère correctement l'absence de sélection.
'------------------------------------------------------------------------------------
Public Sub Test_Empty_List_Transactions()
    wsTest.Range("Z1").Select ' Force la sélection d'une cellule vide
    
    Call AutomatedTestWrapper("Z_IW39_O", "IW39", , , False)
    Call AutomatedTestWrapper("Z_IW29Open", "IW29", , , False)
End Sub

'====================================================================================
' SECTION 6 : POINTS D'ENTRÉE POUR LE RUBAN (CALLBACKS)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURES  : Run_Single_Test_*
' DESCRIPTION : Wrappers publics exposés au Ruban Excel.
'               Ils initialisent l'environnement et lancent le module de test spécifique.
'------------------------------------------------------------------------------------
Public Sub Run_Single_Test_WorkOrder()
    ClearLogFile
    RunSingleTestModule "Test_WorkOrder_Module", "Test_WorkOrder_List_Module"
End Sub
Public Sub Run_Single_Test_Notification()
    ClearLogFile
    RunSingleTestModule "Test_Notification_Module", "Test_Notification_List_Module"
End Sub
Public Sub Run_Single_Test_Equipment()
    ClearLogFile
    RunSingleTestModule "Test_Equipment_Module", "Test_Equipment_List_Module"
End Sub
Public Sub Run_Single_Test_FuncLoc()
    ClearLogFile
    RunSingleTestModule "Test_FuncLoc_Module", "Test_FuncLoc_List_Module"
End Sub
Public Sub Run_Single_Test_M_Plan()
    ClearLogFile
    RunSingleTestModule "Test_M_Plan_Module", "Test_M_Plan_List_Module"
End Sub
Public Sub Run_Single_Test_Materials()
    ClearLogFile
    RunSingleTestModule "Test_Materials_Module", "Test_Materials_List_Module"
End Sub
Public Sub Run_Single_Test_PurchOrder()
    ClearLogFile
    RunSingleTestModule "Test_PurchOrder_Module", "Test_PurchOrder_List_Module"
End Sub
Public Sub Run_Single_Test_Requisition()
    ClearLogFile
    RunSingleTestModule "Test_Requisition_Module", "Test_Requisition_List_Module"
End Sub
Public Sub Run_Single_Test_Documents()
    ClearLogFile
    RunSingleTestModule "Test_Documents_Module", "Test_Document_List_Module"
End Sub
Public Sub Run_Single_Test_Project()
    ClearLogFile
    RunSingleTestModule "Test_Project_Module", "Test_Project_List_Module"
End Sub
Public Sub Run_Single_Test_Network()
    ClearLogFile
    RunSingleTestModule "Test_Network_Module", "Test_Network_List_Module"
End Sub
Public Sub Run_Single_Test_WBS()
    ClearLogFile
    RunSingleTestModule "Test_WBS_Module", "Test_WBS_List_Module"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : RunSingleTestModule (Privée)
' DESCRIPTION : Logique générique d'exécution d'un module de test.
'               Gère la récupération de l'échantillon et l'appel aux sous-routines.
'------------------------------------------------------------------------------------
Private Sub RunSingleTestModule(ByVal singleTestProc As String, ByVal listTestProc As String)
    Dim testRows As Variant
    Dim wasHidden As Boolean
    
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets(TEST_SHEET_NAME)
    On Error GoTo 0
    
    If Not wsTest Is Nothing And wsTest.Visible <> xlSheetVisible Then
        wasHidden = True
        wsTest.Visible = xlSheetVisible
    End If
    
    testRows = GetTestSampleRows() ' Fonction privée pour obtenir l'échantillon
    Call ClearLogFile ' Nettoyage du fichier de log pour chaque exécution de module unique
    
    Dim rowNum As Variant
    For Each rowNum In testRows
        Application.Run singleTestProc, rowNum
    Next rowNum
    
    If listTestProc <> "" Then Application.Run listTestProc, testRows
    
    MsgBox "Test pour le module '" & singleTestProc & "' terminé.", vbInformation
    
    If wasHidden Then wsTest.Visible = xlSheetHidden
End Sub

'------------------------------------------------------------------------------------
' FONCTION    : GetTestSampleRows (Privée)
' DESCRIPTION : Génère un échantillon aléatoire de lignes à tester.
' RETOUR      : Un tableau (Array) contenant les numéros de lignes sélectionnés.
'------------------------------------------------------------------------------------
Private Function GetTestSampleRows() As Variant
    Dim lastrow As Long
    
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets(TEST_SHEET_NAME)
    On Error GoTo 0
    If wsTest Is Nothing Then
        GetTestSampleRows = Array() ' Retourne un tableau vide si la feuille n'existe pas
        Exit Function
    End If
    
    ' wsTest.Visible = xlSheetVisible ' Supprimé : géré par l'appelant
    
    lastrow = wsTest.Cells(wsTest.Rows.count, 1).End(xlUp).row
    
    Randomize
    
    Const SAMPLE_SIZE As Long = 2
    
    If lastrow < START_ROW Then
        GetTestSampleRows = Array() ' Aucune donnée
    ElseIf lastrow < START_ROW + SAMPLE_SIZE - 1 Then ' Pas assez de données pour un échantillon complet
        Dim arr() As Variant, c As Long
        ReDim arr(0 To lastrow - START_ROW)
        For c = START_ROW To lastrow
            arr(c - START_ROW) = c
        Next
        GetTestSampleRows = arr
    Else
        ' Sélection aléatoire distribuée
        Dim i As Long
        Dim tempRows As New Collection
        Dim stepSize As Long
        stepSize = 10 'Application.Max(10, Int((lastRow - START_ROW + 1) / SAMPLE_SIZE))
        
        For i = 0 To SAMPLE_SIZE - 1
            tempRows.add Int((stepSize * Rnd) + (START_ROW + i * stepSize))
        Next i
        GetTestSampleRows = CollectionToArray(tempRows)
    End If
End Function

'------------------------------------------------------------------------------------
' FONCTION    : CollectionToArray (Privée)
' DESCRIPTION : Utilitaire de conversion Collection -> Array.
'------------------------------------------------------------------------------------
Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To col.count - 1)
    For i = 1 To col.count
        arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function

'
'------------------------------------------------------------------------------------
' FONCTION    : BuildSampleRange (Privée)
' DESCRIPTION : Construit un objet Range Excel discontinu correspondant à l'échantillon.
'               Nécessaire pour simuler une sélection multiple par l'utilisateur.
'------------------------------------------------------------------------------------
Private Function BuildSampleRange(sampleRows As Variant, ByVal columnNum As Long) As Range
    Dim resultRange As Range
    Dim i As Long
    
    If Not IsArray(sampleRows) Or UBound(sampleRows) < LBound(sampleRows) Then Exit Function
    
    Set resultRange = wsTest.Cells(sampleRows(LBound(sampleRows)), columnNum)
    
    For i = LBound(sampleRows) + 1 To UBound(sampleRows)
        Set resultRange = Union(resultRange, wsTest.Cells(sampleRows(i), columnNum))
    Next i
    
    Set BuildSampleRange = resultRange
End Function

'------------------------------------------------------------------------------------
' PROCÉDURE   : ClearLogFile (Privée)
' DESCRIPTION : Réinitialise le fichier journal pour une nouvelle session de test.
'------------------------------------------------------------------------------------
Private Sub ClearLogFile()
    Dim logFilePath As String
    logFilePath = GetSetting("PATH_LOGS") & "ExcelToSAP_Log.txt"
    On Error Resume Next
    Kill logFilePath
    On Error GoTo 0
    LogMessage "Ancien fichier de log supprimé. Début de la nouvelle suite de tests."
End Sub

'====================================================================================
' SECTION 8 : MOTEUR D'EXÉCUTION DES TESTS (CORE)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : AutomatedTestWrapper (Privée)
' DESCRIPTION : Cœur du système de test. Cette procédure encapsule l'appel à la macro
'               à tester, gère la connexion SAP, mesure le temps d'exécution, et
'               valide le résultat (changement d'écran, messages d'erreur SAP).
' PARAMÈTRES  :
'   - procedureToTest (String) : Le nom de la macro à exécuter (ex: "equipmentInfo").
'   - expectedTCode (String)   : Le code de transaction attendu après l'exécution.
'   - fieldId (String)         : L'ID du champ SAP à vérifier (optionnel).
'   - expectedValue (String)   : La valeur attendue dans le champ (optionnel).
'   - checkForGrid (Boolean)   : Indique s'il faut vérifier la présence d'une grille (optionnel).
'------------------------------------------------------------------------------------
Private Sub AutomatedTestWrapper(ByVal procedureToTest As String, Optional ByVal expectedTCode As String = "", Optional ByVal fieldId As String = "", Optional ByVal expectedValue As String = "", Optional ByVal checkForGrid As Boolean = False)
    Dim success As Boolean
    Dim currentTCode As String
    Dim fieldValue As String
    Dim statusType As String, statusText As String
    Dim startTime As Single, endTime As Single
    Dim errorMessage As String
    Dim startTestMessage As String
    Dim testErrorNumber As Long, testErrorDescription As String
    Dim windowTitleBefore As String, windowTitleAfter As String
    Const TIMEOUT_SECONDS As Long = 30
    success = True
    g_IsTestMode = True ' Mode silencieux
    
    ' Étape 1 : Nettoyage de l'environnement SAP
    Call CloseSecondaryWindows

    On Error GoTo TestWrapperErrorHandler

    startTestMessage = "--- Test en cours : " & procedureToTest & " ---"

    ' Étape 2 : Connexion à SAP
    If Not onSAP() Then ' Ouvre et prépare la session SAP
        errorMessage = "Impossible de se connecter à SAP."
        success = False
        GoTo EndTest
    End If

    ' Capture le titre de la fenêtre AVANT l'exécution
    windowTitleBefore = g_Session.ActiveWindow.text

    ' Étape 3 : Exécution de la macro et chronométrage
    startTime = Timer

    ' Exécute la procédure à tester
    Application.Run procedureToTest

    ' Arrête le chronomètre et vérifie le dépassement
    endTime = Timer
    If endTime - startTime > TIMEOUT_SECONDS Then
        testErrorDescription = "Le test a dépassé le temps imparti de " & TIMEOUT_SECONDS & " secondes."
        LogMessage startTestMessage
        LogMessage "    -> ÉCHEC (Timeout) : " & testErrorDescription & vbCrLf
        success = False
        GoTo EndTest ' Passe directement à la fin sans faire les autres vérifications
    End If

    ' Étape 4 : Validation des résultats

    ' Validation 1 : Vérifier les messages d'erreur critiques dans la barre de statut
    If Not g_Session Is Nothing Then
        On Error Resume Next
        statusType = g_Session.findById("wnd[0]/sbar").MessageType
        statusText = g_Session.findById("wnd[0]/sbar").text
        If statusType = "E" Or statusType = "A" Then
            testErrorDescription = "Message d'erreur SAP détecté. Type: " & statusType & ", Texte: " & statusText
            success = False
        End If
        On Error GoTo TestWrapperErrorHandler
    Else
        testErrorDescription = "La session SAP a été perdue après l'exécution de la procédure."
        success = False
    End If


    ' Validation 2 : Vérifier si la fenêtre SAP a changé (signe d'une action)
    windowTitleAfter = g_Session.ActiveWindow.text
    If windowTitleBefore = windowTitleAfter Then
        testErrorDescription = "La fenêtre SAP n'a pas changé après l'exécution. Test considéré comme un échec."
        success = False
    End If

    On Error GoTo 0

EndTest:
    ' Étape 5 : Enregistrement des résultats dans le log
    If success Then
        ' Succès : Pas de log pour éviter le bruit
    Else
        ' Échec : Log détaillé
        LogMessage startTestMessage & " (Procédure: " & procedureToTest & ")"
        LogMessage "    -> ÉCHEC : " & testErrorDescription & vbCrLf
    End If


    ' Étape 6 : Nettoyage après le test
    offSAP
    
    Exit Sub
    
TestWrapperErrorHandler:
    testErrorNumber = Err.Number
    testErrorDescription = Err.Description
    On Error GoTo 0
    
    LogMessage startTestMessage & " (Procédure: " & procedureToTest & ")"
    LogMessage "    -> ÉCHEC (Erreur VBA) : Erreur " & testErrorNumber & ": " & testErrorDescription & vbCrLf
    success = False
    
    On Error Resume Next
    offSAP
    On Error GoTo 0
End Sub

'====================================================================================
' SECTION 9 : IMPORTATION DE DONNÉES (DATA SEEDING)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Populate_Test_Sheet_From_Exports
' DESCRIPTION : Alimente la feuille "Tests" avec des données réelles.
'               Les données sont extraites des fichiers Excel générés par les
'               procédures d'export (dossier PATH_EXPORT).
'------------------------------------------------------------------------------------
Public Sub Populate_Test_Sheet_From_Exports()
    Dim exportPath As String
    Dim dashboardPath As String
    Dim wasHidden As Boolean
    
    ' Chargement de la configuration
    Call LoadConfiguration
    exportPath = GetSetting("PATH_EXPORT")
    dashboardPath = GetSetting("DASHBOARD_PATH")
    
    ' Vérification de la feuille de test
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets(TEST_SHEET_NAME)
    If wsTest Is Nothing Then
        MsgBox "La feuille '" & TEST_SHEET_NAME & "' n'existe pas.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If wsTest.Visible <> xlSheetVisible Then
        wasHidden = True
        wsTest.Visible = xlSheetVisible
    End If
    wsTest.Activate
    
    ' Nettoyage des données existantes (sauf en-têtes)
    wsTest.Range("A" & START_ROW - 1 & ":N" & wsTest.Rows.count).ClearContents

    
    ' Importation des données depuis les fichiers d'export (B_Load_Files)
    ' WO.xlsx -> Work Order, Functional Location, Equipment
    ImportDataFromExcel exportPath & "WO.xlsx", Array("Order", "Aufnr", "Ordre"), COL_WO
    ImportDataFromExcel exportPath & "WO.xlsx", Array("Functional location", "Tplnr", "Loc. fonction."), COL_FL
    ImportDataFromExcel exportPath & "WO.xlsx", Array("Equipment", "Equnr", "Equipement"), COL_EQ
    
    ' MR.xlsx -> Notification
    ImportDataFromExcel exportPath & "MR.xlsx", Array("Notification", "Qmnum", "Avis"), COL_MR
    
    ' PMR.xlsx -> Maintenance Plan, Item
    ImportDataFromExcel exportPath & "PMR.xlsx", Array("Maintenance Plan", "Warpl", "Plan entretien"), COL_PMR
    ImportDataFromExcel exportPath & "PMR.xlsx", Array("Maintenance Item", "Wapos", "Poste entretien"), COL_ITEM
    
    ' PR.xlsx -> Purchase Requisition
    ImportDataFromExcel exportPath & "PR.xlsx", Array("Purchase requisition", "Banfn", "Dde d'achat"), COL_PR
    
    ' PO.xlsx -> Purchase Order
    ImportDataFromExcel exportPath & "PO.xlsx", Array("Purchasing Document", "Ebeln", "Doc.achat"), COL_PO
    
    ' GM.XLSX -> Material
    ImportDataFromExcel exportPath & "GM.XLSX", Array("Material", "Matnr", "Article"), COL_MAT
    
    ' RSV.XLSX -> Material
    ImportDataFromExcel exportPath & "RSV.XLSX", Array("Reservation", "Matnr", "Article"), COL_RSV
    
    ' WO.XLSX -> Material
    ImportDataFromExcel exportPath & "WO.XLSX", Array("WBS order header", "Posid", "Element OTP"), COL_WBS
    
    MsgBox "La feuille de tests a été remplie avec les données des fichiers d'export.", vbInformation
    If wasHidden Then wsTest.Visible = xlSheetHidden
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ImportDataFromExcel (Privée)
' DESCRIPTION : Utilitaire d'importation de colonne.
'               Recherche une colonne dans un fichier source via des mots-clés,
'               et copie les 50 premières valeurs vers la colonne de destination.
'------------------------------------------------------------------------------------
Private Sub ImportDataFromExcel(filePath As String, headerKeywords As Variant, destCol As Long)
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim headerCell As Range
    Dim keyword As Variant
    Dim lastrow As Long
    
    If Dir(filePath) = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    If Not wbSource Is Nothing Then
        Set wsSource = wbSource.Sheets(1)
        For Each keyword In headerKeywords
            Set headerCell = wsSource.Rows(1).Find(What:=keyword, LookIn:=xlValues, LookAt:=xlWhole)
            If Not headerCell Is Nothing Then Exit For
        Next keyword
        
        If Not headerCell Is Nothing Then
            ' Ecrit l'entête standard (premier mot-clé)
            wsTest.Cells(1, destCol).value = headerKeywords(0)
            
            lastrow = wsSource.Cells(wsSource.Rows.count, headerCell.Column).End(xlUp).row
            If lastrow > 1 Then
                If lastrow > 51 Then lastrow = 51 ' Limite à 50 lignes pour les tests
                wsSource.Range(wsSource.Cells(1, headerCell.Column), wsSource.Cells(lastrow, headerCell.Column)).Copy
                wsTest.Cells(1, destCol).PasteSpecial xlPasteValues
            End If
        End If
        wbSource.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    With wsTest.Rows(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    wsTest.Rows(1).Font.Bold = True
    
    wsTest.Cells.EntireColumn.AutoFit
    ' Range("A1").Select ' Supprimé pour éviter l'erreur si masqué
    
    
    
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Toggle_Test_Sheet
' DESCRIPTION : Affiche ou masque la feuille de tests.
'------------------------------------------------------------------------------------
Public Sub Toggle_Test_Sheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TEST_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille de test '" & TEST_SHEET_NAME & "' est introuvable.", vbExclamation
        Exit Sub
    End If
    
    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetHidden
    Else
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
End Sub

