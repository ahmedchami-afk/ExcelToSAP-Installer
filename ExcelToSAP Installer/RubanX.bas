Attribute VB_Name = "RubanX"
'====================================================================================
' MODULE      : RubanX------------------
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module gère toutes les interactions avec le ruban personnalisé de
'               l'application ExcelToSAP. Il agit comme un contrôleur qui lie les
'               boutons, menus et autres contrôles de l'interface utilisateur (définis
'               dans le fichier XML du ruban) aux macros VBA correspondantes.
'
'               Chaque procédure publique est un "callback" déclenché par une action
'               de l'utilisateur sur le ruban (ex: clic sur un bouton).
'====================================================================================
Option Explicit

'====================================================================================
' SECTION 1 : DÉCLARATIONS GLOBALES
'====================================================================================

Public g_ribbonUI As IRibbonUI              ' Objet pour manipuler le ruban (ex: rafraîchir un contrôle).
Public g_chkDoNotRun_Pressed As Boolean     ' État de la case à cocher "Do Not Run".
Public g_isSingleCellSelection As Boolean   ' Indicateur pour savoir si la sélection est une cellule unique.

'====================================================================================
' SECTION 1 : INITIALISATION DU RUBAN
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnRibbonLoad
' DESCRIPTION : Callback exécuté une seule fois lors du chargement du ruban.
'               Initialise l'objet ruban global et lance la procédure d'ouverture.
' PARAMÈTRES  : ribbon (IRibbonUI) - Objet ruban fourni par Excel.
'------------------------------------------------------------------------------------
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set g_ribbonUI = ribbon
    LogMessage "Ruban chargé et objet g_ribbonUI initialisé."
    auto_open
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetTabVisibility
' DESCRIPTION : Contrôle la visibilité des onglets en fonction de l'utilisateur connecté.
'               Implémente la logique "User-Based" :
'               - FULL (Admin/Ahmed) : Tous les onglets.
'               - MEDIUM (Planificateurs) : ExcelToSAP + Week KPIs.
'               - MINIMAL (Autres) : ExcelToSAP uniquement.
' APPELÉE PAR : Le ruban (attribut getVisible sur les balises <tab>).
'------------------------------------------------------------------------------------
Sub GetTabVisibility(control As IRibbonControl, ByRef returnedVal)
    Dim currentUser As String
    Dim userRole As String
    Dim currentPlant As String
    
    ' S'assurer que la configuration est chargée (sécurité si le code a été réinitialisé)
    If ConfigSettings Is Nothing Then LoadConfiguration
    
    ' Récupère le nom d'utilisateur Windows (en minuscules pour la comparaison)
    currentUser = LCase(Environ("USERNAME"))
    
    ' --- DÉFINITION DES RÔLES ---
    Select Case currentUser
        Case "ahmchami", "admin" ' Liste des administrateurs (FULL ACCESS)
            userRole = "FULL"
            
        Case Else
            ' Vérifie si l'usine configurée est l'une des usines autorisées pour le niveau MEDIUM
            currentPlant = UCase(GetSetting("SAP_PLANT_PF"))
            If currentPlant = "0P1D" Then 'Or currentPlant = "0D1D" Or currentPlant = "011D" Then
                userRole = "MEDIUM"
            Else
                ' Par défaut pour tous les autres utilisateurs
                userRole = "MINIMAL"
            End If
    End Select
    
    ' --- LOGIQUE D'AFFICHAGE ---
    Select Case control.ID
        Case "customTab" ' Onglet Principal
            returnedVal = True ' Toujours visible
            
        Case "kpiTab" ' Onglet KPIs
            ' Visible pour FULL et MEDIUM
            returnedVal = (userRole = "FULL" Or userRole = "MEDIUM")
            
        Case "reportingTab" ' Onglet Reporting
            ' Visible uniquement pour FULL
            returnedVal = (userRole = "FULL")
            
        Case Else
            returnedVal = True
    End Select
End Sub

'------------------------------------------------------------------------------------
' FONCTION    : GetRibbon
' DESCRIPTION : Assure la robustesse de l'objet ruban. Si g_ribbonUI a été perdu
'               (par exemple, après une erreur VBA non interceptée), cette fonction
'               tente de le recharger.
' RETOUR      : L'objet IRibbonUI.
'------------------------------------------------------------------------------------
Public Function GetRibbon() As IRibbonUI
    ' Si l'objet ruban a été perdu, on le recharge.
    If g_ribbonUI Is Nothing Then
        ' Cette action force Excel à rappeler le callback OnRibbonLoad
        ' qui va réassigner la variable globale g_ribbonUI.
        ' Application.SendKeys "%{F10}", True ' Désactivé : Provoque l'affichage du volet de sélection (sidebar droite)
    End If
    Set GetRibbon = g_ribbonUI
End Function

'====================================================================================
' SECTION 2 : GESTION DES SESSIONS SAP (GROUPE "SAP CONNECTION")
'====================================================================================

'====================================================================================
' SECTION 2 : GESTION DES SESSIONS SAP (GROUPE "SAP CONNECTION")
'====================================================================================

'--- Callbacks pour la ComboBox de sélection de session (cmbSessions) ---

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetSessionCount
' DESCRIPTION : Retourne le nombre de sessions SAP disponibles pour peupler la ComboBox.
' CONTRÔLE    : cmbSessions (attribut getItemCount)
'------------------------------------------------------------------------------------
Sub GetSessionCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetAvailableSessionsCount()
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetSessionLabel
' DESCRIPTION : Retourne le libellé d'une session SAP à un index donné.
' CONTRÔLE    : cmbSessions (attribut getItemLabel)
'------------------------------------------------------------------------------------
Sub GetSessionLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = GetAvailableSessionLabel(index)
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnSessionChange
' DESCRIPTION : Gère l'événement de changement de sélection dans la ComboBox des sessions.
' CONTRÔLE    : cmbSessions (attribut onAction)
'------------------------------------------------------------------------------------
Sub OnSessionChange(control As IRibbonControl, selectedId As String)
    Dim i As Integer
    Dim selectedIndex As Integer ' Correction : Déclaration de la variable
    selectedIndex = -1

    ' Trouve l'index correspondant à l'ID de l'élément sélectionné
    For i = 0 To GetAvailableSessionsCount() - 1
        If GetAvailableSessionLabel(i) = selectedId Then
            selectedIndex = i
            Exit For
        End If
    Next
    
    If selectedIndex <> -1 Then
        SelectSessionByIndex selectedIndex
        If Not GetRibbon() Is Nothing Then
            GetRibbon().InvalidateControl "cmbSessions"
            GetRibbon().InvalidateControl "cmbSessions_KPI"
        End If
        Run ("onSAP")
    End If
End Sub
 
'--- Actions des boutons liés aux sessions ---

'------------------------------------------------------------------------------------
' PROCÉDURE   : LogON
' DESCRIPTION : Ouvre une nouvelle session SAP via l'écran de connexion.
' CONTRÔLE    : Bouton "Log ON" (btnLogON)
'------------------------------------------------------------------------------------
Sub LogON(control As IRibbonControl)
    SAP_Openg_SessionFromLogon
    
    ' Rafraîchit la ComboBox pour afficher la nouvelle session
    If Not GetRibbon() Is Nothing Then
        GetRibbon().InvalidateControl "cmbSessions"
        GetRibbon().InvalidateControl "cmbSessions_KPI"
    End If
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : NewSAPSession
' DESCRIPTION : Crée une nouvelle session SAP à partir d'une connexion existante.
' CONTRÔLE    : Bouton "New Session" (btnNewSession)
'------------------------------------------------------------------------------------
Sub NewSAPSession(control As IRibbonControl)
    Call SAP_CreateNewSession
    
    ' Rafraîchit la ComboBox
    If Not GetRibbon() Is Nothing Then
        GetRibbon().InvalidateControl "cmbSessions"
        GetRibbon().InvalidateControl "cmbSessions_KPI"
    End If
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : LogOFF
' DESCRIPTION : Ferme toutes les sessions SAP ouvertes et nettoie les variables associées.
' CONTRÔLE    : Bouton "Log OFF" (btnLogOFF)
'------------------------------------------------------------------------------------
Sub LogOFF(control As IRibbonControl)
    On Error GoTo ErrorHandler

    SAP_CloseAllSessions

    ' Nettoyage des variables globales de session
    Set g_Session = Nothing
    Erase g_Sessions

    ' Pause pour laisser le temps à SAP de fermer les processus
    ' Invalide tout le ruban pour forcer une mise à jour complète
    If Not GetRibbon() Is Nothing Then
        GetRibbon().Invalidate
        LogMessage "Ruban rechargé après déconnexion."
    End If
    Exit Sub

ErrorHandler:
    LogMessage "Erreur dans LogOFF : " & Err.Description
    MsgBox "Une erreur est survenue lors de la déconnexion : " & vbCrLf & Err.Description, vbExclamation
    
    ' Tentative de rafraîchissement du ruban malgré l'erreur pour éviter l'incohérence visuelle
    On Error Resume Next
    If Not GetRibbon() Is Nothing Then GetRibbon().Invalidate
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : CloseSelectedSession
' DESCRIPTION : Ferme uniquement la session SAP actuellement sélectionnée dans la ComboBox.
' CONTRÔLE    : Bouton "Close Session" (btnCloseSession)
'------------------------------------------------------------------------------------
Sub CloseSelectedSession(control As IRibbonControl)
    SAP_CloseSelectedSession

    ' Nettoyage de la variable de session active
    Set g_Session = Nothing

    ' Rafraîchit la ComboBox pour retirer la session fermée
    If Not GetRibbon() Is Nothing Then
        GetRibbon().InvalidateControl "cmbSessions"
        GetRibbon().InvalidateControl "cmbSessions_KPI"
    End If
End Sub

'====================================================================================
' SECTION 3 : GROUPES D'ACTIONS PRINCIPALES
'====================================================================================

'--- Groupe "Automatic Data" ---

'------------------------------------------------------------------------------------
' PROCÉDURE   : Main
' DESCRIPTION : Exécute la routine principale de traitement de données.
' CONTRÔLE    : Bouton "Main" (btnMain)
'------------------------------------------------------------------------------------
Sub Main(control As IRibbonControl)
    Main2
End Sub

'--- Groupe "PM Options" ---

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnDoNotRunChange
' DESCRIPTION : Met à jour la variable globale `g_DoNotRun` lorsque l'état de la
'               case à cocher change.
' CONTRÔLE    : CheckBox "Do Not Run" (chkDoNotRun)
'------------------------------------------------------------------------------------
Sub OnDoNotRunChange(control As IRibbonControl, pressed As Boolean)
    g_chkDoNotRun_Pressed = pressed
    If pressed Then g_DoNotRun = True Else g_DoNotRun = False
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetDoNotRunPressed
' DESCRIPTION : Retourne l'état actuel de la case "Do Not Run" pour le ruban.
' CONTRÔLE    : CheckBox "Do Not Run" (chkDoNotRun)
'------------------------------------------------------------------------------------
Sub GetDoNotRunPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = g_chkDoNotRun_Pressed
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetPlanningPlant
' DESCRIPTION : Récupère la valeur de "Planning Plant" depuis la configuration.
' CONTRÔLE    : EditBox "Planning Plant" (txtPlanningPlant)
'------------------------------------------------------------------------------------
Sub GetPlanningPlant(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSetting("SAP_PLANT_PF")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnPlanningPlantChange
' DESCRIPTION : Met à jour le paramètre "SAP_PLANT_PF" lors de la modification.
' CONTRÔLE    : EditBox "Planning Plant" (txtPlanningPlant)
'------------------------------------------------------------------------------------
Sub OnPlanningPlantChange(control As IRibbonControl, text As String)
    Call UpdateSetting("SAP_PLANT_PF", text)
    Call SaveSettings
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetMaintenancePlant
' DESCRIPTION : Récupère la valeur de "Maintenance Plant" depuis la configuration.
' CONTRÔLE    : EditBox "Maintenance Plant" (txtMaintenancePlant)
'------------------------------------------------------------------------------------
Sub GetMaintenancePlant(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSetting("SAP_PLANT_MF")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnMaintenancePlantChange
' DESCRIPTION : Met à jour le paramètre "SAP_PLANT_MF" lors de la modification.
' CONTRÔLE    : EditBox "Maintenance Plant" (txtMaintenancePlant)
'------------------------------------------------------------------------------------
Sub OnMaintenancePlantChange(control As IRibbonControl, text As String)
    Call UpdateSetting("SAP_PLANT_MF", text)
    Call SaveSettings
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetStructur
' DESCRIPTION : Récupère la valeur de "Structur" depuis la configuration.
' CONTRÔLE    : EditBox "Structur" (txtStructur)
'------------------------------------------------------------------------------------
Sub GetStructur(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSetting("SAP_PLANT_AF")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnStructurChange
' DESCRIPTION : Met à jour le paramètre "SAP_PLANT_AF" lors de la modification.
' CONTRÔLE    : EditBox "Structur" (txtStructur)
'------------------------------------------------------------------------------------
Sub OnStructurChange(control As IRibbonControl, text As String)
    Call UpdateSetting("SAP_PLANT_AF", text)
    Call SaveSettings
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetMaintStore
' DESCRIPTION : Récupère la valeur de "Maint Store" depuis la configuration.
' CONTRÔLE    : EditBox "Maint Store" (txtMaintStore)
'------------------------------------------------------------------------------------
Sub GetMaintStore(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSetting("Storage_Location")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnMaintStoreChange
' DESCRIPTION : Met à jour le paramètre "Storage_Location" lors de la modification.
' CONTRÔLE    : EditBox "Maint Store" (txtMaintStore)
'------------------------------------------------------------------------------------
Sub OnMaintStoreChange(control As IRibbonControl, text As String)
    Call UpdateSetting("Storage_Location", text)
    Call SaveSettings
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetPurchOrg
' DESCRIPTION : Récupère la valeur de "Prch Org" depuis la configuration.
' CONTRÔLE    : EditBox "Prch Org" (txtPurchOrg)
'------------------------------------------------------------------------------------
Sub GetPurchOrg(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetSetting("Purchasing_Org")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : OnPurchOrgChange
' DESCRIPTION : Met à jour le paramètre "Purchasing_Org" lors de la modification.
' CONTRÔLE    : EditBox "Prch Org" (txtPurchOrg)
'------------------------------------------------------------------------------------
Sub OnPurchOrgChange(control As IRibbonControl, text As String)
    Call UpdateSetting("Purchasing_Org", text)
    Call SaveSettings
End Sub

'====================================================================================
' SECTION 4 : GROUPE "PLANT MAINTENANCE" (ONGLET "PM")
'====================================================================================

'--- Sous-section : Menu "Work Order" (menuWOSingle) ---

Sub R_Z_PMO_IW33_1_Main(control As IRibbonControl)
    Run ("Z_PMO_IW33_1_Main")
End Sub

Sub R_Z_PMO_IW33_2_Operations(control As IRibbonControl)
    Run ("Z_PMO_IW33_2_Operations")
End Sub

Sub R_Z_PMO_IW33_3_Components(control As IRibbonControl)
    Run ("Z_PMO_IW33_3_Components")
End Sub

Sub R_Z_PMO_IW33_4_Costs(control As IRibbonControl)
    Run ("Z_PMO_IW33_4_Costs")
End Sub

Sub R_Z_PMO_IW33_5_Planning(control As IRibbonControl)
    Run ("Z_PMO_IW33_5_Planning")
End Sub

Sub R_Z_PMO_IW33_6_Enhancement(control As IRibbonControl)
    Run ("Z_PMO_IW33_6_Enhancement")
End Sub

Sub R_Z_PMO_IW33_7_LOG(control As IRibbonControl)
    Run ("Z_PMO_IW33_7_LOG")
End Sub

Sub R_Z_PMO_IW32_1_Main(control As IRibbonControl)
    Run ("Z_PMO_IW32_1_Main")
End Sub

Sub R_Z_PMO_IW32_2_Operations(control As IRibbonControl)
    Run ("Z_PMO_IW32_2_Operations")
End Sub

Sub R_Z_PMO_IW32_3_Components(control As IRibbonControl)
    Run ("Z_PMO_IW32_3_Components")
End Sub

Sub R_Z_PMO_IW49N_Operations(control As IRibbonControl)
    Run ("Z_PMO_IW49N_Operations")
End Sub

Sub R_Z_PMO_IW40(control As IRibbonControl)
    Run ("Z_PMO_IW40")
End Sub

Sub R_Z_PMO_IWBK(control As IRibbonControl)
    Run ("Z_PMO_IWBK")
End Sub

Sub R_Z_PMO_IW47(control As IRibbonControl)
    Run ("Z_PMO_IW47")
End Sub

Sub R_Z_PMO_MB51_MatChar(control As IRibbonControl)
    Run ("Z_PMO_MB51_MatChar")
End Sub

Sub R_Z_PMO_ME5A(control As IRibbonControl)
    Run ("Z_PMO_ME5A")
End Sub

Sub R_Z_PMO_IW32_15_Print_Orders(control As IRibbonControl)
    Run ("Z_PMO_IW32_15_Print_Orders")
End Sub

Sub R_Z_PMO_IW32_16_WO_to_BOOM(control As IRibbonControl)
    Run ("Z_PMO_IW32_16_WO_to_BOOM")
End Sub

Sub R_Z_PMO_IW32_17_CNF(control As IRibbonControl)
    Run ("Z_PMO_IW32_17_CNF")
End Sub

Sub R_Z_PMO_IW32_18_Cancel_CNF(control As IRibbonControl)
    ' 21
    Run ("Z_PMO_IW32_18_Cancel_CNF")
End Sub

'--- Sous-section : Menu "Work Order List" (menuWOList) ---

Sub R_Z_IW39_O(control As IRibbonControl)
    ' 22 ' Index pour "Display Open List (IW39)"
    Run ("Z_IW39_O")
End Sub

Sub R_Z_IW39_H(control As IRibbonControl)
    Run ("Z_IW39_H")
End Sub

Sub R_Z_IW38O(control As IRibbonControl)
    Run ("Z_IW38O")
End Sub

Sub R_Z_IW49NL(control As IRibbonControl)
    Run ("Z_IW49NL")
End Sub

Sub R_Z_IW40_L(control As IRibbonControl)
    Run ("Z_IW40_L")
End Sub

Sub R_Z_IWBK_L(control As IRibbonControl)
    Run ("Z_IWBK_L")
End Sub

Sub R_Z_IW47_L(control As IRibbonControl)
    Run ("Z_IW47_L")
End Sub

Sub R_Z_MB51_L(control As IRibbonControl)
    Run ("Z_MB51_L")
End Sub

Sub R_Z_ME5A_L(control As IRibbonControl)
    Run ("Z_ME5A_L")
End Sub

Sub R_Z_PMO_Gantt(control As IRibbonControl)
    Run ("Z_PMO_Gantt")
End Sub

Sub R_Z_PMO_IW38_MassRelease(control As IRibbonControl)
    Run ("Z_PMO_IW38_MassRelease")
End Sub

Sub R_Z_PMO_IW38_MassTECO(control As IRibbonControl)
    Run ("Z_PMO_IW38_MassTECO")
End Sub

Sub R_Z_PMO_IW38_4SCH(control As IRibbonControl)
    Run ("Z_PMO_IW38_4SCH")
End Sub

Sub R_Z_PMO_IW38_5COM(control As IRibbonControl)
    Run ("Z_PMO_IW38_5COM")
End Sub

Sub R_Z_PMO_IW38_User_Status_mass_Change(control As IRibbonControl)
    Run ("Z_PMO_IW38_User_Status_mass_Change")
End Sub

Sub R_Z_PMO_IW38_ChangeDate(control As IRibbonControl)
    Run ("Z_PMO_IW38_Change_Date")
End Sub

Sub R_Z_PMO_Print_Orders(control As IRibbonControl)
    Run ("Z_PMO_Print_Orders")
End Sub

Sub R_Z_PMO_Mass_Confirmations(control As IRibbonControl)
    Run ("Z_PMO_IW38_MassConfirmation")
End Sub

Sub R_Z_PMO_Cancel_Confirmations(control As IRibbonControl)
    Run ("Z_PMO_Cancel_Confirmations")
End Sub

Sub R_Z_PMO_SCHEDULER(control As IRibbonControl)
    Run ("Z_PMO_SCHEDULER")
End Sub

'--- Sous-section : Menu "Notification" (menuNotifSingle) ---

Sub R_Z_IW23NotM(control As IRibbonControl)
    Run ("Z_IW23NotM")
End Sub

Sub R_Z_IW23NotSum(control As IRibbonControl)
    Run ("Z_IW23NotSum")
End Sub

Sub R_Z_IW23NotLoc(control As IRibbonControl)
    Run ("Z_IW23NotLoc")
End Sub

Sub R_Z_IW23NotMul(control As IRibbonControl)
    Run ("Z_IW23NotMul")
End Sub

Sub R_Z_IW23NotLog(control As IRibbonControl)
    Run ("Z_IW23NotLog")
End Sub

Sub R_Z_PMO_IW30_Single(control As IRibbonControl)
    Run ("Z_PMO_IW30_Single")
End Sub

Sub R_Z_IW22NotM(control As IRibbonControl)
    Run ("Z_IW22NotM")
End Sub

'--- Sous-section : Menu "Notification List" (menuNotifList) ---

Sub R_Z_IW29Open(control As IRibbonControl)
    Run ("Z_IW29Open")
End Sub

Sub R_Z_IW29Closed(control As IRibbonControl)
    Run ("Z_IW29Closed")
End Sub

Sub R_Z_PMO_IW30(control As IRibbonControl)
    Run ("Z_PMO_IW30")
End Sub

Sub R_Z_IW28Open(control As IRibbonControl)
    Run ("Z_IW28Open")
End Sub

Sub R_Z_IW39NOpen(control As IRibbonControl)
    Run ("Z_IW39NOpen")
End Sub

Sub R_Z_IW39NClosed(control As IRibbonControl)
    Run ("Z_IW39NClosed")
End Sub

Sub R_Z_IW38NOpen(control As IRibbonControl)
    Run ("Z_IW38NOpen")
End Sub

Sub R_Z_IWBKN(control As IRibbonControl)
    Run ("Z_IWBKN")
End Sub

Sub R_Z_IW49NN(control As IRibbonControl)
    Run ("Z_IW49NN")
End Sub

'====================================================================================
' SECTION 5 : GROUPE "SAP P2P" (PROCURE-TO-PAY) (ONGLET "P2P")
'====================================================================================

'--- Sous-section : Menus "PO" et "PO List" ---

Sub R_Z_ME23M(control As IRibbonControl)
    Run ("Z_ME23M")
End Sub

Sub R_Z_ME23NM(control As IRibbonControl)
    Run ("Z_ME23NM")
End Sub

Sub R_Z_ME80FNM(control As IRibbonControl)
    Run ("Z_ME80FNM")
End Sub

Sub R_Z_ME91FM(control As IRibbonControl)
    Run ("Z_ME91FM")
End Sub

Sub R_Z_ME2NM(control As IRibbonControl)
    Run ("Z_ME2NM")
End Sub

Sub R_Z_ME2NML(control As IRibbonControl)
    Run ("Z_ME2NML")
End Sub

Sub R_Z_ME80FNML(control As IRibbonControl)
    Run ("Z_ME80FNML")
End Sub

Sub R_Z_ME91FML(control As IRibbonControl)
    Run ("Z_ME91FML")
End Sub

'--- Sous-section : Menus "PR" et "PR List" ---

Sub R_Z_ME53M(control As IRibbonControl)
    Run ("Z_ME53M")
End Sub

Sub R_Z_ME53NM(control As IRibbonControl)
    Run ("Z_ME53NM")
End Sub

Sub R_Z_ME5AReq(control As IRibbonControl)
    Run ("Z_ME5AReq")
End Sub

Sub R_Z_ME5AReqL(control As IRibbonControl)
    Run ("Z_ME5AReqL")
End Sub

'--- Sous-section : Menus "Material" et "Material List" ---

Sub R_stock(control As IRibbonControl)
    Run ("stock")
End Sub

Sub R_basdata(control As IRibbonControl)
    Run ("basdata")
End Sub

Sub R_matOrders(control As IRibbonControl)
    Run ("matOrders")
End Sub

Sub R_matlt(control As IRibbonControl)
    Run ("matlt")
End Sub

Sub R_matcons(control As IRibbonControl)
    Run ("matcons")
End Sub

Sub R_mrpe1(control As IRibbonControl)
    Run ("mrpe1")
End Sub

Sub R_docdata(control As IRibbonControl)
    Run ("docdata")
End Sub

Sub R_matmov(control As IRibbonControl)
    Run ("matmov")
End Sub

Sub R_matmb52(control As IRibbonControl)
    Run ("matmb52")
End Sub

Sub R_maatreservations(control As IRibbonControl)
    Run ("maatreservations")
End Sub

Sub R_matnotes(control As IRibbonControl)
    Run ("matnotes")
End Sub

Sub R_mattied(control As IRibbonControl)
    Run ("mattied")
End Sub

Sub R_IW38MatL(control As IRibbonControl)
    Run ("IW38MatL")
End Sub

Sub R_IWBKmatL(control As IRibbonControl)
    Run ("IWBKmatL")
End Sub

Sub R_matOrdersL(control As IRibbonControl)
    Run ("matOrdersL")
End Sub

Sub R_mattiedcs(control As IRibbonControl)
    Run ("mattiedcs")
End Sub

Sub R_IW39MatL(control As IRibbonControl)
    Run ("IW39MatL")
End Sub

Sub R_IW39Mat(control As IRibbonControl)
    Run ("IW39Mat")
End Sub

Sub R_IW38Mat(control As IRibbonControl)
    Run ("IW38Mat")
End Sub

Sub R_IWBKmat(control As IRibbonControl)
    Run ("IWBKmat")
End Sub

Sub R_maatreservationsL(control As IRibbonControl)
    Run ("maatreservationsL")
End Sub

Sub R_matmb52L(control As IRibbonControl)
    Run ("matmb52L")
End Sub

Sub R_matmovL(control As IRibbonControl)
    Run ("matmovL")
End Sub

'====================================================================================
' SECTION 6 : GROUPE "PMRS" (PLANNED MAINTENANCE ROUTINES) (ONGLET "PMR")
'====================================================================================

'--- Sous-section : Menu "PMR" (menuPlanSingle) ---

Sub R_displayPLanM(control As IRibbonControl)
    Run ("displayPLan")
End Sub

Sub R_ChangePLan(control As IRibbonControl)
    Run ("ChangePLan")
End Sub

Sub R_displayPLanOP(control As IRibbonControl)
    Run ("displayPLanOP")
End Sub

Sub R_displayPLanCalls(control As IRibbonControl)
    Run ("displayPLanCalls")
End Sub

Sub R_IP24plan(control As IRibbonControl)
    Run ("IP24plan")
End Sub

Sub R_callH(control As IRibbonControl)
    Run ("callH")
End Sub

Sub R_lastcall(control As IRibbonControl)
    Run ("lastcall")
End Sub

Sub R_enh(control As IRibbonControl)
    Run ("enh")
End Sub

Sub R_changeTLgivenPlan(control As IRibbonControl)
    Run ("changeTLgivenPlan")
End Sub

Sub R_pmohistplan(control As IRibbonControl)
    Run ("pmohistplan")
End Sub

'--- Sous-section : Menu "PMR List" (menuPlanList) ---

Sub R_ip16_L(control As IRibbonControl)
    Run ("ip16_L")
End Sub

Sub R_ip24_L(control As IRibbonControl)
    Run ("ip24_L")
End Sub

Sub R_Z_IW39OL(control As IRibbonControl)
    Run ("Z_IW39OL")
End Sub

Sub R_pmohistplanL(control As IRibbonControl)
    Run ("pmohistplanL")
End Sub

Sub R_Z_IW38OL(control As IRibbonControl)
    Run ("Z_IW38OL")
End Sub

Sub R_Manual_Call_PMRs(control As IRibbonControl)
    Run ("Manual_Call_PMRs")
End Sub

Sub R_Change_StartCycle_PMRs(control As IRibbonControl)
    Run ("Change_StartCycle_PMRs")
End Sub

Sub R_Start_PMRs(control As IRibbonControl)
    Run ("Start_PMRs")
End Sub

Sub R_NewStart_PMRs(control As IRibbonControl)
    Run ("NewStart_PMRs")
End Sub

Sub R_Planification_PMRs(control As IRibbonControl)
    Run ("Planification_PMRs")
End Sub

Sub R_MP_Auto_1W(control As IRibbonControl)
    Run ("MP_Auto_1W")
End Sub

Sub R_MP_Auto_1M(control As IRibbonControl)
    Run ("MP_Auto_1M")
End Sub

Sub R_MP_Auto_13W(control As IRibbonControl)
    Run ("MP_Auto_13W")
End Sub

Sub R_MP_Auto_1Y(control As IRibbonControl)
    Run ("MP_Auto_1Y")
End Sub

Sub R_Afficher_Arbo_PMR(control As IRibbonControl)
    Run ("Afficher_Arbo_PMR")
End Sub

Sub R_Afficher_Arbo_TL(control As IRibbonControl)
    Run ("Afficher_Arbo_TL")
End Sub

'--- Sous-section : Menu "Item, Item List" (menuItemList) ---

Sub R_Z_IP06I(control As IRibbonControl)
    Run ("Z_IP06I")
End Sub

Sub R_Z_IP18LI(control As IRibbonControl)
    Run ("Z_IP18LI")
End Sub

Sub R_IP24_LI(control As IRibbonControl)
    Run ("IP24_LI")
End Sub

Sub R_pmohistplanLI(control As IRibbonControl)
    Run ("pmohistplanLI")
End Sub

Sub R_Z_IW39OLI(control As IRibbonControl)
    Run ("Z_IW39OLI")
End Sub

Sub R_Z_IW38OLI(control As IRibbonControl)
    Run ("Z_IW38OLI")
End Sub

'====================================================================================
' SECTION 7 : GROUPE "MASTER DATA" (DONNÉES DE BASE) (ONGLET "MASTER DATA")
'====================================================================================

'--- Sous-section : Menus "Equip." et "Equip. List" ---

Sub R_equipmentInfo(control As IRibbonControl)
    Run ("equipmentInfo")
End Sub

Sub R_equipmentBOM(control As IRibbonControl)
    Run ("equipmentBOM")
End Sub

Sub R_eqchars(control As IRibbonControl)
    Run ("eqchars")
End Sub

Sub R_equipmentPMO(control As IRibbonControl)
    Run ("equipmentPMO")
End Sub

Sub R_equipmentPMO_H(control As IRibbonControl)
    Run ("equipmentPMO_H")
End Sub

Sub R_equipmentPMOc(control As IRibbonControl)
    Run ("equipmentPMOc")
End Sub

Sub R_equipmentNotif(control As IRibbonControl)
    Run ("equipmentNotif")
End Sub

Sub R_equipmentNotifH(control As IRibbonControl)
    Run ("equipmentNotifH")
End Sub

Sub R_equipmentNotifc(control As IRibbonControl)
    Run ("equipmentNotifc")
End Sub

Sub R_equipmentplan(control As IRibbonControl)
    Run ("equipmentplan")
End Sub

Sub R_equipmentplan24(control As IRibbonControl)
    Run ("equipmentplan24")
End Sub

Sub R_equipmentBOMIB03(control As IRibbonControl)
    Run ("equipmentBOMIB03")
End Sub

Sub R_equipmentBOMIB02(control As IRibbonControl)
    Run ("equipmentBOMIB02")
End Sub

Sub R_equipmentBOMIB01(control As IRibbonControl)
    Run ("equipmentBOMIB01")
End Sub

Sub R_equipmentIA03(control As IRibbonControl)
    Run ("equipmentIA03")
End Sub

Sub R_equipmentIA02(control As IRibbonControl)
    Run ("equipmentIA02")
End Sub

Sub R_equipmentIA01(control As IRibbonControl)
    Run ("equipmentIA01")
End Sub

Sub R_Z_IA17_E_G_Single(control As IRibbonControl)
    Run ("Z_IA17_E_G_Single")
End Sub

Sub R_Z_IH08_E(control As IRibbonControl)
    Run ("Z_IH08_E")
End Sub

Sub R_Z_IE07_E(control As IRibbonControl)
    Run ("Z_IE07_E")
End Sub

Sub R_Z_IW39_E(control As IRibbonControl)
    Run ("Z_IW39_E")
End Sub

Sub R_Z_IW39_EC(control As IRibbonControl)
    Run ("Z_IW39_EC")
End Sub

Sub R_Z_IW38_E(control As IRibbonControl)
    Run ("Z_IW38_E")
End Sub

Sub R_Z_IW29_E(control As IRibbonControl)
    Run ("Z_IW29_E")
End Sub

Sub R_Z_IW29_EC(control As IRibbonControl)
    Run ("Z_IW29_EC")
End Sub

Sub R_Z_IW28_E(control As IRibbonControl)
    Run ("Z_IW28_E")
End Sub

Sub R_Z_IP18_E(control As IRibbonControl)
    Run ("Z_IP18_E")
End Sub

Sub R_Z_IP24_E(control As IRibbonControl)
    Run ("Z_IP24_E")
End Sub

Sub R_Z_IA17_E(control As IRibbonControl)
    Run ("Z_IA17_E")
End Sub

Sub R_Z_IA17_E_G(control As IRibbonControl)
    Run ("Z_IA17_E_G")
End Sub


'--- Sous-section : Menus "Floc" et "Floc List" ---

Sub R_Z_FLOC_IL03(control As IRibbonControl)
    Run ("Z_FLOC_IL03")
End Sub
Sub R_Z_FLOC_IL03_c(control As IRibbonControl)
    Run ("Z_FLOC_IL03_c")
End Sub
Sub R_Z_FLOC_IL03_B(control As IRibbonControl)
    Run ("Z_FLOC_IL03_B")
End Sub

Sub R_flocPMO(control As IRibbonControl)
    Run ("flocPMO")
End Sub
Sub R_flocPMO_H(control As IRibbonControl)
    Run ("flocPMO_H")
End Sub
Sub R_flocPMOc(control As IRibbonControl)
    Run ("flocPMOc")
End Sub

Sub R_flocNotif(control As IRibbonControl)
    Run ("flocNotif")
End Sub
Sub R_flocNotifH(control As IRibbonControl)
    Run ("flocNotifH")
End Sub
Sub R_flocNotifc(control As IRibbonControl)
    Run ("flocNotifc")
End Sub

Sub R_flocplan(control As IRibbonControl)
    Run ("flocplan")
End Sub
Sub R_flocplan24(control As IRibbonControl)
    Run ("flocplan24")
End Sub

Sub R_FLOCBOMIB03(control As IRibbonControl)
    Run ("FLOCBOMIB03")
End Sub
Sub R_floCBOMIB02(control As IRibbonControl)
    Run ("floCBOMIB02")
End Sub
Sub R_flocOMIB01(control As IRibbonControl)
    Run ("flocOMIB01")
End Sub

Sub R_flocIA13(control As IRibbonControl)
    Run ("flocIA13")
End Sub
Sub R_flocIA12(control As IRibbonControl)
    Run ("flocIA12")
End Sub
Sub R_flocIA11(control As IRibbonControl)
    Run ("flocIA11")
End Sub

Sub R_Z_IH06_F(control As IRibbonControl)
    Run ("Z_IH06_F")
End Sub
Sub R_Z_IH08_F(control As IRibbonControl)
    Run ("Z_IH08_F")
End Sub
Sub R_Z_IL07_F(control As IRibbonControl)
    Run ("Z_IL07_F")
End Sub
Sub R_Z_IW39_F(control As IRibbonControl)
    Run ("Z_IW39_F")
End Sub
Sub R_Z_IW39_FC(control As IRibbonControl)
    Run ("Z_IW39_FC")
End Sub
Sub R_Z_IW38_F(control As IRibbonControl)
    Run ("Z_IW38_F")
End Sub
Sub R_Z_IW29_F(control As IRibbonControl)
    Run ("Z_IW29_F")
End Sub
Sub R_Z_IW29_FC(control As IRibbonControl)
    Run ("Z_IW29_FC")
End Sub
Sub R_Z_IW28_F(control As IRibbonControl)
    Run ("Z_IW28_F")
End Sub
Sub R_Z_IP18_F(control As IRibbonControl)
    Run ("Z_IP18_F")
End Sub
Sub R_Z_IP24_F(control As IRibbonControl)
    Run ("Z_IP24_F")
End Sub
Sub R_Z_IA17_F(control As IRibbonControl)
    Run ("Z_IA17_F")
End Sub

Sub R_Z_IA17_F_G(control As IRibbonControl)
    Run ("Z_IA17_F_G")
End Sub

Sub R_IH01(control As IRibbonControl)
    Run ("Z_FLOC_IH01")
End Sub

'====================================================================================
' SECTION 8 : GROUPE "CAPEX" (ONGLET "CAPEX")
'====================================================================================

'--- Sous-section : Menus "Document" et "Doc. List" ---

Sub R_Z_CV03NDoc(control As IRibbonControl)
    Run ("Z_CV03NDoc")
End Sub

Sub R_Z_CV02NDoc(control As IRibbonControl)
    Run ("Z_CV02NDoc")
End Sub

Sub R_Z_CV04NDoc(control As IRibbonControl)
    Run ("Z_CV04NDoc")
End Sub

Sub R_Z_CV04NDocL(control As IRibbonControl)
    Run ("Z_CV04NDocL")
End Sub

'--- Sous-section : Menus "Project" et "Project List" ---

Sub R_Z_CN42NProjL(control As IRibbonControl)
    Run ("Z_CN42NProjL")
End Sub

Sub R_Z_CN43NProjL(control As IRibbonControl)
    Run ("Z_CN43NProjL")
End Sub

Sub R_Z_CNS41ProjL(control As IRibbonControl)
    Run ("Z_CNS41ProjL")
End Sub

Sub R_Z_CN46NProjL(control As IRibbonControl)
    Run ("Z_CN46NProjL")
End Sub

Sub R_Z_CN47ProjL(control As IRibbonControl)
    Run ("Z_CN47ProjL")
End Sub

Sub R_Z_CN47NProjL(control As IRibbonControl)
    Run ("Z_CN47NProjL")
End Sub

Sub R_Z_CN48NProjL(control As IRibbonControl)
    Run ("Z_CN48NProjL")
End Sub

'--- Sous-section : Menus "Network" et "Net. List" ---

Sub R_Z_CN46NNetL(control As IRibbonControl)
    Run ("Z_CN46NNetL")
End Sub

Sub R_Z_CNS41NetL(control As IRibbonControl)
    Run ("Z_CNS41NetL")
End Sub

Sub R_Z_ME5ANetL(control As IRibbonControl)
    Run ("Z_ME5ANetL")
End Sub

Sub R_Z_CN47NetL(control As IRibbonControl)
    Run ("Z_CN47NetL")
End Sub

Sub R_Z_CN47NNetL(control As IRibbonControl)
    Run ("Z_CN47NNetL")
End Sub

Sub R_Z_CN48NNetL(control As IRibbonControl)
    Run ("Z_CN48NNetL")
End Sub

'--- Sous-section : Menus "WBS" et "WBS List" ---

Sub R_Z_CN46N_WBSL(control As IRibbonControl)
    Run ("Z_CN46N_WBSL")
End Sub

Sub R_Z_ME5A_WBSL(control As IRibbonControl)
    Run ("Z_ME5A_WBSL")
End Sub

Sub R_Z_CN47_WBSL(control As IRibbonControl)
    Run ("Z_CN47_WBSL")
End Sub

Sub R_Z_CN47N_WBSL(control As IRibbonControl)
    Run ("Z_CN47N_WBSL")
End Sub

Sub R_Z_CN48N_WBSL(control As IRibbonControl)
    Run ("Z_CN48N_WBSL")
End Sub

Sub R_Z_CN43N_WBSL(control As IRibbonControl)
    Run ("Z_CN43N_WBSL")
End Sub

Sub R_Z_IW38_WBSL(control As IRibbonControl)
    Run ("Z_IW38_WBSL")
End Sub

'====================================================================================
' SECTION 9 : GROUPE "SETUP" (CONFIGURATION) (ONGLET "SETUP")
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : showsetup
' DESCRIPTION : Rend la feuille "Setup" visible et éditable en désactivant
'               temporairement le mode Add-in du classeur.
' CONTRÔLE    : Bouton "Show Setup" (btnShowSetup)
'------------------------------------------------------------------------------------
Sub showsetup(control As IRibbonControl)
    MsgBox ("As soon the changes are applied, please use the Hide Setup button to save the changes and activate the Add-In. If Excel closes without doing that step, the ExcelToSAP ribbon won't show up, unless the add-in is activated manually.")
    If Right(LCase(ThisWorkbook.name), 5) = ".xlam" Then
        ThisWorkbook.IsAddin = False
    End If
    ThisWorkbook.Sheets("Setup").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Setup").Activate
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : hidesetups
' DESCRIPTION : Masque le classeur en réactivant le mode Add-in et sauvegarde
'               les modifications apportées à la configuration.
' CONTRÔLE    : Bouton "Hide Setup" (btnHideSetup)
'------------------------------------------------------------------------------------
Sub hidesetups(control As IRibbonControl)
    If Right(LCase(ThisWorkbook.name), 5) = ".xlam" Then
        ThisWorkbook.IsAddin = True
    End If
    ThisWorkbook.Save
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_Install_AddIn
' DESCRIPTION : Lance la procédure d'installation de l'add-in.
' CONTRÔLE    : Bouton "Install AddIn" (btnInstallAddIn)
'------------------------------------------------------------------------------------
Sub R_Install_AddIn(control As IRibbonControl)
    Call Install_AddIn
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : techsup
' DESCRIPTION : Affiche le formulaire d'aide ou redirige vers le site de support.
' CONTRÔLE    : Bouton "Technical Support" (techsupbtn)
'------------------------------------------------------------------------------------
Sub techsup(control As IRibbonControl)
    ' Affiche la feuille d'aide générée par code.
    ShowHelpPage
End Sub

Sub HideHelpPage()
    ' Affiche la feuille d'aide générée par code.
        If Right(LCase(ThisWorkbook.name), 5) = ".xlam" Then
        ThisWorkbook.IsAddin = True
    End If
    ThisWorkbook.Save
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Aboutp
' DESCRIPTION : Affiche la fenêtre "À propos" de l'application.
' CONTRÔLE    : Bouton "About" (Aboutpbtn)
'------------------------------------------------------------------------------------
Sub Aboutp(control As IRibbonControl)
    Dim msg As String
    msg = "ExcelToSAP Add-in" & vbCrLf & _
          "--------------------------------------------------" & vbCrLf & _
          "Original Project:" & vbCrLf & _
          "Created by Jordi N (admin) - SourceForge" & vbCrLf & _
          " " & vbCrLf & _
          "Description:" & vbCrLf & _
          "This Add-in enables to use a collection of macros for SAP on any existing or new Excel worksheet." & vbCrLf & _
          " " & vbCrLf & _
          "The macros automate the following steps:" & vbCrLf & _
          "- Opens the transaction." & vbCrLf & _
          "- Enters the selected data from Excel." & vbCrLf & _
          "- Optional: enters the plant/layout and executes." & vbCrLf & _
          " " & vbCrLf & _
          "--------------------------------------------------" & vbCrLf & _
          "Adapted for Holcim Algeria by:" & vbCrLf & _
          "Ahmed CHAMI"
          
    MsgBox msg, vbInformation, "About ExcelToSAP"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ShowHelpPage (Privée)
' DESCRIPTION : Génère et affiche une feuille d'aide dynamique pour l'utilisateur.
'------------------------------------------------------------------------------------
Private Sub ShowHelpPage()
    On Error GoTo ErrHandler
    
    Dim wsHelp As Worksheet
    Dim row As Long
    
    ' --- Préparation de la feuille ---
    Application.ScreenUpdating = False
    
    ' Si le classeur est un Add-in (masqué), on le rend visible pour pouvoir afficher la feuille
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
    End If
    
    On Error Resume Next
    Set wsHelp = ThisWorkbook.Sheets("Help")
    If wsHelp Is Nothing Then
        ' Si la feuille n'existe pas, on la crée
        Set wsHelp = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsHelp.name = "Help"
    End If
    On Error GoTo ErrHandler
    
    wsHelp.Visible = xlSheetVisible
    wsHelp.Activate
    
    ' --- Nettoyage et formatage de base ---
    With wsHelp.Cells
        .ClearContents
        .ClearFormats
        .Font.name = "Calibri"
        .Font.Size = 11
        .VerticalAlignment = xlTop
    End With
    
    wsHelp.Columns("A").ColumnWidth = 3
    wsHelp.Columns("B").ColumnWidth = 90
    wsHelp.Columns("C").ColumnWidth = 30
    
    row = 2 ' Ligne de départ
    
    ' --- Écriture du contenu ---
    
    ' Titre principal
    With wsHelp.Range("B" & row)
        .value = "Guide d'Utilisation - ExcelToSAP pour Holcim Algérie"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(0, 80, 120) ' Bleu foncé
    End With
    row = row + 2
    
    ' Section 1
    WriteTitle wsHelp, row, "1. À quoi sert cet outil ?"
    WriteBullet wsHelp, row, "Permet de lancer des transactions SAP directement depuis n'importe quelle feuille Excel."
    WriteBullet wsHelp, row, "Évite d'avoir à ouvrir SAP manuellement, copier/coller des données, et naviguer dans les menus."
    WriteBullet wsHelp, row, "Fonctionne avec la plupart des données SAP courantes : Ordres, Avis, Équipements, Articles, etc."
    row = row + 2
    
    ' Section 2
    WriteTitle wsHelp, row, "2. Comment l'utiliser ? (La méthode rapide)"
    With wsHelp.Range("B" & row)
        .value = "Le Clic-Droit est votre meilleur ami !"
        .Font.Italic = True
    End With
    row = row + 1
    WriteBullet wsHelp, row, "Étape 1 : Sélectionnez une cellule (ou une colonne de cellules) contenant un numéro SAP (ex: un numéro d'ordre)."
    WriteBullet wsHelp, row, "Étape 2 : Faites un clic-droit sur votre sélection."
    WriteBullet wsHelp, row, "Étape 3 : Un menu ""ExcelToSAP"" apparaît avec les commandes pertinentes."
    row = row + 1
    WriteBullet wsHelp, row, "Scénario 1 (Détection auto) : Si l'outil reconnaît le type de donnée (ex: un Ordre), le menu affichera directement les commandes pour cet ordre."
    WriteBullet wsHelp, row, "Scénario 2 (Aucune détection) : Si l'outil ne reconnaît pas la donnée, le menu affichera des sous-menus pour chaque catégorie (Plant Maintenance, Master Data...). Vous pouvez alors naviguer pour trouver la commande."
    row = row + 1
    WriteBullet wsHelp, row, "Note : Les commandes disponibles sont différentes si vous sélectionnez une seule cellule (actions individuelles) ou une liste (actions en masse)."
    row = row + 2
    
    ' Section 5
    WriteTitle wsHelp, row, "3. Support Technique"
    WriteBullet wsHelp, row, "En cas de problème, d'erreur ou de question, veuillez contacter :"
    row = row + 1
    WriteInfo wsHelp, row, "Nom :       Ahmed CHAMI", ""
    WriteInfo wsHelp, row, "Plant :     CILAS Plant Algeria", ""
    WriteInfo wsHelp, row, "Email :     Ahmed.chami@lafarge.com", ""
    WriteInfo wsHelp, row, "Téléphone : (+213) 7078 34 95", ""
    row = row + 1
    WriteBullet wsHelp, row, "Pensez à préparer une capture d'écran du message d'erreur s'il y en a un."
    
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Une erreur est survenue lors de la génération de la page d'aide : " & Err.Description, vbCritical, "Erreur d'aide"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_LoadConfiguration
' DESCRIPTION : Relance la procédure de chargement de la configuration depuis la
'               feuille "Setup" et affiche un message de confirmation.
' CONTRÔLE    : Bouton "Reload Config" (btnReloadConfig)
'------------------------------------------------------------------------------------
Sub R_LoadConfiguration(control As IRibbonControl)
    Call LoadConfiguration
    MsgBox "La configuration a été rechargée avec succès depuis la feuille 'Setup'.", vbInformation, "Configuration Rechargée"
End Sub

'====================================================================================
' SECTION 10 : ONLGET "WEEK KPIS" (INDICATEURS DE PERFORMANCE) (ONGLET "WEEK KPIS")
'====================================================================================

'--- Sous-section : Groupe "Operations" (kpiOpr) ---

Sub R_Z_KPIs_List(control As IRibbonControl)
    '-------------------------------------------------------------------------------
    ' ACTION : Bouton "Refresh KPI Data"
    ' APPELLE: Z_KPIs_List du module KPI_MP pour rafraîchir les données principales.
    '-------------------------------------------------------------------------------
    Run ("Z_KPIs_List")
End Sub
Sub KPIs_RefresheTCD_SCH(control As IRibbonControl)
    Run ("Z_Confirmation_List")
End Sub

'#############################################################3
Sub KPI_Accuracy(control As IRibbonControl)
    Run ("Z_KPI_Schudled")
End Sub

Sub KPI_Compliance(control As IRibbonControl)
    Run ("Z_KPI_No_Cmopliance")
End Sub

Sub KPI_SCH_Ratio(control As IRibbonControl)
    Run ("Z_KPI_SchRatio")
End Sub

Sub KPI_Unplanned(control As IRibbonControl)
    Run ("Z_KPIs_Unplanned")
End Sub

Sub KPI_Overdue(control As IRibbonControl)
    Run ("Z_KPI_Overdue")
End Sub

'--- Sous-section : Groupe "PMRs" (kpiPMRs) ---
'#############################################################3
Sub KPI_PMR_Percentage(control As IRibbonControl)
    Run ("Z_KPIs_PMR_Confirmed")
End Sub

Sub KPI_PMR_NotPerformed(control As IRibbonControl)
    Run ("Z_KPIs_PMR_not_Performed")
End Sub

Sub KPI_PMR_ManualCall(control As IRibbonControl)
    Run ("Z_KPIs_PMR_ManualCall")
End Sub

Sub KPI_PMR_Efficiency(control As IRibbonControl)
    Run ("Z_KPI_MR_Effeciency")
End Sub

'--- Sous-section : Groupe "MR" (kpiMR) ---

Sub KPI_MR_no_Mobile(control As IRibbonControl)
    Run ("Z_KPI_MR_No_Mobile")
End Sub

Sub KPI_CPM_NMobile(control As IRibbonControl)
    Run ("Z_KPI_MR_CPM_no_Mobile")
End Sub

'--- Sous-section : Groupe "Planning" (kpiPlanning) ---

Sub KPI_Wo_MR(control As IRibbonControl)
    Run ("Z_KPI_Wo_MR")
End Sub

Sub KPI_Aging_Wo(control As IRibbonControl)
    Run ("Z_KPI_AgingWo")
End Sub

Sub KPI_Aging_MR(control As IRibbonControl)
    Run ("Z_KPI_AgingMR")
End Sub

'--- Sous-section : Groupe "Gantt Week" ---
Sub KPI_Gantt_LastWeek(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_LastWeek")
End Sub
Sub KPI_Gantt_ThisWeek(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled")
End Sub
Sub KPI_Gantt_NextWeek(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_NextWeek")
End Sub

'--- Sous-section : Groupe "Gantt Month" ---
Sub KPI_Gantt_LastMonth(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_LastMonth")
End Sub
Sub KPI_Gantt_ThisMonth(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_ThisMonth")
End Sub
Sub KPI_Gantt_NextMonth(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_NextMonth")
End Sub

'--- Sous-section : Groupe "Gantt Year" ---
Sub KPI_Gantt_ThisYear(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_ThisYear")
End Sub
Sub KPI_Gantt_NextYear(control As IRibbonControl)
    Run ("Z_KPI_Gantt_Scheduled_NextYear")
End Sub

'====================================================================================
' SECTION 11 : GROUPE "TESTS"
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_Populate_Test_Sheet
' DESCRIPTION : Remplit la feuille de tests avec les données exportées.
' CONTRÔLE    : Bouton "Populate Test Data" (btnPopulateTests)
'------------------------------------------------------------------------------------
Sub R_Populate_Test_Sheet(control As IRibbonControl)
    Run ("Populate_Test_Sheet_From_Exports")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_ToggleTestSheet
' DESCRIPTION : Affiche ou masque la feuille de tests.
' CONTRÔLE    : Bouton "Show/Hide Test Sheet" (btnToggleTestSheet)
'------------------------------------------------------------------------------------
Sub R_ToggleTestSheet(control As IRibbonControl)
    Run ("Toggle_Test_Sheet")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_RunAllTests
' DESCRIPTION : Lance la suite de tests complète du projet.
' CONTRÔLE    : Bouton "Run All Tests" (btnRunAllTests)
'------------------------------------------------------------------------------------
Sub R_RunAllTests(control As IRibbonControl)
    Run ("Run_All_Project_Tests")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : GetTestMenuContent
' DESCRIPTION : Génère dynamiquement le contenu XML pour le menu de test.
'               Ceci permet d'avoir un menu avec des boutons et une action sur le
'               bouton principal du menu lui-même.
' CONTRÔLE    : dynamicMenu "menuRunTests" (attribut getContent)
'------------------------------------------------------------------------------------
Sub GetTestMenuContent(control As IRibbonControl, ByRef content)
    Dim xml As String
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
            "<button id=""btnRunAllTestsMenu"" label=""Run All Project Tests"" onAction=""R_RunAllTests"" imageMso=""MacroPlay"" />" & _
            "<menuSeparator id=""sepTests1""/>" & _
            "<menu id=""menuTestWO"" label=""Test Work Orders"" imageMso=""Paste"">" & _
                "<button id=""btnRunAllWOTests"" label=""Run All Work Order Tests"" onAction=""R_RunTestWorkOrder"" />" & _
                "<menuSeparator id=""sepWO1""/>" & _
                "<button id=""btnTestWO_1"" label=""Test IW33 Main"" onAction=""R_Test_WO_Z_PMO_IW33_1_Main"" />" & _
            "</menu>" & _
            "<button id=""btnTestNotif"" label=""Test Notifications"" onAction=""R_RunTestNotification"" imageMso=""NewTaskRequest2"" />" & _
            "<button id=""btnTestEquip"" label=""Test Equipments"" onAction=""R_RunTestEquipment"" imageMso=""Camera"" />" & _
            "<button id=""btnTestFloc"" label=""Test Func. Locations"" onAction=""R_RunTestFuncLoc"" imageMso=""Pushpin"" />" & _
            "<button id=""btnTestMPlan"" label=""Test Maint. Plans"" onAction=""R_RunTestMPlan"" imageMso=""CalendarInsert"" />" & _
            "<button id=""btnTestMat"" label=""Test Materials"" onAction=""R_RunTestMaterials"" imageMso=""ShapeSeal24"" />" & _
            "<button id=""btnTestPO"" label=""Test Purchase Orders"" onAction=""R_RunTestPurchOrder"" imageMso=""AccountingFormat"" />" & _
            "<button id=""btnTestPR"" label=""Test Purchase Reqs"" onAction=""R_RunTestRequisition"" imageMso=""SharingOpenWssTaskList"" />" & _
            "<button id=""btnTestDoc"" label=""Test Documents"" onAction=""R_RunTestDocuments"" imageMso=""MailMergeResultsPreview"" />" & _
            "<button id=""btnTestProj"" label=""Test Projects"" onAction=""R_RunTestProject"" imageMso=""MicrosoftProject"" />" & _
            "<button id=""btnTestNet"" label=""Test Networks"" onAction=""R_RunTestNetwork"" imageMso=""ShapeLightningBolt"" />" & _
          "</menu>"
    content = xml
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURES   : R_RunTest...
' DESCRIPTION : Lance les suites de tests individuelles depuis le menu de test.
' CONTRÔLE    : Boutons dans le menu "menuRunTests".
'------------------------------------------------------------------------------------

Sub R_RunTestWorkOrder(control As IRibbonControl)
    Run "Run_Single_Test_WorkOrder"
End Sub

Sub R_RunTestNotification(control As IRibbonControl)
    Run "Run_Single_Test_Notification"
End Sub

Sub R_RunTestEquipment(control As IRibbonControl)
    Run "Run_Single_Test_Equipment"
End Sub

Sub R_RunTestFuncLoc(control As IRibbonControl)
    Run "Run_Single_Test_FuncLoc"
End Sub

Sub R_RunTestMPlan(control As IRibbonControl)
    Run "Run_Single_Test_M_Plan"
End Sub

Sub R_RunTestMaterials(control As IRibbonControl)
    Run "Run_Single_Test_Materials"
End Sub

Sub R_RunTestPurchOrder(control As IRibbonControl)
    Run "Run_Single_Test_PurchOrder"
End Sub

Sub R_RunTestRequisition(control As IRibbonControl)
    Run "Run_Single_Test_Requisition"
End Sub

Sub R_RunTestDocuments(control As IRibbonControl)
    Run "Run_Single_Test_Documents"
End Sub

Sub R_RunTestProject(control As IRibbonControl)
    Run "Run_Single_Test_Project"
End Sub

Sub R_RunTestNetwork(control As IRibbonControl)
    Run "Run_Single_Test_Network"
End Sub

Sub R_RunTestWBS(control As IRibbonControl)
    Run "Run_Single_Test_WBS"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : R_OpenLogFile
' DESCRIPTION : Ouvre le fichier de log des tests dans le Bloc-notes.
' CONTRÔLE    : Bouton "Open Log File" (btnOpenLogFile)
'------------------------------------------------------------------------------------
Sub R_OpenLogFile(control As IRibbonControl)
    Dim logFilePath As String

    ' S'assure que la configuration est chargée pour obtenir le chemin
    If ConfigSettings Is Nothing Or ConfigSettings.count = 0 Then
        LoadConfiguration
    End If

    logFilePath = GetSetting("PATH_LOGS") & "ExcelToSAP_Log.txt"

    If Dir(logFilePath) <> "" Then
        Shell "notepad.exe " & logFilePath, vbNormalFocus
    Else
        MsgBox "Le fichier de log n'a pas encore été créé ou est introuvable." & vbCrLf & vbCrLf & _
               "Chemin attendu : " & logFilePath, vbInformation, "Fichier de Log Introuvable"
    End If
End Sub

'====================================================================================
' SECTION 14 : PROCÉDURES UTILITAIRES POUR LA PAGE D'AIDE
'====================================================================================

Private Sub WriteTitle(ws As Worksheet, ByRef row As Long, text As String)
    With ws.Range("B" & row)
        .value = text
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(47, 117, 181) ' Bleu
    End With
    row = row + 1
End Sub

Private Sub WriteBullet(ws As Worksheet, ByRef row As Long, text As String)
    With ws.Range("B" & row)
        .value = "•  " & text
        .WrapText = True
    End With
    row = row + 1
End Sub

Private Sub WriteInfo(ws As Worksheet, ByRef row As Long, label As String, value As String)
    With ws.Range("B" & row)
        .value = "    " & label
        .Font.Bold = True
    End With
    ws.Range("C" & row).value = value
    row = row + 1
End Sub
'====================================================================================
' SECTION 12 : ONGLET "REPORTING"
'====================================================================================

Sub R_Load_data_Dashboard(control As IRibbonControl)
    Run ("Load_data_Dashboard")
End Sub

Private Sub ClearFileAndRun(ByVal macroName As String, ByVal fileName As String)
    Dim dashboardPath As String
    dashboardPath = GetSetting("DASHBOARD_PATH")
    If dashboardPath <> "" And fileName <> "" Then
        ClearFileIfExists dashboardPath, fileName
    End If
    Run macroName
End Sub

Sub R_Z_KPIs_List_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPIs_List_Dash", "01_Schudled_Op_KPI.XLSX"
End Sub

Sub R_Z_Schudled_Confirmation_File_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_Schudled_Confirmation_File_Dash", "02_Schudled_Confirmation.XLSX"
End Sub

Sub R_Z_Confirmed_Confirmation_File_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_Confirmed_Confirmation_File_Dash", "03_Confirmed_Confirmation.XLSX"
End Sub

Sub R_Z_unplanned_Confirmation_File_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_unplanned_Confirmation_File_Dash", "06_Unplanned_Confirmation.XLSX"
End Sub

Sub R_Z_KPI_Schudled_Op_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_Schudled_Op_Dash", "04_Schudled_Operations.XLSX"
End Sub

Sub R_Z_KPI_Schudled_non_Confirmed_Op_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_Schudled_non_Confirmed_Op_Dash", "05_Schudled_no_Confirmed.XLSX"
End Sub

Sub R_Z_KPIs_Unplanned_Op_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPIs_Unplanned_Op_Dash", "07_Unplanned.XLSX"
End Sub

Sub R_Z_KPI_Accuracy_Dash(control As IRibbonControl)
    Run ("Z_KPI_Accuracy_Dash")
End Sub

Sub R_Z_KPI_SchRatio_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_SchRatio_Dash", "17_SchRatio.XLSX"
End Sub

Sub R_Z_KPI_Overdue_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_Overdue_Dash", "08_Overdue.XLSX"
End Sub

Sub R_Z_KPIs_PMR_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPIs_PMR_Dash", "09_PMR.XLSX"
End Sub

Sub R_Z_KPIs_PMR_not_Performed_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPIs_PMR_not_Performed_Dash", "10_PMR_not_Performed.XLSX"
End Sub

Sub R_Z_KPIs_PMR_ManualCall_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPIs_PMR_ManualCall_Dash", "11_PMR_Manual_Call.XLSX"
End Sub

Sub R_Z_KPI_MR_Created_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_MR_Created_Dash", "12_Notification_Created.XLSX"
End Sub

Sub R_Z_KPI_MR_Created_CPM_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_MR_Created_CPM_Dash", "16_Notification_Created_CPM.XLSX"
End Sub

Sub R_Z_KPI_Wo_MR_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_Wo_MR_Dash", "13_PM01_Wo_MR.XLSX"
End Sub

Sub R_Z_KPI_AgingWo_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_AgingWo_Dash", "14_Aging_Wo.XLSX"
End Sub

Sub R_Z_KPI_AgingMR_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_AgingMR_Dash", "15_Aging_MR.XLSX"
End Sub

Sub R_Z_Load_Open_PMwo_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_Load_Open_PMwo_Dash", "Open_PMOrder.XLSX"
End Sub

Sub R_Z_KPI_Open_MR_Dash(control As IRibbonControl)
    ClearFileAndRun "Z_KPI_Open_MR_Dash", "Open_Notifications.XLSX"
End Sub

'====================================================================================
' SECTION 13 : ONGLET "REPORTING" - GROUPE "FULL DATA EXPORTS"
'====================================================================================

Sub R_Export_AllReports(control As IRibbonControl)
    Run ("Export_AllReports")
End Sub

Private Sub ClearExportFileAndRun(ByVal macroName As String, ByVal fileName As String)
    Dim exportPath As String
    exportPath = GetSetting("PATH_EXPORT")
    If exportPath <> "" And fileName <> "" Then
        ClearFileIfExists exportPath, fileName
    End If
    Run macroName
End Sub

Sub R_Z_Load_MR(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_MR", "MR.xlsx"
End Sub

Sub R_Z_Load_WO(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_WO", "WO.xlsx"
End Sub

Sub R_Z_Load_Op(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_Op", "OP.xlsx"
End Sub

Sub R_Z_Load_GM(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_GM", "GM.xlsx"
End Sub

Sub R_Z_Load_KPI(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_KPI", "KPI.xlsx"
End Sub

Sub R_Z_Load_PR(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_PR", "PR.xlsx"
End Sub

Sub R_Z_Load_PO(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_PO", "PO.xlsx"
End Sub

Sub R_Z_Load_SES(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_SES", "SES.xlsx"
End Sub

Sub R_Z_Load_RSV(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_RSV", "RSV.XLSX"
End Sub

Sub R_Z_Load_CNF(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_CNF", "CNF.XLSX"
End Sub

Sub R_Z_Load_PMR(control As IRibbonControl)
    ClearExportFileAndRun "Z_Load_PMR", "PMR.xlsx"
End Sub

