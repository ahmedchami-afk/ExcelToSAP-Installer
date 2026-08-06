Attribute VB_Name = "M_Plan"
'==============================================================
' MODULE      : M_Plan
' VERSION     : 1.0
' AUTEUR      : CHAMI Ahmed / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module centralise les macros pour la gestion des plans de maintenance
'               (Maintenance Plans), des PMRs et des gammes (Task Lists).
'
' FONCTIONS PRINCIPALES :
'   - Consultation et affichage de plans (IP03, IP02).
'   - Historique et appels de plan (IW39, IP18, IP24).
'   - Modification ou ouverture de gammes associées (IAxx).
'   - Opérations en masse sur les plans et PMRs.
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session : Objet de session SAP GUI active.
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'   - Z_... : Macros de transactions SAP prédéfinies.
'==============================================================
Option Explicit

'====================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteSinglePlanAction(ByVal transactionWrapper As String, ByVal procName As String, Optional ByVal actionId As String = "", Optional ByVal actionType As String = "ButtonPress", Optional ByVal postActionKey As Integer = -1)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    g_Session.findById("wnd[0]/usr/ctxtRMIPM-WARPL").text = Cells(ActiveCell.row, ActiveCell.Column).value
    g_Session.findById("wnd[0]/usr/ctxtRMIPM-WARPL").caretPosition = 12
    g_Session.findById("wnd[0]").sendVKey 0
    
    If actionId <> "" Then
        If actionType = "ButtonPress" Then
            g_Session.findById(actionId).press
        ElseIf actionType = "TabSelect" Then
            g_Session.findById(actionId).Select
        End If
    End If
    
    If postActionKey <> -1 Then
        g_Session.findById("wnd[0]").sendVKey postActionKey
    End If
    
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSinglePlanReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldId As String, Optional ByVal checkboxesToDeselect As Variant, Optional ByVal checkboxesToSelect As Variant)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    If Not IsMissing(checkboxesToDeselect) Then
        Dim chk As Variant
        For Each chk In checkboxesToDeselect
            g_Session.findById(chk).Selected = False
        Next chk
    End If
    
    If Not IsMissing(checkboxesToSelect) Then
        Dim chk2 As Variant
        For Each chk2 In checkboxesToSelect
            g_Session.findById(chk2).Selected = True
        Next chk2
    End If
    
    g_Session.findById(fieldId).text = Cells(ActiveCell.row, ActiveCell.Column).value
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub PrepareListPlanTransaction(ByVal transactionWrapper As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False)
    Run ("onSAP")
    Run (transactionWrapper)
    
    g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
    If clearDates Then
        g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    End If
    
    Run ("Z_F8")
End Sub

Private Sub ExecuteListPlanReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False)
    On Error GoTo SapErrorHandler
    PrepareListPlanTransaction transactionWrapper, fieldToClear, buttonToFill, clearDates
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'====================================================================================
' SECTION 1 : CONSULTATION ET MODIFICATION DE PLANS UNIQUES
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : displayPLan
' DESCRIPTION : Ouvre un plan de maintenance (IP03) et affiche le texte long du plan.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub displayPLan()
    ExecuteSinglePlanAction "Z_IP03", "displayPLan", "wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/btnLANGTEXT_APLAN"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ChangePLan
' DESCRIPTION : Ouvre un plan de maintenance en mode modification (IP02) et affiche
'               le texte long du plan.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP02, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub ChangePLan()
    ExecuteSinglePlanAction "Z_IP02", "ChangePLan", "wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/btnLANGTEXT_APLAN"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : displayPLanOP
' DESCRIPTION : Affiche les opérations (gamme) du plan de maintenance.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub displayPLanOP()
    ExecuteSinglePlanAction "Z_IP03", "displayPLanOP", "wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/btnARBEITSPLAN"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : IP24plan
' DESCRIPTION : Exécute IP24 (Planification de maintenance) pour un plan spécifique.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP24_NoOL, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub IP24plan()
    ExecuteSinglePlanReport "Z_IP24_NoOL", "IP24plan", "wnd[0]/usr/ctxtWARPL-LOW"
End Sub


'====================================================================================
' SECTION 2 : TRANSACTIONS EN MASSE (LISTES DE PLANS ET POSTES)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : ip18_L
' DESCRIPTION : Exécute IP18 (Liste des plans) et injecte la liste des plans depuis Excel.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IP18_All, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub ip18_L()
    ExecuteListPlanReport "Z_IP18_All", "ip18_L", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IP06I
' DESCRIPTION : Exécute IP06 (Afficher poste de maintenance) et appuie sur "Entrée".
' CONTEXTE    : Poste unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP06, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IP06I()
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run ("Z_IP06")

    Run ("Z_Enter")
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_IP06I", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IP18LI
' DESCRIPTION : Exécute IP18 sur les postes de plan (WAPOS) sélectionnés depuis Excel.
' CONTEXTE    : Liste de postes de plan sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IP18_All, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IP18LI()
    ExecuteListPlanReport "Z_IP18_All", "Z_IP18LI", "wnd[0]/usr/ctxtWAPOS-LOW", "wnd[0]/usr/btn%_WAPOS_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : IP24_LI
' DESCRIPTION : Exécute IP24 sur les postes de plan (WAPOS) sélectionnés depuis Excel.
' CONTEXTE    : Liste de postes de plan sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IP24_NoOL, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub IP24_LI()
    ExecuteListPlanReport "Z_IP24_NoOL", "IP24_LI", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WAPOS_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : ip16_L
' DESCRIPTION : Exécute IP16 (Modification en masse des plans) pour une liste de plans.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IP16, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub ip16_L()
    ExecuteListPlanReport "Z_IP16", "ip16_L", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39OL
' DESCRIPTION : Exécute IW39 pour afficher les ordres ouverts liés aux plans de
'               maintenance sélectionnés.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39OL()
    ExecuteListPlanReport "Z_IW39", "Z_IW39OL", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : pmohistplanL
' DESCRIPTION : Exécute IW39H (historique complet) pour une liste de plans.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW39H, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub pmohistplanL()
    ExecuteListPlanReport "Z_IW39H", "pmohistplanL", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : pmohistplanLI
' DESCRIPTION : Version de `pmohistplanL` filtrée par postes de plan (WAPOS).
' CONTEXTE    : Liste de postes de plan sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW39H, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub pmohistplanLI()
    ExecuteListPlanReport "Z_IW39H", "pmohistplanLI", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WAPOS_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW38OL
' DESCRIPTION : Exécute IW38 pour afficher les ordres en cours liés à plusieurs plans.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW38OL()
    ExecuteListPlanReport "Z_IW38", "Z_IW38OL", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH", True
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39OLI
' DESCRIPTION : Exécute IW39 sur les postes de plan (WAPOS) sélectionnés.
' CONTEXTE    : Liste de postes de plan sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39OLI()
    ExecuteListPlanReport "Z_IW39", "Z_IW39OLI", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WAPOS_%_APP_%-VALU_PUSH", True
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW38OLI
' DESCRIPTION : Exécute IW38 sur les postes de plan (WAPOS) sélectionnés.
' CONTEXTE    : Liste de postes de plan sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW38OLI()
    ExecuteListPlanReport "Z_IW38", "Z_IW38OLI", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WAPOS_%_APP_%-VALU_PUSH", True
End Sub


'====================================================================================
' SECTION 3 : ANALYSE DE PLANS UNIQUES (APPELS, HISTORIQUE, GAMMES)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : displayPLanCalls
' DESCRIPTION : Ouvre la section des appels de plan (Calls) dans la transaction IP03.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub displayPLanCalls()
    ExecuteSinglePlanAction "Z_IP03", "displayPLanCalls", "wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\04", "TabSelect"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : callH
' DESCRIPTION : Ouvre l'onglet de l'historique des appels de plan (Calls History) dans IP03.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub callH()
    ExecuteSinglePlanAction "Z_IP03", "callH", "wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\02", "TabSelect"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : lastcall
' DESCRIPTION : Affiche le dernier appel ("Last Call") du plan dans IP03.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub lastcall()
    ExecuteSinglePlanAction "Z_IP03", "lastcall", "wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_MAINT_ITEM_TEXT:SAPLIWP3:6005/btnLAST_CALL"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : enh
' DESCRIPTION : Ouvre l'onglet "Enhancements" du plan dans IP03.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP03, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub enh()
    ExecuteSinglePlanAction "Z_IP03", "enh"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : changeTLgivenPlan
' DESCRIPTION : Ouvre la gamme (Task List) associée au plan de maintenance en mode
'               modification (IP02).
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IP02, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub changeTLgivenPlan()
    ExecuteSinglePlanAction "Z_IP02", "changeTLgivenPlan", "wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/btn%#AUTOTEXT002", "ButtonPress", 0
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : pmohistplan
' DESCRIPTION : Affiche l'historique des ordres associés à un plan de maintenance via IW39.
' CONTEXTE    : Plan unique dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IW39H, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub pmohistplan()
    ExecuteSinglePlanReport "Z_IW39H", "pmohistplan", "wnd[0]/usr/ctxtWARPL-LOW", Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR"), Array("wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ip24_L
' DESCRIPTION : Exécute IP24 sur la liste des plans (WARPL) copiée depuis Excel.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_IP24_NoOL, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub ip24_L()
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run ("Z_IP24_NoOL")

    ' Vide le champ et remplit la liste de sélection
    g_Session.findById("wnd[0]/usr/ctxtWARPL-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH", Selection

    Run ("Z_F8")
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "ip24_L", Err
End Sub

'====================================================================================
' SECTION 4 : MODIFICATIONS EN MASSE DES PLANS
'====================================================================================

Private Sub ExecuteMassChangeStrategy(ByVal procName As String, ByVal colIndex As Integer)
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    PrepareListPlanTransaction "Z_IP16", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
    
    g_Session.findById("wnd[0]/tbar[1]/btn[17]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[2]/menu[5]").Select
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").GetAbsoluteRow(6).Selected = True
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    g_Session.findById("wnd[2]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[1]").sendVKey 4
    Run ("Z_Enter")
    
    g_Session.findById("wnd[2]/usr/lbl[1," & colIndex & "]").SetFocus
    g_Session.findById("wnd[2]").sendVKey 2
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    g_Session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    RestoreExcel
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : MP_Auto_1W
' DESCRIPTION : Configure les plans de maintenance sélectionnés pour une planification
'               automatique hebdomadaire (1W - 1 Week).
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, IP16_L, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub MP_Auto_1W()
    ExecuteMassChangeStrategy "MP_Auto_1W", 3
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : MP_Auto_1M
' DESCRIPTION : Configure la planification automatique mensuelle (1M - 1 Month)
'               pour les plans de maintenance sélectionnés.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, IP16_L, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub MP_Auto_1M()
    ExecuteMassChangeStrategy "MP_Auto_1M", 4
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : MP_Auto_13W
' DESCRIPTION : Configure une planification automatique sur 13 semaines.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, IP16_L, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub MP_Auto_13W()
    ExecuteMassChangeStrategy "MP_Auto_13W", 5
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : MP_Auto_1Y
' DESCRIPTION : Configure la planification automatique annuelle (1Y - 1 Year).
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, IP16_L, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub MP_Auto_1Y()
    ExecuteMassChangeStrategy "MP_Auto_1Y", 6
End Sub

'====================================================================================
' SECTION 5 : OPÉRATIONS EN MASSE SUR LES PMR
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Afficher_Arbo_PMR
' DESCRIPTION : Affiche les PMRs sous forme arborescente dans SAP via ZPM004.
' CONTEXTE    : Liste de PMRs sélectionnés dans Excel.
' DÉPENDANCES : DateSemain, onSAP, Z_zpm004, Z_F8, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub Afficher_Arbo_PMR()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    Run ("DateSemain")
    Run ("onSAP")
    Run ("Z_zpm004")
    
    ' Options de sélection
    g_Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[0]/usr/ctxtP_SDATE").text = g_DebutAnnee

    ' Collage de la liste des PMRs
    FillSAPSelectionList "wnd[0]/usr/btn%_S_WARPL_%_APP_%-VALU_PUSH", Selection
    
    g_Session.findById("wnd[0]/usr/radP_TREE").Select ' Sélection arborescence
    Run ("Z_F8")
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Développe tous les nœuds
    g_Session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").SelectNode "          1"
    g_Session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
    
    ' Charge le layout
    g_Session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").PressContextButton "&LOAD"
    g_Session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem "&LOAD"
    g_Session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").currentCellRow = 21
    g_Session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").FirstVisibleRow = 2
    g_Session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "21"
    g_Session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell
    
    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Afficher_Arbo_PMR", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Manual_Call_PMRs
' DESCRIPTION : Lance manuellement les PMRs sélectionnés dans SAP.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, ip16_L, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub Manual_Call_PMRs()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    PrepareListPlanTransaction "Z_IP16", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"
    
    ' La session SAP est déjà active
    g_Session.findById("wnd[0]/tbar[1]/btn[17]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[0]").Select
    
    'bouton manual call
    g_Session.findById("wnd[0]/tbar[1]/btn[18]").press
    Run ("Z_Enter")
    
    ' OK pour la date
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press

    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Manual_Call_PMRs", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Change_StartCycle_PMRs
' DESCRIPTION : Modifie la date de début de cycle pour chaque PMR sélectionné.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, ip16_L, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub Change_StartCycle_PMRs()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    PrepareListPlanTransaction "Z_IP16", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"

    g_Session.findById("wnd[0]/tbar[1]/btn[17]").press ' Sélectionne tout
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[2]/menu[0]").Select
    
    Dim NbLign As Long, i As Long
    NbLign = Selection.Rows.count
    
    For i = 1 To NbLign
        ' Navigue vers l'onglet de planification et met à jour la date de début de cycle
        g_Session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\03").Select
        On Error Resume Next
        g_Session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY1:SAPLIWP3:8012/subSUBSCREEN_PARAMETER:SAPLIWP3:0113/ctxtRMIPM-STTAG").text = Format(GetSetting("PMR_START_CYCLE_DATE"), "dd.mm.yyyy")
        On Error GoTo SapErrorHandler
        g_Session.findById("wnd[0]/tbar[0]/btn[11]").press
    Next i

    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Change_StartCycle_PMRs", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Start_PMRs
' DESCRIPTION : Démarre les PMRs sélectionnés de manière automatique.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, ip16_L, Z_Enter, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub Start_PMRs()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    PrepareListPlanTransaction "Z_IP16", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"

    g_Session.findById("wnd[0]/tbar[1]/btn[17]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[0]").Select
    
    Dim NbLign As Long, i As Long
    NbLign = Selection.Rows.count
    
    For i = 1 To NbLign
        
        ' Bouton start
        g_Session.findById("wnd[0]/tbar[1]/btn[9]").press
        
        ' Sélection "Newstart start date"
        On Error Resume Next
        Run ("Z_Enter")
        g_Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        g_Session.findById("wnd[1]/usr/ctxtRMIPM-STADT").text = Format(GetSetting("PMR_START_DATE"), "dd.mm.yyyy")
        g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        ' Save
        On Error GoTo SapErrorHandler
        g_Session.findById("wnd[0]/tbar[0]/btn[11]").press
    Next i

    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Start_PMRs", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : NewStart_PMRs
' DESCRIPTION : Relance les PMRs sélectionnés avec un nouveau démarrage.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, ip16_L, Z_Enter, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub NewStart_PMRs()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    PrepareListPlanTransaction "Z_IP16", "wnd[0]/usr/ctxtWARPL-LOW", "wnd[0]/usr/btn%_WARPL_%_APP_%-VALU_PUSH"

    g_Session.findById("wnd[0]/tbar[1]/btn[17]").press
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[0]").Select
    
    Dim NbLign As Long, i As Long
    NbLign = Selection.Rows.count
    
    For i = 1 To NbLign
        
        ' Bouton "New Start"
        g_Session.findById("wnd[0]/tbar[1]/btn[19]").press
        
        ' Sélection "Newstart start date"
        
        On Error Resume Next
        Run ("Z_Enter")
        g_Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        g_Session.findById("wnd[1]/usr/ctxtRMIPM-STADT").text = Format(GetSetting("PMR_NEW_START_DATE"), "dd.mm.yyyy")
        g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        ' Save
        On Error GoTo SapErrorHandler
        g_Session.findById("wnd[0]/tbar[0]/btn[11]").press
    Next i


    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "NewStart_PMRs", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Planification_PMRs
' DESCRIPTION : Exécute la planification automatique des PMRs via IP30.
' CONTEXTE    : Liste de plans sélectionnés dans Excel.
' DÉPENDANCES : onSAP, Z_F8, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Public Sub Planification_PMRs()
    On Error GoTo SapErrorHandler
    OptimizeExcel
    
    Run ("onSAP")
    
    g_Session.findById("wnd[0]/tbar[0]/okcd").text = "/n ip30"
    g_Session.findById("wnd[0]").sendVKey 0
    
    g_Session.findById("wnd[0]/usr/ctxtWPLAN-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_WPLAN_%_APP_%-VALU_PUSH", Selection
    
    ' Utilise le paramètre de configuration pour la période de planification
    g_Session.findById("wnd[0]/usr/txtOFF_FREI").text = GetSetting("PMR_PLANNING_PERIOD")
    g_Session.findById("wnd[0]/usr/chkNTERM").Selected = False
    g_Session.findById("wnd[0]/usr/chkSTART").Selected = False
    
    Run ("Z_F8")
    ' La macro s'arrête ici pour laisser l'utilisateur exécuter manuellement.
    
    Run ("offSAP")
    RestoreExcel
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Planification_PMRs", Err
End Sub
