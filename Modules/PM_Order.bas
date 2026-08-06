Attribute VB_Name = "PM_Order"
'================================================================================
' MODULE      : PM_Order
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe toutes les procédures d'interaction avec SAP
'               pour la gestion des ordres de maintenance (PM Orders).
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session     : Objet de session SAP GUI active.
'   - Z_...         : Macros de transactions SAP prédéfinies.
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'   - GetSetting    : Lit les paramètres depuis la configuration.
'================================================================================
Option Explicit

'================================================================================
' SECTION 0 : PROCÉDURE UTILITAIRE PRIVÉE
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExecuteSingleOrderAction (Privée)
' DESCRIPTION : Factorise la logique commune pour les actions sur un ordre unique.
'               Gère le cycle de vie de la connexion, le lancement de la transaction,
'               et l'exécution d'une action optionnelle (sélection d'onglet, menu, etc.).
' PARAMÈTRES  :
'   - transactionWrapper (String) : Le nom de la macro SAP à lancer (ex: "Z_IW33").
'   - procName (String)           : Le nom de la procédure appelante (pour le log d'erreur).
'   - actionId (String)           : L'ID de l'objet SAP sur lequel agir (optionnel).
'   - actionType (String)         : Le type d'action à effectuer ("TabSelect", "MenuSelect").
'   - fieldToFill (String)        : ID du champ à remplir avec la valeur de la cellule active (optionnel).
'   - executeKey (String)         : Touche de validation ("Z_Enter" par défaut, ou "Z_F8").
'   - clearStandardDates (Boolean): Si True, vide les champs DATUV et DATUB.
'------------------------------------------------------------------------------------
Private Sub ExecuteSingleOrderAction(ByVal transactionWrapper As String, ByVal procName As String, Optional ByVal actionId As String = "", Optional ByVal actionType As String = "TabSelect", Optional ByVal fieldToFill As String = "", Optional ByVal executeKey As String = "Z_Enter", Optional ByVal clearStandardDates As Boolean = False)
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run (transactionWrapper)
    
    If fieldToFill <> "" Then
        g_Session.findById(fieldToFill).text = Cells(ActiveCell.row, ActiveCell.Column).value
    End If
    
    If clearStandardDates Then
        On Error Resume Next
        g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
        On Error GoTo SapErrorHandler
    End If
    
    Run (executeKey)
    
    If actionId <> "" Then
        Select Case actionType
            Case "TabSelect", "MenuSelect"
                g_Session.findById(actionId).Select
            Case "ButtonPress"
                g_Session.findById(actionId).press
        End Select
    End If
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : PrepareListOrderTransaction (Privée)
' DESCRIPTION : Prépare et lance une transaction sur une liste d'ordres.
'               Gère la connexion, les filtres (checkboxes), l'injection de la liste et l'exécution.
'               Ne ferme PAS la session (pour permettre des actions post-exécution).
'------------------------------------------------------------------------------------
Private Sub PrepareListOrderTransaction(ByVal transactionWrapper As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False, Optional ByVal checkboxesToSelect As Variant)
    Run ("onSAP")
    Run (transactionWrapper)
    
    ' Gestion des cases à cocher (filtres de statut)
    If Not IsMissing(checkboxesToSelect) Then
        Dim chk As Variant
        If IsArray(checkboxesToSelect) Then
            For Each chk In checkboxesToSelect
                g_Session.findById(chk).Selected = True
            Next chk
        End If
    End If
    
    If fieldToClear <> "" Then g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
    If clearDates Then
        On Error Resume Next
        g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
        On Error GoTo 0
    End If
    
    Run ("Z_F8")
End Sub

Private Sub ExecuteListOrderAction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False, Optional ByVal checkboxesToSelect As Variant)
    On Error GoTo SapErrorHandler
    PrepareListOrderTransaction transactionWrapper, fieldToClear, buttonToFill, clearDates, checkboxesToSelect
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteIW49N(ByVal procName As String, ByVal isList As Boolean)
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_IW49N")
    
    ' Tab 1: Selection
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    
    If isList Then
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUFNR-LOW").text = ""
        FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_AUFNR_%_APP_%-VALU_PUSH", Selection
    Else
        g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUFNR-LOW").text = Cells(ActiveCell.row, ActiveCell.Column).value
    End If
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_DATUM-LOW").text = ""
    
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9").Select
    If GetSetting("LAY_IW49N") <> "" Then g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB9/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1900/ctxtSP_VARI").text = GetSetting("LAY_IW49N")
    
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'================================================================================
' SECTION 1 : TRANSACTIONS IW33 (AFFICHER ORDRE) - ORDRE UNIQUE
'================================================================================

'------------------------------------------------------------------------------------
Sub Z_PMO_IW33_1_Main(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_1_Main": End Sub
Sub Z_PMO_IW33_2_Operations(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_2_Operations", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE": End Sub
Sub Z_PMO_IW33_3_Components(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_3_Components", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB": End Sub
Sub Z_PMO_IW33_4_Costs(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_4_Costs", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpKOAU": End Sub
Sub Z_PMO_IW33_5_Planning(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_5_Planning", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHPL": End Sub
Sub Z_PMO_IW33_6_Enhancement(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_6_Enhancement", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK": End Sub
Sub Z_PMO_IW33_7_LOG(): ExecuteSingleOrderAction "Z_IW33", "Z_PMO_IW33_7_LOG", "wnd[0]/mbar/menu[3]/menu[12]/menu[5]", "MenuSelect": End Sub

'================================================================================
' SECTION 2 : TRANSACTIONS IW32 (MODIFIER ORDRE) - ORDRE UNIQUE
'================================================================================

'------------------------------------------------------------------------------------
Sub Z_PMO_IW32_1_Main(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_1_Main": End Sub
Sub Z_PMO_IW32_2_Operations(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_2_Operations", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE": End Sub
Sub Z_PMO_IW32_3_Components(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_3_Components", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpMUEB": End Sub
Sub Z_PMO_IW32_4_Costs(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_4_Costs", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpKOAU": End Sub
Sub Z_PMO_IW32_5_Planning(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_5_Planning", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHPL": End Sub
Sub Z_PMO_IW32_6_Enhancement(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_6_Enhancement", "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK": End Sub
Sub Z_PMO_IW32_7_LOG(): ExecuteSingleOrderAction "Z_IW32", "Z_PMO_IW32_7_LOG", "wnd[0]/mbar/menu[3]/menu[12]/menu[5]", "MenuSelect": End Sub


'================================================================================
' SECTION 3 : RAPPORTS DIVERS SUR ORDRE UNIQUE
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW49N_Operations
' DESCRIPTION : Exécute IW49N (Opérations) pour l'ordre dans la cellule active.
'               Utilise le layout défini dans la configuration ("LAY_IW49N").
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW49N, Z_F8, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW49N_Operations()
    ExecuteIW49N "Z_PMO_IW49N_Operations", False
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW40
' DESCRIPTION : Exécute IW40 (Liste d'ordres PM) pour l'ordre de la cellule active.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW40, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW40()
    ExecuteSingleOrderAction "Z_IW40", "Z_PMO_IW40", , , "wnd[0]/usr/ctxtAUFNR-LOW", "Z_F8", True
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IWBK
' DESCRIPTION : Exécute IWBK (Composants) pour l'ordre de la cellule active.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IWBK, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IWBK()
    ExecuteSingleOrderAction "Z_IWBK", "Z_PMO_IWBK", , , "wnd[0]/usr/ctxtAUFNR-LOW", "Z_F8"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW47
' DESCRIPTION : Exécute IW47 (Confirmations) pour l'ordre de la cellule active.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW47, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW47()
    ExecuteSingleOrderAction "Z_IW47", "Z_PMO_IW47", , , "wnd[0]/usr/ctxtAUFNR_O-LOW", "Z_F8"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_MB51_MatChar
' DESCRIPTION : Exécute MB51 (Mouvements de stock) pour l'ordre de la cellule active.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_MB51, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_MB51_MatChar()
    ExecuteSingleOrderAction "Z_MB51", "Z_PMO_MB51_MatChar", , , "wnd[0]/usr/ctxtAUFNR-LOW", "Z_F8"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_ME5A
' DESCRIPTION : Exécute ME5A (Demandes d'achat) pour l'ordre de la cellule active.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_ME5A, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_ME5A()
    ExecuteSingleOrderAction "Z_ME5A", "Z_PMO_ME5A", , , "wnd[0]/usr/ctxtS_AUFNR-LOW", "Z_F8"
End Sub

'================================================================================
' SECTION 4 : ACTIONS AUTOMATISÉES SUR ORDRE UNIQUE (IW32)
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW32_15_Print_Orders
' DESCRIPTION : Imprime un ordre de travail via la transaction IW32.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW32, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW32_15_Print_Orders()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_IW32")
    Run ("Z_Enter")
    
    ' --- Print WO ---
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK").Select
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/radRADIO_BUTTON-WOPM").Select
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/radRADIO_BUTTON-WOPM").SetFocus
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/btn%#AUTOTEXT002").press
    g_Session.findById("wnd[0]/tbar[1]/btn[31]").press
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW32_15_Print_Orders", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW32_16_WO_to_BOOM
' DESCRIPTION : Crée une nomenclature (BOM) à partir d'un ordre de travail via IW32.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW32, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW32_16_WO_to_BOOM()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_IW32")
    Run ("Z_Enter")
    
    ' --- Print WO ---
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK").Select
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/radRADIO_BUTTON-O2B").Select
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/radRADIO_BUTTON-O2B").SetFocus
    g_Session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabp+CUK/ssubSUB_AUFTRAG:SAPLCOIH:1180/ssubCUSTSCR1:SAPLXWOC:0900/btn%#AUTOTEXT002").press
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW32_16_WO_to_BOOM", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW32_17_CNF
' DESCRIPTION : Lance la transaction IW41 (Saisie de confirmation) pour l'ordre actif.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW41, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW32_17_CNF()
    ExecuteSingleOrderAction "Z_IW41", "Z_PMO_IW32_17_CNF"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW44
' DESCRIPTION : Lance la transaction IW44 (Afficher Confirmation) pour la confirmation
'               dont le numéro est dans la cellule active.
' CONTEXTE    : Confirmation unique.
' DÉPENDANCES : onSAP, SAP.Z_IW44, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW44()
    ExecuteSingleOrderAction "Z_IW44", "Z_IW44"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW32_18_Cancel_CNF
' DESCRIPTION : Annule une confirmation via la transaction IW45.
' CONTEXTE    : Ordre unique.
' DÉPENDANCES : onSAP, Z_IW45, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW32_18_Cancel_CNF()
    On Error GoTo SapErrorHandler
    
    ' Détermine le champ à remplir selon le type de donnée (Ordre ou Confirmation)
    Dim fieldId As String
    If g_DataType = "CNF" Then
        fieldId = "ctxtCORUF-RUECK" ' Champ Confirmation
    Else
        fieldId = "ctxtCORUF-AUFNR" ' Champ Ordre (défaut)
    End If
    
    ' Utilise RunSAPTransaction directement pour cibler le bon champ
    RunSAPTransaction "IW45", fieldId
    Run ("Z_Enter")
    
    ' Gérer la popup de confirmation (si elle apparaît)
    On Error Resume Next
    g_Session.findById("wnd[1]/usr/btnOPTION2").press
    On Error GoTo SapErrorHandler
    
    
    ' Sauvegarder l'annulation
    g_Session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save
    
    ' Revenir en arrière pour la prochaine itération
    g_Session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW32_18_Cancel_CNF", Err
End Sub

'================================================================================
' SECTION 5 : RAPPORTS SUR LISTES D'ORDRES (SÉLECTION MULTIPLE)
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39_O
' DESCRIPTION : Exécute IW39 (Liste d'ordres PM) pour une liste d'ordres.
'               Vide les champs de date pour inclure tout l'historique.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39_O()
    ExecuteListOrderAction "Z_IW39", "Z_IW39_O", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39_H
' DESCRIPTION : Exécute IW39H (Historique) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW39H, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39_H()
    ExecuteListOrderAction "Z_IW39H", "Z_IW39_H", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW38O
' DESCRIPTION : Exécute IW38 (Liste d'ordres PM) pour une liste d'ordres ouverts.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW38O()
    ExecuteListOrderAction "Z_IW38", "Z_IW38O", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW38ALL
' DESCRIPTION : Exécute IW38 (Liste d'ordres PM) pour une liste d'ordres, tous statuts confondus.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW38ALL()
    ExecuteListOrderAction "Z_IW38", "Z_IW38ALL", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW49NL
' DESCRIPTION : Exécute IW49N (Opérations) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW49N, Z_F8, offSAP, GetSetting, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW49NL()
    ExecuteIW49N "Z_IW49NL", True
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW40_L
' DESCRIPTION : Exécute IW40 (Liste d'ordres PM) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW40, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW40_L()
    ExecuteListOrderAction "Z_IW40", "Z_IW40_L", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IWBK_L
' DESCRIPTION : Exécute IWBK (Composants) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IWBK, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IWBK_L()
    ExecuteListOrderAction "Z_IWBK", "Z_IWBK_L", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW47_L
' DESCRIPTION : Exécute IW47 (Confirmations) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_IW47, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW47_L()
    ExecuteListOrderAction "Z_IW47", "Z_IW47_L", "wnd[0]/usr/ctxtAUFNR_O-LOW", "wnd[0]/usr/btn%_AUFNR_O_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_MB51_L
' DESCRIPTION : Exécute MB51 (Mouvements de stock) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_MB51, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_MB51_L()
    ExecuteListOrderAction "Z_MB51", "Z_MB51_L", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME5A_L
' DESCRIPTION : Exécute ME5A (Demandes d'achat) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_ME5A, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME5A_L()
    ExecuteListOrderAction "Z_ME5A", "Z_ME5A_L", "wnd[0]/usr/ctxtS_AUFNR-LOW", "wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH"
End Sub

'================================================================================
' SECTION 6 : ACTIONS EN MASSE SUR LISTES D'ORDRES
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_MassRelease
' DESCRIPTION : Lance la transaction IW38, colle une liste d'ordres,
'               et effectue une libération en masse (Mass Release).
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38O, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_MassRelease()
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    ' Libération en masse
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press   'select all
    g_Session.findById("wnd[0]/tbar[1]/btn[44]").press  'Release
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW38_MassRelease", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_MassTECO
' DESCRIPTION : Lance la transaction IW38, colle une liste d'ordres,
'               et effectue une clôture technique (TECO) en masse.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38O, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_MassTECO()
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    ' Clôture technique (TECO)
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press           ' Select all
    g_Session.findById("wnd[0]/tbar[1]/btn[45]").press          ' TECO
    g_Session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press   ' with notifications
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press           ' Exécuter TECO
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW38_MassTECO", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_User_Status_mass_Change
' DESCRIPTION : Lance IW38, colle une liste d'ordres, et modifie en masse
'               le statut utilisateur.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38ALL, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_User_Status_mass_Change()
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    ' Actions pour le changement de statut en masse
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[14]").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[5]").press
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 12
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").GetAbsoluteRow(16).Selected = True
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,4]").SetFocus
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,4]").caretPosition = 0
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    g_Session.findById("wnd[2]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[1]").sendVKey 4
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW38_User_Status_mass_Change", Err
End Sub

Private Sub ExecuteMassStatusChange(ByVal procName As String, ByVal columnIndex As Integer)
    On Error GoTo SapErrorHandler
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[14]").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[5]").press
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").verticalScrollbar.Position = 12
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").GetAbsoluteRow(16).Selected = True
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,4]").SetFocus
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,4]").caretPosition = 0
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    g_Session.findById("wnd[2]/tbar[0]/btn[0]").press
    g_Session.findById("wnd[1]").sendVKey 4
    g_Session.findById("wnd[2]/usr/lbl[1," & columnIndex & "]").SetFocus
    g_Session.findById("wnd[2]").sendVKey 2
    g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_4SCH
' DESCRIPTION : Applique le statut utilisateur "4SCH" en masse à une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38ALL, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_4SCH()
    ExecuteMassStatusChange "Z_PMO_IW38_4SCH", 7
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_5COM
' DESCRIPTION : Applique le statut utilisateur "5COM" en masse à une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38ALL, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_5COM()
    ExecuteMassStatusChange "Z_PMO_IW38_5COM", 8
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_Change_Date
' DESCRIPTION : Ouvre l'interface de modification en masse des dates pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38ALL, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_Change_Date()
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    'change date
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[14]").Select
    g_Session.findById("wnd[1]/tbar[0]/btn[5]").press
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/cmbGRUPPEN-TEXT").SetFocus
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/cmbGRUPPEN-TEXT").key = "Date fields"
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").GetAbsoluteRow(1).Selected = True
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,1]").SetFocus
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER/txtALLE_FELDER-SCRTEXT[0,1]").caretPosition = 0
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subALLE_FELDER:SAPLCNFA:0130/tblSAPLCNFATC_ALLE_FELDER").GetAbsoluteRow(0).Selected = True
    g_Session.findById("wnd[2]/usr/ssubRAHMEN:SAPLCNFA:0111/subAUSWAHL:SAPLCNFA:0140/btnAUSWAEHLEN").press
    g_Session.findById("wnd[2]/tbar[0]/btn[0]").press
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW38_Change_Date", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_Print_Orders
' DESCRIPTION : Imprime en masse une liste d'ordres via la transaction /PROGROUP/PPM.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_PPM, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_Print_Orders()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_PPM") ' Lance /PROGROUP/PPM
    
    ' Copier/Coller la liste d'ordres
    g_Session.findById("wnd[0]/usr/ctxtP_AUFNR-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_P_AUFNR_%_APP_%-VALU_PUSH", Selection
    g_Session.findById("wnd[0]/tbar[1]/btn[8]").press ' Probablement le bouton OK/Continuer de la popup
    
    Run ("Z_F8") ' Exécuter
    
    ' Ouvrir la sélection des papiers d'atelier
    g_Session.findById("wnd[0]/tbar[1]/btn[31]").press
    
    
    ' Lancer l'aperçu avant impression
    g_Session.findById("wnd[0]/tbar[1]/btn[9]").press
    
    ' La procédure s'arrête sur l'aperçu, l'utilisateur ferme manuellement.
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_Print_Orders", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW38_MassConfirmation
' DESCRIPTION : Lance une confirmation en masse pour une liste d'ordres via IW38.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW38O, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW38_MassConfirmation()
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW38", "wnd[0]/usr/ctxtAUFNR-LOW", "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH", True, Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR", "wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS")
    
    ' Confirmation en masse
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[4]/menu[0]").Select
    g_Session.findById("wnd[0]/tbar[1]/btn[9]").press
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW38_MassConfirmation", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_Cancel_Confirmations
' DESCRIPTION : Annule les confirmations pour une liste d'ordres via IW47 et IW45.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : Z_IW47_L, onSAP, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_Cancel_Confirmations()
Dim i As Integer
    On Error GoTo SapErrorHandler
    
    PrepareListOrderTransaction "Z_IW47", "wnd[0]/usr/ctxtAUFNR_O-LOW", "wnd[0]/usr/btn%_AUFNR_O_%_APP_%-VALU_PUSH"
    
    Run ("onSAP")
    
    ' Annulation de la confirmation
    g_Session.findById("wnd[0]/tbar[1]/btn[5]").press
    g_Session.findById("wnd[0]/mbar/menu[3]/menu[0]").Select
    On Error Resume Next
    g_Session.findById("wnd[1]").Close
    On Error GoTo 0
    
    For i = 1 To Selection.Rows.count
    Z_Enter
    On Error Resume Next
    g_Session.findById("wnd[1]/usr/btnOPTION2").press
    g_Session.findById("wnd[0]/tbar[0]/btn[11]").press
    g_Session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    Next i

Cleanup:
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_Cancel_Confirmations", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_Gantt
' DESCRIPTION : Affiche le diagramme de Gantt pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_Gannt_Display, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_Gantt()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_Gannt_Display")
    
    ' Options d'affichage
    If GetSetting("LAY_DSP-Gannt") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Gannt")
    If GetSetting("LAY_PGPNL/GS_D") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_PGPNL/GS_D")
    
    ' Période de 14 jours avant et après
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = "7"
    
    ' Usine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = "0P1D"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = "0P1D"
    
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
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUFNR-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_AUFNR_%_APP_%-VALU_PUSH", Selection
    
    Run ("Z_F8") ' Exécuter
    
    'Zoom out
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "ZOOMOUT"
    'g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "ZOOMOUT"

    
    ' La procédure s'arrête sur le Gantt, l'utilisateur ferme manuellement.
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_Gantt", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_SCHEDULER
' DESCRIPTION : Affiche le planificateur (Scheduler) pour une liste d'ordres.
' CONTEXTE    : Liste d'ordres.
' DÉPENDANCES : onSAP, Z_Gannt_Scheduler, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_SCHEDULER()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_Gannt_Scheduler")

    ' Options d'affichage
    If GetSetting("LAY_DSP-Schudler") <> "" Then g_Session.findById("wnd[0]/usr/ctxtR_DSPSET").text = GetSetting("LAY_DSP-Schudler")
    If GetSetting("LAY_/PGP/SCHEDULER") <> "" Then g_Session.findById("wnd[0]/usr/ctxtPS_VARI").text = GetSetting("LAY_/PGP/SCHEDULER")
    
    ' Période de 14 jours avant et après
    g_Session.findById("wnd[0]/usr/txtP_DAYS").text = "7"
    g_Session.findById("wnd[0]/usr/txtP_LOWDAY").text = "7"
    
    ' Usine
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtIWERK-LOW").text = "0P1D"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtS_WERKS-LOW").text = "0P1D"
    
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
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/ctxtAUFNR-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_PMPS_002/tabpTAB_PM/ssub%_SUBSCREEN_PMPS_002:/PGPNL/GPSS_PM:0212/btn%_AUFNR_%_APP_%-VALU_PUSH", Selection
    
    Run ("Z_F8") ' Exécuter
    
    'Zoom out
    g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "ZOOMOUT"
    'g_Session.findById("wnd[0]/shellcont/shellcont/shell").pressButton "ZOOMOUT"

    
    ' La procédure s'arrête sur le Gantt, l'utilisateur ferme manuellement.
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_SCHEDULER", Err
End Sub
