Attribute VB_Name = "notifications"
'================================================================================
' MODULE      : notifications
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe toutes les procédures d'interaction avec SAP
'               pour la gestion des avis de maintenance (Notifications).
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session     : Objet de session SAP GUI active.
'   - Z_...         : Macros de transactions SAP prédéfinies.
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'================================================================================
Option Explicit

'================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExecuteSingleNotifAction (Privée)
' DESCRIPTION : Factorise la logique pour les actions sur un avis unique.
'------------------------------------------------------------------------------------
Private Sub ExecuteSingleNotifAction(ByVal transactionWrapper As String, ByVal procName As String, Optional ByVal actionId As String = "", Optional ByVal actionType As String = "TabSelect", Optional ByVal useOnErrorResumeNext As Boolean = False)
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_Enter")
    
    If actionId <> "" Then
        If useOnErrorResumeNext Then On Error Resume Next
        
        Select Case actionType
            Case "TabSelect", "MenuSelect"
                g_Session.findById(actionId).Select
            Case "ButtonPress"
                g_Session.findById(actionId).press
        End Select
        
        If useOnErrorResumeNext Then On Error GoTo SapErrorHandler
    End If
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExecuteListNotifAction (Privée)
' DESCRIPTION : Factorise la logique pour les rapports de liste d'avis.
'------------------------------------------------------------------------------------
Private Sub ExecuteListNotifAction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False, Optional ByVal setFocusField As String = "", Optional ByVal extraFieldToClear As String = "")
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run (transactionWrapper)
    
    If extraFieldToClear <> "" Then g_Session.findById(extraFieldToClear).text = ""
    
    If fieldToClear <> "" Then g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
    If setFocusField <> "" Then g_Session.findById(setFocusField).SetFocus
    
    If clearDates Then
        g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    End If
    
    Run ("Z_F8")
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'================================================================================
' SECTION 1 : TRANSACTIONS IW23 (AFFICHER AVIS) - AVIS UNIQUE
'================================================================================

'------------------------------------------------------------------------------------
Sub Z_IW23NotM(): ExecuteSingleNotifAction "Z_IW23", "Z_IW23NotM": End Sub
Sub Z_IW23NotSum(): ExecuteSingleNotifAction "Z_IW23", "Z_IW23NotSum", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02": End Sub
Sub Z_IW23NotLoc(): ExecuteSingleNotifAction "Z_IW23", "Z_IW23NotLoc", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0100/btnPRED", "ButtonPress": End Sub
Sub Z_IW23NotMul(): ExecuteSingleNotifAction "Z_IW23", "Z_IW23NotMul", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB19", "TabSelect", True: End Sub
Sub Z_IW23NotLog(): ExecuteSingleNotifAction "Z_IW23", "Z_IW23NotLog", "wnd[0]/mbar/menu[3]/menu[4]/menu[2]", "MenuSelect": End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW30_Single
' DESCRIPTION : Exécute IW30 (Affichage multi-niveaux) pour un avis unique.
' CONTEXTE    : Avis unique.
' DÉPENDANCES : onSAP, Z_IW30, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW30_Single()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_IW30")
    
    ' Remplit le numéro d'avis et vide les autres champs de sélection
    g_Session.findById("wnd[0]/usr/ctxtQMNUM-LOW").text = Cells(ActiveCell.row, ActiveCell.Column).value
    g_Session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    
    Run ("Z_F8")
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_PMO_IW30_Single", Err
End Sub

'================================================================================
' SECTION 2 : TRANSACTIONS IW22 (MODIFIER AVIS) - AVIS UNIQUE
'================================================================================

Sub Z_IW22NotM(): ExecuteSingleNotifAction "Z_IW22", "Z_IW22NotM": End Sub
Sub Z_IW22NotSum(): ExecuteSingleNotifAction "Z_IW22", "Z_IW22NotSum", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02": End Sub
Sub Z_IW22NotLoc(): ExecuteSingleNotifAction "Z_IW22", "Z_IW22NotLoc", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0100/btnPRED", "ButtonPress": End Sub
Sub Z_IW22NotMul(): ExecuteSingleNotifAction "Z_IW22", "Z_IW22NotMul", "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB19": End Sub
Sub Z_IW22NotLog(): ExecuteSingleNotifAction "Z_IW22", "Z_IW22NotLog", "wnd[0]/mbar/menu[3]/menu[4]/menu[2]", "MenuSelect": End Sub

'================================================================================
' SECTION 3 : RAPPORTS SUR LISTES D'AVIS (SÉLECTION MULTIPLE)
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW29Open
' DESCRIPTION : Exécute IW29 (Avis ouverts) pour la sélection (liste).
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW29, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW29Open()
    ExecuteListNotifAction "Z_IW29", "Z_IW29Open", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW29Closed
' DESCRIPTION : Exécute IW29 (Avis clôturés/historique) pour la sélection (liste).
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW29Hist, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW29Closed()
    ExecuteListNotifAction "Z_IW29Hist", "Z_IW29Closed", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", False, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW28Open
' DESCRIPTION : Exécute IW28 (Avis ouverts) pour la sélection (liste).
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW28, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW28Open()
    ExecuteListNotifAction "Z_IW28", "Z_IW28Open", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", False, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39NOpen
' DESCRIPTION : Exécute IW39 (Ordres) filtré par la liste d'avis (ouverts).
'               Vide les champs de date pour inclure tout l'historique.
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39NOpen()
    ExecuteListNotifAction "Z_IW39", "Z_IW39NOpen", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", True, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW39NClosed
' DESCRIPTION : Exécute IW39H (Historique ordres) filtré par la liste d'avis (clôturés).
'               Vide les champs de date.
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW39H, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW39NClosed()
    ExecuteListNotifAction "Z_IW39H", "Z_IW39NClosed", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", True, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW38NOpen
' DESCRIPTION : Exécute IW38 (Ordres) filtré par la liste d'avis (ouverts).
'               Vide les champs de date.
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW38NOpen()
    ExecuteListNotifAction "Z_IW38", "Z_IW38NOpen", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", True, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IWBKN
' DESCRIPTION : Exécute IWBK (Composants) filtré par la liste d'avis.
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IWBK, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IWBKN()
    ExecuteListNotifAction "Z_IWBK", "Z_IWBKN", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", False, "wnd[0]/usr/ctxtQMNUM-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IW49NN
' DESCRIPTION : Exécute IW49N (Opérations) filtré par la liste d'avis.
'               Vide le champ "Ordre".
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW49N, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IW49NN()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")
    Run ("Z_IW49N")
    
    ' Vider le champ "Ordre"
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1").Select
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_AUFNR-LOW").text = ""
    
    ' Copie/Colle la liste d'avis depuis Excel
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/ctxtS_QMNUM-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB1/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1100/btn%_S_QMNUM_%_APP_%-VALU_PUSH", Selection
    
    Run ("Z_F8")
    
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_IW49NN", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_PMO_IW30
' DESCRIPTION : Exécute IW30 (Affichage multi-niveaux) pour une liste d'avis.
' CONTEXTE    : Liste d'avis.
' DÉPENDANCES : onSAP, Z_IW30, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_PMO_IW30()
    ExecuteListNotifAction "Z_IW30", "Z_PMO_IW30", "wnd[0]/usr/ctxtQMNUM-LOW", "wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH", True, , "wnd[0]/usr/ctxtAUFNR-LOW"
End Sub
