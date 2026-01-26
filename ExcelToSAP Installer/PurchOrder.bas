Attribute VB_Name = "PurchOrder"
'====================================================================================
' MODULE      : PurchOrder
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module gère les interactions entre Excel et SAP pour la consultation,
'               l'affichage et la gestion des commandes d'achat (ME23, ME23N, ME80FN, ME91F, ME2N).
'               Il automatise l'exécution de transactions SAP en utilisant la cellule
'               active ou une sélection de cellules comme entrée.
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session     : Objet de session SAP GUI active.
'   - Z_...         : Macros de transactions SAP prédéfinies.
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'====================================================================================
Option Explicit

'====================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteSimplePurchaseOrderAction(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_Enter")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSinglePurchaseOrderReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldId As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    g_Session.findById(fieldId).text = ActiveCell.value
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteListPurchaseOrderReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal buttonToFill As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    FillSAPSelectionList buttonToFill, Selection
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'====================================================================================
' SECTION 1 : TRANSACTIONS SUR COMMANDE D'ACHAT UNIQUE
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME23M
' DESCRIPTION : Ouvre la transaction ME23 (Afficher commande d'achat) et appuie sur "Entrée".
' CONTEXTE    : Commande d'achat unique (le numéro est lu depuis la cellule active).
' DÉPENDANCES : onSAP, Z_ME23, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME23M()
    ExecuteSimplePurchaseOrderAction "Z_ME23", "Z_ME23M"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME23NM
' DESCRIPTION : Ouvre la transaction ME23N (Afficher commande d'achat - moderne) et appuie sur "Entrée".
' CONTEXTE    : Commande d'achat unique (le numéro est lu depuis la cellule active).
' DÉPENDANCES : onSAP, Z_ME23N, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME23NM()
    ExecuteSimplePurchaseOrderAction "Z_ME23N", "Z_ME23NM"
End Sub

'====================================================================================
' SECTION 2 : RAPPORTS SUR COMMANDE D'ACHAT Liste
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME80FNM
' DESCRIPTION : Affiche les informations ME80FN (Analyse générale des documents d'achat)
'               pour un numéro de commande d'achat donné (lu depuis la cellule active).
' CONTEXTE    : Commande d'achat unique.
' DÉPENDANCES : ExecuteSinglePurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME80FNM()
    ExecuteSinglePurchaseOrderReport "Z_ME80FN", "Z_ME80FNM", "wnd[0]/usr/ctxtSP$00003-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME91FM
' DESCRIPTION : Affiche les relances (ME91F) pour un numéro de commande d'achat
'               (lu depuis la cellule active).
' CONTEXTE    : Commande d'achat unique.
' DÉPENDANCES : ExecuteSinglePurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME91FM()
    ExecuteSinglePurchaseOrderReport "Z_ME91F", "Z_ME91FM", "wnd[0]/usr/ctxtEN_EBELN-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME2NM
' DESCRIPTION : Ouvre la liste des commandes (ME2N) pour un numéro de commande d'achat
'               donné (lu depuis la cellule active).
' CONTEXTE    : Commande d'achat unique.
' DÉPENDANCES : ExecuteSinglePurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME2NM()
    ExecuteSinglePurchaseOrderReport "Z_ME2N", "Z_ME2NM", "wnd[0]/usr/ctxtEN_EBELN-LOW"
End Sub

'====================================================================================
' SECTION 3 : RAPPORTS SUR LISTES DE COMMANDES D'ACHAT
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME80FNML
' DESCRIPTION : Affiche les informations ME80FN (Analyse générale des documents d'achat)
'               pour une liste de numéros de commandes d'achat copiée depuis Excel.
' CONTEXTE    : Liste de commandes d'achat.
' DÉPENDANCES : ExecuteListPurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME80FNML()
    ExecuteListPurchaseOrderReport "Z_ME80FN", "Z_ME80FNML", "wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME91FML
' DESCRIPTION : Lance la transaction ME91F (Relances) pour une liste de numéros de
'               commandes d'achat copiée depuis Excel.
' CONTEXTE    : Liste de commandes d'achat.
' DÉPENDANCES : ExecuteListPurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME91FML()
    ExecuteListPurchaseOrderReport "Z_ME91F", "Z_ME91FML", "wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME2NML
' DESCRIPTION : Ouvre la transaction ME2N (Liste des commandes d'achat) pour une
'               liste de numéros de commandes d'achat copiée depuis Excel.
' CONTEXTE    : Liste de commandes d'achat.
' DÉPENDANCES : ExecuteListPurchaseOrderReport.
'------------------------------------------------------------------------------------
Sub Z_ME2NML()
    ExecuteListPurchaseOrderReport "Z_ME2N", "Z_ME2NML", "wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH"
End Sub
