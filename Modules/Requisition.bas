Attribute VB_Name = "Requisition"
'====================================================================================
' MODULE      : Requisition
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module automatise les transactions SAP pour la gestion et la
'               consultation des demandes d'achat (Requisitions). Il permet de lancer
'               les transactions ME53, ME53N et ME5A à partir de la cellule active
'               ou d'une sélection de cellules.
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

Private Sub ExecuteSimpleRequisitionAction(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_Enter")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSingleRequisitionReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldId As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    g_Session.findById(fieldId).text = Cells(ActiveCell.row, ActiveCell.Column).value
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteListRequisitionReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal buttonToFill As String)
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
' SECTION 1 : AFFICHAGE DE DEMANDE D'ACHAT UNIQUE
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME53M
' DESCRIPTION : Ouvre la transaction ME53 (Afficher demande d'achat - ancienne version)
'               et appuie sur "Entrée".
' CONTEXTE    : Demande d'achat unique (le numéro est lu depuis la cellule active).
' DÉPENDANCES : onSAP, Z_ME53, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME53M()
    ExecuteSimpleRequisitionAction "Z_ME53", "Z_ME53M"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME53NM
' DESCRIPTION : Ouvre la transaction ME53N (Afficher demande d'achat - moderne)
'               et appuie sur "Entrée".
' CONTEXTE    : Demande d'achat unique (le numéro est lu depuis la cellule active).
' DÉPENDANCES : onSAP, Z_ME53N, Z_Enter, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME53NM()
    ExecuteSimpleRequisitionAction "Z_ME53N", "Z_ME53NM"
End Sub

'====================================================================================
' SECTION 2 : RAPPORTS SUR DEMANDES D'ACHAT (ME5A)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME5AReq
' DESCRIPTION : Exécute la transaction ME5A pour une demande d'achat unique.
' CONTEXTE    : Demande d'achat unique (le numéro est lu depuis la cellule active).
' DÉPENDANCES : onSAP, Z_ME5A, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME5AReq()
    ExecuteSingleRequisitionReport "Z_ME5A", "Z_ME5AReq", "wnd[0]/usr/ctxtBA_BANFN-LOW"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME5AReqL
' DESCRIPTION : Exécute la transaction ME5A pour une liste de demandes d'achat.
' CONTEXTE    : Liste de demandes d'achat sélectionnées dans Excel.
' DÉPENDANCES : onSAP, Z_ME5A, Z_F8, offSAP, IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_ME5AReqL()
    ExecuteListRequisitionReport "Z_ME5A", "Z_ME5AReqL", "wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH"
End Sub
