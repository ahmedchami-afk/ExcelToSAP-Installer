Attribute VB_Name = "Documents"
'==============================================================
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe les procédures VBA pour interagir avec les
'               transactions de gestion de documents (DMS) dans SAP. Il permet de
'               lancer les transactions CV02N, CV03N et CV04N en gérant le cycle
'               de connexion et de déconnexion à SAP pour chaque opération.
'==============================================================
Option Explicit


'====================================================================================
' SECTION : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteSimpleDocumentAction(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Run "onSAP"
    Run transactionWrapper
    Run "Z_Enter"
    Run "offSAP"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSingleDocumentReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldId As String)
    On Error GoTo SapErrorHandler
    Run "onSAP"
    Run transactionWrapper
    g_Session.findById(fieldId).text = ActiveCell.value
    Run "Z_F8"
    Run "offSAP"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteListDocumentReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String)
    On Error GoTo SapErrorHandler
    Run "onSAP"
    Run transactionWrapper
    g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    Run "Z_F8"
    Run "offSAP"
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'====================================================================================
' PROCÉDURES DE GESTION DE DOCUMENTS (DMS)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CV02NDoc
' DESCRIPTION : Ouvre la transaction SAP CV02N (Modifier document) et appuie sur "Entrée".
' CONTEXTE    : Utile pour accéder rapidement à l'écran de modification de document.
' DÉPENDANCES : onSAP, Z_CV02N, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_CV02NDoc()
    ExecuteSimpleDocumentAction "Z_CV02N", "Z_CV02NDoc"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CV03NDoc
' DESCRIPTION : Ouvre la transaction SAP CV03N (Afficher document) et appuie sur "Entrée".
' CONTEXTE    : Utile pour accéder rapidement à l'écran de consultation de document.
' DÉPENDANCES : onSAP, Z_CV03N, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_CV03NDoc()
    ExecuteSimpleDocumentAction "Z_CV03N", "Z_CV03NDoc"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CV04NDoc
' DESCRIPTION : Exécute la transaction CV04N (Rechercher document) pour un document
'               unique dont le numéro est dans la cellule active.
' CONTEXTE    : Conçu pour une recherche rapide à partir d'une seule valeur.
' DÉPENDANCES : onSAP, Z_CV04N, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_CV04NDoc()
    ExecuteSingleDocumentReport "Z_CV04N", "Z_CV04NDoc", "wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKNR-LOW"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CV04NDocL
' DESCRIPTION : Exécute la transaction CV04N (Rechercher document) pour une liste
'               de numéros de documents copiée depuis la sélection Excel.
' CONTEXTE    : Conçu pour une recherche en masse à partir d'une colonne de valeurs.
' DÉPENDANCES : onSAP, Z_CV04N, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_CV04NDocL()
    ExecuteListDocumentReport "Z_CV04N", "Z_CV04NDocL", "wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKNR-LOW", "wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/btn%_STDOKNR_%_APP_%-VALU_PUSH"
End Sub
