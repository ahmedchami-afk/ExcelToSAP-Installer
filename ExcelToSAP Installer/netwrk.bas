Attribute VB_Name = "netwrk"
'====================================================================================
' MODULE      : netwrk
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe les macros d'automatisation SAP destinées à la
'               gestion des réseaux (transactions CN46N, CN47N, CN48N, CNS41, ME5A, etc.).
'               Il fournit un accès rapide et automatisé aux principales transactions
'               SAP liées aux réseaux pour consultation, extraction ou analyse.
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session     : Objet de session SAP GUI active.
'   - Z_F8          : Macro pour simuler la touche F8 (Exécuter).
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'====================================================================================
Option Explicit


'====================================================================================
' SECTION : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteNetworkListTransaction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    If fieldToClear <> "" Then g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'====================================================================================
' SECTION : TRANSACTIONS EN MASSE POUR LES RÉSEAUX
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN46NNetL
' DESCRIPTION : Exécute la transaction CN46N (Liste des réseaux) pour une liste
'               de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN46NNetL()
    ExecuteNetworkListTransaction "Z_CN46N", "Z_CN46NNetL", "wnd[0]/usr/ctxtCN_NETNR-LOW", "wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CNS41NetL
' DESCRIPTION : Exécute la transaction CNS41 (Synthèse des réseaux) pour une liste
'               de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CNS41NetL()
    ExecuteNetworkListTransaction "Z_CNS41", "Z_CNS41NetL", "wnd[0]/usr/ctxtCN_NETNR-LOW", "wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_ME5ANetL
' DESCRIPTION : Exécute la transaction ME5A (Demandes d'achat) pour une liste
'               de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_ME5ANetL()
    ExecuteNetworkListTransaction "Z_ME5A", "Z_ME5ANetL", "wnd[0]/usr/ctxtS_NPLNR-LOW", "wnd[0]/usr/btn%_S_NPLNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN47NetL
' DESCRIPTION : Exécute la transaction CN47 (Activités réseau) pour une liste
'               de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN47NetL()
    ExecuteNetworkListTransaction "Z_CN47", "Z_CN47NetL", "wnd[0]/usr/ctxtCN_NETNR-LOW", "wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN47NNetL
' DESCRIPTION : Exécute la transaction CN47N (Activités réseau - nouvelle version)
'               pour une liste de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN47NNetL()
    ExecuteNetworkListTransaction "Z_CN47N", "Z_CN47NNetL", "wnd[0]/usr/ctxtCN_NETNR-LOW", "wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN48NNetL
' DESCRIPTION : Exécute la transaction CN48N (Suivi des composants réseau) pour une liste
'               de numéros de réseaux copiée depuis la sélection Excel.
' CONTEXTE    : Liste de réseaux sélectionnés dans Excel.
' DÉPENDANCES : ExecuteNetworkListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN48NNetL()
    ExecuteNetworkListTransaction "Z_CN48N", "Z_CN48NNetL", "wnd[0]/usr/ctxtCN_NETNR-LOW", "wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH"
End Sub
