Attribute VB_Name = "projectM"
'====================================================================================
' MODULE      : projectM
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module automatise l'ouverture et la gestion des rapports SAP liés
'               aux projets (transactions CN** et CNS**). Il permet d'exécuter des
'               transactions sur une sélection multiple de projets directement depuis Excel.
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

Private Sub ExecuteProjectListTransaction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String)
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
' SECTION 1 : TRANSACTIONS EN MASSE POUR LES PROJETS
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN46NProjL
' DESCRIPTION : Exécute la transaction CN46N (Liste des réseaux de projet) pour une
'               liste de projets copiée depuis la sélection Excel.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN46NProjL()
    ExecuteProjectListTransaction "Z_CN46N", "Z_CN46NProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN42NProjL
' DESCRIPTION : Exécute la transaction CN42N (Structure du projet) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN42NProjL()
    ExecuteProjectListTransaction "Z_CN42N", "Z_CN42NProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CNS41ProjL
' DESCRIPTION : Exécute la transaction CNS41 (Synthèse des coûts) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CNS41ProjL()
    ExecuteProjectListTransaction "Z_CNS41", "Z_CNS41ProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN47ProjL
' DESCRIPTION : Exécute la transaction CN47 (Liste des activités réseau) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN47ProjL()
    ExecuteProjectListTransaction "Z_CN47", "Z_CN47ProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN47NProjL
' DESCRIPTION : Exécute la transaction CN47N (Activités réseau - nouvelle version) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN47NProjL()
    ExecuteProjectListTransaction "Z_CN47N", "Z_CN47NProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN43NProjL
' DESCRIPTION : Exécute la transaction CN43N (Vue hiérarchique) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN43NProjL()
    ExecuteProjectListTransaction "Z_CN43N", "Z_CN43NProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_CN48NProjL
' DESCRIPTION : Exécute la transaction CN48N (Réseaux détaillés) pour une liste de projets.
' CONTEXTE    : Liste de projets sélectionnés dans Excel.
' DÉPENDANCES : ExecuteProjectListTransaction.
'------------------------------------------------------------------------------------
Sub Z_CN48NProjL()
    ExecuteProjectListTransaction "Z_CN48N", "Z_CN48NProjL", "wnd[0]/usr/ctxtCN_PROJN-LOW", "wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH"
End Sub
