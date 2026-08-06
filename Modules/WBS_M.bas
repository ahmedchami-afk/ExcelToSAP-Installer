Attribute VB_Name = "WBS_M"
'====================================================================================
' MODULE      : WBS_M (Work Breakdown Structure - Macros)
' AUTEUR      : Votre Nom / Gemini Code Assist
' DATE        : 02/12/2025
' DESCRIPTION :
'   Ce module regroupe les macros permettant de lancer des transactions SAP
'   en utilisant une liste d'éléments WBS (Work Breakdown Structure, ou OTP en français)
'   sélectionnée directement depuis une feuille Excel.
'
'   Chaque procédure suit le même modèle :
'     1. Établit la connexion à SAP (`onSAP`).
'     2. Lance la transaction SAP de base (ex: `Z_CN46N`).
'     3. Copie la sélection de cellules depuis Excel.
'     4. Colle les données dans le champ de sélection multiple de la transaction.
'     5. Exécute la transaction (`Z_F8`).
'     6. Ferme la connexion SAP (`offSAP`).
'
' DÉPENDANCES :
'   - Module `SAP` : pour les fonctions `onSAP`, `offSAP`, `Z_F8`, `IsSAPConnectionAlive`
'     et les wrappers de transactions (ex: `Z_CN46N`).
'====================================================================================
Option Explicit
'====================================================================================

'====================================================================================
' SECTION : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteWBSListTransaction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
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

'====================================================================================
' SECTION : TRANSACTIONS POUR LISTES D'ÉLÉMENTS WBS (OTP)
'====================================================================================

'------------------------------------------------------------------------------
' PROCEDURE   : Z_CN46N_WBSL
' DESCRIPTION :
'   Exécute la transaction CN46N (Vue d'ensemble réseau) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CN46N_WBSL()
    ExecuteWBSListTransaction "Z_CN46N", "Z_CN46N_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_CNS41_WBSL
' DESCRIPTION :
'   Exécute la transaction CNS41 (Synthèse des composants) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CNS41_WBSL()
    ExecuteWBSListTransaction "Z_CNS41", "Z_CNS41_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_ME5A_WBSL
' DESCRIPTION :
'   Exécute la transaction ME5A (Liste des demandes d'achat) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_ME5A_WBSL()
    ExecuteWBSListTransaction "Z_ME5A", "Z_ME5A_WBSL", "wnd[0]/usr/ctxtS_PSEXT-LOW", "wnd[0]/usr/btn%_S_PSEXT_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_CN47_WBSL
' DESCRIPTION :
'   Exécute la transaction CN47 (Liste des activités) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CN47_WBSL()
    ExecuteWBSListTransaction "Z_CN47", "Z_CN47_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_CN47N_WBSL
' DESCRIPTION :
'   Exécute la transaction CN47N (Liste des activités) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CN47N_WBSL()
    ExecuteWBSListTransaction "Z_CN47N", "Z_CN47N_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_CN48N_WBSL
' DESCRIPTION :
'   Exécute la transaction CN48N (Liste des confirmations) pour une liste d'éléments WBS
'   sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CN48N_WBSL()
    ExecuteWBSListTransaction "Z_CN48N", "Z_CN48N_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_CN43N_WBSL
' DESCRIPTION :
'   Exécute la transaction CN43N (Liste hiérarchique des documents) pour une liste
'   d'éléments WBS sélectionnés dans Excel.
'------------------------------------------------------------------------------
Sub Z_CN43N_WBSL()
    ExecuteWBSListTransaction "Z_CN43N", "Z_CN43N_WBSL", "wnd[0]/usr/ctxtCN_PSPNR-LOW", "wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH"
End Sub


'------------------------------------------------------------------------------
' PROCEDURE   : Z_IW38_WBSL
' DESCRIPTION :
'   Exécute la transaction IW38 (Liste des ordres de travail) pour une liste d'éléments WBS
'   sélectionnés dans Excel. Les dates sont effacées pour une recherche étendue.
'------------------------------------------------------------------------------
Sub Z_IW38_WBSL()
    ExecuteWBSListTransaction "Z_IW38", "Z_IW38_WBSL", "wnd[0]/usr/ctxtPROID-LOW", "wnd[0]/usr/btn%_PROID_%_APP_%-VALU_PUSH", True
End Sub
