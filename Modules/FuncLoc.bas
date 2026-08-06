Attribute VB_Name = "FuncLoc"
'==============================================================
' MODULE      : FuncLoc
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe les macros d'interaction avec SAP pour la gestion
'               des emplacements fonctionnels (Functional Locations).
'
'               Chaque procédure est autonome et gère le cycle de vie de la connexion SAP :
'                 - Connexion à SAP (onSAP).
'                 - Exécution d'une transaction (IL03, IW39, IP18, etc.).
'                 - Déconnexion de SAP (offSAP).
'
'               Les procédures sont organisées par type d'action :
'                 - Consultation d'informations (IL03).
'                 - Affichage des ordres de maintenance (IW39, IW38).
'                 - Consultation des nomenclatures (BOM) via IB0x.
'                 - Gestion des listes de tâches (IAxx).
'==============================================================

Option Explicit

'====================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'====================================================================================

Private Sub ExecuteSingleFlocAction(ByVal transactionWrapper As String, ByVal procName As String, Optional ByVal actionId As String = "", Optional ByVal actionType As String = "ButtonPress")
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_Enter")
    
    If actionId <> "" Then
        If actionType = "ButtonPress" Then
            g_Session.findById(actionId).press
        ElseIf actionType = "TabSelect" Then
            WaitForSAP
            g_Session.findById(actionId).Select
        End If
    End If
    
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSingleFlocReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToFill As String, Optional ByVal checkboxesToSelect As Variant, Optional ByVal checkboxesToDeselect As Variant)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    If Not IsMissing(checkboxesToSelect) Then
        Dim chk As Variant
        For Each chk In checkboxesToSelect
            g_Session.findById(chk).Selected = True
        Next chk
    End If
    
    If Not IsMissing(checkboxesToDeselect) Then
        Dim chk2 As Variant
        For Each chk2 In checkboxesToDeselect
            g_Session.findById(chk2).Selected = False
        Next chk2
    End If
    
    g_Session.findById(fieldToFill).text = ActiveCell.value
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteFlocBOM(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Dim EQNMB As String
    Run ("onSAP")
    Run ("Z_IL03")
    Run ("Z_Enter")
    
    ' Lecture de l'équipement installé
    g_Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\04").Select
    EQNMB = g_Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\04/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102B:SAPLITO0:1061/subSUB_1061A:SAPLIEL2:0110/tblSAPLIEL2TCTRL_0110/ctxtIEQINSTALL-EQUNR[1,0]").text
    
    Run (transactionWrapper)
    g_Session.findById("wnd[0]/usr/ctxtRC29N-EQUNR").text = EQNMB
    
    If GetSetting("SAP_PLANT_PF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = GetSetting("SAP_PLANT_PF")
    Else
        If CheckSAPError(transactionWrapper) Then GoTo CleanExit
    End If
    
    g_Session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "4"
    Run ("Z_Enter")
    
CleanExit:
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteIA17FlocList(ByVal procName As String, ByVal radioButtonId As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run ("Z_IA17")
    g_Session.findById("wnd[0]/usr/ctxtPN_STRNO-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_PN_STRNO_%_APP_%-VALU_PUSH", Selection
    g_Session.findById(radioButtonId).Select
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'====================================================================================
' SECTION 1 : CONSULTATION DE BASE (TRANSACTION IL03)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_FLOC_IL03
' DESCRIPTION : Exécute la transaction SAP IL03 (Afficher un emplacement fonctionnel)
'               et appuie sur "Entrée".
' DÉPENDANCES : onSAP, Z_IL03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_FLOC_IL03()
    ExecuteSingleFlocAction "Z_IL03", "Z_FLOC_IL03"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_FLOC_IL03_c
' DESCRIPTION : Exécute IL03 et affiche les caractéristiques associées à l'emplacement fonctionnel.
' DÉPENDANCES : onSAP, Z_IL03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_FLOC_IL03_c()
    ExecuteSingleFlocAction "Z_IL03", "Z_FLOC_IL03_c", "wnd[0]/usr/tabsTABSTRIP/tabpT\03", "TabSelect"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_FLOC_IL03_B
' DESCRIPTION : Ouvre IL03 puis accède aux vues BOM (Nomenclature) et structure technique.
' DÉPENDANCES : onSAP, Z_IL03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_FLOC_IL03_B()
    On Error GoTo SapErrorHandler

    Run ("onSAP")
    Run ("Z_IL03")
    Run ("Z_Enter")
    
    ' Appuie sur le bouton "Nomenclature de structure" (BOM)
    g_Session.findById("wnd[0]/tbar[1]/btn[32]").press
    ' Appuie sur le bouton "Structure"
    g_Session.findById("wnd[0]/tbar[1]/btn[16]").press

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_FLOC_IL03_B", Err
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : IH01
' DESCRIPTION : Exécute la transaction SAP IH01 (Arborescence Structurelle).
' DÉPENDANCES : onSAP, Z_IH01, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_FLOC_IH01()
    On Error GoTo SapErrorHandler

    Run ("onSAP")
    Run ("Z_IH01")
    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "Z_FLOC_IH01", Err
End Sub


'====================================================================================
' SECTION 2 : CONSULTATION DES ORDRES ET AVIS (POUR UN FLOC UNIQUE)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : flocPMO
' DESCRIPTION : Exécute IW39 pour afficher les ordres de maintenance liés à un
'               emplacement fonctionnel dont l'ID est dans la cellule active.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocPMO()
    ExecuteSingleFlocReport "Z_IW39", "flocPMO", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocPMO_H
' DESCRIPTION : Variante de `flocPMO` qui affiche les ordres historiques/terminés.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocPMO_H()
    ExecuteSingleFlocReport "Z_IW39", "flocPMO_H", "wnd[0]/usr/ctxtSTRNO-LOW", Array("wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS"), Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR")
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocPMOc
' DESCRIPTION : Exécute IW38 pour afficher les ordres en cours de l'emplacement actif.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocPMOc()
    ExecuteSingleFlocReport "Z_IW38", "flocPMOc", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocNotif
' DESCRIPTION : Exécute IW29 pour afficher les notifications associées à un FLoc.
' DÉPENDANCES : onSAP, Z_IW29, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocNotif()
    ExecuteSingleFlocReport "Z_IW29", "flocNotif", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocNotifH
' DESCRIPTION : Affiche uniquement les notifications historiques (terminées) du FLoc.
' DÉPENDANCES : onSAP, Z_IW29, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocNotifH()
    ExecuteSingleFlocReport "Z_IW29", "flocNotifH", "wnd[0]/usr/ctxtSTRNO-LOW", Array("wnd[0]/usr/chkDY_MAB"), Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR")
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocNotifc
' DESCRIPTION : Exécute IW28 pour consulter les notifications en cours du FLoc.
' DÉPENDANCES : onSAP, Z_IW28, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocNotifc()
    ExecuteSingleFlocReport "Z_IW28", "flocNotifc", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub


'====================================================================================
' SECTION 3 : CONSULTATION DES PLANS DE MAINTENANCE (POUR UN FLOC UNIQUE)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : flocplan
' DESCRIPTION : Exécute IP18 pour afficher les plans de maintenance liés à un FLoc.
' DÉPENDANCES : onSAP, Z_IP18, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocplan()
    ExecuteSingleFlocReport "Z_IP18", "flocplan", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocplan24
' DESCRIPTION : Exécute IP24 pour consulter la planification de maintenance d'un FLoc.
' DÉPENDANCES : onSAP, Z_IP24, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocplan24()
    ExecuteSingleFlocReport "Z_IP24", "flocplan24", "wnd[0]/usr/ctxtSTRNO-LOW"
End Sub

'====================================================================================
' SECTION 4 : GESTION DES NOMENCLATURES (BOM) (POUR UN FLOC UNIQUE)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : FLOCBOMIB03
' DESCRIPTION : Ouvre IL03, lit l'équipement installé, puis exécute IB03
'               pour afficher la nomenclature (BOM) associée.
' DÉPENDANCES : onSAP, Z_IL03, Z_IB03, Z_Enter, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub FLOCBOMIB03()
    ExecuteFlocBOM "Z_IB03", "FLOCBOMIB03"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : floCBOMIB02
' DESCRIPTION : Ouvre IL03 pour lire l'équipement, puis exécute IB02
'               pour modifier la nomenclature correspondante.
' DÉPENDANCES : onSAP, Z_IL03, Z_IB02, Z_Enter, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub floCBOMIB02()
    ExecuteFlocBOM "Z_IB02", "floCBOMIB02"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocOMIB01
' DESCRIPTION : Ouvre IL03, lit l'équipement, puis exécute IB01 pour créer une nomenclature (BOM).
' DÉPENDANCES : onSAP, Z_IL03, Z_IB01, Z_Enter, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocOMIB01()
    ExecuteFlocBOM "Z_IB01", "flocOMIB01"
End Sub


'====================================================================================
' SECTION 5 : GESTION DES LISTES DE TÂCHES (TASK LISTS) (POUR UN FLOC UNIQUE)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : flocIA13
' DESCRIPTION : Ouvre la transaction IA13 (Afficher la liste de tâches).
' DÉPENDANCES : onSAP, Z_IA13, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocIA13()
    ExecuteSingleFlocAction "Z_IA13", "flocIA13"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocIA12
' DESCRIPTION : Ouvre IA12 (Modifier la liste de tâches) et active le profil "PM tasklist".
' DÉPENDANCES : onSAP, Z_IA12, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocIA12()
    On Error GoTo SapErrorHandler

    Run ("onSAP")
    Run ("Z_IA12")
    Run ("Z_Enter")

    ' Sélection du profil "PM tasklist"
    g_Session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[0]").Select
    g_Session.findById("wnd[1]/usr/ctxtTCA41-GRPRF_GRUP").text = "PM"
    g_Session.findById("wnd[1]/usr/ctxtTCA41-GRPRF_NAME").text = "tasklist"
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "flocIA12", Err
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : flocIA11
' DESCRIPTION : Crée une nouvelle liste de tâches (Task List) via la transaction IA11.
' DÉPENDANCES : onSAP, Z_IA11, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub flocIA11()
    On Error GoTo SapErrorHandler

    Run ("onSAP")
    Run ("Z_IA11")
    Run ("Z_Enter")

    ' Activation du profil "PM tasklist"
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRCR01-WERKS").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[0]/usr/ctxtPLKOD-VERWE").text = "4"
    g_Session.findById("wnd[0]/usr/ctxtPLKOD-STATU").text = "4"
    
    MsgBox "remplisser ele champ Wok Center et cliquer sur Enter.", vbOK, "Champ manquant"

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "flocIA11", Err
End Sub


'====================================================================================
' SECTION 6 : TRANSACTIONS EN MASSE (POUR LISTES DE FLOCs)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IA17_F
' DESCRIPTION : Liste les listes de tâches via IA17 pour une sélection d'emplacements fonctionnels.
' DÉPENDANCES : onSAP, Z_IA17, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IA17_F()
    ExecuteIA17FlocList "Z_IA17_F", "wnd[0]/usr/radPN_IFLO"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : Z_IA17_F_G
' DESCRIPTION : Liste les listes de tâches générales (General Task Lists) via IA17
'               pour une liste de FLocs.
' DÉPENDANCES : onSAP, Z_IA17, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub Z_IA17_F_G()
    ExecuteIA17FlocList "Z_IA17_F_G", "wnd[0]/usr/radPN_IHAN"
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : ExecuteListTransaction (Privée)
' DESCRIPTION : Procédure générique pour exécuter une transaction en lot en important
'               une liste de FLocs depuis la sélection Excel.
' PARAMÈTRES  :
'   - sapCommand (String) : Nom de la procédure VBA qui lance la transaction SAP.
'   - optionalFlags (Variant, Optional) : Tableau d'ID de checkboxes à cocher.
'------------------------------------------------------------------------------------
Private Sub ExecuteListTransaction(ByVal sapCommand As String, Optional ByVal optionalFlags As Variant)
    On Error GoTo SapErrorHandler

    Run ("onSAP")
    Run (sapCommand)

    g_Session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_STRNO_%_APP_%-VALU_PUSH", Selection

    ' Application des options spécifiques (flags) si fournies
    If Not IsMissing(optionalFlags) Then
        Dim flag As Variant
        For Each flag In optionalFlags
            g_Session.findById(flag).Selected = True
        Next flag
    End If

    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "ExecuteListTransaction (" & sapCommand & ")", Err
End Sub


'------------------------------------------------------------------------------------
' Procédures publiques appelant la routine générique ExecuteListTransaction
'------------------------------------------------------------------------------------
Sub Z_IH08_F(): ExecuteListTransaction "Z_IH08": End Sub
Sub Z_IH06_F(): ExecuteListTransaction "Z_IH06": End Sub
Sub Z_IL07_F(): ExecuteListTransaction "Z_IL07": End Sub
Sub Z_IW39_F(): ExecuteListTransaction "Z_IW39": End Sub
Sub Z_IW39_FC(): ExecuteListTransaction "Z_IW39H": End Sub
Sub Z_IW38_F(): ExecuteListTransaction "Z_IW38": End Sub
Sub Z_IW29_F(): ExecuteListTransaction "Z_IW29": End Sub
Sub Z_IW29_FC(): ExecuteListTransaction "Z_IW29Hist": End Sub
Sub Z_IW28_F(): ExecuteListTransaction "Z_IW28": End Sub

Sub Z_IP18_F(): ExecuteListTransaction "Z_IP18_All", Array("wnd[0]/usr/chkSPERRE", "wnd[0]/usr/chkOBLIS"): End Sub
Sub Z_IP24_F(): ExecuteListTransaction "Z_IP24_NoOL", Array("wnd[0]/usr/chkOBLIS"): End Sub
