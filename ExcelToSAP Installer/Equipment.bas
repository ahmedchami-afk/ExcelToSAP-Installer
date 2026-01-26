Attribute VB_Name = "Equipment"
Option Explicit
'==============================================================
' MODULE      : Equipment
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module regroupe toutes les procédures d'interaction SAP pour les
'               équipements techniques (Master Data). Chaque procédure est autonome
'               et gère le cycle de vie complet de la connexion SAP (ouverture,
'               exécution, fermeture).
'
'               Les procédures sont organisées par type d'action :
'                 - Consultation d'informations (IE03).
'                 - Affichage de la nomenclature (BOM).
'                 - Consultation des ordres de maintenance (PMO) et notifications.
'                 - Consultation des plans de maintenance (IP18, IP24).
'                 - Création/Modification de nomenclatures (IB01, IB02).
'                 - Gestion des listes de tâches (IA01, IA02, IA03).
'                 - Procédures pour listes d'équipements (suffixe _L ou _E).
'==============================================================

'================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'================================================================================

Private Sub ExecuteSingleEquipAction(ByVal transactionWrapper As String, ByVal procName As String, Optional ByVal actionId As String = "", Optional ByVal actionType As String = "ButtonPress")
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_Enter")
    If actionId <> "" Then
        If actionType = "ButtonPress" Then g_Session.findById(actionId).press
    End If
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSingleEquipReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToFill As String, Optional ByVal checkboxesToSelect As Variant, Optional ByVal checkboxesToDeselect As Variant)
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

Private Sub ExecuteListEquipAction(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String, Optional ByVal clearDates As Boolean = False, Optional ByVal checkboxesToSelect As Variant)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    
    g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    
    If clearDates Then
        g_Session.findById("wnd[0]/usr/ctxtDATUV").text = ""
        g_Session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    End If
    
    If Not IsMissing(checkboxesToSelect) Then
        Dim chk As Variant
        For Each chk In checkboxesToSelect
            g_Session.findById(chk).Selected = True
        Next chk
    End If
    
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteBOMAction(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    g_Session.findById("wnd[0]/usr/ctxtRC29N-EQUNR").text = ActiveCell.value
    
    If GetSetting("SAP_PLANT_PF") <> "" Then
        g_Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = GetSetting("SAP_PLANT_PF")
    Else
        If CheckSAPError(transactionWrapper) Then GoTo CleanExit
    End If
    
    g_Session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "4"
    Run ("Z_Enter")
    Run ("offSAP")
    Exit Sub
CleanExit:
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteIA17Single(ByVal procName As String, ByVal radioButtonId As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run ("Z_IA17")
    
    g_Session.findById("wnd[0]/usr/ctxtPN_EQUNR-LOW").text = ActiveCell.value
    g_Session.findById(radioButtonId).Select
    
    g_Session.findById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = "*"
    g_Session.findById("wnd[0]/usr/txtPN_PLNAL-LOW").text = "*"
    g_Session.findById("wnd[0]/usr/ctxtPN_DATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtPN_VAGRP-LOW").text = "*"
    
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'================================================================================
' SECTION 1 : CONSULTATION ÉQUIPEMENT (IE03)
'================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : equipmentInfo
' DESCRIPTION : Lance la transaction IE03 (Afficher équipement) et appuie sur "Entrée".
' CONTEXTE    : Accès rapide à l'écran de base de la transaction.
' DÉPENDANCES : onSAP, Z_IE03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub equipmentInfo()
    ExecuteSingleEquipAction "Z_IE03", "equipmentInfo"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : equipmentBOM
' DESCRIPTION : Lance IE03, saisit le numéro d'équipement de la cellule active,
'               puis navigue vers l'affichage de la nomenclature (BOM).
' CONTEXTE    : Consultation de la nomenclature pour un équipement unique.
' DÉPENDANCES : onSAP, Z_IE03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub equipmentBOM()
    ExecuteSingleEquipAction "Z_IE03", "equipmentBOM", "wnd[0]/tbar[1]/btn[32]"
End Sub


'------------------------------------------------------------------------------------
' PROCÉDURE   : eqchars
' DESCRIPTION : Lance IE03, saisit le numéro d'équipement, puis navigue vers l'écran
'               des caractéristiques techniques et applique un tri.
' CONTEXTE    : Consultation des caractéristiques pour un équipement unique.
' DÉPENDANCES : onSAP, Z_IE03, Z_Enter, offSAP, SAP.IsSAPConnectionAlive.
'------------------------------------------------------------------------------------
Sub eqchars()
    On Error GoTo SapErrorHandler
    Run ("onSAP")

    Run ("Z_IE03")
    Run ("Z_Enter")

    ' Ajoute une pause pour laisser le temps à l'écran principal de se charger
    WaitForSAP

    ' Navigation vers les caractéristiques via le menu "Infos supplémentaires"
    g_Session.findById("wnd[0]/tbar[1]/btn[20]").press
    g_Session.findById("wnd[0]/mbar/menu[4]/menu[5]").Select

    ' Désactive une option dans les paramètres (checkbox MEOHN)
    'g_Session.findById("wnd[1]/usr/tabsPARAM/tabpWERT/ssubSUB:SAPLCLPR:0100/chkRMCLPAR-MEOHN").Selected = False
    'g_Session.findById("wnd[1]/usr/tabsPARAM/tabpWERT/ssubSUB:SAPLCLPR:0100/chkRMCLPAR-MEOHN").SetFocus
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' Ajoute une pause pour laisser le temps à l'écran des caractéristiques de s'afficher
    WaitForSAP

    ' Sélection de la première caractéristique et tri ascendant
    g_Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S").Columns.elementAt(0).Selected = True
    g_Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MNAME[0,0]").SetFocus
    g_Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MNAME[0,0]").caretPosition = 0
    g_Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/btnRCTMS-CHAR_SORTUP").press

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "eqchars", Err
End Sub

'================================================================================
' SECTION 2 : RAPPORTS SUR ÉQUIPEMENT UNIQUE
'================================================================================

'--------------------------------------------------------------
' PROCEDURE   : equipmentPMO
' DESCRIPTION :
'   Execute la transaction IW39 pour afficher les ordres de maintenance
'   lies a l'equipement dont le numero est dans la cellule active.
'--------------------------------------------------------------
Sub equipmentPMO()
    ExecuteSingleEquipReport "Z_IW39", "equipmentPMO", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentPMO_H
' DESCRIPTION :
'   Variante historique de IW39 : filtre sur les ordres termines ou historiques pour l'equipement.
'--------------------------------------------------------------
Sub equipmentPMO_H()
    ExecuteSingleEquipReport "Z_IW39", "equipmentPMO_H", "wnd[0]/usr/ctxtEQUNR-LOW", Array("wnd[0]/usr/chkDY_MAB", "wnd[0]/usr/chkDY_HIS"), Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR")
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentPMOc
' DESCRIPTION :
'   Variante avec la transaction IW38 : affiche les ordres en cours pour l'equipement actif.
'--------------------------------------------------------------
Sub equipmentPMOc()
    ExecuteSingleEquipReport "Z_IW38", "equipmentPMOc", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentNotif
' DESCRIPTION :
'   Lance IW29 pour afficher les notifications liees a l'equipement.
'   Le numero d'equipement est lu depuis la cellule active.
'--------------------------------------------------------------
Sub equipmentNotif()
    ExecuteSingleEquipReport "Z_IW29", "equipmentNotif", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentNotifH
' DESCRIPTION :
'   Variante historique de IW29 : affiche les notifications terminees ou historiques.
'--------------------------------------------------------------
Sub equipmentNotifH()
    ExecuteSingleEquipReport "Z_IW29", "equipmentNotifH", "wnd[0]/usr/ctxtEQUNR-LOW", Array("wnd[0]/usr/chkDY_MAB"), Array("wnd[0]/usr/chkDY_OFN", "wnd[0]/usr/chkDY_IAR")
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentNotifc
' DESCRIPTION :
'   Execute IW28 pour voir les notifications en cours pour l'equipement.
'--------------------------------------------------------------
Sub equipmentNotifc()
    ExecuteSingleEquipReport "Z_IW28", "equipmentNotifc", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentplan
' DESCRIPTION :
'   Lance IP18 pour visualiser les plans de maintenance lies a l'equipement.
'   Recoit le numero d'equipement depuis la cellule active.
'--------------------------------------------------------------
Sub equipmentplan()
    ExecuteSingleEquipReport "Z_IP18", "equipmentplan", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentplan24
' DESCRIPTION :
'   Lance la transaction IP24 pour visualiser la planification de maintenance
'   sans ordres ouverts pour l'equipement actif.
'--------------------------------------------------------------
Sub equipmentplan24()
    ExecuteSingleEquipReport "Z_IP24", "equipmentplan24", "wnd[0]/usr/ctxtEQUNR-LOW"
End Sub

'================================================================================
' SECTION 3 : GESTION DES NOMENCLATURES (BOM)
'================================================================================


'--------------------------------------------------------------
' PROCEDURE   : equipmentBOMIB03
' DESCRIPTION :
'   Lance la transaction IB03 pour afficher la nomenclature (BOM) de l'equipement.
'   Remplit les champs EQUNR et WERKS (usine) depuis la configuration.
'--------------------------------------------------------------
Sub equipmentBOMIB03()
    ExecuteBOMAction "Z_IB03", "equipmentBOMIB03"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentBOMIB02
' DESCRIPTION :
'   Lance IB02 pour modifier une nomenclature (BOM) existante.
'--------------------------------------------------------------
Sub equipmentBOMIB02()
    ExecuteBOMAction "Z_IB02", "equipmentBOMIB02"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentBOMIB01
' DESCRIPTION :
'   Lance IB01 pour creer une nouvelle nomenclature (BOM) pour l'equipement selectionne.
'--------------------------------------------------------------
Sub equipmentBOMIB01()
    ExecuteBOMAction "Z_IB01", "equipmentBOMIB01"
End Sub

'================================================================================
' SECTION 4 : GESTION DES GAMMES (TASK LISTS)
'================================================================================


'--------------------------------------------------------------
' PROCEDURE   : equipmentIA03
' DESCRIPTION :
'   Lance IA03 pour afficher une liste de taches liee a l'equipement actif.
'--------------------------------------------------------------
Sub equipmentIA03()
    ExecuteSingleEquipAction "Z_IA03", "equipmentIA03", "wnd[0]/usr/ctxtRC27E-EQUNR", "FillOnly"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentIA02
' DESCRIPTION :
'   Lance IA02 pour modifier une liste de taches d'equipement existante.
'--------------------------------------------------------------
Sub equipmentIA02()
    ExecuteSingleEquipAction "Z_IA02", "equipmentIA02", "wnd[0]/usr/ctxtRC27E-EQUNR", "FillOnly"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : equipmentIA01
' DESCRIPTION :
'   Lance IA01 pour creer une nouvelle liste de taches d'equipement.
'--------------------------------------------------------------
Sub equipmentIA01()
    On Error GoTo SapErrorHandler
    
    Run ("onSAP")

    Run ("Z_IA01")
    g_Session.findById("wnd[0]/usr/ctxtRC27E-EQUNR").text = ActiveCell.value
    
    Run ("Z_Enter")
    
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtRCR01-WERKS").text = GetSetting("SAP_PLANT_PF")
    g_Session.findById("wnd[0]/usr/ctxtPLKOD-VERWE").text = "4"
    g_Session.findById("wnd[0]/usr/ctxtPLKOD-STATU").text = "4"

    MsgBox "remplisser ele champ Wok Center et cliquer sur Enter.", vbOK, "Champ manquant"

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "equipmentIA01", Err
End Sub

'================================================================================
' SECTION 5 : RAPPORTS SUR LISTES D'ÉQUIPEMENTS
'================================================================================

'--------------------------------------------------------------
' PROCEDURE   : Z_IH08_L
' DESCRIPTION :
'   Lance la transaction IH08 pour rechercher une liste d'equipements
'   importee depuis Excel.
'--------------------------------------------------------------
Sub Z_IH08_L()
    ExecuteListEquipAction "Z_IH08", "Z_IH08_L", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IH08_E
' DESCRIPTION :
'   Variante de Z_IH08_L : importe une liste d'equipements et execute la recherche.
'--------------------------------------------------------------
Sub Z_IH08_E()
    ExecuteListEquipAction "Z_IH08", "Z_IH08_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IE07_E
' DESCRIPTION :
'   Lance IE07 pour afficher une liste d'equipements importee depuis Excel.
'--------------------------------------------------------------
Sub Z_IE07_E()
    ExecuteListEquipAction "Z_IE07", "Z_IE07_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW39_E
' DESCRIPTION :
'   Lance IW39 pour extraire les ordres lies a la liste d'equipements importee.
'   Efface les dates de filtre avant execution.
'--------------------------------------------------------------
Sub Z_IW39_E()
    ExecuteListEquipAction "Z_IW39", "Z_IW39_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH", True
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW39_EC
' DESCRIPTION :
'   Variante avec historique (utilise Z_IW39H) et effacement des dates pour une liste d'equipements.
'--------------------------------------------------------------
Sub Z_IW39_EC()
    ExecuteListEquipAction "Z_IW39H", "Z_IW39_EC", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH", True
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW38_E
' DESCRIPTION :
'   Lance IW38 pour afficher la liste des ordres pour les equipements importes.
'   Efface les dates de filtre avant execution.
'--------------------------------------------------------------
Sub Z_IW38_E()
    ExecuteListEquipAction "Z_IW38", "Z_IW38_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH", True
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW29_E
' DESCRIPTION :
'   Lance IW29 pour une liste d'equipements importee depuis Excel.
'--------------------------------------------------------------
Sub Z_IW29_E()
    ExecuteListEquipAction "Z_IW29", "Z_IW29_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW29_EC
' DESCRIPTION :
'   Lance IW29 en mode historique (utilise Z_IW29Hist) avec une liste d'equipements.
'--------------------------------------------------------------
Sub Z_IW29_EC()
    ExecuteListEquipAction "Z_IW29Hist", "Z_IW29_EC", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IW28_E
' DESCRIPTION :
'   Lance IW28 pour consulter les notifications en cours via
'   une liste d'equipements importee depuis Excel.
'--------------------------------------------------------------
Sub Z_IW28_E()
    ExecuteListEquipAction "Z_IW28", "Z_IW28_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH"
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IP18_E
' DESCRIPTION :
'   Lance IP18_All pour consulter les plans de maintenance
'   d'une liste d'equipements avec les filtres SPERRE et OBLIS.
'--------------------------------------------------------------
Sub Z_IP18_E()
    ExecuteListEquipAction "Z_IP18_All", "Z_IP18_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH", False, Array("wnd[0]/usr/chkSPERRE", "wnd[0]/usr/chkOBLIS")
End Sub


'--------------------------------------------------------------
' PROCEDURE   : Z_IP24_E
' DESCRIPTION :
'   Lance IP24_NoOL pour afficher les planifications sans ordres ouverts.
'   Importe une liste d'equipements et coche le filtre OBLIS.
'--------------------------------------------------------------
Sub Z_IP24_E()
    ExecuteListEquipAction "Z_IP24_NoOL", "Z_IP24_E", "wnd[0]/usr/ctxtEQUNR-LOW", "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH", False, Array("wnd[0]/usr/chkOBLIS")
End Sub

Private Sub ExecuteIA17List(ByVal procName As String, ByVal radioButtonId As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run ("Z_IA17")
    g_Session.findById("wnd[0]/usr/ctxtPN_EQUNR-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_PN_EQUNR_%_APP_%-VALU_PUSH", Selection
    
    g_Session.findById(radioButtonId).Select
    g_Session.findById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = "*"
    g_Session.findById("wnd[0]/usr/txtPN_PLNAL-LOW").text = "*"
    g_Session.findById("wnd[0]/usr/ctxtPN_DATUV").text = ""
    g_Session.findById("wnd[0]/usr/ctxtPN_VAGRP-LOW").text = "*"
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'--------------------------------------------------------------
' PROCEDURE   : Z_IA17_E
' DESCRIPTION :
'   Lance IA17 pour consulter les listes de taches des equipements importes.
'   Colle la liste depuis Excel et selectionne le mode "Equipement".
'--------------------------------------------------------------
Sub Z_IA17_E()
    ExecuteIA17List "Z_IA17_E", "wnd[0]/usr/radPN_EQUI"
End Sub

'--------------------------------------------------------------
' PROCEDURE   : Z_IA17_E_G
' DESCRIPTION :
'   Lance IA17 pour consulter les listes de taches generales (General Task Lists).
'   Colle la liste d'equipements depuis Excel et execute la recherche.
'--------------------------------------------------------------
Sub Z_IA17_E_G()
    ExecuteIA17List "Z_IA17_E_G", "wnd[0]/usr/radPN_IHAN"
End Sub

'--------------------------------------------------------------
' PROCEDURE   : Z_IA17_E_G_Single
' DESCRIPTION :
'   Lance IA17 pour consulter les listes de taches generales (General Task List)
'   pour un equipement unique.
'   Le numero d'equipement est lu depuis la cellule active.
'--------------------------------------------------------------
Sub Z_IA17_E_G_Single()
    ExecuteIA17Single "Z_IA17_E_G_Single", "wnd[0]/usr/radPN_IHAN"
End Sub

'==============================================================
' Fin du module Equipment
'==============================================================
