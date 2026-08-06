Attribute VB_Name = "materials"
'==============================================================
' MODULE      : materials
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DATE    : [Date d?application]
'
' DESCRIPTION :
'   Ce module regroupe les proc?dures d?acc?s rapide aux ?crans
'   SAP li?s aux stocks et aux donn?es de base des mat?riaux :
'       - MMBE : Vue stock par type
'       - MM03 : Fiche info mat?riel
'       - ME2N : Liste des commandes d'achat.
'
' STRUCTURE GÉNÉRALE :
'       1. Initialisation de la connexion SAP
'       2. Lancement de la transaction
'       3. Actions sur l'interface (sélection, navigation)
'       4. Fermeture de la session SAP
'
' DÉPENDANCES :
'   - onSAP, offSAP : Gestion de la connexion SAP.
'   - g_Session : Objet de session SAP GUI active.
'   - IsSAPConnectionAlive : Vérifie l'état de la connexion SAP.
'==============================================================

Option Explicit

'================================================================================
' SECTION 0 : PROCÉDURES UTILITAIRES PRIVÉES
'================================================================================

Private Sub ExecuteSimpleMaterialTransaction(ByVal transactionWrapper As String, ByVal procName As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

Private Sub ExecuteSingleMaterialReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldId As String)
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

Private Sub ExecuteListMaterialReport(ByVal transactionWrapper As String, ByVal procName As String, ByVal fieldToClear As String, ByVal buttonToFill As String)
    On Error GoTo SapErrorHandler
    Run ("onSAP")
    Run (transactionWrapper)
    g_Session.findById(fieldToClear).text = ""
    FillSAPSelectionList buttonToFill, Selection
    Run ("Z_F8")
    Run ("offSAP")
    Exit Sub
SapErrorHandler:
    DisplayAndLogError procName, Err
End Sub

'==============================================================
' PROC?DURE : stock
' DESCRIPTION : Ouvre la transaction Z_MMBE pour afficher la disponibilité
'               des stocks d'un matériau.
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub stock()
    ExecuteSimpleMaterialTransaction "Z_MMBE", "stock"
End Sub


'==============================================================
' DESCRIPTION : Affiche les commandes d'achat liées au matériau actif (ME2N).
' DÉPENDANCES : onSAP, Z_ME2N, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matOrders()
    ExecuteSimpleMaterialTransaction "Z_ME2N", "matOrders"
End Sub


'==============================================================
' DESCRIPTION : Ouvre la fiche technique du matériau sélectionné "MRP1"
'               à partir de la vue de stock MMBE (MM03).
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matlt()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MMBE")
    Run ("Z_F8")

    ' S?lection de la premi?re ligne du tableau stock
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1", "C          1"

    ' Ouverture du menu contextuel -> transaction MM03
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").itemContextMenu "          1", "C          1"
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").selectContextMenuItem "MM03"

    ' S?lection de l?onglet "Achats"
    g_Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").Select ' Onglet Achats
    g_Session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = GetSetting("SAP_PLANT_MF") ' Utilise le paramètre du dictionnaire
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    On Error GoTo 0
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "matlt", Err
End Sub


'==============================================================
' DESCRIPTION : Accède à la vue de consommation du matériau dans MM03.
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matcons()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MMBE")
    Run ("Z_F8")

    On Error Resume Next
    ' S?lection de la 4e ligne du tableau MMBE
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").selectItem "          4", "C          1"
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          4", "C          1"

    ' Ouverture du menu contextuel
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").itemContextMenu "          4", "C          1"
    g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]").selectContextMenuItem "MM03"

    ' Navigation vers l?onglet des consommations
    g_Session.findById("wnd[0]/tbar[1]/btn[30]").press
    g_Session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU08").Select
    On Error GoTo 0

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "matcons", Err
End Sub


'==============================================================
' DESCRIPTION : Affiche les données de base du matériau (MM03),
'               en déroulant l'arborescence de MMBE.
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub basdata()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MMBE")
    Run ("Z_F8")

    ' Fermeture et r?ouverture de la hi?rarchie
    With g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]")
        .itemContextMenu "          1", "&Hierarchy"
        .selectContextMenuItem "&COLLAPSE"
        .selectItem "          1", "&Hierarchy"
        .ensureVisibleHorizontalItem "          1", "&Hierarchy"
        .itemContextMenu "          1", "&Hierarchy"
        .selectContextMenuItem "&EXPAND"
        .selectItem "          1", "C          1"
        .ensureVisibleHorizontalItem "          1", "C          1"
        .itemContextMenu "          1", "C          1"
        .selectContextMenuItem "MM03"
    End With
    On Error GoTo 0

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "basdata", Err
End Sub


'==============================================================
' DESCRIPTION : Variante simplifiée d'accès aux données de base (onglet SP02)
'               à partir de la vue stock MMBE.
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub basdata2()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MMBE")
    Run ("Z_F8")

    ' S?lection directe du premier ?l?ment
    With g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]")
        .selectItem "          1", "C          1"
        .ensureVisibleHorizontalItem "          1", "C          1"
        .itemContextMenu "          1", "C          1"
        .selectContextMenuItem "MM03"
    End With

    ' Acc?s ? l?onglet Donn?es de base
    g_Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02").Select

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "basdata2", Err
End Sub

'==============================================================
' DESCRIPTION : Accède à la vue des documents associés à un matériau.
' DÉPENDANCES : onSAP, Z_MMBE, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub docdata()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MMBE")
    Run ("Z_F8")

    On Error Resume Next
    ' S?lection du premier mat?riau et ouverture de MM03
    With g_Session.findById("wnd[0]/usr/cntlCC_CONTAINER/shellcont/shell/shellcont[1]/shell[1]")
        .selectItem "          1", "C          1"
        .ensureVisibleHorizontalItem "          1", "C          1"
        .itemContextMenu "          1", "C          1"
        .selectContextMenuItem "MM03"
    End With

    ' Navigation vers l?onglet des documents
    g_Session.findById("wnd[0]/tbar[1]/btn[30]").press
    On Error GoTo 0

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "docdata", Err
End Sub


'==============================================================
' DESCRIPTION : Affiche les mouvements de stock du matériau actif (MB51).
' DÉPENDANCES : onSAP, Z_MB51, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matmov()
    ExecuteSimpleMaterialTransaction "Z_MB51", "matmov"
End Sub


'==============================================================
' DESCRIPTION : Liste les réservations du matériau actif (MB24).
' DÉPENDANCES : onSAP, Z_MB24, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub maatreservations()
    ExecuteSimpleMaterialTransaction "Z_MB24", "maatreservations"
End Sub


'==============================================================
' DESCRIPTION : Affiche la liste des stocks par entrepôt pour un matériau (MB52).
' DÉPENDANCES : onSAP, Z_MB52, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matmb52()
    ExecuteSingleMaterialReport "Z_MB52", "matmb52", "wnd[0]/usr/ctxtMATNR-LOW"
End Sub


'==============================================================
' DESCRIPTION : Accède aux notes internes d'un matériau (MB53).
' DÉPENDANCES : onSAP, Z_MB53, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matnotes()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_MB53")
    Run ("Z_F8")

    On Error Resume Next
    ' Affiche les notes via le menu contextuel SAP
    g_Session.findById("wnd[0]/tbar[1]/btn[44]").press
    g_Session.findById("wnd[0]/usr/subINCLUDE8XX:SAPMM61R:0800/btnRM61R-MNTXT").press
    On Error GoTo 0
   

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "matnotes", Err
End Sub

'====================================================================================
' SECTION 3 : STRUCTURE ET ANALYSE DES NOMENCLATURES (BOM)
'====================================================================================
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE    : [Date d?application]
'
' DESCRIPTION :
'   Ce bloc permet d?analyser la structure des mat?riaux
'   et d?explorer leurs liens avec les ?quipements et
'   emplacements fonctionnels via les transactions SAP :
'       - CC04 : Visualisation de structure
'       - CS15 : Explosion de nomenclature
'
'==============================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : mattied
' DESCRIPTION : Ouvre la transaction SAP CC04 pour afficher la structure des
'               matériaux et développe les nœuds principaux.
' DÉPENDANCES : onSAP, Z_CC04, offSAP, SAP.IsSAPConnectionAlive.
'
' S?QUENCE :
'   1. Connexion SAP
'   2?? Lancement de CC04
'   3?? Navigation dans l?arborescence
'   4?? Fermeture propre de session
'==============================================================
Sub mattied()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_cc04")
    Run ("Z_Enter")

    With g_Session.findById("wnd[0]/usr/cntlCNTL_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell[1]")
        ' D?veloppement de certains n?uds de structure
        .expandNode "          3"
        .topNode = "          1"
        .expandNode "          5"
        .topNode = "          1"
    End With

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "mattied", Err
End Sub


'==============================================================
'##############################################################
'##############################################################
'################## continue testes ###########################
'==============================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : mattiedcs
' DESCRIPTION : Effectue une explosion de nomenclature via la transaction CS15.
'               Si l'utilisateur le souhaite, récupère aussi les informations
'               des équipements liés à la nomenclature via IH08.
' DÉPENDANCES : onSAP, Z_CS15, Z_IH08, offSAP, SAP.IsSAPConnectionAlive.
'
' PARAM?TRES / INTERACTIONS :
'   - Demande utilisateur via MsgBox (Oui/Non)
'   - Appelle la transaction IH08 si choix Oui
'==============================================================
Sub mattiedcs()
    On Error GoTo SapErrorHandler
    Dim floctrans As VbMsgBoxResult
    Dim coll1 As Variant
    
    '==== 1?? Demande utilisateur ====
    floctrans = MsgBox( _
        "Do you want to see the Functional Location information? " & _
        "If yes, the results are limited to less than 200 equipment, " & _
        "not including material BOMS.", vbYesNo)

    '==== 2?? Lancement SAP et transaction CS15 ====
    Run ("onsap")
    Run ("Z_CS15")
    Run ("Z_F8")

    '==== 3?? Extraction conditionnelle ====
    On Error Resume Next
    If floctrans = vbYes Then
        
        'Add all feild
        g_Session.findById("wnd[0]/tbar[1]/btn[32]").press
        g_Session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 130
        g_Session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").FirstVisibleRow = 120
        g_Session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0-130"
        g_Session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
        g_Session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        'optimise colomn and let equipement first colomn
        g_Session.findById("wnd[0]/mbar/menu[4]/menu[4]/menu[0]").Select
        g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "EQUNR"
        
        ' S?lection des ?quipements dans la grille de r?sultat
        With g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
            .selectColumn "EQUNR"
            .contextMenu
            .selectContextMenuItemByPosition "0"
        End With
        
        If Selection Is Null Then
                floctrans = MsgBox( _
            "No Equipements Linked", vbYes)
            Exit Sub
        End If
        
        ' Lancement de la transaction IH08
        Run ("Z_IH08")

        ' Navigation dans les fen?tres SAP
        g_Session.findById("wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH").press
        g_Session.findById("wnd[1]/tbar[0]/btn[24]").press
        g_Session.findById("wnd[1]/tbar[0]/btn[8]").press
        g_Session.findById("wnd[0]/usr/ctxtEQUNR-HIGH").SetFocus

        ' Ex?cution
        Run ("Z_Enter")
        Run ("Z_F8")

    End If
    On Error GoTo 0

    '==== 4?? Fermeture propre ====
    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "mattiedcs", Err
End Sub

'====================================================================================
' SECTION 4 : LIENS ENTRE MATÉRIAUX ET ORDRES DE MAINTENANCE
' AUTEUR  : [Ton Nom]
' DATE    : [Date d?application]
'
' DESCRIPTION :
'   Ce bloc regroupe les macros permettant d?afficher,
'   depuis un mat?riau donn?, les ordres ou historiques
'   de maintenance associ?s dans SAP :
'
'       - IW39 : Liste des ordres de maintenance
'       - IW38 : Ordres de travail
'       - IWBK : Historique des notifications li?es
'
'==============================================================
'Option Explicit


'==============================================================
' DESCRIPTION : Ouvre la transaction SAP IW39 pour afficher les ordres
'   de maintenance associ?s ? un mat?riau donn?.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'
' D?TAIL :
'   - Le code mat?riau est lu dans la cellule active.
'   - Les filtres usine sont pris depuis la feuille ?Setup?.
'==============================================================
Sub IW39Mat()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IW39")

    '--- Affectation du mat?riau s?lectionn? ---
    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = Cells(ActiveCell.row, ActiveCell.Column)

    '--- Application des filtres d?usine depuis Setup ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")

    '--- Ex?cution de la recherche ---
    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IW39Mat", Err
End Sub


'==============================================================
' DESCRIPTION : Exécute la transaction SAP IW38 pour afficher les ordres
'   de travail li?s ? un mat?riau sp?cifique.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'
' D?TAIL :
'   - Lecture de la cellule active.
'   - Filtrage usine depuis ?Setup?.
'==============================================================
Sub IW38Mat()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IW38")


    '--- Affectation du mat?riau ---
    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = Cells(ActiveCell.row, ActiveCell.Column)

    '--- Filtres d?usine optionnels ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")

    '--- Ex?cution ---
    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IW38Mat", Err
End Sub


'==============================================================
' DESCRIPTION : Lance la transaction IWBK pour consulter les notifications
'   et historiques associ?s ? un mat?riau donn?.
' DÉPENDANCES : onSAP, Z_IWBK, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'
'==============================================================
Sub IWBKmat()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IWBK")


    '--- Affectation du code mat?riau ---
    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = Cells(ActiveCell.row, ActiveCell.Column)

    '--- Application des filtres ?Setup? si d?finis ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_PT") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = GetSetting("SAP_PLANT_PT")

    '--- Ex?cution ---
    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IWBKmat", Err
End Sub

'==============================================================
' MODULE : Material_Stock
' BLOC 5  : Transactions multi-mat?riaux (suffixe _L)
' AUTEUR  : [Ton Nom]
' DATE    : [Date d?application]
'
' DESCRIPTION :
'   Ce bloc regroupe toutes les macros permettant
'   d?ex?cuter des transactions SAP sur plusieurs
'   mat?riaux s?lectionn?s dans Excel.
'
'   Transactions couvertes :
'       - MB52  : Stocks
'       - ME2N  : Commandes d?achat
'       - MB51  : Mouvements de stocks
'       - MB24  : R?servations
'       - IW39 / IW38 / IWBK : Ordres & historiques maintenance
'
'   Chaque macro suit la m?me s?quence :
'       1. Ouverture de SAP et de la transaction.
'       2. Nettoyage du champ principal (MATNR ou SERMAT).
'       3. Copie de la s?lection Excel dans SAP.
'       4. Validation des valeurs s?lectionn?es.
'       5. Ex?cution (F8) et fermeture de SAP.
'==============================================================

'--------------- LIST ----------------------

'==============================================================
' PROC?DURE : matmb52L
' DESCRIPTION :
'   Affiche les stocks de plusieurs mat?riaux
'   dans la transaction SAP MB52.
'==============================================================
Sub matmb52L()
    ExecuteListMaterialReport "Z_MB52", "matmb52L", "wnd[0]/usr/ctxtMATNR-LOW", "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
End Sub


'==============================================================
' DESCRIPTION : Ouvre la transaction ME2N pour consulter les
'   commandes d?achat li?es ? une liste de mat?riaux.
' DÉPENDANCES : onSAP, Z_ME2N, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matOrdersL()
    ExecuteListMaterialReport "Z_ME2N", "matOrdersL", "wnd[0]/usr/ctxtS_MATNR-LOW", "wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH"
End Sub


'==============================================================
' DESCRIPTION : Affiche les mouvements de stock pour plusieurs
'   mat?riaux via la transaction MB51.
' DÉPENDANCES : onSAP, Z_MB51, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub matmovL()
    ExecuteListMaterialReport "Z_MB51", "matmovL", "wnd[0]/usr/ctxtMATNR-LOW", "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
End Sub


'==============================================================
' DESCRIPTION : Liste toutes les réservations d'un ensemble
'   de mat?riaux via la transaction MB24.
' DÉPENDANCES : onSAP, Z_MB24, Z_F8, offSAP, SAP.IsSAPConnectionAlive.
'==============================================================
Sub maatreservationsL()
    ExecuteListMaterialReport "Z_MB24", "maatreservationsL", "wnd[0]/usr/ctxtMATNR-LOW", "wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH"
End Sub


'==============================================================
' DESCRIPTION : Exécute la transaction IW39 pour une liste de
'   mat?riaux s?lectionn?s dans Excel.
' DÉPENDANCES : onSAP, Z_IW39, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'==============================================================
Sub IW39MatL()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IW39")


    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_SERMAT_%_APP_%-VALU_PUSH", Selection

    '--- Filtres suppl?mentaires depuis ?Setup? ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_PT") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = GetSetting("SAP_PLANT_PT")

    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IW39MatL", Err
End Sub


'==============================================================
' DESCRIPTION : Ouvre la transaction IW38 pour plusieurs
'   mat?riaux s?lectionn?s.
' DÉPENDANCES : onSAP, Z_IW38, Z_F8, offSAP, GetSetting, SAP.IsSAPConnectionAlive.
'==============================================================
Sub IW38MatL()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IW38")


    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_SERMAT_%_APP_%-VALU_PUSH", Selection

    '--- Filtres suppl?mentaires ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = GetSetting("SAP_PLANT_MF")

    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IW38MatL", Err
End Sub


'==============================================================
' PROC?DURE : IWBKmatL
' DESCRIPTION :
'   Ex?cute la transaction IWBK sur plusieurs
'   mat?riaux list?s dans la s?lection Excel.
'==============================================================
Sub IWBKmatL()
    On Error GoTo SapErrorHandler
    Run ("onsap")
    Run ("Z_IWBK")

    g_Session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = ""
    FillSAPSelectionList "wnd[0]/usr/btn%_SERMAT_%_APP_%-VALU_PUSH", Selection

    '--- Filtres d?usine optionnels ---
    If GetSetting("SAP_PLANT_PF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = GetSetting("SAP_PLANT_PF")
    If GetSetting("SAP_PLANT_MF") <> "" Then g_Session.findById("wnd[0]/usr/ctxtIWERK-HIGH").text = GetSetting("SAP_PLANT_MF")

    Run ("Z_F8")

    Run ("offSAP")
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "IWBKmatL", Err
End Sub

