Attribute VB_Name = "modMain"
'====================================================================================
' MODULE      : modMain
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module contient la logique principale d'exécution de l'add-in ExcelToSAP.
'               Il agit comme un contrôleur qui :
'                 1. Détecte le type de donnée SAP sélectionnée par l'utilisateur.
'                 2. Identifie si la sélection est une cellule unique ou une liste.
'                 3. Appelle dynamiquement le formulaire (UserForm) approprié.
'====================================================================================
Const CUSTOM_MENU_TAG As String = "ExcelToSAP_CustomMenu"
Public g_DetectedTypes As Object
'====================================================================================
' SECTION 1 : POINT D'ENTRÉE PRINCIPAL
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Main2
' DESCRIPTION :
'   Point d'entrée principal de l'application, généralement appelé par un raccourci
'   clavier. Cette procédure prépare les variables globales, détecte le type de
'   donnée sélectionné, et appelle le formulaire approprié.
' DÉPENDANCES : LogMessage, HandleError, Detect_g_LastRow, Detect_g_DataType,
'               GetSetting, UpdateRibbonState, Handle_Form_Call.
'------------------------------------------------------------------------------------
Public Sub Main2()

    On Error GoTo ErrHandler
    LogMessage "D?marrage de Main2."

    ' ---------------------------------------------------------------------------
    ' ?? Initialisation des variables
    ' ---------------------------------------------------------------------------'
    ' Réinitialise les variables d'état globales pour cette nouvelle exécution.
    
    LoadConfiguration
    
    g_DataType = ""
    g_FirstRow = Selection.row
    Call Detect_g_LastRow

    ' Détermine si la sélection est une cellule unique ou une liste.
    Dim isSingleCell As Boolean
    isSingleCell = (g_LastRow = g_FirstRow)

    LogMessage "Sélection détectée : " & IIf(isSingleCell, "Cellule unique", "Liste de valeurs")

    ' ---------------------------------------------------------------------------
    ' ?? Détection du type de données
    ' ---------------------------------------------------------------------------'
    If g_DataType = "" Then Call Detect_g_DataType

    ' Réinitialise le flag pour l'exécution SAP.
    g_DoNotRun = False
    'g_DataType = g_DataType
    ' ---------------------------------------------------------------------------
    ' ?? Appel automatique du formulaire adapté
    ' ---------------------------------------------------------------------------'
    ' Met à jour l'état du ruban avant d'afficher le formulaire.
    Call UpdateRibbonState(isSingleCell)
    
    LogMessage "Main2 terminé avec succès."
    Exit Sub

ErrHandler:
    DisplayAndLogError "Main2", Err
End Sub



'====================================================================================
' SECTION 2 : LOGIQUE DE DÉTECTION
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : Detect_g_DataType
' DESCRIPTION :
'   Analyse la valeur de la cellule sélectionnée pour déterminer le type de
'   donnée SAP (par exemple, Equipement, FLOC, Notification).
'
'   ATTENTION : Cette fonction est complexe et s'appuie sur des règles lues
'   directement depuis la feuille "Setup". Pour une meilleure maintenance,
'   elle devrait être réécrite pour utiliser exclusivement le dictionnaire
'   `ConfigSettings` chargé par `LoadConfiguration`.
' DÉPENDANCES : LogMessage, HandleError, GetSetting.
'------------------------------------------------------------------------------------
Public Sub Detect_g_DataType()

    On Error GoTo ErrHandler
    LogMessage "Détection du type de donnée..."

    Dim selCol As Long
    Dim gto As Variant
    
    selCol = Selection.Column
    
    If IsError(Cells(g_FirstRow, selCol).value) Then
        LogMessage "La cellule sélectionnée contient une erreur. Détection annulée."
        Exit Sub
    End If
    
    gto = Cells(g_FirstRow, selCol).value

    ' Initialisation
    g_DataType = ""
    
    ' --- Gestion des conflits : Collecte des types correspondants ---
    Dim detectedTypes As Object
    Set detectedTypes = CreateObject("Scripting.Dictionary")

    ' S'assurer que la configuration est chargée
    If ConfigSettings Is Nothing Then LoadConfiguration
    
    ' Vérification séquentielle des plages définies dans le Setup
    ' 1. Equipment
    If CheckRange("EQUI MIN", "EQUI MAX", gto) Then detectedTypes.add "Equip", "Equipement"
    
    ' 2. Functional Location (FLOC)
    If CheckRange("FLOC MIN", "FLOC MAX", gto) Then detectedTypes.add "FLOC", "Functional Location"
    
    ' 3. Material
    If CheckRange("Material MIN", "Material MAX", gto) Then detectedTypes.add "Material", "Material"
    
    ' 4. Maintenance Plan
    If CheckRange("M.Plan MIN", "M.Plan MAX", gto) Then detectedTypes.add "M.Plan", "Maintenance Plan"
    
    ' 5. Notification
    If CheckRange("Notification MIN", "Notification MAX", gto) Then detectedTypes.add "Notif", "MR"
    
    ' 6. PM Order
    If CheckRange("PM Order MIN", "PM Order MAX", gto) Then detectedTypes.add "Pmord", "PM Order"
    
    ' 7. Purchase Requisition
    If CheckRange("Purchase Requisition MIN", "Purchase Requisition MAX", gto) Then detectedTypes.add "Requis", "PR"
    
    ' 8. Purchase Order
    If CheckRange("Purchase Order MIN", "Purchase Order MAX", gto) Then detectedTypes.add "PurOrder", "Purchase Order"
    
    ' 9. Network
    If CheckRange("Network MIN", "Network MAX", gto) Then detectedTypes.add "Network", "Network"
    
    ' 10. Project Definition
    If CheckRange("Project Definition MIN", "Project Definition MAX", gto) Then detectedTypes.add "Project", "Project Definition"
    
    ' 11. WBS
    If CheckRange("WBS MIN", "WBS MAX", gto) Then detectedTypes.add "WBS", "WBS Element"

    ' 12. Item
    If CheckRange("Item MIN", "Item MAX", gto) Then detectedTypes.add "Item", "Item Element"

    ' 13. SES
    If CheckRange("SES MIN", "SES MAX", gto) Then detectedTypes.add "SES", "SES Element"

    ' 14. Task List
    'If CheckRange("Task List MIN", "Task List MAX", gto) Then detectedTypes.add "TaskL", "TaskL Element"

    ' 15. Confirmation
    If CheckRange("Confirmation MIN", "Confirmation MAX", gto) Then detectedTypes.add "CNF", "CNF Element"
    
    ' Sauvegarde des types détectés pour le menu contextuel
    Set g_DetectedTypes = detectedTypes

    ' --- Analyse des résultats ---
    If detectedTypes.count = 0 Then
        LogMessage "Aucun type détecté automatiquement."
        ' g_DataType reste vide, ce qui déclenchera SwFRM dans Main2
    ElseIf detectedTypes.count = 1 Then
        Dim vKeys As Variant
        vKeys = detectedTypes.Keys
        g_DataType = vKeys(0)
        LogMessage "Type unique détecté : " & g_DataType
    Else
        ' Conflit détecté
        'LogMessage "Conflit détecté (plusieurs types possibles). La sélection manuelle sera requise."
        ' g_DataType reste vide, ce qui forcera une sélection manuelle dans l'interface.
    End If
    
    Exit Sub

ErrHandler:
    DisplayAndLogError "Detect_g_DataType", Err
End Sub

'------------------------------------------------------------------------------------
' FONCTION    : CheckRange (Privée)
' DESCRIPTION : Vérifie si une valeur se trouve dans une plage définie par deux clés
'               de configuration (MIN et MAX). Gère les comparaisons numériques et texte.
'------------------------------------------------------------------------------------
Private Function CheckRange(keyMin As String, keyMax As String, val As Variant) As Boolean
    
    Dim minVal As Variant
    Dim maxVal As Variant
    
    minVal = GetSetting(keyMin)
    maxVal = GetSetting(keyMax)
    
    ' Si l'une des bornes n'est pas définie, on ignore cette plage
    If minVal = "" Or maxVal = "" Then Exit Function
    
    ' --- Logique de comparaison ---
    ' Si la valeur, min et max sont tous convertibles en nombres purs
    If IsNumeric(val) And IsNumeric(minVal) And IsNumeric(maxVal) Then
        ' --- PARTIE NUMERIQUE ---
        ' Compare la valeur entre les bornes numériques min et max.
        If CDec(val) >= CDec(minVal) And CDec(val) <= CDec(maxVal) Then
            CheckRange = True
        End If
    Else
        ' --- PARTIE TEXTE (logique de validation de format) ---
        Dim valStr As String, minStr As String, maxStr As String
        valStr = CStr(val)
        minStr = CStr(minVal)
        maxStr = CStr(maxVal)
        
        Dim lenVal As Long, lenMin As Long, lenMax As Long
        lenVal = Len(valStr)
        lenMin = Len(minStr)
        lenMax = Len(maxStr)
        
        ' 1. Vérification de la longueur
        If lenVal < lenMin Or lenVal > lenMax Then Exit Function
        
        ' 2. Vérification du préfixe (basé sur minVal)
        If Left(valStr, lenMin) <> minStr Then Exit Function
        
        ' 3. Vérification des caractères littéraux de la borne MAX
        Dim i As Long
        For i = 1 To lenVal
            If i > lenMax Then Exit For ' Ne pas comparer au-delà de la longueur du pattern
            
            If UCase(Mid(maxStr, i, 1)) <> "*" Then
                If Mid(valStr, i, 1) <> Mid(maxStr, i, 1) Then Exit Function
            End If
        Next i
        
        ' Si toutes les vérifications précédentes ont réussi, c'est un match.
        CheckRange = True
    End If
End Function

'------------------------------------------------------------------------------------
' PROCÉDURE   : Detect_g_LastRow
' DESCRIPTION :
'   Identifie la dernière ligne de la sélection actuelle de l'utilisateur dans Excel
'   et met à jour la variable globale `g_LastRow`.
' DÉPENDANCES : HandleError.
'------------------------------------------------------------------------------------
Public Sub Detect_g_LastRow()
    Dim lastCell As Range
    On Error GoTo ErrHandler

    ' Utilise la méthode Find pour trouver la dernière cellule non vide dans la sélection.
    ' C'est plus robuste que de se fier à la dernière ligne de l'objet Selection,
    ' surtout si la sélection contient des lignes vides à la fin.
    Set lastCell = Selection.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    ' Si une cellule non vide est trouvée, sa ligne devient la dernière ligne.
    If Not lastCell Is Nothing Then
        g_LastRow = lastCell.row
    Else
        ' Si la sélection est complètement vide, on considère que la dernière ligne
        ' est la même que la première ligne de la sélection.
        g_LastRow = Selection.row
    End If

    LogMessage "Dernière ligne détectée (robuste) : " & g_LastRow
    Exit Sub

ErrHandler:
    DisplayAndLogError "Detect_g_LastRow", Err
End Sub


'====================================================================================
' SECTION 3 : ROUTAGE ET AFFICHAGE DES FORMULAIRES
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : UpdateRibbonState
' DESCRIPTION :
'   Active ou désactive les menus du ruban en fonction du type de sélection.
'   - Si une seule cellule est sélectionnée, active les menus "Single".
'   - Si une liste est sélectionnée, active les menus "List".
'   Ceci est réalisé en invalidant le ruban pour forcer l'appel des callbacks "getEnabled".
' PARAMÈTRES  : isSingleCell (Boolean) - True si la sélection est une cellule unique.
' DÉPENDANCES : HandleError, g_ribbonUI (objet Ruban).
'------------------------------------------------------------------------------------
Public Sub UpdateRibbonState(ByVal isSingleCell As Boolean)
    On Error GoTo ErrHandler
    LogMessage "Mise à jour de l'état du ruban. isSingleCell = " & isSingleCell

    ' Met à jour la variable globale utilisée par les callbacks du ruban.
    g_isSingleCellSelection = isSingleCell

    ' Invalide le ruban pour forcer l'appel des callbacks 'getEnabled'.
    If GetRibbon() Is Nothing Then Exit Sub
    GetRibbon().Invalidate

    Exit Sub
ErrHandler:
    DisplayAndLogError "UpdateRibbonState", Err
End Sub

'====================================================================================
' SECTION 4 : GESTION DU MENU CONTEXTUEL (CLIC-DROIT)
'====================================================================================


'------------------------------------------------------------------------------------
' PROCÉDURE   : CreateDynamicContextMenu
' DESCRIPTION :
'   Point d'entrée pour la création du menu contextuel. Appelé par l'événement
'   SheetBeforeRightClick, il analyse la sélection et construit le menu approprié.
' PARAMÈTRES : Target (Range) - La cellule sur laquelle l'utilisateur a cliqué.
'------------------------------------------------------------------------------------
Public Sub CreateDynamicContextMenu(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    ' S'assurer que la cellule cliquée est bien la cellule active pour les macros
    Target.Cells(1).Activate
    
    ' --- 1. Détection du type de données et de la sélection (unique/liste) ---
    g_FirstRow = Selection.row
    Call Detect_g_LastRow
    g_IsList = (g_LastRow <> g_FirstRow) Or (Selection.Cells.count > 1)
    
    Call Detect_g_DataType
    
    ' --- 2. Construction du menu ---
    Dim cellMenu As CommandBar
    Dim customMenu As CommandBarControl
    Dim vKeys As Variant
    
    Set cellMenu = Application.CommandBars("Cell")
    
    ' Crée le menu principal "ExcelToSAP"
    Set customMenu = cellMenu.Controls.add(Type:=msoControlPopup, Before:=1, Temporary:=True)
    customMenu.Tag = CUSTOM_MENU_TAG ' Tag pour pouvoir le retrouver et le supprimer
    
    ' --- 3. Remplissage du menu en fonction du type de données ---
    If Not g_DetectedTypes Is Nothing Then
        If g_DetectedTypes.count = 1 Then
            ' Cas 1: Un type de données a été trouvé
            vKeys = g_DetectedTypes.Keys
            customMenu.caption = "ExcelToSAP (" & g_DetectedTypes(vKeys(0)) & ")"
            Call AddTypedMenuItems(customMenu.CommandBar, CStr(vKeys(0)), g_IsList)
        ElseIf g_DetectedTypes.count > 1 Then
            ' Cas 2: Plusieurs types trouvés -> Créer des sous-menus
            customMenu.caption = "ExcelToSAP (Multiple)"
            vKeys = g_DetectedTypes.Keys
            Dim i As Long
            Dim subMenu As CommandBarPopup
            For i = LBound(vKeys) To UBound(vKeys)
                Set subMenu = customMenu.CommandBar.Controls.add(Type:=msoControlPopup)
                subMenu.caption = g_DetectedTypes(vKeys(i))
                Call AddTypedMenuItems(subMenu.CommandBar, CStr(vKeys(i)), g_IsList)
            Next i
        Else
            ' Cas 3: Aucun type trouvé -> Afficher tous les menus
            customMenu.caption = "ExcelToSAP"
            Call AddAllMenuItems(customMenu.CommandBar, g_IsList)
        End If
    Else
        ' Fallback
        customMenu.caption = "ExcelToSAP"
        Call AddAllMenuItems(customMenu.CommandBar, g_IsList)
    End If

    Exit Sub
ErrHandler:
    ' Ne pas afficher d'erreur si l'utilisateur annule, mais nettoyer le menu
    Call CleanupContextMenu
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : AddTypedMenuItems (Privée)
' DESCRIPTION : Ajoute les commandes spécifiques à un type de données détecté.
'------------------------------------------------------------------------------------
Private Sub AddTypedMenuItems(ByVal parentMenu As CommandBar, ByVal dataType As String, ByVal isList As Boolean)
    Select Case dataType
        Case "Pmord"
            If isList Then
                AddContextMenuItem parentMenu, "Display Orders (IW39)", "'PM_Order.Z_IW39_O'", 25
                AddContextMenuItem parentMenu, "Display Orders History (IW39)", "'PM_Order.Z_IW39_H'", 173
                AddContextMenuItem parentMenu, "Display Orders Open (IW38)", "'PM_Order.Z_IW38O'", 25
                AddContextMenuItem parentMenu, "Operations List (IW49N)", "'PM_Order.Z_IW49NL'", 59, True
                AddContextMenuItem parentMenu, "Multi Level (IW40)", "'PM_Order.Z_IW40_L'", 25
                AddContextMenuItem parentMenu, "Components List (IWBK)", "'PM_Order.Z_IWBK_L'", 211
                AddContextMenuItem parentMenu, "Confirmations List (IW47)", "'PM_Order.Z_IW47_L'", 71
                AddContextMenuItem parentMenu, "Material Movements (MB51)", "'PM_Order.Z_MB51_L'", 485
                AddContextMenuItem parentMenu, "Purchase Requisitions (ME5A)", "'PM_Order.Z_ME5A_L'", 486
                AddContextMenuItem parentMenu, "Mass Release (IW38)", "'PM_Order.Z_PMO_IW38_MassRelease'", 101, True
                AddContextMenuItem parentMenu, "Mass TECO (IW38)", "'PM_Order.Z_PMO_IW38_MassTECO'", 102
                AddContextMenuItem parentMenu, "Mass Status 4SCH (IW38)", "'PM_Order.Z_PMO_IW38_4SCH'", 71
                AddContextMenuItem parentMenu, "Mass Status 5COM (IW38)", "'PM_Order.Z_PMO_IW38_5COM'", 71
                AddContextMenuItem parentMenu, "Mass User Status (IW38)", "'PM_Order.Z_PMO_IW38_User_Status_mass_Change'", 71
                AddContextMenuItem parentMenu, "Mass Change Date (IW38)", "'PM_Order.Z_PMO_IW38_Change_Date'", 71
                AddContextMenuItem parentMenu, "Mass Print Order", "'PM_Order.Z_PMO_Print_Orders'", 4
                AddContextMenuItem parentMenu, "Mass Confirmation (IW38)", "'PM_Order.Z_PMO_IW38_MassConfirmation'", 71
                AddContextMenuItem parentMenu, "Cancel Confirmations", "'PM_Order.Z_PMO_Cancel_Confirmations'", 22
                AddContextMenuItem parentMenu, "Gantt", "'PM_Order.Z_PMO_Gantt'", 573, True
                AddContextMenuItem parentMenu, "Scheduler", "'PM_Order.Z_PMO_SCHEDULER'", 573
            Else
                AddContextMenuItem parentMenu, "Display Order (IW33)", "'PM_Order.Z_PMO_IW33_1_Main'", 25
                AddContextMenuItem parentMenu, "Change Order (IW32)", "'PM_Order.Z_PMO_IW32_1_Main'", 71
                AddContextMenuItem parentMenu, "Operations (IW33)", "'PM_Order.Z_PMO_IW33_2_Operations'", 59
                AddContextMenuItem parentMenu, "Components (IW33)", "'PM_Order.Z_PMO_IW33_3_Components'", 211
                AddContextMenuItem parentMenu, "Costs (IW33)", "'PM_Order.Z_PMO_IW33_4_Costs'", 1085
                AddContextMenuItem parentMenu, "Planning (IW33)", "'PM_Order.Z_PMO_IW33_5_Planning'", 0, True
                AddContextMenuItem parentMenu, "Enhancement (IW33)", "'PM_Order.Z_PMO_IW33_6_Enhancement'", 0
                AddContextMenuItem parentMenu, "Log (IW33)", "'PM_Order.Z_PMO_IW33_7_LOG'", 0
                AddContextMenuItem parentMenu, "Print (IW32)", "'PM_Order.Z_PMO_IW32_15_Print_Orders'", 0, True
                AddContextMenuItem parentMenu, "WO to BOM (IW32)", "'PM_Order.Z_PMO_IW32_16_WO_to_BOOM'", 0
                AddContextMenuItem parentMenu, "CNF (IW41)", "'PM_Order.Z_PMO_IW32_17_CNF'", 0
                AddContextMenuItem parentMenu, "Cancel CNF (IW45)", "'PM_Order.Z_PMO_IW32_18_Cancel_CNF'", 0
                AddContextMenuItem parentMenu, "Operations List (IW49N)", "'PM_Order.Z_PMO_IW49N_Operations'", 0, True
                AddContextMenuItem parentMenu, "Multi Level (IW40)", "'PM_Order.Z_PMO_IW40'", 0
                AddContextMenuItem parentMenu, "Components List (IWBK)", "'PM_Order.Z_PMO_IWBK'"
                AddContextMenuItem parentMenu, "Confirmations List (IW47)", "'PM_Order.Z_PMO_IW47'"
                AddContextMenuItem parentMenu, "Material Movements (MB51)", "'PM_Order.Z_PMO_MB51_MatChar'"
                AddContextMenuItem parentMenu, "Purchase Requisitions (ME5A)", "'PM_Order.Z_PMO_ME5A'"
            End If
            
        Case "Notif"
            If isList Then
                AddContextMenuItem parentMenu, "Display Open Notifications (IW29)", "'notifications.Z_IW29Open'", 25
                AddContextMenuItem parentMenu, "Display Closed Notifications (IW29)", "'notifications.Z_IW29Closed'", 173
                AddContextMenuItem parentMenu, "Display Open Notifications (IW28)", "'notifications.Z_IW28Open'", 25
                AddContextMenuItem parentMenu, "Multi-Level (IW30)", "'notifications.Z_PMO_IW30'", 1080, True
                AddContextMenuItem parentMenu, "Orders Open (IW39)", "'notifications.Z_IW39NOpen'", 22
                AddContextMenuItem parentMenu, "Orders Closed (IW39)", "'notifications.Z_IW39NClosed'", 173
                AddContextMenuItem parentMenu, "Orders Open (IW38)", "'notifications.Z_IW38NOpen'", 22
                AddContextMenuItem parentMenu, "Components (IWBK)", "'notifications.Z_IWBKN'", 211
                AddContextMenuItem parentMenu, "Operations (IW49N)", "'notifications.Z_IW49NN'", 59
            Else
                AddContextMenuItem parentMenu, "Display Notification (IW23)", "'notifications.Z_IW23NotM'", 25
                AddContextMenuItem parentMenu, "Change Notification (IW22)", "'notifications.Z_IW22NotM'", 71
                AddContextMenuItem parentMenu, "Synthesis (IW23)", "'notifications.Z_IW23NotSum'", 487, True
                AddContextMenuItem parentMenu, "Location (IW23)", "'notifications.Z_IW23NotLoc'", 36
                AddContextMenuItem parentMenu, "Multi-Level (IW23)", "'notifications.Z_IW23NotMul'", 1080
                AddContextMenuItem parentMenu, "Log (IW23)", "'notifications.Z_IW23NotLog'", 173
                AddContextMenuItem parentMenu, "Multi-Level (IW30)", "'notifications.Z_PMO_IW30_Single'", 1080
            End If
            
        Case "Equip"
            If isList Then
                AddContextMenuItem parentMenu, "List Equipments (IH08)", "'Equipment.Z_IH08_E'", 1081
                AddContextMenuItem parentMenu, "List Equipments (IE07)", "'Equipment.Z_IE07_E'", 1081
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'Equipment.Z_IW39_E'", 22, True
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'Equipment.Z_IW39_EC'", 173
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'Equipment.Z_IW38_E'", 22
                AddContextMenuItem parentMenu, "Notifications (IW29)", "'Equipment.Z_IW29_E'", 1644
                AddContextMenuItem parentMenu, "Notifications History (IW29)", "'Equipment.Z_IW29_EC'", 173
                AddContextMenuItem parentMenu, "Notifications (IW28)", "'Equipment.Z_IW28_E'", 1644
                AddContextMenuItem parentMenu, "Maintenance Plans (IP18)", "'Equipment.Z_IP18_E'", 1085, True
                AddContextMenuItem parentMenu, "Maintenance Scheduling (IP24)", "'Equipment.Z_IP24_E'", 1085
                AddContextMenuItem parentMenu, "Task Lists (IA17)", "'Equipment.Z_IA17_E'", 19
                AddContextMenuItem parentMenu, "General Task Lists (IA17)", "'Equipment.Z_IA17_E_G'", 19
            Else
                AddContextMenuItem parentMenu, "Display Equipment (IE03)", "'Equipment.equipmentInfo'", 37
                AddContextMenuItem parentMenu, "BOM (IE03)", "'Equipment.equipmentBOM'", 1080
                AddContextMenuItem parentMenu, "Characteristics (IE03)", "'Equipment.eqchars'", 78, True
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'Equipment.equipmentPMO'", 22
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'Equipment.equipmentPMO_H'", 173
                AddContextMenuItem parentMenu, "PM Orders Current (IW38)", "'Equipment.equipmentPMOc'", 22
                AddContextMenuItem parentMenu, "Notifications (IW29)", "'Equipment.equipmentNotif'", 1644
                AddContextMenuItem parentMenu, "Notifications History (IW29)", "'Equipment.equipmentNotifH'", 173
                AddContextMenuItem parentMenu, "Notifications Current (IW28)", "'Equipment.equipmentNotifc'", 1644
                AddContextMenuItem parentMenu, "Maintenance Plans (IP18)", "'Equipment.equipmentplan'", 1085, True
                AddContextMenuItem parentMenu, "Maintenance Scheduling (IP24)", "'Equipment.equipmentplan24'", 1085
                AddContextMenuItem parentMenu, "Display BOM (IB03)", "'Equipment.equipmentBOMIB03'", 25, True
                AddContextMenuItem parentMenu, "Change BOM (IB02)", "'Equipment.equipmentBOMIB02'", 71
                AddContextMenuItem parentMenu, "Create BOM (IB01)", "'Equipment.equipmentBOMIB01'", 1079
                AddContextMenuItem parentMenu, "Display Task List (IA03)", "'Equipment.equipmentIA03'", 25, True
                AddContextMenuItem parentMenu, "Change Task List (IA02)", "'Equipment.equipmentIA02'", 71
                AddContextMenuItem parentMenu, "Create Task List (IA01)", "'Equipment.equipmentIA01'", 1079
                AddContextMenuItem parentMenu, "General Task List (IA17)", "'Equipment.Z_IA17_E_G_Single'", 19
            End If

        Case "FLOC"
            If isList Then
                AddContextMenuItem parentMenu, "List FLOCs (IH06)", "'FuncLoc.Z_IH06_F'", 36
                AddContextMenuItem parentMenu, "List FLOCs (IH08)", "'FuncLoc.Z_IH08_F'", 36
                AddContextMenuItem parentMenu, "List FLOCs (IL07)", "'FuncLoc.Z_IL07_F'", 36
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'FuncLoc.Z_IW39_F'", 22, True
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'FuncLoc.Z_IW39_FC'", 173
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'FuncLoc.Z_IW38_F'", 22
                AddContextMenuItem parentMenu, "Notifications (IW29)", "'FuncLoc.Z_IW29_F'", 1644
                AddContextMenuItem parentMenu, "Notifications History (IW29)", "'FuncLoc.Z_IW29_FC'", 173
                AddContextMenuItem parentMenu, "Notifications (IW28)", "'FuncLoc.Z_IW28_F'", 1644
                AddContextMenuItem parentMenu, "Maintenance Plans (IP18)", "'FuncLoc.Z_IP18_F'", 1085, True
                AddContextMenuItem parentMenu, "Maintenance Scheduling (IP24)", "'FuncLoc.Z_IP24_F'", 1085
                AddContextMenuItem parentMenu, "Task Lists (IA17)", "'FuncLoc.Z_IA17_F'", 19
                AddContextMenuItem parentMenu, "General Task Lists (IA17)", "'FuncLoc.Z_IA17_F_G'", 19
            Else
                AddContextMenuItem parentMenu, "Display FLOC (IL03)", "'FuncLoc.Z_FLOC_IL03'", 36
                AddContextMenuItem parentMenu, "Characteristics (IL03)", "'FuncLoc.Z_FLOC_IL03_c'", 78
                AddContextMenuItem parentMenu, "BOM (IL03)", "'FuncLoc.Z_FLOC_IL03_B'", 1080
                AddContextMenuItem parentMenu, "Structure (IH01)", "'FuncLoc.Z_FLOC_IH01'", 1082, True
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'FuncLoc.flocPMO'", 22
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'FuncLoc.flocPMO_H'", 173
                AddContextMenuItem parentMenu, "PM Orders Current (IW38)", "'FuncLoc.flocPMOc'", 22
                AddContextMenuItem parentMenu, "Notifications (IW29)", "'FuncLoc.flocNotif'", 1644
                AddContextMenuItem parentMenu, "Notifications History (IW29)", "'FuncLoc.flocNotifH'", 173
                AddContextMenuItem parentMenu, "Notifications Current (IW28)", "'FuncLoc.flocNotifc'", 1644
                AddContextMenuItem parentMenu, "Maintenance Plans (IP18)", "'FuncLoc.flocplan'", 1085, True
                AddContextMenuItem parentMenu, "Maintenance Scheduling (IP24)", "'FuncLoc.flocplan24'", 1085
                AddContextMenuItem parentMenu, "Display BOM (IB03)", "'FuncLoc.FLOCBOMIB03'", 25, True
                AddContextMenuItem parentMenu, "Change BOM (IB02)", "'FuncLoc.floCBOMIB02'", 71
                AddContextMenuItem parentMenu, "Create BOM (IB01)", "'FuncLoc.flocOMIB01'", 1079
                AddContextMenuItem parentMenu, "Display Task List (IA13)", "'FuncLoc.flocIA13'", 25, True
                AddContextMenuItem parentMenu, "Change Task List (IA12)", "'FuncLoc.flocIA12'", 71
                AddContextMenuItem parentMenu, "Create Task List (IA11)", "'FuncLoc.flocIA11'", 1079
            End If

        Case "M.Plan"
            If isList Then
                AddContextMenuItem parentMenu, "List Plans (IP16)", "'M_Plan.ip16_L'", 577
                AddContextMenuItem parentMenu, "Scheduling (IP24)", "'M_Plan.ip24_L'", 1085
                AddContextMenuItem parentMenu, "PM Orders Open (IW39)", "'M_Plan.Z_IW39OL'", 22, True
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'M_Plan.pmohistplanL'", 173
                AddContextMenuItem parentMenu, "PM Orders Current (IW38)", "'M_Plan.Z_IW38OL'", 22
                AddContextMenuItem parentMenu, "Manual Call", "'M_Plan.Manual_Call_PMRs'", 1091, True
                AddContextMenuItem parentMenu, "Change Start Cycle", "'M_Plan.Change_StartCycle_PMRs'", 1085
                AddContextMenuItem parentMenu, "Start PMRs", "'M_Plan.Start_PMRs'", 206
                AddContextMenuItem parentMenu, "New Start PMRs", "'M_Plan.NewStart_PMRs'", 24
                AddContextMenuItem parentMenu, "Planification (IP30)", "'M_Plan.Planification_PMRs'", 1092
                AddContextMenuItem parentMenu, "Auto 1W", "'M_Plan.MP_Auto_1W'", 24, True
                AddContextMenuItem parentMenu, "Auto 1M", "'M_Plan.MP_Auto_1M'", 24
                AddContextMenuItem parentMenu, "Auto 13W", "'M_Plan.MP_Auto_13W'", 24
                AddContextMenuItem parentMenu, "Auto 1Y", "'M_Plan.MP_Auto_1Y'", 24
                AddContextMenuItem parentMenu, "Tree View (ZPM004)", "'M_Plan.Afficher_Arbo_PMR'", 1082, True
            Else
                AddContextMenuItem parentMenu, "Display Plan (IP03)", "'M_Plan.displayPLan'", 25
                AddContextMenuItem parentMenu, "Change Plan (IP02)", "'M_Plan.ChangePLan'", 71
                AddContextMenuItem parentMenu, "Operations (IP03)", "'M_Plan.displayPLanOP'", 59, True
                AddContextMenuItem parentMenu, "Calls (IP03)", "'M_Plan.displayPLanCalls'", 1091
                AddContextMenuItem parentMenu, "Scheduling (IP24)", "'M_Plan.IP24plan'", 1085
                AddContextMenuItem parentMenu, "Call History (IP03)", "'M_Plan.callH'", 173
                AddContextMenuItem parentMenu, "Last Call (IP03)", "'M_Plan.lastcall'", 161
                AddContextMenuItem parentMenu, "Enhancements (IP03)", "'M_Plan.enh'", 71
                AddContextMenuItem parentMenu, "Change Task List (IP02)", "'M_Plan.changeTLgivenPlan'", 71
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'M_Plan.pmohistplan'", 173
            End If

        Case "Item"
            If isList Then
                AddContextMenuItem parentMenu, "List Items (IP18)", "'M_Plan.Z_IP18LI'", 1081
                AddContextMenuItem parentMenu, "Scheduling (IP24)", "'M_Plan.IP24_LI'", 1085, True
                AddContextMenuItem parentMenu, "PM Orders History (IW39)", "'M_Plan.pmohistplanLI'", 173
                AddContextMenuItem parentMenu, "PM Orders Open (IW39)", "'M_Plan.Z_IW39OLI'", 22
                AddContextMenuItem parentMenu, "PM Orders Current (IW38)", "'M_Plan.Z_IW38OLI'", 22
            Else
                AddContextMenuItem parentMenu, "Display Item (IP06)", "'M_Plan.Z_IP06I'", 1645
            End If

        Case "Material"
            If isList Then
                AddContextMenuItem parentMenu, "Stock (MB52)", "'materials.matmb52L'", 576
                AddContextMenuItem parentMenu, "Movements (MB51)", "'materials.matmovL'", 161
                AddContextMenuItem parentMenu, "Orders (ME2N)", "'materials.matOrdersL'", 169, True
                AddContextMenuItem parentMenu, "Reservations (MB24)", "'materials.maatreservationsL'", 486
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'materials.IW39MatL'", 22, True
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'materials.IW38MatL'", 22
                AddContextMenuItem parentMenu, "Components (IWBK)", "'materials.IWBKmatL'", 211
            Else
                AddContextMenuItem parentMenu, "Stock (MMBE)", "'materials.stock'", 576
                AddContextMenuItem parentMenu, "Basic Data (MM03)", "'materials.basdata'", 78, True
                AddContextMenuItem parentMenu, "Orders (ME2N)", "'materials.matOrders'", 169
                AddContextMenuItem parentMenu, "Long Text (MM03)", "'materials.matlt'", 213
                AddContextMenuItem parentMenu, "Consumption (MM03)", "'materials.matcons'", 1085
                AddContextMenuItem parentMenu, "Documents (MM03)", "'materials.docdata'", 213
                AddContextMenuItem parentMenu, "Movements (MB51)", "'materials.matmov'", 161
                AddContextMenuItem parentMenu, "Stock (MB52)", "'materials.matmb52'", 576
                AddContextMenuItem parentMenu, "Reservations (MB24)", "'materials.maatreservations'", 486
                AddContextMenuItem parentMenu, "Notes (MB53)", "'materials.matnotes'", 1645
                AddContextMenuItem parentMenu, "Structure (CC04)", "'materials.mattied'", 1082, True
                AddContextMenuItem parentMenu, "Structure CS15", "'materials.mattiedcs'", 1082
                AddContextMenuItem parentMenu, "PM Orders (IW39)", "'materials.IW39Mat'", 22, True
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'materials.IW38Mat'", 22
                AddContextMenuItem parentMenu, "Components (IWBK)", "'materials.IWBKmat'", 211
            End If

        Case "Requis"
            If isList Then
                AddContextMenuItem parentMenu, "List PR (ME5A)", "'Requisition.Z_ME5AReqL'", 486
            Else
                AddContextMenuItem parentMenu, "Display PR (ME53)", "'Requisition.Z_ME53M'", 25
                AddContextMenuItem parentMenu, "Display PR (ME53N)", "'Requisition.Z_ME53NM'", 25
                AddContextMenuItem parentMenu, "List PR (ME5A)", "'Requisition.Z_ME5AReq'", 486, True
            End If

        Case "PurOrder"
            If isList Then
                AddContextMenuItem parentMenu, "List PO (ME2N)", "'PurchOrder.Z_ME2NML'", 108
                AddContextMenuItem parentMenu, "Analysis (ME80FN)", "'PurchOrder.Z_ME80FNML'", 487, True
                AddContextMenuItem parentMenu, "Reminders (ME91F)", "'PurchOrder.Z_ME91FML'", 161
            Else
                AddContextMenuItem parentMenu, "Display PO (ME23)", "'PurchOrder.Z_ME23M'", 25
                AddContextMenuItem parentMenu, "Display PO (ME23N)", "'PurchOrder.Z_ME23NM'", 25
                AddContextMenuItem parentMenu, "Analysis (ME80FN)", "'PurchOrder.Z_ME80FNM'", 487, True
                AddContextMenuItem parentMenu, "Reminders (ME91F)", "'PurchOrder.Z_ME91FM'", 161
                AddContextMenuItem parentMenu, "List PO (ME2N)", "'PurchOrder.Z_ME2NM'", 108
            End If

        Case "Project"
            If isList Then
                AddContextMenuItem parentMenu, "Structure (CN42N)", "'projectM.Z_CN42NProjL'", 1080
                AddContextMenuItem parentMenu, "Documents (CN43N)", "'projectM.Z_CN43NProjL'", 213
                AddContextMenuItem parentMenu, "Components (CNS41)", "'projectM.Z_CNS41ProjL'", 211, True
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'projectM.Z_CN46NProjL'", 137
                AddContextMenuItem parentMenu, "Activities (CN47)", "'projectM.Z_CN47ProjL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'projectM.Z_CN47NProjL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'projectM.Z_CN48NProjL'", 15
            
            Else
                AddContextMenuItem parentMenu, "Structure (CN42N)", "'projectM.Z_CN42NProjL'", 1080
                AddContextMenuItem parentMenu, "Documents (CN43N)", "'projectM.Z_CN43NProjL'", 213
                AddContextMenuItem parentMenu, "Components (CNS41)", "'projectM.Z_CNS41ProjL'", 211
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'projectM.Z_CN46NProjL'", 137
                AddContextMenuItem parentMenu, "Activities (CN47)", "'projectM.Z_CN47ProjL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'projectM.Z_CN47NProjL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'projectM.Z_CN48NProjL'", 15
            
            End If

        Case "Network"
            If isList Then
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'netwrk.Z_CN46NNetL'", 137
                AddContextMenuItem parentMenu, "Components (CNS41)", "'netwrk.Z_CNS41NetL'", 211
                AddContextMenuItem parentMenu, "PRs (ME5A)", "'netwrk.Z_ME5ANetL'", 486, True
                AddContextMenuItem parentMenu, "Activities (CN47)", "'netwrk.Z_CN47NetL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'netwrk.Z_CN47NNetL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'netwrk.Z_CN48NNetL'", 15
            
            Else
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'netwrk.Z_CN46NNetL'", 137
                AddContextMenuItem parentMenu, "Components (CNS41)", "'netwrk.Z_CNS41NetL'", 211
                AddContextMenuItem parentMenu, "PRs (ME5A)", "'netwrk.Z_ME5ANetL'", 486
                AddContextMenuItem parentMenu, "Activities (CN47)", "'netwrk.Z_CN47NetL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'netwrk.Z_CN47NNetL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'netwrk.Z_CN48NNetL'", 15
                        
            End If

        Case "WBS"
            If isList Then
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'WBS_M.Z_CN46N_WBSL'", 137
                AddContextMenuItem parentMenu, "Activities (CN47)", "'WBS_M.Z_CN47_WBSL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'WBS_M.Z_CN47N_WBSL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'WBS_M.Z_CN48N_WBSL'", 15
                AddContextMenuItem parentMenu, "PRs (ME5A)", "'WBS_M.Z_ME5A_WBSL'", 486, True
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'WBS_M.Z_IW38_WBSL'", 22
                AddContextMenuItem parentMenu, "Components (CNS41)", "'WBS_M.Z_CNS41_WBSL'", 211, True
                AddContextMenuItem parentMenu, "Documents (CN43N)", "'WBS_M.Z_CN43N_WBSL'", 213
            Else
                AddContextMenuItem parentMenu, "Networks (CN46N)", "'WBS_M.Z_CN46N_WBSL'", 137
                AddContextMenuItem parentMenu, "PRs (ME5A)", "'WBS_M.Z_ME5A_WBSL'", 486
                AddContextMenuItem parentMenu, "Activities (CN47)", "'WBS_M.Z_CN47_WBSL'", 169
                AddContextMenuItem parentMenu, "Activities (CN47N)", "'WBS_M.Z_CN47N_WBSL'", 169
                AddContextMenuItem parentMenu, "Confirmations (CN48N)", "'WBS_M.Z_CN48N_WBSL'", 15
                AddContextMenuItem parentMenu, "Documents (CN43N)", "'WBS_M.Z_CN43N_WBSL'", 213
                AddContextMenuItem parentMenu, "PM Orders (IW38)", "'WBS_M.Z_IW38_WBSL'", 22
                AddContextMenuItem parentMenu, "Components (CNS41)", "'WBS_M.Z_CNS41_WBSL'", 211
            
            End If

        Case "CNF"
            If isList Then
                AddContextMenuItem parentMenu, "Cancel Confirmations", "'PM_Order.Z_PMO_Cancel_Confirmations'", 22
            Else
                AddContextMenuItem parentMenu, "Confirmation (IW44)", "'PM_Order.Z_PMO_IW44'", 22
                AddContextMenuItem parentMenu, "Cancel Confirmation (IW45)", "'PM_Order.Z_PMO_IW32_18_Cancel_CNF'", 22
            End If
        
    End Select
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : AddAllMenuItems (Privée)
' DESCRIPTION : Construit le menu complet avec toutes les commandes, groupées par type.
'------------------------------------------------------------------------------------
Private Sub AddAllMenuItems(ByVal parentMenu As CommandBar, ByVal isList As Boolean)
    Dim subMenu As CommandBarPopup
    
    ' Work Order
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Work Order"
    Call AddTypedMenuItems(subMenu.CommandBar, "Pmord", isList)
    
    ' Notification
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Notification"
    Call AddTypedMenuItems(subMenu.CommandBar, "Notif", isList)
    
    ' Confirmation
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Confirmation"
    Call AddTypedMenuItems(subMenu.CommandBar, "CNF", isList)
    
    ' Equipment
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.beginGroup = True
    subMenu.caption = "Equipment"
    Call AddTypedMenuItems(subMenu.CommandBar, "Equip", isList)
    
    ' Functional Location
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Functional Location"
    Call AddTypedMenuItems(subMenu.CommandBar, "FLOC", isList)
    
    ' Maintenance Plan
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.beginGroup = True
    subMenu.caption = "Maintenance Plan"
    Call AddTypedMenuItems(subMenu.CommandBar, "M.Plan", isList)
    
    ' Item
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Item"
    Call AddTypedMenuItems(subMenu.CommandBar, "Item", isList)
    
    ' Material
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.beginGroup = True
    subMenu.caption = "Material"
    Call AddTypedMenuItems(subMenu.CommandBar, "Material", isList)
    
    ' Purchase Requisition
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Purchase Requisition"
    Call AddTypedMenuItems(subMenu.CommandBar, "Requis", isList)
    
    ' Purchase Order
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Purchase Order"
    Call AddTypedMenuItems(subMenu.CommandBar, "PurOrder", isList)
    
    ' Project
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.beginGroup = True
    subMenu.caption = "Project"
    Call AddTypedMenuItems(subMenu.CommandBar, "Project", isList)
    
    ' Network
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "Network"
    Call AddTypedMenuItems(subMenu.CommandBar, "Network", isList)
    
    ' WBS
    Set subMenu = parentMenu.Controls.add(Type:=msoControlPopup)
    subMenu.caption = "WBS"
    Call AddTypedMenuItems(subMenu.CommandBar, "WBS", isList)
    
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : AddContextMenuItem (Privée)
' DESCRIPTION : Ajoute un bouton au menu contextuel.
'------------------------------------------------------------------------------------
Private Sub AddContextMenuItem(ByVal parentMenu As CommandBar, ByVal caption As String, ByVal onAction As String, Optional ByVal iconId As Long = 0, Optional ByVal beginGroup As Boolean = False)
    Dim newItem As CommandBarButton
    Dim cleanAction As String
    Set newItem = parentMenu.Controls.add(Type:=msoControlButton, Temporary:=True)
    With newItem
        .caption = caption
        ' Nettoie les apostrophes existantes et ajoute le nom du classeur courant pour éviter les liens rompus
        cleanAction = Replace(onAction, "'", "")
        .onAction = "'" & ThisWorkbook.name & "'!" & cleanAction
        .beginGroup = beginGroup
        ' Si un iconId est fourni, on l'utilise et on change le style pour afficher l'icône et le texte.
        If iconId <> 0 Then
            .FaceId = iconId
            .Style = msoButtonIconAndCaption
        Else
            .Style = msoButtonCaption
        End If
    End With
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : CleanupContextMenu
' DESCRIPTION :
'   Supprime le menu contextuel personnalisé ajouté précédemment pour éviter
'   les doublons lors de clics droits successifs.
'------------------------------------------------------------------------------------
Public Sub CleanupContextMenu()
    On Error Resume Next
    ' Supprime le contrôle en utilisant le Tag que nous avons défini
    Application.CommandBars("Cell").FindControl(Tag:=CUSTOM_MENU_TAG).Delete
    On Error GoTo 0
End Sub
