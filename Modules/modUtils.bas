Attribute VB_Name = "modUtils"
'====================================================================================
' MODULE      : modUtils
' VERSION     : 1.0
' AUTEUR      : [Votre Nom] / Révisé par Gemini
' DATE        : 02/12/2025
' DESCRIPTION : Ce module centralise toutes les fonctions utilitaires transverses
'               au projet ExcelToSAP. Il fournit des outils réutilisables pour :
'                 - La journalisation des événements.
'                 - La gestion centralisée des erreurs.
'                 - Les temporisations et la réactivité de l'interface.
'                 - Les optimisations de performance d'Excel.
'                 - La manipulation de fichiers Excel (export, chargement).
'                 - Le calcul et la gestion des dates et périodes.
'
'               Ce module ne contient aucune logique métier spécifique à SAP,
'               mais des fonctions génériques utilisées par d'autres modules.
'====================================================================================

Option Explicit

'====================================================================================
' SECTION 1 : DÉCLARATIONS GLOBALES ET API
'====================================================================================

' --- 1.1 Variables Globales de Date et Périodes (calculées par DateSemain) ---
' Ces variables sont calculées dynamiquement et représentent l'état temporel de l'application.
Public SemaineKPIs As String
Public SemaineKPIs_Avant As String
Public DebutSemaine As String
Public DebutSemaine_Avant As String
Public FinSemaine As String
Public AgingDate As String
Public g_DebutMois As Date ' Date de début du mois courant.
Public g_FinMois As Date   ' Date de fin du mois courant.
Public g_DebutAnnee As String
Public g_FinAnnee As String

' --- 1.2 Déclarations API Windows ---
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If


'====================================================================================
' SECTION 2 : JOURNALISATION (LOGGING)
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : InitLogs
' DESCRIPTION : Crée le dossier de logs s'il n'existe pas et initialise le fichier
'               journal pour la journée en cours, en écrasant le contenu précédent.
' DÉPENDANCES : GetSetting (pour PATH_LOGS).
'------------------------------------------------------------------------------------
Public Sub InitLogs()
    Dim logFile As String
    Dim logPath As String
    
    logPath = GetSetting("PATH_LOGS")
    If logPath = "" Then
        Debug.Print "AVERTISSEMENT: Le parametre PATH_LOGS n'est pas defini. Logs desactives."
        Exit Sub
    End If
    
    logFile = logPath & "SAP_Export_" & Format(Date, "yyyy-mm-dd") & ".txt"
    If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    ' Ouvre le fichier en mode écriture (Output) pour écraser son contenu.
    Open logFile For Output As #1
    Print #1, "===== Debut du log " & Now & " ====="
    Close #1
End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : LogMessage
' DESCRIPTION : Écrit un message horodaté dans la fenêtre d'exécution (Immediate Window)
'               et, optionnellement, dans un fichier journal centralisé.
' PARAMÈTRES  :
'   - txt (String) : Le message à enregistrer.
' DÉPENDANCES : GetSetting (pour PATH_LOGS).
'------------------------------------------------------------------------------------
Public Sub LogMessage(ByVal txt As String)
    Dim logLine As String
    logLine = Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & txt

    ' -- Écriture dans la fenêtre Immediate (pour le débogage)
    Debug.Print logLine

    ' -- Écriture optionnelle dans un fichier log
    '    Utilise On Error Resume Next pour éviter de planter si le fichier est inaccessible.
    On Error Resume Next
    Dim fn As String, h As Integer
    fn = GetSetting("PATH_LOGS") & "ExcelToSAP_Log.txt"
    h = FreeFile

    ' Cree le dossier si necessaire
    If Dir(GetSetting("PATH_LOGS"), vbDirectory) = "" Then MkDir GetSetting("PATH_LOGS")

    Open fn For Append As #h
    Print #h, logLine
    Close #h
    On Error GoTo 0

End Sub

'====================================================================================
' SECTION 3 : GESTION DES ERREURS
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : HandleError
' DESCRIPTION : Centralise la gestion des erreurs pour le projet. Affiche un message
'               clair à l'utilisateur et enregistre les détails techniques dans le journal.
' PARAMÈTRES  :
'   - procedureName (String) : Nom de la procédure où l'erreur est survenue.
'   - errObj (ErrObject)     : L'objet Err système contenant les détails de l'erreur.
'------------------------------------------------------------------------------------
Public Sub HandleError(ByVal procedureName As String, ByVal errObj As ErrObject)
    Dim errMsg As String
    errMsg = "ERREUR dans [" & procedureName & "] : " & errObj.Description & _
             " (Code " & errObj.Number & ")"

    ' Ecrire dans le log
    LogMessage errMsg
    
    ' Avertir l'utilisateur
    MsgBox errMsg, vbCritical, "Erreur - ExcelToSAP"

End Sub

'------------------------------------------------------------------------------------
' PROCÉDURE   : DisplayAndLogError
' DESCRIPTION : Centralise la gestion des erreurs. Affiche un message clair à
'               l'utilisateur, enregistre les détails techniques, et nettoie la session.
' PARAMÈTRES  :
'   - procedureName (String) : Nom de la procédure où l'erreur est survenue.
'   - errObj (ErrObject)     : L'objet Err système.
'------------------------------------------------------------------------------------
Public Sub DisplayAndLogError(ByVal procedureName As String, ByVal errObj As ErrObject)
    Dim userMsg As String
    Dim logMsg As String
    
    logMsg = "ERREUR dans [" & procedureName & "] : " & errObj.Description & " (Code " & errObj.Number & ")"
    LogMessage logMsg ' Journalise l'erreur technique en premier.

    ' Analyse l'erreur pour fournir un message utilisateur plus pertinent,
    ' mais seulement si nous ne sommes PAS en mode test automatisé.
    If Not g_IsTestMode Then
    If Not IsSAPConnectionAlive() Then
        userMsg = "La connexion à SAP a été perdue pendant l'opération." & vbCrLf & vbCrLf & _
                  "Veuillez vérifier votre connexion SAP et réessayer."
    ElseIf errObj.Number = 619 Then ' Erreur typique de SAP GUI Scripting : élément non trouvé.
        userMsg = "Un élément de l'interface SAP n'a pas pu être trouvé." & vbCrLf & vbCrLf & _
                  "Cause possible : Une mise à jour de SAP a modifié l'écran de la transaction." & vbCrLf & _
                  "Veuillez contacter le support technique avec le nom de la procédure : '" & procedureName & "'." & vbCrLf & vbCrLf & _
                  "Détail technique : " & errObj.Description
    Else
        ' Message générique pour toutes les autres erreurs.
        userMsg = "Une erreur inattendue est survenue dans la procédure '" & procedureName & "'." & vbCrLf & vbCrLf & _
                  "Détail : " & errObj.Description & " (Erreur " & errObj.Number & ")"
    End If
    MsgBox userMsg, vbCritical, "Erreur - ExcelToSAP"

    ' Nettoyage systématique après une erreur
    RestoreExcel
    offSAP
    End If
End Sub

'====================================================================================
' SECTION 4 : TEMPORISATION ET INTERACTION CLAVIER
'====================================================================================

'------------------------------------------------------------------------------------
' PROCÉDURE   : WaitSeconds
' DESCRIPTION : Met en pause l'exécution de la macro pour une durée spécifiée en secondes.
'               Utilise `DoEvents` pour que l'interface Excel reste réactive pendant l'attente.
' PARAMÈTRES  :
'   - seconds (Double) : Le nombre de secondes à attendre.
'------------------------------------------------------------------------------------
Public Sub WaitSeconds(ByVal seconds As Double)
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + seconds
        ' Permet à Excel de traiter d'autres événements pendant l'attente.
        DoEvents
    Loop
End Sub

'====================================================================================
' FONCTION    : NumLock
' DESCRIPTION :
'   Verifie l'etat de la touche de verrouillage numerique (NumLock).
'   Si elle est desactivee, cette fonction l'active automatiquement pour eviter
'   des problemes de saisie avec SendKeys.
'====================================================================================
Public Function NumLock() As Boolean
    NumLock = KeyState(c_Key_NumLock) ' Lit l'etat actuel de NumLock

    If (NumLock = False) Then
        SendKeys "{NUMLOCK}", True ' Active NumLock si necessaire
    End If

End Function

'====================================================================================
' FONCTION    : KeyState
' DESCRIPTION :
'   Retourne l'etat (vrai/faux) d'une touche virtuelle Windows en interrogeant
'   l'API GetKeyState.
'====================================================================================
Public Function KeyState(lKey As Long) As Boolean
    KeyState = CBool(GetKeyState(lKey)) ' Conversion du resultat API en booleen
End Function

'====================================================================================
' PROCEDURE   : OptimizeExcel
' DESCRIPTION :
'   Desactive les fonctionnalites d'Excel qui ralentissent l'execution des macros
'   (mise a jour de l'ecran, evenements, calculs automatiques).
'   A appeler au debut des procedures longues.
'====================================================================================
Public Sub OptimizeExcel()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

'====================================================================================
' PROCEDURE   : RestoreExcel
' DESCRIPTION :
'   Retablit les parametres Excel par defaut apres l'execution d'une macro.
'   A appeler a la fin des procedures longues, y compris dans les blocs de
'   gestion d'erreurs.
'====================================================================================
Public Sub RestoreExcel()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'----------------------------------------------------
'------------- Fonctions et procedures utilitaires --
'----------------------------------------------------

Public Sub DateSemain()
'calcule les dates de la semaine en cours.

    ' --- D?claration des variables locales pour les calculs ---
    Dim localDebutSemaine As Date
    Dim localDebutSemaine_Avant As Date
    Dim localFinSemaine As Date
    Dim localAgingDate As Date
    Dim localDebutAnnee As Date
    Dim localFinAnnee As Date


    '----------- Pour une seule semaine -----------
    Dim numeroJour As Integer
    numeroJour = Weekday(Date)
    
    'D?finir la date pour laquelle vous voulez d?tecter le num?ro de la semaine
    Dim MaDate As Date
    MaDate = Date '- 7
    
    'D?terminer le num?ro de la semaine
    Dim NumSemaine As Integer
    ' Utilisation de IsoWeekNum pour gerer correctement les semaines a cheval sur deux annees
    NumSemaine = Application.WorksheetFunction.IsoWeekNum(MaDate)
    
    
    ' Lire la semaine S-1 si le jour est un dimanche
    ' If numeroJour = 1 Then NumSemaine = NumSemaine - 1
     
    ' Obtenir l'ann?e en cours
    Dim AnneeEnCours As Integer
    AnneeEnCours = Year(MaDate)
    
    ' Ajustement de l'annee si la semaine ISO appartient a l'annee suivante ou precedente
    If NumSemaine = 1 And Month(MaDate) = 12 Then
        AnneeEnCours = AnneeEnCours + 1
    ElseIf NumSemaine >= 52 And Month(MaDate) = 1 Then
        AnneeEnCours = AnneeEnCours - 1
    End If

    ' Concat?ner le num?ro de semaine et l'ann?e en cours
    SemaineKPIs = Format(AnneeEnCours, "0000") & Format(NumSemaine, "00") ' Format YYYYWW
    
    
    '--------------------------------------------------
    '-------------- commande manuelle -----------------
    
    SemaineKPIs = Format(AnneeEnCours, "0000") & Format(NumSemaine, "00")
    'SemaineKPIs = "202401"
    
    '--------------------------------------------------
    '--------------------------------------------------
    
    ' --- Calcul de la semaine pr?c?dente (SemaineKPIs_Avant) ---
    Dim MaDate_Avant As Date
    MaDate_Avant = MaDate - 7 ' Date il y a 7 jours
    
    Dim NumSemaine_Avant As Integer
    NumSemaine_Avant = Application.WorksheetFunction.IsoWeekNum(MaDate_Avant)
    
    Dim AnneeEnCours_Avant As Integer
    AnneeEnCours_Avant = Year(MaDate_Avant)
    
    ' Ajustement de l'ann?e pour la semaine pr?c?dente, si n?cessaire
    If NumSemaine_Avant = 1 And Month(MaDate_Avant) = 12 Then
        AnneeEnCours_Avant = AnneeEnCours_Avant + 1
    ElseIf NumSemaine_Avant >= 52 And Month(MaDate_Avant) = 1 Then
        AnneeEnCours_Avant = AnneeEnCours_Avant - 1
    End If
    SemaineKPIs_Avant = Format(AnneeEnCours_Avant, "0000") & Format(NumSemaine_Avant, "00")
    
    ' Trouver le d?but de la semaine (lundi) pour la date sp?cifi?e
    localDebutSemaine = MaDate - Weekday(MaDate, vbMonday) + 1
    localDebutSemaine_Avant = localDebutSemaine '- 7
    
    ' Calculer la date de debut de semaine en fonction du numero de semaine
    'DebutSemaine = DateAdd("d", (1 - Weekday(MaDate, vbMonday)) + (NumSemaine - 1) * 7, MaDate)

    ' Ajouter dix jours dans le passe
    localDebutSemaine = localDebutSemaine ' - 10


    ' Trouver le dimanche de la meme semaine
    localFinSemaine = localDebutSemaine + 6
    
    ' Ajouter dix jours dans le futur
    localFinSemaine = localFinSemaine ' + 10
    
    ' Date de reference pour calculer l'anciennete des ordres (aging wo)
    localAgingDate = localFinSemaine - 10

    ' Debut et fin d'un mois glissant pour les PR, PO, SES
    g_DebutMois = Date - 30
    g_FinMois = Date
    
    ' Debut et fin d'annee
    localDebutAnnee = DateSerial(Year(Date), 1, 1)
    localFinAnnee = DateSerial(Year(Date), 12, 31)

    ' --- Assignation et formatage final vers les variables globales ---
    DebutSemaine = Format(localDebutSemaine, "dd.mm.yyyy")
    DebutSemaine_Avant = Format(localDebutSemaine_Avant, "dd.mm.yyyy")
    FinSemaine = Format(localFinSemaine, "dd.mm.yyyy")
    AgingDate = Format(localAgingDate, "dd.mm.yyyy")
    g_DebutAnnee = Format(localDebutAnnee, "dd.mm.yyyy")
    g_FinAnnee = Format(localFinAnnee, "dd.mm.yyyy")
End Sub

Public Sub FermerFichierExcel(ByVal nomFichier As String)
    Dim classeur As Workbook
    Dim fichierTrouve As Boolean
    Dim startTime As Double
    Const attenteMaxEnSecondes As Double = 2
    
    ' Initialiser la variable a False
    fichierTrouve = False

    ' Parcourir tous les classeurs ouverts
    For Each classeur In Workbooks
        ' Verifier si le classeur a le meme nom que celui recherche
        If classeur.name = nomFichier Then
            ' Le fichier est trouve
            fichierTrouve = True

            ' Fermer le classeur sans sauvegarder les modifications
            classeur.Close SaveChanges:=False
            Exit For
        End If
    Next classeur

    ' Si le fichier n'est pas trouve, attendre un certain temps
    If Not fichierTrouve Then
        startTime = Timer
        Do While Timer < startTime + attenteMaxEnSecondes
            DoEvents
        Loop
    End If
    
    ' Rechercher a nouveau et fermer le fichier s'il est ouvert
    On Error Resume Next
    Set classeur = Workbooks(nomFichier)
    If Not classeur Is Nothing Then
        classeur.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub

Public Sub RenommerColonnes(chemin As String, nomFichier As String)
    Dim classeur As Workbook
    Dim feuille As Worksheet

    ' Specifier le chemin complet du fichier Excel
    Dim cheminFichier As String
    cheminFichier = chemin & "\" & nomFichier

    ' Ouvrir le fichier Excel
    Set classeur = Workbooks.Open(cheminFichier)

    ' Referencer la premiere feuille du classeur
    Set feuille = classeur.Sheets(1)

    ' Renommer les colonnes
    feuille.Cells(1, 1).value = "Planning plant"
    feuille.Cells(1, 2).value = "Operation WorkCenter"
    feuille.Cells(1, 3).value = "Work"
    feuille.Cells(1, 4).value = "Created on"

    ' Enregistrer et fermer le fichier Excel
    classeur.Close SaveChanges:=True
End Sub

Public Sub ChargerDonnees(path As String, fileName As String, sheetName As String)
    ' D?claration des variables
    Dim cheminSource As String
    Dim cheminDestination As String
    Dim feuilleDestination As Worksheet
    
    ' Definir les valeurs des variables
    cheminSource = path & "\" & fileName
    cheminDestination = ThisWorkbook.path ' Utilise le chemin du classeur add-in
    Set feuilleDestination = ThisWorkbook.Sheets(sheetName)
    
    ThisWorkbook.Save
    
    ' Effacer le contenu de la feuille destination
    feuilleDestination.Cells.ClearContents
    
    ' Charger les donnees depuis le fichier source vers la feuille destination
    ChargerDonneesFeuille cheminSource, cheminDestination, feuilleDestination
End Sub

Private Sub ChargerDonneesFeuille(cheminSource As String, cheminDestination As String, feuilleDestination As Worksheet)
    ' Charger les donnees depuis une feuille source vers une feuille destination
    
    Dim classeurSource As Workbook
    Dim wsSource As Worksheet
    Dim rngSource As Range
    
    ' Ouvrir le classeur source en tant que workbook
    Set classeurSource = Workbooks.Open(cheminSource, ReadOnly:=True)
    
    ' Utilise la première feuille (index 1) pour plus de robustesse que le nom "sheet1"
    Set wsSource = classeurSource.Sheets(1)
    Set rngSource = wsSource.UsedRange
    
    ' Transfert direct des valeurs pour éviter l'usage du presse-papiers
    ' Note: On utilise Value2 pour plus de rapidité et éviter les conversions de format
    If rngSource.Cells.CountLarge > 0 Then
        feuilleDestination.Range(feuilleDestination.Cells(1, 1), _
            feuilleDestination.Cells(rngSource.Rows.count, rngSource.Columns.count)).value = rngSource.Value2
    End If
    
    ' Fermer le classeur source sans enregistrer les modifications
    classeurSource.Close False
    
End Sub

Public Sub RecreerPivotTable(PTSheet As String, sheetDestination As String, NamePT As String)
    ' Declarer les variables
    Dim feuilleSource As Worksheet
    Dim tableauDynamique As PivotTable
    Dim champConfirmation As PivotField
    Dim plageSource As Range

    ' Specifier la feuille source
    Set feuilleSource = ThisWorkbook.Sheets(sheetDestination)

    ' Supprimer le tableau croise dynamique s'il existe
    On Error Resume Next
    feuilleSource.PivotTables(NamePT).TableRange2.Clear
    feuilleSource.PivotTables(NamePT).PivotCache.Refresh
    feuilleSource.PivotTables(NamePT).Delete
    On Error GoTo 0
    
    ' Specifier la plage source (Used Range) dans la feuille "KPIs"
    Set plageSource = ThisWorkbook.Sheets(PTSheet).UsedRange

    ' Creer un nouveau tableau croise dynamique a partir de la cellule A1
    Set tableauDynamique = feuilleSource.PivotTableWizard(TableDestination:=feuilleSource.Range("A1"), _
                                                           TableName:=NamePT, SourceType:=xlDatabase, SourceData:=plageSource)

    ' Ajouter le champ "confirmation" en tant que champ de ligne dans le tableau crois? dynamique
    Set champConfirmation = tableauDynamique.PivotFields("confirmation")
    champConfirmation.Orientation = xlRowField
    
    ' Desactiver les totaux generaux et les totaux partiels
    tableauDynamique.RowGrand = False
    tableauDynamique.ColumnGrand = False
    champConfirmation.Subtotals(1) = False ' Totals for Rows
    champConfirmation.Subtotals(2) = False ' Totals for Columns


    ' Ajouter le champ "confirmation" en tant que seule valeur dans le tableau crois? dynamique
    champConfirmation.Orientation = xlDataField
    champConfirmation.Function = xlCount ' Vous pouvez ajuster la fonction selon vos besoins

    On Error Resume Next
    With ActiveSheet.PivotTables("TCD_CNF").PivotFields("Confirmation")
        .PivotItems("(blank)").Visible = False
    End With

    ' Actualiser le tableau croise dynamique
    tableauDynamique.RefreshTable
    
End Sub

'-------------------------------------------------------------------------------
' SUB : ExportToExcel
' DESCRIPTION : Exporte les donnees d'une grille SAP vers un fichier Excel.
'               Utilise le menu contextuel "&XXL".
'-------------------------------------------------------------------------------
Public Sub ExportToExcel(ByVal fileName As String)
    On Error GoTo SapErrorHandler

    If Not SAP.IsSAPConnectionAlive() Then Exit Sub

    ' Tente d'exporter depuis une grille. Si la grille n'existe pas,
    ' le menu contextuel echouera gracieusement grace a "On Error Resume Next".
    On Error Resume Next
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    g_Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    On Error GoTo 0 ' Reactiver la gestion d'erreur normale
    
    g_Session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"
    g_Session.findById("wnd[1]/usr/chkCB_ALWAYS").Selected = True
    g_Session.findById("wnd[1]/tbar[0]/btn[0]").press

    modUtils.SaveFile GetSetting("PATH_EXPORT"), fileName
    Exit Sub

SapErrorHandler:
    DisplayAndLogError "ExportToExcel", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : SaveFile
' DESCRIPTION : Sauvegarde un fichier extrait par SAP dans un chemin specifie.
'               Utilise g_Session pour interagir avec la boite de dialogue SAP.
'-------------------------------------------------------------------------------
Public Sub SaveFile(ByVal path As String, ByVal File_Name As String)
    On Error GoTo SapErrorHandler
    Dim cheminFichier As String
    
    ' Verifier si le dossier de destination existe, sinon le creer
    ' --- AJOUT : Nettoie le fichier de destination s'il existe déjà ---
    ClearFileIfExists path, File_Name
    
    If Dir(path, vbDirectory) = "" Then
        MkDir path
        LogMessage ">> Dossier cree : " & path
    End If

    ' Interagir avec la boite de dialogue de sauvegarde SAP (wnd[1])
    g_Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = File_Name
    g_Session.findById("wnd[1]/usr/ctxtDY_PATH").text = path ' Definir le chemin
    g_Session.findById("wnd[1]/tbar[0]/btn[11]").press ' Cliquer sur "Enregistrer" ou "Remplacer"

    LogMessage ">> Fichier exporte avec succes : " & path & File_Name
    Exit Sub
    
SapErrorHandler:
    DisplayAndLogError "SaveFile", Err
End Sub

'-------------------------------------------------------------------------------
' SUB : ClearFileIfExists (Privée)
' DESCRIPTION : Vérifie si un fichier existe et le supprime si c'est le cas.
'               Utilisée pour garantir que les exports SAP sont toujours frais.
'-------------------------------------------------------------------------------
Public Sub ClearFileIfExists(ByVal path As String, ByVal fileName As String)
    Dim fullPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCell As Range
    
    ' S'assure que le chemin se termine bien par un antislash
    If Right(path, 1) <> "\" Then
        fullPath = path & "\" & fileName
    Else
        fullPath = path & fileName
    End If

    ' Si le fichier existe, on le vide au lieu de le supprimer
    If Dir(fullPath) <> "" Then
        On Error Resume Next ' Gère le cas où le fichier serait en lecture seule ou verrouillé
        Set wb = Workbooks.Open(fullPath)
        If Not wb Is Nothing Then
            Set ws = wb.Sheets(1)
            Set lastCell = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not lastCell Is Nothing Then
                If lastCell.row > 1 Then
                    ws.Rows("2:" & lastCell.row).Delete
                End If
            End If
            wb.Close SaveChanges:=True
        End If
        On Error GoTo 0
    End If
End Sub

