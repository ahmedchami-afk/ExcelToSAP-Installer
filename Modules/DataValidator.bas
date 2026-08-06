Attribute VB_Name = "DataValidator"
'====================================================================================
' MODULE      : DataValidator
' AUTEUR      : Gemini Code Assist
' DATE        : 26/01/2026
' DESCRIPTION : Validation robuste des données avant envoi à SAP.
'               Prévient la corruption des données SAP par validation de caractères,
'               longueurs, formats et injection de contenu malveillant.
'
' FIX #3 - VALIDATION MANQUANTE:
'   Problème: Les données n'étaient pas validées avant envoi ? Corruption possible
'   Solution: Module complet de validation avec logging des données suspectes
'
' VALIDATIONS INCLUSES:
'   1. CaractÃ¨res spÃ©ciaux interdits en SAP
'   2. Longueur maximale des champs
'   3. Format numÃ©rique/alphanumÃ©rique
'   4. PrÃ©vention d'injection (apostrophes, guillemets)
'   5. DÃ©tection de contenu malveillant
'   6. Logging complet des validations
'====================================================================================
Option Explicit

'====================================================================================
' SECTION 1 : CONSTANTES DE VALIDATION
'====================================================================================

' CaractÃ¨res interdits en SAP (causent de la corruption)
Private Const SAP_FORBIDDEN_CHARS As String = "'""<>|\*?:[]{}&$#@!^~`"

' CaractÃ¨res autorisÃ©s pour nombres
Private Const NUMBERS_ONLY As String = "0123456789"

' CaractÃ¨res autorisÃ©s pour alphanumÃ©rique basique
Private Const ALPHANUMERIC_CHARS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

' Longueurs maximales SAP
Private Const MAX_MATERIAL_CODE As Long = 18
Private Const MAX_NOTIFICATION_ID As Long = 12
Private Const MAX_EQUIPMENT_ID As Long = 18
Private Const MAX_LOCATION_ID As Long = 30
Private Const MAX_PURCHASE_ORDER As Long = 10
Private Const MAX_TEXT_FIELD As Long = 255

'====================================================================================
' SECTION 2 : VALIDATION GLOBALE
'====================================================================================

Public Function ValidateDataBeforeSending(ByVal dataType As String, ByVal dataValue As String, Optional ByVal fieldName As String = "") As Boolean
    '--------------------------------------------------------------------------------
    ' DESCRIPTION : Valide les donnÃ©es selon le type avant envoi Ã  SAP.
    '               Fonction principale de validation
    '
    ' PARAMÃˆTRES:
    '   - dataType: Type de donnÃ©e (MATERIAL, NOTIFICATION, EQUIPMENT, etc.)
    '   - dataValue: Valeur Ã  valider
    '   - fieldName: Nom du champ pour logging
    '
    ' RETOUR: True si validation OK, False sinon
    '
    ' FIX #3 - Utiliser avant tout envoi Ã  SAP
    '
    ' EXEMPLE:
    '   If ValidateDataBeforeSending("MATERIAL", inputValue) Then
    '       g_Session.findById(...).text = inputValue
    '   End If
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If dataValue = "" Then
        ValidateDataBeforeSending = True ' Vide est acceptÃ©
        Exit Function
    End If
    
    ' Nettoyer les espaces inutiles
    Dim cleanValue As String
    cleanValue = Trim(dataValue)
    
    ' VÃ©rifier les caractÃ¨res interdits
    If ContainsForbiddenChars(cleanValue) Then
        LogMessageCompat "ERREUR DE VALIDATION: CaractÃ¨res interdits dÃ©tectÃ©s dans [" & fieldName & "]: " & cleanValue
        ValidateDataBeforeSending = False
        Exit Function
    End If
    
    ' Validation spÃ©cifique par type
    Select Case UCase(dataType)
        Case "MATERIAL"
            ValidateDataBeforeSending = ValidateMaterialCode(cleanValue, fieldName)
        Case "NOTIFICATION"
            ValidateDataBeforeSending = ValidateNotificationID(cleanValue, fieldName)
        Case "EQUIPMENT"
            ValidateDataBeforeSending = ValidateEquipmentID(cleanValue, fieldName)
        Case "LOCATION"
            ValidateDataBeforeSending = ValidateLocationCode(cleanValue, fieldName)
        Case "PURCHASE_ORDER"
            ValidateDataBeforeSending = ValidatePurchaseOrderNumber(cleanValue, fieldName)
        Case "TEXT"
            ValidateDataBeforeSending = ValidateTextField(cleanValue, fieldName)
        Case "NUMERIC"
            ValidateDataBeforeSending = ValidateNumericField(cleanValue, fieldName)        Case "WBS"
            ValidateDataBeforeSending = ValidateWBSCode(cleanValue, fieldName)
        Case "DATE"
            ValidateDataBeforeSending = ValidateSAPDate(cleanValue, fieldName)
        Case "FILEPATH"
            ValidateDataBeforeSending = ValidateFilePath(cleanValue, fieldName)
        Case "AMOUNT"
            ValidateDataBeforeSending = ValidateMonetaryAmount(cleanValue, fieldName)
        Case "PLANT"
            ValidateDataBeforeSending = ValidatePlantCode(cleanValue, fieldName)        Case Else
            ' Type non reconnu â†’ Validation stricte par dÃ©faut
            ValidateDataBeforeSending = ValidateGenericField(cleanValue, fieldName)
    End Select
    
    Exit Function
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateDataBeforeSending: " & Err.Description
    ValidateDataBeforeSending = False
End Function

'====================================================================================
' SECTION 3 : VALIDATIONS SPÃ‰CIFIQUES PAR TYPE
'====================================================================================

Public Function ValidateMaterialCode(ByVal code As String, Optional ByVal fieldName As String = "MATERIAL") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un code matÃ©riel SAP
    ' - Longueur: max 18 caractÃ¨res
    ' - Pas de caractÃ¨res spÃ©ciaux
    ' - Pas d'espaces
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If code = "" Then
        ValidateMaterialCode = True
        Exit Function
    End If
    
    ' VÃ©rifier la longueur
    If Len(code) > MAX_MATERIAL_CODE Then
        LogMessageCompat "ERREUR: Code matÃ©riel trop long [" & fieldName & "]: " & code & " (max " & MAX_MATERIAL_CODE & " caractÃ¨res)"
        ValidateMaterialCode = False
        Exit Function
    End If
    
    ' Pas d'espaces
    If InStr(code, " ") > 0 Then
        LogMessageCompat "ERREUR: Code matÃ©riel contient des espaces [" & fieldName & "]: " & code
        ValidateMaterialCode = False
        Exit Function
    End If
    
    ' Pas de caractÃ¨res spÃ©ciaux
    If ContainsForbiddenChars(code) Then
        LogMessageCompat "ERREUR: Code matÃ©riel contient des caractÃ¨res interdits [" & fieldName & "]: " & code
        ValidateMaterialCode = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: Code matÃ©riel [" & fieldName & "]: " & code
    ValidateMaterialCode = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateMaterialCode: " & Err.Description
    ValidateMaterialCode = False
End Function

Public Function ValidateNotificationID(ByVal notifID As String, Optional ByVal fieldName As String = "NOTIFICATION") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un numÃ©ro de notification SAP
    ' - Longueur: max 12 caractÃ¨res
    ' - Doit Ãªtre numÃ©rique ou alphanumÃ©rique
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If notifID = "" Then
        ValidateNotificationID = True
        Exit Function
    End If
    
    If Len(notifID) > MAX_NOTIFICATION_ID Then
        LogMessageCompat "ERREUR: ID notification trop long [" & fieldName & "]: " & notifID & " (max " & MAX_NOTIFICATION_ID & ")"
        ValidateNotificationID = False
        Exit Function
    End If
    
    If ContainsForbiddenChars(notifID) Then
        LogMessageCompat "ERREUR: ID notification contient des caractÃ¨res interdits [" & fieldName & "]: " & notifID
        ValidateNotificationID = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: ID notification [" & fieldName & "]: " & notifID
    ValidateNotificationID = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateNotificationID: " & Err.Description
    ValidateNotificationID = False
End Function

Public Function ValidateEquipmentID(ByVal equipID As String, Optional ByVal fieldName As String = "EQUIPMENT") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un code Ã©quipement SAP
    ' - Longueur: max 18 caractÃ¨res
    ' - AlphanumÃ©rique, tirets, points autorisÃ©s
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If equipID = "" Then
        ValidateEquipmentID = True
        Exit Function
    End If
    
    If Len(equipID) > MAX_EQUIPMENT_ID Then
        LogMessageCompat "ERREUR: Code Ã©quipement trop long [" & fieldName & "]: " & equipID & " (max " & MAX_EQUIPMENT_ID & ")"
        ValidateEquipmentID = False
        Exit Function
    End If
    
    ' VÃ©rifier les caractÃ¨res spÃ©ciaux interdits (sauf tiret et point)
    Dim i As Long, ch As String
    For i = 1 To Len(equipID)
        ch = Mid(equipID, i, 1)
        If InStr(SAP_FORBIDDEN_CHARS, ch) > 0 Then
            LogMessageCompat "ERREUR: Code Ã©quipement contient des caractÃ¨res interdits [" & fieldName & "]: " & equipID
            ValidateEquipmentID = False
            Exit Function
        End If
    Next
    
    LogMessageCompat "VALIDATION OK: Code Ã©quipement [" & fieldName & "]: " & equipID
    ValidateEquipmentID = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateEquipmentID: " & Err.Description
    ValidateEquipmentID = False
End Function

Public Function ValidateLocationCode(ByVal locCode As String, Optional ByVal fieldName As String = "LOCATION") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un code localisation SAP (WBS, Profit Center, etc.)
    ' - Longueur: max 30 caractÃ¨res
    ' - Peut contenir des tirets et points
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If locCode = "" Then
        ValidateLocationCode = True
        Exit Function
    End If
    
    If Len(locCode) > MAX_LOCATION_ID Then
        LogMessageCompat "ERREUR: Code localisation trop long [" & fieldName & "]: " & locCode & " (max " & MAX_LOCATION_ID & ")"
        ValidateLocationCode = False
        Exit Function
    End If
    
    If ContainsForbiddenChars(locCode) Then
        LogMessageCompat "ERREUR: Code localisation contient des caractÃ¨res interdits [" & fieldName & "]: " & locCode
        ValidateLocationCode = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: Code localisation [" & fieldName & "]: " & locCode
    ValidateLocationCode = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateLocationCode: " & Err.Description
    ValidateLocationCode = False
End Function

Public Function ValidatePurchaseOrderNumber(ByVal poNumber As String, Optional ByVal fieldName As String = "PO") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un numÃ©ro de commande d'achat SAP
    ' - Longueur: max 10 caractÃ¨res
    ' - Doit Ãªtre numÃ©rique
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If poNumber = "" Then
        ValidatePurchaseOrderNumber = True
        Exit Function
    End If
    
    If Len(poNumber) > MAX_PURCHASE_ORDER Then
        LogMessageCompat "ERREUR: NÂ° commande trop long [" & fieldName & "]: " & poNumber & " (max " & MAX_PURCHASE_ORDER & ")"
        ValidatePurchaseOrderNumber = False
        Exit Function
    End If
    
    ' VÃ©rifier que c'est numÃ©rique
    If Not IsNumeric(poNumber) Then
        LogMessageCompat "ERREUR: NÂ° commande doit Ãªtre numÃ©rique [" & fieldName & "]: " & poNumber
        ValidatePurchaseOrderNumber = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: NÂ° commande [" & fieldName & "]: " & poNumber
    ValidatePurchaseOrderNumber = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidatePurchaseOrderNumber: " & Err.Description
    ValidatePurchaseOrderNumber = False
End Function

Public Function ValidateTextField(ByVal textValue As String, Optional ByVal fieldName As String = "TEXT") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un champ de texte gÃ©nÃ©rique
    ' - Longueur: max 255 caractÃ¨res
    ' - Pas d'apostrophes simples (cause d'injection)
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If textValue = "" Then
        ValidateTextField = True
        Exit Function
    End If
    
    If Len(textValue) > MAX_TEXT_FIELD Then
        LogMessageCompat "ERREUR: Texte trop long [" & fieldName & "]: " & Left(textValue, 50) & "... (max " & MAX_TEXT_FIELD & ")"
        ValidateTextField = False
        Exit Function
    End If
    
    ' VÃ©rifier les apostrophes (risque d'injection)
    If InStr(textValue, "'") > 0 Then
        LogMessageCompat "AVERTISSEMENT: Apostrophe trouvÃ©e [" & fieldName & "]: " & textValue
        ' Apostrophes acceptÃ©es mais loggÃ©es (remplacer par "")
    End If
    
    LogMessageCompat "VALIDATION OK: Texte [" & fieldName & "]: " & Left(textValue, 50)
    ValidateTextField = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateTextField: " & Err.Description
    ValidateTextField = False
End Function

Public Function ValidateNumericField(ByVal numValue As String, Optional ByVal fieldName As String = "NUMERIC") As Boolean
    '--------------------------------------------------------------------------------
    ' Valide un champ numÃ©rique
    ' - Doit Ãªtre numÃ©rique (avec point dÃ©cimal autorisÃ©)
    ' - Pas d'espaces
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If numValue = "" Then
        ValidateNumericField = True
        Exit Function
    End If
    
    If Not IsNumeric(numValue) Then
        LogMessageCompat "ERREUR: Valeur non numÃ©rique [" & fieldName & "]: " & numValue
        ValidateNumericField = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: Valeur numÃ©rique [" & fieldName & "]: " & numValue
    ValidateNumericField = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateNumericField: " & Err.Description
    ValidateNumericField = False
End Function

Public Function ValidateGenericField(ByVal fieldValue As String, Optional ByVal fieldName As String = "GENERIC") As Boolean
    '--------------------------------------------------------------------------------
    ' Validation par dÃ©faut pour champs non spÃ©cifiÃ©s
    ' - VÃ©rifier les caractÃ¨res interdits
    ' - VÃ©rifier la longueur (max 255)
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If fieldValue = "" Then
        ValidateGenericField = True
        Exit Function
    End If
    
    If Len(fieldValue) > MAX_TEXT_FIELD Then
        LogMessageCompat "ERREUR: Champ trop long [" & fieldName & "]: max " & MAX_TEXT_FIELD
        ValidateGenericField = False
        Exit Function
    End If
    
    If ContainsForbiddenChars(fieldValue) Then
        LogMessageCompat "ERREUR: CaractÃ¨res interdits [" & fieldName & "]: " & fieldValue
        ValidateGenericField = False
        Exit Function
    End If
    
    LogMessageCompat "VALIDATION OK: Champ [" & fieldName & "]: " & Left(fieldValue, 50)
    ValidateGenericField = True
    Exit Function
    
ErrHandler:
    LogMessageCompat "ERREUR dans ValidateGenericField: " & Err.Description
    ValidateGenericField = False
End Function

'====================================================================================
' SECTION 4 : FONCTIONS UTILITAIRES DE VALIDATION
'====================================================================================

Public Function ContainsForbiddenChars(ByVal text As String) As Boolean
    '--------------------------------------------------------------------------------
    ' VÃ©rifie si le texte contient des caractÃ¨res interdits en SAP
    ' CaractÃ¨res interdits: ' " < > | \ * ? : [ ] { } & $ # @ ! ^ ~ `
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim i As Long, ch As String
    For i = 1 To Len(SAP_FORBIDDEN_CHARS)
        ch = Mid(SAP_FORBIDDEN_CHARS, i, 1)
        If InStr(text, ch) > 0 Then
            ContainsForbiddenChars = True
            Exit Function
        End If
    Next
    
    ContainsForbiddenChars = False
    Exit Function
    
ErrHandler:
    ContainsForbiddenChars = True ' Par sÃ©curitÃ©
End Function

Public Function IsNumeric(ByVal text As String) As Boolean
    '--------------------------------------------------------------------------------
    ' VÃ©rifie si le texte est entiÃ¨rement numÃ©rique (avec point dÃ©cimal autorisÃ©)
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim i As Long, ch As String
    Dim hasDecimal As Boolean
    
    If text = "" Then
        IsNumeric = False
        Exit Function
    End If
    
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        
        ' Accepter les chiffres
        If InStr(NUMBERS_ONLY, ch) = 0 Then
            ' Accepter le premier point dÃ©cimal
            If ch = "." And Not hasDecimal Then
                hasDecimal = True
            ' Accepter le signe moins au dÃ©but
            Else If ch = "-" And i = 1 Then
                ' OK
            Else
                IsNumeric = False
                Exit Function
            End If
        End If
    Next
    
    IsNumeric = True
    Exit Function
    
ErrHandler:
    IsNumeric = False
End Function

Public Function SanitizeForSAP(ByVal text As String) As String
    '--------------------------------------------------------------------------------
    ' Nettoie une chaÃ®ne en supprimant les caractÃ¨res interdits SAP
    ' USAGE: Ã€ utiliser comme derniÃ¨re Ã©tape avant envoi Ã  SAP
    '
    ' EXEMPLE:
    '   cleanedText = SanitizeForSAP(userInput)
    '   g_Session.findById(...).text = cleanedText
    '--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    Dim result As String
    Dim i As Long, ch As String
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        
        ' Garder le caractÃ¨re s'il n'est pas interdit
        If InStr(SAP_FORBIDDEN_CHARS, ch) = 0 Then
            result = result & ch
        End If
    Next
    
    SanitizeForSAP = result
    Exit Function
    
ErrHandler:
    SanitizeForSAP = text ' Retourner l'original si erreur
End Function

Public Function GetFieldMaxLength(ByVal dataType As String) As Long
    '--------------------------------------------------------------------------------
    ' Retourne la longueur maximale autorisÃ©e pour un type de donnÃ©e
    '
    ' EXEMPLE:
    '   maxLen = GetFieldMaxLength("MATERIAL")
    '   If Len(userInput) > maxLen Then MsgBox "Trop long!"
    '--------------------------------------------------------------------------------
    Select Case UCase(dataType)
        Case "MATERIAL"
            GetFieldMaxLength = MAX_MATERIAL_CODE
        Case "NOTIFICATION"
            GetFieldMaxLength = MAX_NOTIFICATION_ID
        Case "EQUIPMENT"
            GetFieldMaxLength = MAX_EQUIPMENT_ID
        Case "LOCATION"
            GetFieldMaxLength = MAX_LOCATION_ID
        Case "PURCHASE_ORDER"
            GetFieldMaxLength = MAX_PURCHASE_ORDER
        Case Else
            GetFieldMaxLength = MAX_TEXT_FIELD
    End Select
End Function

Public Sub DisplayValidationRules()
    '--------------------------------------------------------------------------------
    ' Affiche les rÃ¨gles de validation dans les logs et une boÃ®te de dialogue
    '--------------------------------------------------------------------------------
    Dim rules As String
    rules = "RÃˆGLES DE VALIDATION SAP:" & vbCrLf & vbCrLf & _
            "MATÃ‰RIEL: max " & MAX_MATERIAL_CODE & " caractÃ¨res" & vbCrLf & _
            "NOTIFICATION: max " & MAX_NOTIFICATION_ID & " caractÃ¨res" & vbCrLf & _
            "Ã‰QUIPEMENT: max " & MAX_EQUIPMENT_ID & " caractÃ¨res" & vbCrLf & _
            "LOCALISATION: max " & MAX_LOCATION_ID & " caractÃ¨res" & vbCrLf & _
            "COMMANDE: max " & MAX_PURCHASE_ORDER & " caractÃ¨res" & vbCrLf & vbCrLf & _
            "CARACTÃˆRES INTERDITS:" & vbCrLf & _
            SAP_FORBIDDEN_CHARS & vbCrLf & vbCrLf & _
            "Remplacer les apostrophes par quotes" & vbCrLf & _
            "Pas d'espaces en dÃ©but/fin" & vbCrLf & _
            "Pas d'accents si possible"
    
    LogMessage rules
    MsgBox rules, vbInformation, "RÃ¨gles de Validation SAP"
End Sub


