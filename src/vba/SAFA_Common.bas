Attribute VB_Name = "SAFA_Common"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SAFA_COMMON v10.0
' ==============================================================================
' Description: Module utilitaire commun - Fonctions partagées centralisées
' Ce module élimine la duplication de code entre les différents modules
' ==============================================================================

' --- CONSTANTES GLOBALES CENTRALISÉES ---
Public Const SAFA_VERSION As String = "10.0"
Public Const SAFA_BUILD As String = "2024.12.001"

' Constantes de configuration par défaut
Public Const DEFAULT_TOLERANCE As Double = 100
Public Const DEFAULT_SOL_ID As String = "799"
Public Const DEFAULT_CURRENCY As String = "XAF"
Public Const MAX_ROWS_MEMORY As Long = 500000

' Constantes Forensiques
Public Const BENFORD_THRESHOLD As Double = 0.15
Public Const BENFORD_MAD_EXCELLENT As Double = 0.006
Public Const BENFORD_MAD_ACCEPTABLE As Double = 0.012
Public Const BENFORD_MAD_MARGINAL As Double = 0.015
Public Const Z_SCORE_WARNING As Double = 2.5
Public Const Z_SCORE_CRITICAL As Double = 3.5
Public Const WEEKEND_AMOUNT_THRESHOLD As Double = 50000
Public Const ROUND_NUMBER_THRESHOLD As Double = 100000
Public Const SPLIT_DETECTION_THRESHOLD As Double = 50000
Public Const LAYERING_MIN_COUNT As Integer = 4

' Constantes Réglementaires
Public Const COBAC_SUSPENS_LIMIT_DAYS As Long = 90
Public Const COBAC_TRANSIT_LIMIT_DAYS As Long = 7
Public Const LAB_THRESHOLD_XAF As Double = 5000000
Public Const LARGE_EXPOSURE_RATIO As Double = 0.25

' Constantes Sécurité
Public Const SESSION_TIMEOUT_MINUTES As Integer = 30
Public Const MAX_LOGIN_ATTEMPTS As Integer = 3
Public Const HASH_SEED As Long = 5381

' Constantes Scoring Risque
Public Const RISK_CRITICAL_THRESHOLD As Integer = 70
Public Const RISK_HIGH_THRESHOLD As Integer = 50
Public Const RISK_MEDIUM_THRESHOLD As Integer = 30

' Couleurs standard
Public Const COLOR_HEADER As Long = 3368754      ' RGB(0, 51, 102)
Public Const COLOR_CRITICAL As Long = 255        ' Rouge
Public Const COLOR_HIGH As Long = 42495          ' Orange
Public Const COLOR_MEDIUM As Long = 65535        ' Jaune
Public Const COLOR_LOW As Long = 5287936         ' Vert
Public Const COLOR_COMPLIANT As Long = 5287936   ' Vert
Public Const COLOR_WARNING As Long = 65535       ' Jaune
Public Const COLOR_ERROR As Long = 255           ' Rouge

' --- TYPES PARTAGÉS ---
Public Type AlertRecord
    Reference As String
    Category As String
    Description As String
    Severity As String
    Account As String
    Details As String
    Amount As Double
    TransactionDate As Date
    RiskScore As Integer
    SLA As String
End Type

Public Type AuditLogEntry
    Timestamp As Date
    LogType As String
    Username As String
    Action As String
    Details As String
    Hash As String
    PreviousHash As String
End Type

' ==============================================================================
' 1. FONCTION DE HASH UNIFIÉE (CORRIGE BUG-001)
' ==============================================================================

Public Function ComputeHash(text As String) As String
    ' Fonction de hash DJB2 - UNIQUE pour tout le système
    ' Utilisée par: Audit Trail, Intégrité, Mots de passe, Transactions
    Dim i As Long
    Dim h As Double
    Dim c As Long

    h = HASH_SEED

    For i = 1 To Len(text)
        c = Asc(Mid(text, i, 1))
        ' Utiliser Double pour éviter overflow sur Long
        h = ((h * 33) + c)
        ' Modulo pour garder dans les limites
        If h > 2147483647 Then
            h = h - 4294967296#
        End If
    Next i

    ' Convertir en hex 8 caractères (toujours positif)
    ComputeHash = Right("00000000" & Hex(Abs(h) And &H7FFFFFFF), 8)
End Function

Public Function ComputeHashChain(currentData As String, previousHash As String) As String
    ' Hash chainé pour audit trail (blockchain-like)
    ComputeHashChain = ComputeHash(currentData & "|" & previousHash)
End Function

' ==============================================================================
' 2. FONCTIONS DE CONVERSION SÉCURISÉES
' ==============================================================================

Public Function SafeText(val As Variant, Optional maxLength As Long = 0) As String
    ' Conversion sécurisée en texte avec option de troncature
    On Error Resume Next

    If IsError(val) Then
        SafeText = ""
    ElseIf IsEmpty(val) Or IsNull(val) Then
        SafeText = ""
    Else
        SafeText = Trim(CStr(val))
        SafeText = CleanControlChars(SafeText)
        If maxLength > 0 And Len(SafeText) > maxLength Then
            SafeText = Left(SafeText, maxLength)
        End If
    End If

    On Error GoTo 0
End Function

Public Function SafeVal(val As Variant, Optional defaultValue As Double = 0) As Double
    ' Fonction "Bulldozer" : Convertit n'importe quoi en nombre
    On Error Resume Next

    If IsError(val) Then SafeVal = defaultValue: Exit Function
    If IsEmpty(val) Or Trim(CStr(val)) = "" Then SafeVal = defaultValue: Exit Function

    If IsNumeric(val) And Not IsDate(val) Then
        SafeVal = CDbl(val)
        Exit Function
    End If

    Dim s As String
    s = CStr(val)

    ' Nettoyage avancé
    s = Replace(s, Chr(160), "")
    s = Replace(s, Chr(8239), "")
    s = Replace(s, " ", "")

    ' Devises multiples
    Dim currencies As Variant
    currencies = Array("XAF", "EUR", "USD", "GBP", "XOF", "NGN", "GHS", "KES", "ZAR", "MAD", "TND", "EGP", "FCFA", "F CFA")
    Dim curr As Variant
    For Each curr In currencies
        s = Replace(s, CStr(curr), "", , , vbTextCompare)
    Next curr

    ' Indicateurs débit/crédit
    s = Replace(s, "Cr", "", , , vbTextCompare)
    s = Replace(s, "Dr", "", , , vbTextCompare)
    s = Replace(s, """", "")
    s = Replace(s, "'", "")

    ' Gestion du signe
    Dim isNeg As Boolean: isNeg = False

    If Right(s, 1) = "-" Then
        isNeg = True
        s = Left(s, Len(s) - 1)
    ElseIf Left(s, 1) = "-" Then
        isNeg = True
        s = Mid(s, 2)
    ElseIf Left(s, 1) = "(" And Right(s, 1) = ")" Then
        isNeg = True
        s = Mid(s, 2, Len(s) - 2)
    ElseIf Left(s, 1) = "<" And Right(s, 1) = ">" Then
        isNeg = True
        s = Mid(s, 2, Len(s) - 2)
    End If

    ' Normaliser format numérique
    s = NormalizeNumberFormat(s)

    If IsNumeric(s) Then
        SafeVal = CDbl(s)
        If isNeg Then SafeVal = SafeVal * -1
    Else
        SafeVal = defaultValue
    End If

    On Error GoTo 0
End Function

Public Function SafeDate(val As Variant, Optional defaultDate As Date = 0) As Date
    ' Conversion sécurisée de date
    On Error Resume Next

    If IsError(val) Or IsEmpty(val) Or IsNull(val) Then
        SafeDate = defaultDate
        Exit Function
    End If

    If IsDate(val) Then
        SafeDate = CDate(val)
        Exit Function
    End If

    Dim s As String
    s = Trim(CStr(val))

    If IsDate(s) Then
        SafeDate = CDate(s)
    Else
        SafeDate = defaultDate
    End If

    On Error GoTo 0
End Function

Public Function SafeLong(val As Variant, Optional defaultValue As Long = 0) As Long
    On Error Resume Next
    If IsNumeric(val) Then
        SafeLong = CLng(val)
    Else
        SafeLong = defaultValue
    End If
    On Error GoTo 0
End Function

Private Function NormalizeNumberFormat(s As String) As String
    Dim dotPos As Long, commaPos As Long
    Dim dotCount As Long, commaCount As Long
    Dim i As Long

    For i = 1 To Len(s)
        If Mid(s, i, 1) = "." Then
            dotCount = dotCount + 1
            dotPos = i
        ElseIf Mid(s, i, 1) = "," Then
            commaCount = commaCount + 1
            commaPos = i
        End If
    Next i

    If dotCount = 0 And commaCount = 0 Then
        NormalizeNumberFormat = s
    ElseIf dotCount = 1 And commaCount = 0 Then
        NormalizeNumberFormat = s
    ElseIf dotCount = 0 And commaCount = 1 Then
        NormalizeNumberFormat = Replace(s, ",", ".")
    ElseIf dotCount > 1 And commaCount = 0 Then
        NormalizeNumberFormat = Replace(s, ".", "")
    ElseIf dotCount = 0 And commaCount > 1 Then
        NormalizeNumberFormat = Replace(s, ",", "")
    ElseIf dotCount >= 1 And commaCount >= 1 Then
        If dotPos > commaPos Then
            NormalizeNumberFormat = Replace(s, ",", "")
        Else
            s = Replace(s, ".", "")
            NormalizeNumberFormat = Replace(s, ",", ".")
        End If
    Else
        NormalizeNumberFormat = s
    End If
End Function

Private Function CleanControlChars(s As String) As String
    Dim i As Long, c As Long, result As String

    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        If c >= 32 Or c = 9 Or c = 10 Or c = 13 Then
            result = result & Mid(s, i, 1)
        End If
    Next i

    CleanControlChars = result
End Function

' ==============================================================================
' 3. FONCTIONS UTILITAIRES DE FEUILLES
' ==============================================================================

Public Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not ws Is Nothing
End Function

Public Function GetOrCreateSheet(nom As String, Optional clearIfExists As Boolean = False) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = Left(nom, 31)
    ElseIf clearIfExists Then
        ws.Cells.Clear
    End If

    Set GetOrCreateSheet = ws
End Function

Public Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If UCase(CStr(arr(i))) = UCase(val) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

' ==============================================================================
' 4. FORMATAGE UNIFIÉ
' ==============================================================================

Public Sub FormatHeader(rng As Range, Optional bgColor As Long = -1)
    If bgColor = -1 Then bgColor = COLOR_HEADER

    With rng
        .Font.Bold = True
        .Interior.Color = bgColor
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub FormatAlertCell(cell As Range, severity As String)
    Select Case UCase(severity)
        Case "CRITICAL"
            cell.Interior.Color = COLOR_CRITICAL
            cell.Font.Color = vbWhite
            cell.Font.Bold = True
        Case "FRAUD"
            cell.Interior.Color = RGB(139, 0, 0)
            cell.Font.Color = vbWhite
            cell.Font.Bold = True
        Case "HIGH"
            cell.Interior.Color = COLOR_HIGH
            cell.Font.Bold = True
        Case "MAJOR"
            cell.Interior.Color = RGB(255, 192, 0)
        Case "MEDIUM"
            cell.Interior.Color = COLOR_MEDIUM
        Case "COMPLIANCE"
            cell.Interior.Color = RGB(173, 216, 230)
        Case "LOW"
            cell.Interior.Color = COLOR_LOW
            cell.Font.Color = vbWhite
        Case Else
            cell.Interior.Color = RGB(211, 211, 211)
    End Select
End Sub

Public Sub ApplyConditionalFormatting(ws As Worksheet, rngPriority As Range)
    ' Supprimer formats existants
    rngPriority.FormatConditions.Delete

    ' CRITICAL = Rouge
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""CRITICAL""")
        .Interior.Color = COLOR_CRITICAL
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    ' HIGH = Orange
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""HIGH""")
        .Interior.Color = COLOR_HIGH
        .Font.Bold = True
    End With

    ' MEDIUM = Jaune
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""MEDIUM""")
        .Interior.Color = COLOR_MEDIUM
    End With

    ' LOW = Vert
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""LOW""")
        .Interior.Color = COLOR_LOW
        .Font.Color = vbWhite
    End With
End Sub

' ==============================================================================
' 5. SYSTÈME D'ALERTES UNIFIÉ
' ==============================================================================

Public Sub AddAlert(ByVal ws As Worksheet, ByRef rowNum As Long, _
                    ByVal reference As String, ByVal category As String, _
                    ByVal description As String, ByVal severity As String, _
                    ByVal account As String, ByVal details As String, _
                    ByVal amount As Double, Optional ByVal sla As String = "")

    ws.Cells(rowNum, 1).Value = reference
    ws.Cells(rowNum, 2).Value = category
    ws.Cells(rowNum, 3).Value = description
    ws.Cells(rowNum, 4).Value = severity
    ws.Cells(rowNum, 5).Value = account
    ws.Cells(rowNum, 6).Value = details
    ws.Cells(rowNum, 7).Value = amount

    If sla <> "" Then
        ws.Cells(rowNum, 8).Value = sla
    End If

    ' Formatage selon sévérité
    FormatAlertCell ws.Cells(rowNum, 4), severity

    rowNum = rowNum + 1
End Sub

' ==============================================================================
' 6. AUDIT TRAIL UNIFIÉ (CORRIGE BUG-002)
' ==============================================================================

Public Sub WriteAuditLog(logType As String, action As String, Optional details As String = "", Optional username As String = "")
    ' Fonction unique pour écrire dans l'audit trail
    ' Format unifié: Timestamp | Type | User | Action | Details | Hash | PrevHash
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim previousHash As String
    Dim currentHash As String
    Dim logData As String

    If username = "" Then username = Environ("USERNAME")

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("AUDIT_TRAIL")

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "AUDIT_TRAIL"
        wsLog.Range("A1:G1").Value = Array("Timestamp", "Type", "User", "Action", "Details", "Hash", "PrevHash")
        FormatHeader wsLog.Range("A1:G1")
        wsLog.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    ' Récupérer le hash précédent
    If nextRow > 2 Then
        previousHash = SafeText(wsLog.Cells(nextRow - 1, 6).Value)
    Else
        previousHash = "GENESIS"
    End If

    ' Construire les données du log
    logData = Format(Now, "yyyy-mm-dd hh:nn:ss.000") & "|" & _
              logType & "|" & username & "|" & action & "|" & details & "|" & previousHash

    ' Calculer le hash chainé
    currentHash = ComputeHash(logData)

    ' Écrire l'entrée
    wsLog.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss.000")
    wsLog.Cells(nextRow, 2).Value = logType
    wsLog.Cells(nextRow, 3).Value = username
    wsLog.Cells(nextRow, 4).Value = Left(action, 500)
    wsLog.Cells(nextRow, 5).Value = Left(details, 1000)
    wsLog.Cells(nextRow, 6).Value = currentHash
    wsLog.Cells(nextRow, 7).Value = previousHash
End Sub

Public Sub LogError(moduleName As String, procName As String, errNumber As Long, errDescription As String, Optional severity As String = "ERROR")
    ' Log centralisé des erreurs - écrit dans AUDIT_TRAIL et Debug
    Dim logEntry As String
    logEntry = moduleName & "." & procName & ": [" & errNumber & "] " & errDescription

    ' Debug output
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & severity & " | " & logEntry

    ' Écrire dans l'audit trail
    WriteAuditLog severity, logEntry, "Module: " & moduleName & ", Proc: " & procName
End Sub

' ==============================================================================
' 7. SCORING DE RISQUE CONFIGURABLE
' ==============================================================================

Public Function CalculateRiskScore(ecart As Double, ageMax As Double, nbTrans As Long, _
                                   status As String, ByRef factors As String, _
                                   Optional customWeights As Variant) As Integer
    ' Calcul du score de risque (0-100) avec poids configurables
    Dim score As Integer
    Dim wEcart10M As Integer, wEcart1M As Integer, wEcart100K As Integer
    Dim wAge1Y As Integer, wAge6M As Integer, wAge90D As Integer
    Dim wOrphelin As Integer, wVolume As Integer

    factors = ""
    score = 0

    ' Poids par défaut ou personnalisés
    If IsMissing(customWeights) Then
        wEcart10M = 30: wEcart1M = 20: wEcart100K = 10
        wAge1Y = 30: wAge6M = 20: wAge90D = 10
        wOrphelin = 20: wVolume = 10
    Else
        ' customWeights = Array(wEcart10M, wEcart1M, wEcart100K, wAge1Y, wAge6M, wAge90D, wOrphelin, wVolume)
        wEcart10M = customWeights(0): wEcart1M = customWeights(1): wEcart100K = customWeights(2)
        wAge1Y = customWeights(3): wAge6M = customWeights(4): wAge90D = customWeights(5)
        wOrphelin = customWeights(6): wVolume = customWeights(7)
    End If

    ' Facteur 1: Montant de l'écart
    If Abs(ecart) > 10000000 Then
        score = score + wEcart10M
        factors = factors & "Ecart >10M;"
    ElseIf Abs(ecart) > 1000000 Then
        score = score + wEcart1M
        factors = factors & "Ecart >1M;"
    ElseIf Abs(ecart) > 100000 Then
        score = score + wEcart100K
        factors = factors & "Ecart >100K;"
    End If

    ' Facteur 2: Âge
    If ageMax > 365 Then
        score = score + wAge1Y
        factors = factors & "Age >1an;"
    ElseIf ageMax > 180 Then
        score = score + wAge6M
        factors = factors & "Age >6mois;"
    ElseIf ageMax > 90 Then
        score = score + wAge90D
        factors = factors & "Age >90j;"
    End If

    ' Facteur 3: Statut orphelin
    If InStr(status, "Absent") > 0 Or InStr(status, "Présent GL") > 0 Or InStr(status, "Présent Bal") > 0 Then
        score = score + wOrphelin
        factors = factors & "Orphelin;"
    End If

    ' Facteur 4: Volume transactions
    If nbTrans > 100 Then
        score = score + wVolume
        factors = factors & "Volume élevé;"
    End If

    ' Plafonner à 100
    If score > 100 Then score = 100

    CalculateRiskScore = score
End Function

Public Function GetPriority(riskScore As Integer) As String
    If riskScore >= RISK_CRITICAL_THRESHOLD Then
        GetPriority = "CRITICAL"
    ElseIf riskScore >= RISK_HIGH_THRESHOLD Then
        GetPriority = "HIGH"
    ElseIf riskScore >= RISK_MEDIUM_THRESHOLD Then
        GetPriority = "MEDIUM"
    Else
        GetPriority = "LOW"
    End If
End Function

' ==============================================================================
' 8. ALGORITHME DE FUZZY MATCHING (INNOV-001)
' ==============================================================================

Public Function LevenshteinDistance(str1 As String, str2 As String) As Long
    ' Calcule la distance d'édition entre deux chaînes
    Dim len1 As Long, len2 As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim matrix() As Long

    len1 = Len(str1)
    len2 = Len(str2)

    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If
    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ReDim matrix(0 To len1, 0 To len2)

    For i = 0 To len1
        matrix(i, 0) = i
    Next i

    For j = 0 To len2
        matrix(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid(str1, i, 1) = Mid(str2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            matrix(i, j) = WorksheetFunction.Min( _
                matrix(i - 1, j) + 1, _
                matrix(i, j - 1) + 1, _
                matrix(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = matrix(len1, len2)
End Function

Public Function FuzzyMatch(str1 As String, str2 As String, Optional threshold As Double = 0.8) As Boolean
    ' Retourne True si les chaînes sont similaires à au moins threshold%
    Dim maxLen As Long
    Dim distance As Long
    Dim similarity As Double

    str1 = UCase(Trim(Replace(str1, " ", "")))
    str2 = UCase(Trim(Replace(str2, " ", "")))

    If str1 = str2 Then
        FuzzyMatch = True
        Exit Function
    End If

    maxLen = WorksheetFunction.Max(Len(str1), Len(str2))
    If maxLen = 0 Then
        FuzzyMatch = False
        Exit Function
    End If

    distance = LevenshteinDistance(str1, str2)
    similarity = 1 - (distance / maxLen)

    FuzzyMatch = (similarity >= threshold)
End Function

Public Function GetSimilarity(str1 As String, str2 As String) As Double
    ' Retourne le pourcentage de similarité entre 0 et 1
    Dim maxLen As Long
    Dim distance As Long

    str1 = UCase(Trim(Replace(str1, " ", "")))
    str2 = UCase(Trim(Replace(str2, " ", "")))

    If str1 = str2 Then
        GetSimilarity = 1
        Exit Function
    End If

    maxLen = WorksheetFunction.Max(Len(str1), Len(str2))
    If maxLen = 0 Then
        GetSimilarity = 0
        Exit Function
    End If

    distance = LevenshteinDistance(str1, str2)
    GetSimilarity = 1 - (distance / maxLen)
End Function

' ==============================================================================
' 9. CLASSIFICATION COMPTABLE OHADA
' ==============================================================================

Public Function GetAccountClass(acctNum As String) As String
    ' Classification par classe comptable OHADA/SYSCOHADA
    Dim firstChar As String
    firstChar = Left(Trim(acctNum), 1)

    Select Case firstChar
        Case "1": GetAccountClass = "Capital"
        Case "2": GetAccountClass = "Immobilisation"
        Case "3": GetAccountClass = "Stock"
        Case "4": GetAccountClass = "Tiers"
        Case "5": GetAccountClass = "Trésorerie"
        Case "6": GetAccountClass = "Charges"
        Case "7": GetAccountClass = "Produits"
        Case "8": GetAccountClass = "Hors Bilan"
        Case "9": GetAccountClass = "Analytique"
        Case Else: GetAccountClass = "Autre"
    End Select
End Function

Public Function GetAccountSubClass(acctNum As String) As String
    ' Sous-classification détaillée
    Dim prefix As String
    prefix = Left(Trim(acctNum), 2)

    Select Case prefix
        ' Classe 1 - Capitaux
        Case "10": GetAccountSubClass = "Capital social"
        Case "11": GetAccountSubClass = "Réserves"
        Case "12": GetAccountSubClass = "Report à nouveau"
        Case "13": GetAccountSubClass = "Résultat"

        ' Classe 4 - Tiers
        Case "40": GetAccountSubClass = "Fournisseurs"
        Case "41": GetAccountSubClass = "Clients"
        Case "42": GetAccountSubClass = "Personnel"
        Case "43": GetAccountSubClass = "Organismes sociaux"
        Case "44": GetAccountSubClass = "Etat"
        Case "46": GetAccountSubClass = "Débiteurs/Créditeurs divers"
        Case "47": GetAccountSubClass = "Comptes transitoires"
        Case "48": GetAccountSubClass = "Comptes de régularisation"
        Case "49": GetAccountSubClass = "Dépréciations"

        ' Classe 5 - Trésorerie
        Case "51": GetAccountSubClass = "Banques"
        Case "52": GetAccountSubClass = "Instruments de trésorerie"
        Case "53": GetAccountSubClass = "Établissements financiers"
        Case "56": GetAccountSubClass = "Banques centrales"
        Case "57": GetAccountSubClass = "Caisse"
        Case "58": GetAccountSubClass = "Virements internes"
        Case "59": GetAccountSubClass = "Dépréciations trésorerie"

        Case Else: GetAccountSubClass = GetAccountClass(acctNum)
    End Select
End Function

' ==============================================================================
' 10. PROVISION IFRS9 UNIFIÉE
' ==============================================================================

Public Function CalculateIFRS9Provision(amount As Double, ageDays As Double, _
                                        Optional ecl12M As Double = 0.01, _
                                        Optional eclLifetime As Double = 0.5) As Double
    ' Calcul provision IFRS 9 selon staging
    ' Stage 1: 12-month ECL (age <= 30j)
    ' Stage 2: Lifetime ECL (30 < age <= 90j) - facteur progressif
    ' Stage 3: Lifetime ECL + impaired (age > 90j) - facteur élevé

    Dim absAmount As Double
    absAmount = Abs(amount)

    If ageDays <= 30 Then
        ' Stage 1
        CalculateIFRS9Provision = absAmount * ecl12M
    ElseIf ageDays <= 90 Then
        ' Stage 2 - progressif
        CalculateIFRS9Provision = absAmount * (ecl12M + (eclLifetime - ecl12M) * ((ageDays - 30) / 60))
    ElseIf ageDays <= 180 Then
        ' Stage 3 early
        CalculateIFRS9Provision = absAmount * eclLifetime
    ElseIf ageDays <= 360 Then
        ' Stage 3 advanced
        CalculateIFRS9Provision = absAmount * (eclLifetime + (1 - eclLifetime) * ((ageDays - 180) / 180))
    Else
        ' Write-off territory
        CalculateIFRS9Provision = absAmount
    End If
End Function

' ==============================================================================
' 11. UTILITAIRES STATISTIQUES
' ==============================================================================

Public Function CalculateMean(values As Variant) As Double
    On Error Resume Next
    Dim total As Double
    Dim count As Long
    Dim i As Long

    total = 0
    count = 0

    For i = LBound(values) To UBound(values)
        If IsNumeric(values(i)) Then
            total = total + CDbl(values(i))
            count = count + 1
        End If
    Next i

    If count > 0 Then
        CalculateMean = total / count
    Else
        CalculateMean = 0
    End If
    On Error GoTo 0
End Function

Public Function CalculateStdDev(values As Variant, Optional meanValue As Double = -999999) As Double
    On Error Resume Next
    Dim mean As Double
    Dim sumSq As Double
    Dim count As Long
    Dim i As Long

    If meanValue = -999999 Then
        mean = CalculateMean(values)
    Else
        mean = meanValue
    End If

    sumSq = 0
    count = 0

    For i = LBound(values) To UBound(values)
        If IsNumeric(values(i)) Then
            sumSq = sumSq + (CDbl(values(i)) - mean) ^ 2
            count = count + 1
        End If
    Next i

    If count > 1 Then
        CalculateStdDev = Sqr(sumSq / (count - 1))
    Else
        CalculateStdDev = 0
    End If
    On Error GoTo 0
End Function

Public Function CalculateZScore(value As Double, mean As Double, stdDev As Double) As Double
    If stdDev > 0 Then
        CalculateZScore = (value - mean) / stdDev
    Else
        CalculateZScore = 0
    End If
End Function

' ==============================================================================
' 12. DISTRIBUTION BENFORD THÉORIQUE
' ==============================================================================

Public Function GetBenfordExpected(digit As Integer) As Double
    ' Retourne la fréquence attendue selon Benford pour un chiffre (1-9)
    If digit >= 1 And digit <= 9 Then
        GetBenfordExpected = WorksheetFunction.Log10(1 + (1 / digit))
    Else
        GetBenfordExpected = 0
    End If
End Function

Public Function GetBenfordArray() As Variant
    ' Retourne un tableau des 9 fréquences Benford attendues
    Dim benford(1 To 9) As Double
    benford(1) = 0.301
    benford(2) = 0.176
    benford(3) = 0.125
    benford(4) = 0.097
    benford(5) = 0.079
    benford(6) = 0.067
    benford(7) = 0.058
    benford(8) = 0.051
    benford(9) = 0.046
    GetBenfordArray = benford
End Function

' ==============================================================================
' 13. GÉNÉRATION ID UNIQUE
' ==============================================================================

Public Function GenerateUniqueID(Optional prefix As String = "") As String
    ' Génère un ID unique basé sur timestamp + random
    ' IMPORTANT: Appeler Randomize une fois au démarrage de l'application
    GenerateUniqueID = prefix & _
                       Format(Now, "yyyymmddhhnnss") & "_" & _
                       Right("0000" & Hex(Int(Rnd * 65535)), 4) & _
                       Right("0000" & Hex(Int(Rnd * 65535)), 4)
End Function

Public Sub InitializeRandomizer()
    ' À appeler une fois au démarrage pour initialiser le générateur aléatoire
    Randomize Timer
End Sub

' ==============================================================================
' 14. NORMALISATION CLÉS DE COMPTE
' ==============================================================================

Public Function NormalizeAccountKey(acct As String) As String
    ' Normalise une clé de compte pour le matching
    NormalizeAccountKey = UCase(Trim(Replace(Replace(Replace(acct, " ", ""), "'", ""), "-", "")))
End Function

Public Function ExtractNumericPart(acct As String) As String
    ' Extrait uniquement les chiffres d'un numéro de compte
    Dim i As Long
    Dim c As String
    Dim result As String

    result = ""
    For i = 1 To Len(acct)
        c = Mid(acct, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        End If
    Next i

    ExtractNumericPart = result
End Function
