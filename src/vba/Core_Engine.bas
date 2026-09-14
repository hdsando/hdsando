Attribute VB_Name = "Core_Engine"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: CORE_ENGINE v10.0
' ==============================================================================
' Description: Moteur principal - Ingestion, Conversion, Rapprochement
' Auteur: Equipe S.A.F.A
' Version: 10.0 (Enhanced)
' ==============================================================================

' --- CONSTANTES GLOBALES ---
Public Const SAFA_VERSION As String = "10.0"
Public Const MAX_ROWS_MEMORY As Long = 500000
Public Const DEFAULT_TOLERANCE As Double = 100

' --- TYPES PERSONNALISÉS ---
Public Type TransactionRecord
    AccountNumber As String
    AccountName As String
    TransactionDate As Date
    Amount As Double
    Narration As String
    Direction As String
    Currency As String
    RiskScore As Integer
End Type

Public Type ReconciliationResult
    AccountNumber As String
    AccountName As String
    BalanceGL As Double
    BalanceProof As Double
    Difference As Double
    Status As String
    TransactionCount As Long
    AgeMax As Long
    AgeAverage As Double
    RiskScore As Integer
    RiskFactors As String
End Type

' --- VARIABLES GLOBALES ---
Public g_StartTime As Double
Public g_ErrorLog As Collection
Public g_ProcessingStats As Object

' ==============================================================================
' 1. FONCTIONS DE SÉCURITÉ & CONVERSION (AMÉLIORÉES)
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
        ' Nettoyage des caractères de contrôle
        SafeText = CleanControlChars(SafeText)
        ' Troncature optionnelle
        If maxLength > 0 And Len(SafeText) > maxLength Then
            SafeText = Left(SafeText, maxLength)
        End If
    End If

    On Error GoTo 0
End Function

Public Function SafeVal(val As Variant, Optional defaultValue As Double = 0) As Double
    ' Fonction "Bulldozer" v2.0 : Convertit n'importe quoi en chiffre
    ' Améliorée avec support multi-devises et formats internationaux
    On Error Resume Next

    ' 1. Sécurité de base
    If IsError(val) Then SafeVal = defaultValue: Exit Function
    If IsEmpty(val) Or Trim(CStr(val)) = "" Then SafeVal = defaultValue: Exit Function

    ' 2. Si c'est déjà un nombre propre
    If IsNumeric(val) And Not IsDate(val) Then
        SafeVal = CDbl(val)
        Exit Function
    End If

    Dim s As String
    s = CStr(val)

    ' 3. Nettoyage avancé (Espaces insécables, Devises, Guillemets)
    s = Replace(s, Chr(160), "")  ' Espace insécable
    s = Replace(s, Chr(8239), "") ' Narrow no-break space
    s = Replace(s, " ", "")

    ' Devises multiples
    Dim currencies As Variant
    currencies = Array("XAF", "EUR", "USD", "GBP", "XOF", "NGN", "GHS", "KES", "ZAR", "MAD", "TND", "EGP", "FCFA", "F CFA")
    Dim curr As Variant
    For Each curr In currencies
        s = Replace(s, curr, "", , , vbTextCompare)
    Next curr

    ' Indicateurs débit/crédit
    s = Replace(s, "Cr", "", , , vbTextCompare)
    s = Replace(s, "Dr", "", , , vbTextCompare)
    s = Replace(s, "CR", "")
    s = Replace(s, "DR", "")
    s = Replace(s, """", "")
    s = Replace(s, "'", "")

    ' 4. Gestion du signe (Multiple formats)
    Dim isNeg As Boolean: isNeg = False

    ' Format: 1000.00-
    If Right(s, 1) = "-" Then
        isNeg = True
        s = Left(s, Len(s) - 1)
    ' Format: -1000.00
    ElseIf Left(s, 1) = "-" Then
        isNeg = True
        s = Mid(s, 2)
    ' Format: (1000.00)
    ElseIf Left(s, 1) = "(" And Right(s, 1) = ")" Then
        isNeg = True
        s = Mid(s, 2, Len(s) - 2)
    ' Format: <1000.00>
    ElseIf Left(s, 1) = "<" And Right(s, 1) = ">" Then
        isNeg = True
        s = Mid(s, 2, Len(s) - 2)
    End If

    ' 5. Gestion intelligente des séparateurs
    s = NormalizeNumberFormat(s)

    ' 6. Conversion finale
    If IsNumeric(s) Then
        SafeVal = CDbl(s)
        If isNeg Then SafeVal = SafeVal * -1
    Else
        SafeVal = defaultValue
    End If

    On Error GoTo 0
End Function

Private Function NormalizeNumberFormat(s As String) As String
    ' Détecte et normalise le format numérique (US vs EU)
    Dim dotPos As Long, commaPos As Long
    Dim dotCount As Long, commaCount As Long
    Dim i As Long

    ' Compter les séparateurs
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "." Then
            dotCount = dotCount + 1
            dotPos = i
        ElseIf Mid(s, i, 1) = "," Then
            commaCount = commaCount + 1
            commaPos = i
        End If
    Next i

    ' Logique de détection
    If dotCount = 0 And commaCount = 0 Then
        ' Pas de séparateur
        NormalizeNumberFormat = s
    ElseIf dotCount = 1 And commaCount = 0 Then
        ' Format US simple: 1234.56
        NormalizeNumberFormat = s
    ElseIf dotCount = 0 And commaCount = 1 Then
        ' Format EU simple: 1234,56
        NormalizeNumberFormat = Replace(s, ",", ".")
    ElseIf dotCount > 1 And commaCount = 0 Then
        ' Format EU milliers: 1.234.567 -> enlever tous les points
        NormalizeNumberFormat = Replace(s, ".", "")
    ElseIf dotCount = 0 And commaCount > 1 Then
        ' Format US milliers: 1,234,567 -> enlever toutes les virgules
        NormalizeNumberFormat = Replace(s, ",", "")
    ElseIf dotCount >= 1 And commaCount >= 1 Then
        ' Format mixte - le dernier séparateur est décimal
        If dotPos > commaPos Then
            ' Format US: 1,234.56
            NormalizeNumberFormat = Replace(s, ",", "")
        Else
            ' Format EU: 1.234,56
            s = Replace(s, ".", "")
            NormalizeNumberFormat = Replace(s, ",", ".")
        End If
    Else
        NormalizeNumberFormat = s
    End If
End Function

Private Function CleanControlChars(s As String) As String
    ' Supprime les caractères de contrôle (ASCII 0-31 sauf tab, newline)
    Dim i As Long, c As Long, result As String

    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        If c >= 32 Or c = 9 Or c = 10 Or c = 13 Then
            result = result & Mid(s, i, 1)
        End If
    Next i

    CleanControlChars = result
End Function

Public Function ToDouble(val As Variant) As Double
    ToDouble = SafeVal(val)
End Function

Public Function SafeDate(val As Variant, Optional defaultDate As Date = 0) As Date
    ' Conversion sécurisée de date avec support multi-format
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

    ' Essayer différents formats
    If IsDate(s) Then
        SafeDate = CDate(s)
    Else
        SafeDate = defaultDate
    End If

    On Error GoTo 0
End Function

Public Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not ws Is Nothing
End Function

Public Function GetOrCreateSheet(nom As String, Optional clearIfExists As Boolean = False) As Worksheet
    ' Delegue a SAFA_Common (implementation unique, nom tronque a 31 caracteres)
    Set GetOrCreateSheet = SAFA_Common.GetOrCreateSheet(nom, clearIfExists)
End Function

' ==============================================================================
' 2. GESTION D'ERREURS CENTRALISÉE
' ==============================================================================

Public Sub InitializeErrorLog()
    Set g_ErrorLog = New Collection
    g_StartTime = Timer
End Sub

Public Sub LogError(moduleName As String, procName As String, errNumber As Long, errDescription As String, Optional severity As String = "ERROR")
    ' Log centralisé des erreurs
    Dim logEntry As String
    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & "|" & severity & "|" & moduleName & "|" & procName & "|" & errNumber & "|" & errDescription

    On Error Resume Next
    If g_ErrorLog Is Nothing Then Set g_ErrorLog = New Collection
    g_ErrorLog.Add logEntry
    On Error GoTo 0

    ' Écrire dans la feuille LOG si elle existe
    WriteToAuditLog "ERROR", moduleName & "." & procName & ": " & errDescription
End Sub

Public Sub WriteToAuditLog(logType As String, message As String, Optional details As String = "")
    ' Ecriture dans le journal d'audit: UN SEUL ecrivain (SAFA_Common.WriteAuditLog), format chaine
    ' Timestamp | Type | User | Action | Details | Hash | PrevHash | Algo
    ' (l'ancienne implementation ecrivait un format non chaine dans la meme feuille,
    '  ce qui rendait la verification d'integrite impossible)
    On Error Resume Next
    Call SAFA_Common.WriteAuditLog(logType, message, details)
End Sub

Private Function SimpleHash(text As String) As String
    ' Hash simple pour l'intégrité (en production, utiliser SHA-256 via API)
    Dim i As Long, h As Long
    h = 5381
    For i = 1 To Len(text)
        h = ((h * 33) Xor Asc(Mid(text, i, 1))) And &H7FFFFFFF
    Next i
    SimpleHash = Hex(h)
End Function

' ==============================================================================
' 3. SÉQUENCE DE DÉMARRAGE & NETTOYAGE (AMÉLIORÉS)
' ==============================================================================

Public Sub Sequence_Demarrage()
    Dim repClean As Integer, repHist As Integer, wsBal As Worksheet

    ' Initialiser le logging
    Call InitializeErrorLog
    Call WriteToAuditLog("SESSION", "S.A.F.A v" & SAFA_VERSION & " démarré")

    If Not FeuilleExiste("BALANCE_RAW") Then Exit Sub

    Set wsBal = ThisWorkbook.Sheets("BALANCE_RAW")
    If Application.CountA(wsBal.Range("A:A")) > 1 Then
        repClean = MsgBox("Données détectées dans l'outil." & vbCrLf & vbCrLf & _
                          "Voulez-vous nettoyer avant une nouvelle analyse ?", _
                          vbYesNo + vbExclamation, "Démarrage S.A.F.A v" & SAFA_VERSION)
        If repClean = vbYes Then
            repHist = MsgBox("Conserver l'historique (Analyse de Vélocité) ?" & vbCrLf & vbCrLf & _
                            "OUI = Audit Mensuel (Recommandé)" & vbCrLf & _
                            "NON = Remise à zéro totale", _
                            vbYesNoCancel + vbQuestion, "Mode de nettoyage")
            If repHist <> vbCancel Then
                Call Reinitialiser_Outil_Parametrable(repHist = vbYes)
            End If
        End If
    End If
End Sub

Public Sub Reinitialiser_Outil_Parametrable(keepHistory As Boolean)
    Dim ws As Worksheet, n As String
    Dim protectedSheets As Variant

    ' Feuilles à ne jamais supprimer
    protectedSheets = Array("PARAM", "MENU", "BALANCE_RAW", "GLPROOF_RAW", "CONFIG", "AUDIT_TRAIL")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call WriteToAuditLog("RESET", "Nettoyage lancé", "KeepHistory=" & keepHistory)

    For Each ws In ThisWorkbook.Worksheets
        n = UCase(ws.name)

        ' Vérifier si c'est une feuille protégée
        If Not IsInArray(n, protectedSheets) Then
            If n = "HISTORY_LOG" Then
                If Not keepHistory Then ws.Delete
            Else
                ws.Delete
            End If
        End If
    Next ws

    ' Nettoyer les feuilles RAW
    On Error Resume Next
    ThisWorkbook.Sheets("BALANCE_RAW").Cells.Clear
    ThisWorkbook.Sheets("GLPROOF_RAW").Cells.Clear
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Call WriteToAuditLog("RESET", "Nettoyage terminé")
End Sub

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i)) = UCase(val) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

' ==============================================================================
' 4. IMPORTS (AMÉLIORÉS AVEC VALIDATION)
' ==============================================================================

Public Sub Importer_Source_Balance()
    Dim fd As FileDialog, filePath As String, wb As Workbook
    Dim rowCount As Long

    ' Créer la feuille si nécessaire
    If Not FeuilleExiste("BALANCE_RAW") Then
        ThisWorkbook.Sheets.Add.name = "BALANCE_RAW"
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Sélectionner le fichier BALANCE FINACLE"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xls;*.xlsx;*.xlsb;*.xlsm"
        .Filters.Add "Fichiers CSV", "*.csv;*.txt"
        .Filters.Add "Tous les fichiers", "*.*"
        .AllowMultiSelect = False
    End With

    If fd.Show <> -1 Then Exit Sub
    filePath = fd.SelectedItems(1)

    ' Logging
    Call WriteToAuditLog("IMPORT", "Import Balance démarré", "Fichier: " & filePath)

    Application.ScreenUpdating = False
    On Error GoTo ImportError

    ThisWorkbook.Sheets("BALANCE_RAW").Cells.Clear

    ' Ouvrir et copier
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    wb.Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("BALANCE_RAW").Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    rowCount = ThisWorkbook.Sheets("BALANCE_RAW").Cells(Rows.count, 1).End(xlUp).Row
    wb.Close False

    Application.ScreenUpdating = True

    Call WriteToAuditLog("IMPORT", "Import Balance terminé", "Lignes: " & rowCount)
    MsgBox "Balance importée avec succès." & vbCrLf & "Lignes: " & rowCount, vbInformation, "Import Balance"
    Exit Sub

ImportError:
    Application.ScreenUpdating = True
    Call LogError("Core_Engine", "Importer_Source_Balance", Err.Number, Err.Description)
    MsgBox "Erreur lors de l'import: " & Err.Description, vbCritical, "Erreur Import"
End Sub

Public Sub Importer_Source_GLProof()
    Dim fd As FileDialog, filePath As String, wb As Workbook
    Dim sh As Worksheet, lr As Long, totalRows As Long

    If Not FeuilleExiste("GLPROOF_RAW") Then
        ThisWorkbook.Sheets.Add.name = "GLPROOF_RAW"
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Sélectionner le fichier GL PROOF"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xls;*.xlsx;*.xlsb;*.xlsm"
        .AllowMultiSelect = False
    End With

    If fd.Show <> -1 Then Exit Sub
    filePath = fd.SelectedItems(1)

    Call WriteToAuditLog("IMPORT", "Import GL Proof démarré", "Fichier: " & filePath)

    Application.ScreenUpdating = False
    On Error GoTo ImportError

    ThisWorkbook.Sheets("GLPROOF_RAW").Cells.Clear

    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

    ' Importer toutes les feuilles
    totalRows = 0
    For Each sh In wb.Worksheets
        If Application.CountA(sh.UsedRange) > 0 Then
            lr = ThisWorkbook.Sheets("GLPROOF_RAW").Cells(Rows.count, 1).End(xlUp).Row
            If lr > 1 Then lr = lr + 1 Else lr = 1

            sh.UsedRange.Copy
            ThisWorkbook.Sheets("GLPROOF_RAW").Cells(lr, 1).PasteSpecial xlPasteValues
            totalRows = totalRows + sh.UsedRange.Rows.count
        End If
    Next sh

    Application.CutCopyMode = False
    wb.Close False

    Application.ScreenUpdating = True

    Call WriteToAuditLog("IMPORT", "Import GL Proof terminé", "Lignes totales: " & totalRows)
    MsgBox "GL Proof importé avec succès." & vbCrLf & "Lignes: " & totalRows, vbInformation, "Import GL Proof"
    Exit Sub

ImportError:
    Application.ScreenUpdating = True
    Call LogError("Core_Engine", "Importer_Source_GLProof", Err.Number, Err.Description)
    MsgBox "Erreur lors de l'import: " & Err.Description, vbCritical, "Erreur Import"
End Sub

' ==============================================================================
' 5. NETTOYAGE BALANCE (LOGIQUE v10 - OPTIMISÉE AVEC ARRAYS)
' ==============================================================================

Public Sub NettoyerBalance()
    Dim wsRaw As Worksheet, wsOut As Worksheet, wsExcl As Worksheet
    Dim dataArray As Variant, outputArray() As Variant, exclArray() As Variant
    Dim lr As Long, i As Long, outRow As Long, exclRow As Long
    Dim cAcct As Integer, cBal As Integer, cSub As Integer, cDesc As Integer
    Dim solId As String, hRow As Long, rngFound As Range

    On Error GoTo CleanError

    ' Récupérer SOL ID
    On Error Resume Next
    solId = ThisWorkbook.Sheets("PARAM").Range("F2").Value
    On Error GoTo CleanError
    If solId = "" Then solId = "799"

    Set wsRaw = ThisWorkbook.Sheets("BALANCE_RAW")
    Set wsOut = GetOrCreateSheet("BALANCE_DATA", True)

    ' Supprimer et recréer COMPTES_EXCLUS
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("COMPTES_EXCLUS").Delete
    On Error GoTo CleanError
    Application.DisplayAlerts = True

    Set wsExcl = ThisWorkbook.Sheets.Add(After:=wsOut)
    wsExcl.name = "COMPTES_EXCLUS"

    ' En-têtes
    wsOut.Range("A1:D1").Value = Array("AccountNumber", "AccountName", "ClosingBalance", "Currency")
    wsOut.Range("A1:D1").Font.Bold = True
    wsOut.Range("A1:D1").Interior.Color = RGB(0, 51, 102)
    wsOut.Range("A1:D1").Font.Color = vbWhite

    wsExcl.Range("A1:D1").Value = Array("Compte", "Libellé", "Montant", "Motif")
    wsExcl.Range("A1:D1").Font.Bold = True

    ' Recherche ligne titre (scan intelligent)
    Set rngFound = FindHeaderRow(wsRaw, Array("Account Number", "Acct Num", "ACCOUNT", "Numéro Compte"))

    If rngFound Is Nothing Then
        MsgBox "ERREUR: Colonne 'Account Number' introuvable dans les 50 premières lignes.", vbCritical, "Erreur Parsing"
        Exit Sub
    End If

    hRow = rngFound.Row
    cAcct = rngFound.Column

    ' Recherche autres colonnes
    cBal = FindColumnInRow(wsRaw, hRow, Array("Closing", "Solde", "Balance", "CLR_BAL_AMT"))
    cSub = FindColumnInRow(wsRaw, hRow, Array("Sub", "Subsidiary", "SOL"))
    cDesc = FindColumnInRow(wsRaw, hRow, Array("Desc", "Name", "Libellé", "ACCT_NAME"))

    If cBal = 0 Then
        MsgBox "ERREUR: Colonne Solde (Closing/Balance) introuvable.", vbCritical, "Erreur Parsing"
        Exit Sub
    End If
    If cSub = 0 Then cSub = cAcct
    If cDesc = 0 Then cDesc = cAcct

    ' Charger données en mémoire pour performance
    lr = wsRaw.Cells(wsRaw.Rows.count, cAcct).End(xlUp).Row

    ' Dimensionner les arrays de sortie
    ReDim outputArray(1 To lr, 1 To 4)
    ReDim exclArray(1 To lr, 1 To 4)

    outRow = 0
    exclRow = 0

    ' Traitement
    For i = hRow + 1 To lr
        Dim acct As String, acctClean As String, descr As String
        Dim balValue As Double

        acct = SafeText(wsRaw.Cells(i, cAcct).Value)

        ' Filtrer lignes invalides
        If Len(acct) > 5 And Left(acct, 1) <> "-" And InStr(1, acct, "Total", vbTextCompare) = 0 Then

            balValue = SafeVal(wsRaw.Cells(i, cBal).Value)
            If SAFA_Common.g_FlipBalanceSign Then balValue = -balValue   ' calibration: convention opposee au GL
            descr = SafeText(wsRaw.Cells(i, cDesc).Value, 200)

            ' Exclure comptes ISO
            If InStr(UCase(acct), "ISO") > 0 Then
                exclRow = exclRow + 1
                exclArray(exclRow, 1) = acct
                exclArray(exclRow, 2) = descr
                exclArray(exclRow, 3) = balValue
                exclArray(exclRow, 4) = "ISO"
            Else
                ' Construire numéro de compte normalisé
                Dim pfx As String, subCode As String, sfx As String
                pfx = Left(acct, 3)

                If cSub <> cAcct Then
                    subCode = Left(Trim(SafeText(wsRaw.Cells(i, cSub).Value)), 1)
                    If subCode = "" Or subCode = "-" Then subCode = "0"
                Else
                    subCode = "0"
                End If

                sfx = Mid(acct, 8)
                acctClean = NormalizeBalanceAccount(acct, solId, subCode)

                outRow = outRow + 1
                outputArray(outRow, 1) = "'" & acctClean
                outputArray(outRow, 2) = descr
                outputArray(outRow, 3) = balValue
                outputArray(outRow, 4) = "XAF" ' Default currency
            End If
        End If
    Next i

    ' Écrire les résultats en bloc
    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, 4).Value = outputArray
    End If

    If exclRow > 0 Then
        wsExcl.Range("A2").Resize(exclRow, 4).Value = exclArray
    End If

    wsOut.Columns("A:D").AutoFit
    wsExcl.Columns("A:D").AutoFit

    Call WriteToAuditLog("PROCESS", "NettoyerBalance terminé", "Comptes traités: " & outRow & ", Exclus: " & exclRow)
    Exit Sub

CleanError:
    Call LogError("Core_Engine", "NettoyerBalance", Err.Number, Err.Description)
    MsgBox "Erreur lors du nettoyage Balance: " & Err.Description, vbCritical
End Sub

Private Function FindHeaderRow(ws As Worksheet, searchTerms As Variant) As Range
    ' Recherche intelligente de la ligne d'en-tête
    Dim term As Variant
    Dim rng As Range
    Dim searchRange As Range

    Set searchRange = ws.Range("A1:Z50")

    For Each term In searchTerms
        Set rng = searchRange.Find(What:=CStr(term), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not rng Is Nothing Then
            Set FindHeaderRow = rng
            Exit Function
        End If
    Next term

    Set FindHeaderRow = Nothing
End Function

Private Function FindColumnInRow(ws As Worksheet, rowNum As Long, searchTerms As Variant) As Integer
    ' Recherche une colonne dans une ligne spécifique
    Dim term As Variant
    Dim rng As Range

    For Each term In searchTerms
        Set rng = ws.Rows(rowNum).Find(What:=CStr(term), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not rng Is Nothing Then
            FindColumnInRow = rng.Column
            Exit Function
        End If
    Next term

    FindColumnInRow = 0
End Function

' ==============================================================================
' 6. NETTOYAGE GL PROOF (LOGIQUE v10 - STATE MACHINE)
' ==============================================================================

Public Sub NettoyerGLProof()
    Dim wsRaw As Worksheet, wsOut As Worksheet
    Dim lr As Long, i As Long, outRow As Long
    Dim outputArray() As Variant

    On Error GoTo CleanError

    Set wsOut = GetOrCreateSheet("GLPROOF_DATA", True)
    Set wsRaw = ThisWorkbook.Sheets("GLPROOF_RAW")

    ' En-têtes améliorés
    wsOut.Range("A1:J1").Value = Array("AccountNumber", "AccountName", "Balance_Proof", "Balance_GL", _
                                        "Difference", "NbTransactions", "AgeMax", "AgeMoyenne", _
                                        "Currency", "LastActivity")
    wsOut.Range("A1:J1").Font.Bold = True
    wsOut.Range("A1:J1").Interior.Color = RGB(0, 51, 102)
    wsOut.Range("A1:J1").Font.Color = vbWhite

    lr = wsRaw.Cells(wsRaw.Rows.count, 1).End(xlUp).Row
    If lr < 10 Then lr = wsRaw.Cells(wsRaw.Rows.count, 2).End(xlUp).Row

    ' Variables d'état (State Machine)
    Dim currentAccount As String, currentName As String
    Dim balProof As Double, balGL As Double, diff As Double
    Dim nbTrans As Long, sumAge As Double, ageMax As Double, lastDate As Date

    ReDim outputArray(1 To lr, 1 To 10)
    outRow = 0
    currentAccount = ""

    ' Parcours avec machine d'état
    For i = 1 To lr
        Dim rowText As String
        Dim c1 As String, c2 As String, c3 As String

        c1 = SafeText(wsRaw.Cells(i, 1).Value)
        c2 = SafeText(wsRaw.Cells(i, 2).Value)
        c3 = SafeText(wsRaw.Cells(i, 3).Value)
        rowText = UCase(c1 & " " & c2 & " " & c3)

        ' État: Nouveau compte détecté
        If InStr(rowText, "ACCOUNT NUMBER") > 0 Then
            currentAccount = ExtractAccountNumber(wsRaw, i)
            currentName = ""
            balProof = 0: balGL = 0: diff = 0
            nbTrans = 0: sumAge = 0: ageMax = 0
            lastDate = 0

        ' État: Nom du compte
        ElseIf InStr(rowText, "ACCOUNT NAME") > 0 Then
            currentName = ExtractAccountName(wsRaw, i)

        ' État: Balance GL
        ElseIf InStr(rowText, "BALANCE AS PER GL") > 0 Then
            balGL = TrouverMontantSurLigne(wsRaw, i)

        ' État: Balance Proof
        ElseIf InStr(rowText, "BALANCE AS PER PROOF") > 0 Then
            balProof = TrouverMontantSurLigne(wsRaw, i)

        ' État: Différence (fin de bloc = sauvegarde)
        ElseIf InStr(rowText, "DIFFERENCE") > 0 And Len(rowText) < 60 Then
            diff = TrouverMontantSurLigne(wsRaw, i)

            If currentAccount <> "" Then
                outRow = outRow + 1
                outputArray(outRow, 1) = "'" & currentAccount
                outputArray(outRow, 2) = currentName
                outputArray(outRow, 3) = balProof
                outputArray(outRow, 4) = balGL
                outputArray(outRow, 5) = diff
                outputArray(outRow, 6) = nbTrans
                outputArray(outRow, 7) = ageMax
                outputArray(outRow, 8) = IIf(nbTrans > 0, sumAge / nbTrans, 0)
                outputArray(outRow, 9) = "XAF"
                outputArray(outRow, 10) = lastDate
            End If

        ' État: Transaction détectée
        Else
            Dim transResult As Variant
            transResult = DetectTransaction(wsRaw, i)

            If Not IsEmpty(transResult) Then
                nbTrans = nbTrans + 1
                sumAge = sumAge + transResult(1) ' Age
                If transResult(1) > ageMax Then ageMax = transResult(1)
                If transResult(0) > lastDate Then lastDate = transResult(0)
            End If
        End If
    Next i

    ' Écrire les résultats
    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, 10).Value = outputArray
    End If

    wsOut.Columns("A:J").AutoFit

    Call WriteToAuditLog("PROCESS", "NettoyerGLProof terminé", "Comptes traités: " & outRow)
    Exit Sub

CleanError:
    Call LogError("Core_Engine", "NettoyerGLProof", Err.Number, Err.Description)
    MsgBox "Erreur lors du nettoyage GL Proof: " & Err.Description, vbCritical
End Sub

Private Function ExtractAccountNumber(ws As Worksheet, rowNum As Long) As String
    ' Extraction intelligente du numéro de compte
    Dim j As Integer
    Dim cellVal As String

    For j = 1 To 6
        cellVal = SafeText(ws.Cells(rowNum, j).Value)
        If InStr(UCase(cellVal), "NUMBER") = 0 And InStr(UCase(cellVal), "ACCOUNT") = 0 _
           And cellVal <> ":" And Len(cellVal) > 3 Then
            ExtractAccountNumber = cellVal
            Exit Function
        End If
    Next j

    ExtractAccountNumber = ""
End Function

Private Function ExtractAccountName(ws As Worksheet, rowNum As Long) As String
    ' Extraction intelligente du nom de compte
    Dim j As Integer
    Dim cellVal As String

    For j = 1 To 6
        cellVal = SafeText(ws.Cells(rowNum, j).Value)
        If InStr(UCase(cellVal), "NAME") = 0 And InStr(UCase(cellVal), "ACCOUNT") = 0 _
           And cellVal <> ":" And Len(cellVal) > 3 Then
            ExtractAccountName = cellVal
            Exit Function
        End If
    Next j

    ExtractAccountName = ""
End Function

Private Function DetectTransaction(ws As Worksheet, rowNum As Long) As Variant
    ' Détecte si une ligne est une transaction et retourne (Date, Age)
    Dim cellVal As Variant
    Dim transDate As Date
    Dim transAge As Double
    Dim j As Integer

    ' Chercher une date dans les premières colonnes
    For j = 1 To 3
        cellVal = ws.Cells(rowNum, j).Value
        If IsDate(cellVal) And Not IsError(cellVal) Then
            transDate = CDate(cellVal)

            ' Chercher l'âge dans les colonnes suivantes: entier plausible (0..3650 jours).
            ' CORRIGE (B5): l'ancien code prenait la premiere valeur > 0, donc le MONTANT
            ' quand la colonne age valait 0 ou etait vide. A defaut, age = aujourd'hui - date.
            transAge = -1
            Dim k As Integer, cand As Double
            For k = j + 1 To j + 4
                If Not IsEmpty(ws.Cells(rowNum, k).Value) And IsNumeric(ws.Cells(rowNum, k).Value) Then
                    cand = SafeVal(ws.Cells(rowNum, k).Value)
                    If cand >= 0 And cand <= 3650 And cand = Int(cand) Then
                        transAge = cand
                        Exit For
                    End If
                End If
            Next k
            If transAge < 0 Then
                transAge = DateDiff("d", transDate, Date)
                If transAge < 0 Then transAge = 0
            End If

            DetectTransaction = Array(transDate, transAge)
            Exit Function
        End If
    Next j

    DetectTransaction = Empty
End Function

Private Function TrouverMontantSurLigne(ws As Worksheet, r As Long) As Double
    Dim c As Integer, v As Variant
    For c = 3 To 12
        v = ws.Cells(r, c).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            If SafeVal(v) <> 0 Then
                TrouverMontantSurLigne = SafeVal(v)
                Exit Function
            End If
        End If
    Next c
    TrouverMontantSurLigne = 0
End Function

' ==============================================================================
' 7. RAPPROCHEMENT (v10 - AVEC RISK SCORING)
' ==============================================================================

Public Sub ConstruireRapprochement()
    Dim wsGL As Worksheet, wsBal As Worksheet, wsRec As Worksheet, wsP As Worksheet
    Dim lGL As Long, lBal As Long, i As Long, outRow As Long
    Dim tolerance As Double
    Dim dictBal As Object
    Dim outputArray() As Variant
    Dim exclusionPatterns As Variant

    On Error GoTo ReconcilError

    Set wsGL = ThisWorkbook.Sheets("GLPROOF_DATA")
    Set wsBal = ThisWorkbook.Sheets("BALANCE_DATA")
    Set wsRec = GetOrCreateSheet("RECONCIL", True)
    Set wsP = ThisWorkbook.Sheets("PARAM")

    ' Paramètres
    tolerance = SafeVal(wsP.Range("B2").Value, DEFAULT_TOLERANCE)

    ' Charger les patterns d'exclusion
    If wsP.Cells(wsP.Rows.count, "D").End(xlUp).Row >= 2 Then
        exclusionPatterns = wsP.Range("D2:D50").Value
    End If

    ' En-têtes améliorés
    wsRec.Range("A1:N1").Value = Array("AccountNumber", "AccountName", "Solde GL Proof", "Solde Balance", _
                                        "Ecart", "Statut", "Nb Trans.", "Age Max", "Age Moy.", _
                                        "Source", "Type", "Risk Score", "Risk Factors", "Priority")

    FormatHeader wsRec.Range("A1:N1")

    lGL = wsGL.Cells(wsGL.Rows.count, 1).End(xlUp).Row
    lBal = wsBal.Cells(wsBal.Rows.count, 1).End(xlUp).Row

    ' Créer dictionnaire Balance
    Set dictBal = CreateObject("Scripting.Dictionary")
    For i = 2 To lBal
        Dim keyBal As String
        keyBal = NormalizeAccountKey(SafeText(wsBal.Cells(i, 1).Value))
        If keyBal <> "" And Not dictBal.Exists(keyBal) Then
            dictBal.Add keyBal, i
        End If
    Next i

    ' Dimensionner output
    ReDim outputArray(1 To lGL + lBal, 1 To 14)
    outRow = 0

    ' Parcours GL Proof
    For i = 2 To lGL
        Dim keyGL As String, status As String, riskScore As Integer, riskFactors As String
        Dim balGL As Double, balBal As Double, ecart As Double
        Dim rowBal As Long

        keyGL = NormalizeAccountKey(SafeText(wsGL.Cells(i, 1).Value))
        If keyGL = "" Then GoTo NextGL

        balGL = SafeVal(wsGL.Cells(i, 4).Value)

        If dictBal.Exists(keyGL) Then
            ' Match trouvé
            rowBal = dictBal(keyGL)
            balBal = SafeVal(wsBal.Cells(rowBal, 3).Value)
            dictBal.Remove keyGL

            ecart = balGL - balBal
            status = DetermineStatus(balGL, balBal, ecart, tolerance)

        Else
            ' Présent GL seulement
            balBal = 0
            ecart = balGL
            If balGL = 0 Then
                status = "Sans mouvement"
            Else
                status = "Présent GL / Absent Bal"
            End If
        End If

        ' Calculer Risk Score
        riskScore = SAFA_Common.CalculateRiskScore(ecart, SafeVal(wsGL.Cells(i, 7).Value), _
                                        SafeVal(wsGL.Cells(i, 6).Value), status, riskFactors)

        ' Vérifier exclusion
        If IsExcluded(keyGL, SafeText(wsGL.Cells(i, 2).Value), exclusionPatterns) Then
            status = "Non-Proofable (Auto)"
            riskScore = 0
        End If

        outRow = outRow + 1
        outputArray(outRow, 1) = "'" & SafeText(wsGL.Cells(i, 1).Value)
        outputArray(outRow, 2) = SafeText(wsGL.Cells(i, 2).Value)
        outputArray(outRow, 3) = balGL
        outputArray(outRow, 4) = balBal
        outputArray(outRow, 5) = ecart
        outputArray(outRow, 6) = status
        outputArray(outRow, 7) = SafeVal(wsGL.Cells(i, 6).Value)
        outputArray(outRow, 8) = SafeVal(wsGL.Cells(i, 7).Value)
        outputArray(outRow, 9) = SafeVal(wsGL.Cells(i, 8).Value)
        outputArray(outRow, 10) = IIf(dictBal.Exists(keyGL), "Match", "GL Only")
        outputArray(outRow, 11) = GetAccountType(keyGL)
        outputArray(outRow, 12) = riskScore
        outputArray(outRow, 13) = riskFactors
        outputArray(outRow, 14) = SAFA_Common.GetPriority(riskScore)

NextGL:
    Next i

    ' Traiter comptes Balance orphelins
    Dim vKey As Variant
    For Each vKey In dictBal.Keys
        rowBal = dictBal(vKey)
        balBal = SafeVal(wsBal.Cells(rowBal, 3).Value)

        If balBal = 0 Then
            status = "Sans mouvement"
        Else
            status = "Présent Bal / Absent GL"
        End If

        riskScore = SAFA_Common.CalculateRiskScore(-balBal, 0, 0, status, riskFactors)

        If IsExcluded(CStr(vKey), SafeText(wsBal.Cells(rowBal, 2).Value), exclusionPatterns) Then
            status = "Non-Proofable (Auto)"
            riskScore = 0
        End If

        outRow = outRow + 1
        outputArray(outRow, 1) = "'" & SafeText(wsBal.Cells(rowBal, 1).Value)
        outputArray(outRow, 2) = SafeText(wsBal.Cells(rowBal, 2).Value)
        outputArray(outRow, 3) = 0
        outputArray(outRow, 4) = balBal
        outputArray(outRow, 5) = -balBal
        outputArray(outRow, 6) = status
        outputArray(outRow, 7) = 0
        outputArray(outRow, 8) = 0
        outputArray(outRow, 9) = 0
        outputArray(outRow, 10) = "Balance Only"
        outputArray(outRow, 11) = GetAccountType(CStr(vKey))
        outputArray(outRow, 12) = riskScore
        outputArray(outRow, 13) = riskFactors
        outputArray(outRow, 14) = SAFA_Common.GetPriority(riskScore)
    Next vKey

    ' Écrire résultats
    If outRow > 0 Then
        wsRec.Range("A2").Resize(outRow, 14).Value = outputArray

        ' Formatage conditionnel
        ApplyConditionalFormatting wsRec, outRow
    End If

    wsRec.Columns("A:N").AutoFit

    Call WriteToAuditLog("PROCESS", "ConstruireRapprochement terminé", "Comptes: " & outRow)
    Exit Sub

ReconcilError:
    Call LogError("Core_Engine", "ConstruireRapprochement", Err.Number, Err.Description)
    MsgBox "Erreur lors du rapprochement: " & Err.Description, vbCritical
End Sub

Public Function NormalizeBalanceAccount(rawAcct As String, solId As String, subCode As String) As String
    ' Construit le numero de compte "GL" a partir du numero de la Balance.
    ' Mode SOL_INJECT (defaut, extraction Finacle UBA Tchad): 3 premiers caracteres & SOL & sous-code & caracteres a partir du 8e
    ' Mode NONE: numero inchange (si Balance et GL Proof partagent deja la meme numerotation)
    ' A VALIDER sur les extractions reelles ("Consolidated GL Balance Report" vs GL Proof) avant chaque mission.
    Dim cfg As Config_Manager.GeneralConfig, mode As String
    On Error Resume Next
    cfg = Config_Manager.GetGeneralConfig()
    mode = UCase(Trim(cfg.AccountNormalization))
    On Error GoTo 0
    ' AUTO (defaut): la transformation est choisie par Auto_Calibration sur les donnees
    If mode = "" Or mode = "AUTO" Then
        If SAFA_Common.g_BalanceTransform = "" Then Call Auto_Calibration.CalibrerNumerotation(True)
        mode = SAFA_Common.g_BalanceTransform
        If mode = "" Then mode = "SOL_INJECT"
    End If
    If mode = "NONE" Then
        NormalizeBalanceAccount = Trim(rawAcct)
    ElseIf Len(rawAcct) >= 8 Then
        NormalizeBalanceAccount = Left(rawAcct, 3) & solId & subCode & Mid(rawAcct, 8)
    Else
        NormalizeBalanceAccount = Trim(rawAcct)
    End If
End Function

Public Function ScanBalanceAccounts() As Object
    ' Numeros bruts -> solde de cloture, lus dans BALANCE_RAW avec la meme detection d'en-tete que NettoyerBalance
    Dim d As Object, wsRaw As Worksheet, rngFound As Range
    Dim hRow As Long, cAcct As Long, cBal As Long, lr As Long, i As Long, acct As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not FeuilleExiste("BALANCE_RAW") Then Set ScanBalanceAccounts = d: Exit Function
    Set wsRaw = ThisWorkbook.Sheets("BALANCE_RAW")
    Set rngFound = FindHeaderRow(wsRaw, Array("Account Number", "Acct Num", "ACCOUNT", "Numéro Compte"))
    If rngFound Is Nothing Then Set ScanBalanceAccounts = d: Exit Function
    hRow = rngFound.Row: cAcct = rngFound.Column
    cBal = FindColumnInRow(wsRaw, hRow, Array("Closing", "Solde", "Balance", "CLR_BAL_AMT"))
    lr = wsRaw.Cells(wsRaw.Rows.count, cAcct).End(xlUp).Row
    For i = hRow + 1 To lr
        acct = SafeText(wsRaw.Cells(i, cAcct).Value)
        If Len(acct) > 5 And Left(acct, 1) <> "-" And InStr(1, acct, "Total", vbTextCompare) = 0 And InStr(UCase(acct), "ISO") = 0 Then
            If Not d.Exists(acct) Then d.Add acct, IIf(cBal > 0, SafeVal(wsRaw.Cells(i, cBal).Value), 0)
        End If
    Next i
    Set ScanBalanceAccounts = d
End Function

Public Function ScanGLProofAccounts() As Object
    ' Numeros de compte -> "Balance as per GL", lus dans GLPROOF_RAW (meme machine d'etat que NettoyerGLProof)
    Dim d As Object, wsRaw As Worksheet, lr As Long, i As Long, rowText As String, cur As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not FeuilleExiste("GLPROOF_RAW") Then Set ScanGLProofAccounts = d: Exit Function
    Set wsRaw = ThisWorkbook.Sheets("GLPROOF_RAW")
    lr = wsRaw.Cells(wsRaw.Rows.count, 1).End(xlUp).Row
    If lr < 10 Then lr = wsRaw.Cells(wsRaw.Rows.count, 2).End(xlUp).Row
    For i = 1 To lr
        rowText = UCase(SafeText(wsRaw.Cells(i, 1).Value) & " " & SafeText(wsRaw.Cells(i, 2).Value) & " " & SafeText(wsRaw.Cells(i, 3).Value))
        If InStr(rowText, "ACCOUNT NUMBER") > 0 Then
            cur = ExtractAccountNumber(wsRaw, i)
            If cur <> "" Then
                If Not d.Exists(cur) Then d.Add cur, 0
            End If
        ElseIf InStr(rowText, "BALANCE AS PER GL") > 0 And cur <> "" Then
            If d.Exists(cur) Then d(cur) = TrouverMontantSurLigne(wsRaw, i)
        End If
    Next i
    Set ScanGLProofAccounts = d
End Function

Private Function NormalizeAccountKey(acct As String) As String
    ' Delegue a SAFA_Common (cle de matching unique pour Balance et GL)
    NormalizeAccountKey = SAFA_Common.NormalizeAccountKey(acct)
End Function

Private Function DetermineStatus(balGL As Double, balBal As Double, ecart As Double, tolerance As Double) As String
    If balGL = 0 And balBal = 0 Then
        DetermineStatus = "Sans mouvement"
    ElseIf Abs(Round(ecart, 0)) <= tolerance Then
        DetermineStatus = "OK"
    Else
        DetermineStatus = "Ecart à analyser"
    End If
End Function


Private Function GetAccountType(acctNum As String) As String
    ' Delegue a SAFA_Common.GetAccountClass (classification OHADA unique)
    GetAccountType = SAFA_Common.GetAccountClass(acctNum)
End Function


Private Function IsExcluded(acct As String, name As String, patterns As Variant) As Boolean
    If IsEmpty(patterns) Then Exit Function

    Dim i As Long
    On Error Resume Next
    For i = 1 To UBound(patterns)
        If patterns(i, 1) <> "" Then
            If InStr(1, acct, patterns(i, 1), vbTextCompare) > 0 Or _
               InStr(1, name, patterns(i, 1), vbTextCompare) > 0 Then
                IsExcluded = True
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
End Function

Private Sub FormatHeader(rng As Range)
    ' Delegue a SAFA_Common (implementation unique)
    SAFA_Common.FormatHeader rng
End Sub

Private Sub ApplyConditionalFormatting(ws As Worksheet, lastRow As Long)
    ' Formatage conditionnel sur la colonne Priority
    Dim rngPriority As Range
    Set rngPriority = ws.Range("N2:N" & lastRow + 1)

    ' Supprimer formats existants
    rngPriority.FormatConditions.Delete

    ' CRITICAL = Rouge
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""CRITICAL""")
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    ' HIGH = Orange
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""HIGH""")
        .Interior.Color = RGB(255, 165, 0)
        .Font.Bold = True
    End With

    ' MEDIUM = Jaune
    With rngPriority.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""MEDIUM""")
        .Interior.Color = RGB(255, 255, 0)
    End With
End Sub

' ==============================================================================
' 8. ORGANISATION DES RÉSULTATS
' ==============================================================================

Public Sub Organiser_Resultats()
    Dim arrOrdre As Variant, i As Integer, ws As Worksheet

    ' Ordre optimal pour l'audit
    arrOrdre = Array("DASHBOARD_RISQUE", "EXECUTIVE_SUMMARY", "AUDIT_REPORT", "RECONCIL", _
                     "ECHANTILLON_TEST", "FORENSIC_ANALYSIS", "COMPLIANCE_CHECK", _
                     "TRANSACTION_DATA", "BALANCE_DATA", "GLPROOF_DATA", _
                     "COMPTES_EXCLUS", "HISTORY_LOG", "BALANCE_RAW", "GLPROOF_RAW", _
                     "PARAM", "CONFIG", "AUDIT_TRAIL", "MENU")

    Application.ScreenUpdating = False

    For i = LBound(arrOrdre) To UBound(arrOrdre)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(arrOrdre(i))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible

            On Error Resume Next
            ws.Move Before:=ThisWorkbook.Sheets(1)
            On Error GoTo 0

            ' Zoom et position
            ws.Activate
            ActiveWindow.Zoom = 85
            ws.Range("A1").Select
        End If
        Set ws = Nothing
    Next i

    ' Cacher les feuilles techniques
    On Error Resume Next
    ThisWorkbook.Sheets("AUDIT_TRAIL").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("CONFIG").Visible = xlSheetHidden
    On Error GoTo 0

    ' Activer le Dashboard
    On Error Resume Next
    ThisWorkbook.Sheets("DASHBOARD_RISQUE").Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub

' ==============================================================================
' 9. ORCHESTRATION PRINCIPALE
' ==============================================================================

Public Sub Afficher_Cockpit()
    ' CORRIGÉ: Utiliser SAFA_Console au lieu de feuille MENU (SAFA_Menu) (formulaire non disponible)
    On Error Resume Next
    Call SAFA_Console.Demarrer
    On Error GoTo 0
End Sub

Public Sub Lancer_Demarrage_Differe()
    Call Sequence_Demarrage
    ' CORRIGÉ: Utiliser SAFA_Console au lieu de feuille MENU (SAFA_Menu)
    Call SAFA_Console.Demarrer
End Sub

Public Sub Lancer_Traitement_Complet()
    ' Point d'entrée principal avec gestion d'erreurs globale
    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo TraitementError

    Call InitializeErrorLog
    Call WriteToAuditLog("PROCESS", "Traitement complet démarré")

    ' Étape 0: Calibration automatique (numerotation Balance/GL, signe) sur les donnees
    Call UpdateProgress("Calibration automatique...", 5)
    On Error Resume Next
    Call Auto_Calibration.CalibrerNumerotation(True)
    If Err.Number <> 0 Then Call WriteToAuditLog("ERROR", "Calibration: " & Err.Description)
    On Error GoTo TraitementError

    ' Étape 1: Nettoyage des données
    Call UpdateProgress("Nettoyage Balance...", 10)
    Call NettoyerBalance

    Call UpdateProgress("Nettoyage GL Proof...", 20)
    Call NettoyerGLProof

    ' Étape 2: Rapprochement
    Call UpdateProgress("Construction Rapprochement...", 30)
    Call ConstruireRapprochement

    ' Étape 3: Analyses Forensiques
    Call UpdateProgress("Génération Base Transactions...", 40)
    Call Forensic_Rules.Generer_Base_Transactions

    Call UpdateProgress("Audit Automatique...", 50)
    Call Forensic_Rules.Lancer_Audit_Automatique

    Call UpdateProgress("Analyses Forensiques...", 60)
    Call Forensic_Rules.Lancer_Analyses_Forensic

    Call UpdateProgress("Analyse Benford...", 70)
    Call Forensic_Rules.Lancer_Benford_Enhanced

    ' Étape 4: Intelligence Artificielle
    Call UpdateProgress("Analyse IA par Compte...", 80)
    Call Advanced_AI.Lancer_IA_Par_Compte

    ' Étape GL Monitoring (règles de la Knowledge Sharing Session DAI du 09/09/2026)
    Call UpdateProgress("GL Monitoring...", 83)
    On Error Resume Next
    Call GL_Monitoring.Lancer_GL_Monitoring(False)
    If Err.Number <> 0 Then Call WriteToAuditLog("ERROR", "GL Monitoring: " & Err.Description)
    On Error GoTo TraitementError

    Call UpdateProgress("Gestion Historique...", 85)
    Call Advanced_AI.Gerer_Historique

    ' Étape 5: Conformité Réglementaire
    Call UpdateProgress("Vérification Conformité...", 90)
    Call Regulatory_Compliance.Lancer_Verification_Conformite

    ' Étape 6: Finalisation
    Call UpdateProgress("Génération Rapports...", 95)
    Call Advanced_AI.Finaliser_Rapport

    Call UpdateProgress("Organisation Résultats...", 98)
    Call Organiser_Resultats

    ' Terminé
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    Call WriteToAuditLog("PROCESS", "Traitement complet terminé", "Durée: " & Format(elapsedTime, "0.00") & "s")
    Call Afficher_Resume_Final(elapsedTime)

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

TraitementError:
    Call LogError("Core_Engine", "Lancer_Traitement_Complet", Err.Number, Err.Description)
    MsgBox "Erreur critique lors du traitement:" & vbCrLf & vbCrLf & _
           "Module: Core_Engine" & vbCrLf & _
           "Erreur: " & Err.Description, vbCritical, "Erreur S.A.F.A"
    Resume Cleanup
End Sub

Private Sub UpdateProgress(message As String, percent As Integer)
    Application.StatusBar = "S.A.F.A [" & percent & "%] - " & message
    DoEvents
End Sub

Private Sub Afficher_Resume_Final(elapsedTime As Double)
    Dim nbEcarts As Long, nbAlertes As Long, nbCritical As Long
    Dim msg As String

    On Error Resume Next
    nbEcarts = Application.CountIf(ThisWorkbook.Sheets("RECONCIL").Range("F:F"), "Ecart à analyser")
    nbAlertes = ThisWorkbook.Sheets("AUDIT_REPORT").Cells(Rows.count, 1).End(xlUp).Row - 1
    nbCritical = Application.CountIf(ThisWorkbook.Sheets("RECONCIL").Range("N:N"), "CRITICAL")
    On Error GoTo 0

    If nbAlertes < 0 Then nbAlertes = 0

    msg = "═══════════════════════════════════════" & vbCrLf & _
          "     ANALYSE S.A.F.A v" & SAFA_VERSION & " TERMINÉE" & vbCrLf & _
          "═══════════════════════════════════════" & vbCrLf & vbCrLf & _
          "Temps d'exécution: " & Format(elapsedTime, "0.00") & " secondes" & vbCrLf & vbCrLf & _
          "RÉSULTATS:" & vbCrLf & _
          "  • Écarts à analyser: " & nbEcarts & vbCrLf & _
          "  • Alertes détectées: " & nbAlertes & vbCrLf & _
          "  • Risques CRITIQUES: " & nbCritical & vbCrLf & vbCrLf & _
          "Consultez le DASHBOARD_RISQUE pour les détails."

    ThisWorkbook.Sheets("DASHBOARD_RISQUE").Activate
    MsgBox msg, vbInformation, "Analyse Terminée"
End Sub

' Alias pour compatibilité
Public Sub Lancer_Traitement_Seul()
    Call Lancer_Traitement_Complet
End Sub
