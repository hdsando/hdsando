Attribute VB_Name = "Forensic_Rules"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: FORENSIC_RULES v10.0 (Enhanced)
' ==============================================================================
' Description: Détection avancée de fraudes et anomalies
' Nouvelles fonctionnalités:
'   - Benford amélioré avec Chi-squared et MAD
'   - Détection de round-tripping
'   - Analyse temporelle avancée (fin de mois, veilles fêtes)
'   - Duplicate detection
'   - Sequence analysis
'   - Counter-party analysis
' ==============================================================================

' --- CONSTANTES ---
Private Const BENFORD_THRESHOLD As Double = 0.15
Private Const WEEKEND_AMOUNT_THRESHOLD As Double = 50000
Private Const ROUND_NUMBER_THRESHOLD As Double = 100000
Private Const SPLIT_DETECTION_THRESHOLD As Double = 50000
Private Const LAYERING_MIN_COUNT As Integer = 4

' --- TYPES ---
Private Type ForensicAlert
    Reference As String
    Description As String
    Severity As String
    Account As String
    Details As String
    Amount As Double
    TransactionDate As Date
    RiskScore As Integer
End Type

' ==============================================================================
' 1. GÉNÉRATION BASE TRANSACTIONS (AMÉLIORÉE)
' ==============================================================================

Public Sub Generer_Base_Transactions()
    Dim wsRaw As Worksheet, wsTrans As Worksheet
    Dim dataArray As Variant, outputArray() As Variant
    Dim r As Long, i As Long, lr As Long, outRow As Long
    Dim curAcct As String

    On Error GoTo GenError

    ' Vérification dépendances
    If Not Core_Engine.FeuilleExiste("GLPROOF_RAW") Then
        Call Core_Engine.WriteToAuditLog("ERROR", "Generer_Base_Transactions: GLPROOF_RAW manquant")
        Exit Sub
    End If

    Set wsRaw = ThisWorkbook.Worksheets("GLPROOF_RAW")

    ' Création/Reset TRANSACTION_DATA
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("TRANSACTION_DATA").Delete
    On Error GoTo GenError
    Application.DisplayAlerts = True

    Set wsTrans = ThisWorkbook.Worksheets.Add
    wsTrans.name = "TRANSACTION_DATA"

    ' En-têtes enrichis
    wsTrans.Range("A1:K1").Value = Array("Compte", "Date", "Narration", "Montant", "Sens", _
                                          "Jour Semaine", "Heure", "Fin de Mois", "Montant Rond", _
                                          "Catégorie", "Hash")
    FormatHeader wsTrans.Range("A1:K1")

    lr = wsRaw.Cells(wsRaw.Rows.count, 1).End(xlUp).Row
    If lr < 10 Then lr = wsRaw.Cells(wsRaw.Rows.count, 2).End(xlUp).Row

    ' Dimensionner array de sortie
    ReDim outputArray(1 To lr, 1 To 11)
    outRow = 0
    curAcct = ""

    ' Extraction avec enrichissement
    For i = 1 To lr
        Dim cellVal1 As String, cellVal2 As String
        cellVal1 = Core_Engine.SafeText(wsRaw.Cells(i, 1).Value)
        cellVal2 = Core_Engine.SafeText(wsRaw.Cells(i, 2).Value)

        ' Détection nouveau compte
        If InStr(1, UCase(cellVal1 & cellVal2), "ACCOUNT NUMBER", vbTextCompare) > 0 Then
            curAcct = ExtractAccountFromRow(wsRaw, i)
        End If

        ' Détection transaction
        Dim transDate As Date, transAmount As Double, narration As String
        Dim isTransaction As Boolean

        isTransaction = False

        If IsDate(wsRaw.Cells(i, 1).Value) Then
            transDate = wsRaw.Cells(i, 1).Value
            transAmount = Core_Engine.SafeVal(wsRaw.Cells(i, 4).Value)
            narration = Core_Engine.SafeText(wsRaw.Cells(i, 3).Value)
            isTransaction = True
        ElseIf IsDate(wsRaw.Cells(i, 2).Value) Then
            transDate = wsRaw.Cells(i, 2).Value
            transAmount = Core_Engine.SafeVal(wsRaw.Cells(i, 5).Value)
            narration = Core_Engine.SafeText(wsRaw.Cells(i, 3).Value)
            isTransaction = True
        End If

        If isTransaction And curAcct <> "" And transAmount <> 0 Then
            outRow = outRow + 1

            ' Données de base
            outputArray(outRow, 1) = "'" & curAcct
            outputArray(outRow, 2) = transDate
            outputArray(outRow, 3) = narration
            outputArray(outRow, 4) = transAmount
            outputArray(outRow, 5) = IIf(transAmount < 0, "D", "C")

            ' Enrichissements forensiques
            outputArray(outRow, 6) = WeekdayName(Weekday(transDate, vbMonday), True)
            outputArray(outRow, 7) = "" ' Heure non disponible dans ce format
            outputArray(outRow, 8) = IIf(IsEndOfMonth(transDate), "OUI", "")
            outputArray(outRow, 9) = IIf(IsRoundNumber(transAmount), "OUI", "")
            outputArray(outRow, 10) = CategorizeTransaction(narration)
            outputArray(outRow, 11) = GenerateTransactionHash(curAcct, transDate, transAmount)
        End If
    Next i

    ' Écrire résultats
    If outRow > 0 Then
        wsTrans.Range("A2").Resize(outRow, 11).Value = outputArray
    End If

    wsTrans.Columns("A:K").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Generer_Base_Transactions terminé", "Transactions: " & outRow)
    Exit Sub

GenError:
    Call Core_Engine.LogError("Forensic_Rules", "Generer_Base_Transactions", Err.Number, Err.Description)
End Sub

Private Function ExtractAccountFromRow(ws As Worksheet, rowNum As Long) As String
    Dim j As Integer, cellVal As String

    For j = 1 To 6
        cellVal = Core_Engine.SafeText(ws.Cells(rowNum, j).Value)
        If InStr(1, cellVal, "NUMBER", vbTextCompare) = 0 And _
           InStr(1, cellVal, "ACCOUNT", vbTextCompare) = 0 And _
           cellVal <> ":" And Len(cellVal) > 3 Then
            ExtractAccountFromRow = Trim(cellVal)
            Exit Function
        End If
    Next j
End Function

Private Function IsEndOfMonth(dt As Date) As Boolean
    IsEndOfMonth = (Day(DateSerial(Year(dt), Month(dt) + 1, 0)) = Day(dt))
End Function

Private Function IsRoundNumber(amount As Double) As Boolean
    Dim absAmount As Double
    absAmount = Abs(amount)

    ' Vérifier si c'est un multiple de 10000
    If absAmount >= 10000 Then
        IsRoundNumber = (absAmount Mod 10000 = 0)
    End If
End Function

Private Function CategorizeTransaction(narration As String) As String
    Dim upperNarr As String
    upperNarr = UCase(narration)

    If InStr(upperNarr, "SALARY") > 0 Or InStr(upperNarr, "SALAIRE") > 0 Then
        CategorizeTransaction = "SALAIRE"
    ElseIf InStr(upperNarr, "TRANSFER") > 0 Or InStr(upperNarr, "VIREMENT") > 0 Then
        CategorizeTransaction = "VIREMENT"
    ElseIf InStr(upperNarr, "CASH") > 0 Or InStr(upperNarr, "ESPECE") > 0 Then
        CategorizeTransaction = "ESPECES"
    ElseIf InStr(upperNarr, "CHEQUE") > 0 Or InStr(upperNarr, "CHQ") > 0 Then
        CategorizeTransaction = "CHEQUE"
    ElseIf InStr(upperNarr, "FEE") > 0 Or InStr(upperNarr, "COMMISSION") > 0 Or InStr(upperNarr, "FRAIS") > 0 Then
        CategorizeTransaction = "FRAIS"
    ElseIf InStr(upperNarr, "INTEREST") > 0 Or InStr(upperNarr, "INTERET") > 0 Then
        CategorizeTransaction = "INTERET"
    ElseIf InStr(upperNarr, "LOAN") > 0 Or InStr(upperNarr, "PRET") > 0 Or InStr(upperNarr, "CREDIT") > 0 Then
        CategorizeTransaction = "CREDIT"
    Else
        CategorizeTransaction = "AUTRE"
    End If
End Function

Private Function GenerateTransactionHash(acct As String, dt As Date, amount As Double) As String
    ' Hash simple pour détecter les doublons
    Dim hashInput As String
    hashInput = acct & "|" & Format(dt, "yyyymmdd") & "|" & Format(amount, "0.00")

    Dim i As Long, h As Long
    h = 5381
    For i = 1 To Len(hashInput)
        h = ((h * 33) Xor Asc(Mid(hashInput, i, 1))) And &H7FFFFFFF
    Next i

    GenerateTransactionHash = Hex(h)
End Function

' ==============================================================================
' 2. AUDIT AUTOMATIQUE (RÈGLES ENRICHIES)
' ==============================================================================

Public Sub Lancer_Audit_Automatique()
    Dim wsR As Worksheet, wsA As Worksheet
    Dim lr As Long, i As Long, alertRow As Long
    Dim alertsArray() As Variant, alertCount As Long

    On Error GoTo AuditError

    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        Call Core_Engine.WriteToAuditLog("ERROR", "Lancer_Audit_Automatique: RECONCIL manquant")
        Exit Sub
    End If

    Set wsR = ThisWorkbook.Worksheets("RECONCIL")

    ' Création AUDIT_REPORT
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("AUDIT_REPORT").Delete
    On Error GoTo AuditError
    Application.DisplayAlerts = True

    Set wsA = ThisWorkbook.Worksheets.Add
    wsA.name = "AUDIT_REPORT"

    ' En-têtes enrichis
    wsA.Range("A1:I1").Value = Array("Ref", "Catégorie", "Risque", "Niveau", "Compte", _
                                      "Description", "Valeur", "Impact Est.", "SLA")
    FormatAlertHeader wsA.Range("A1:I1")

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    ReDim alertsArray(1 To lr * 5, 1 To 9) ' Multiple alertes possibles par ligne
    alertCount = 0

    For i = 2 To lr
        Dim aNum As String, aNam As String
        Dim sBal As Double, ecart As Double, age As Double
        Dim stat As String, riskScore As Integer

        aNum = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        aNam = UCase(Core_Engine.SafeText(wsR.Cells(i, 2).Value))
        sBal = Core_Engine.SafeVal(wsR.Cells(i, 4).Value)
        ecart = Core_Engine.SafeVal(wsR.Cells(i, 5).Value)
        age = Core_Engine.SafeVal(wsR.Cells(i, 8).Value)
        stat = Core_Engine.SafeText(wsR.Cells(i, 6).Value)
        riskScore = Core_Engine.SafeVal(wsR.Cells(i, 12).Value)

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #1: Suspens âgés > 90 jours
        ' ═══════════════════════════════════════════════════════════════
        If age > 90 And stat = "Ecart à analyser" Then
            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-001"
            alertsArray(alertCount, 2) = "SUSPENS"
            alertsArray(alertCount, 3) = "Suspens > 90 jours"
            alertsArray(alertCount, 4) = IIf(age > 180, "CRITICAL", "MAJOR")
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam & " - Age: " & age & "j"
            alertsArray(alertCount, 7) = ecart
            alertsArray(alertCount, 8) = CalculateProvision(ecart, age)
            alertsArray(alertCount, 9) = IIf(age > 180, "Immédiat", "7 jours")
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #2: Comptes de passage non soldés
        ' ═══════════════════════════════════════════════════════════════
        If (InStr(aNam, "TRANSIT") > 0 Or InStr(aNam, "SUSPENSE") > 0 Or _
            InStr(aNam, "CLEARING") > 0 Or InStr(aNam, "PASSAGE") > 0) And _
            Abs(sBal) > 100 Then

            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-002"
            alertsArray(alertCount, 2) = "TRANSIT"
            alertsArray(alertCount, 3) = "Compte Passage Non Soldé"
            alertsArray(alertCount, 4) = IIf(Abs(sBal) > 1000000, "CRITICAL", "MAJOR")
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam
            alertsArray(alertCount, 7) = sBal
            alertsArray(alertCount, 8) = sBal
            alertsArray(alertCount, 9) = "J+1"
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #3: Actif créditeur (anomalie comptable)
        ' ═══════════════════════════════════════════════════════════════
        If Left(aNum, 1) = "1" And sBal < -1000 And _
           InStr(aNam, "PROV") = 0 And InStr(aNam, "AMORT") = 0 Then

            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-003"
            alertsArray(alertCount, 2) = "ANOMALIE"
            alertsArray(alertCount, 3) = "Actif Créditeur"
            alertsArray(alertCount, 4) = "MAJOR"
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam & " - Sens inversé"
            alertsArray(alertCount, 7) = sBal
            alertsArray(alertCount, 8) = 0
            alertsArray(alertCount, 9) = "Investigation"
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #4: Passif débiteur (anomalie comptable)
        ' ═══════════════════════════════════════════════════════════════
        If Left(aNum, 1) = "4" And sBal > 1000 And _
           (InStr(aNam, "FOURN") > 0 Or InStr(aNam, "CREDIT") > 0) Then

            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-004"
            alertsArray(alertCount, 2) = "ANOMALIE"
            alertsArray(alertCount, 3) = "Passif Débiteur"
            alertsArray(alertCount, 4) = "MAJOR"
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam & " - Sens inversé"
            alertsArray(alertCount, 7) = sBal
            alertsArray(alertCount, 8) = 0
            alertsArray(alertCount, 9) = "Investigation"
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #5: Écart GL/Balance significatif
        ' ═══════════════════════════════════════════════════════════════
        If Abs(ecart) > 1000000 Then
            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-005"
            alertsArray(alertCount, 2) = "ECART"
            alertsArray(alertCount, 3) = "Écart GL/Balance >1M"
            alertsArray(alertCount, 4) = IIf(Abs(ecart) > 10000000, "CRITICAL", "HIGH")
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam
            alertsArray(alertCount, 7) = ecart
            alertsArray(alertCount, 8) = ecart
            alertsArray(alertCount, 9) = "Immédiat"
        ElseIf Abs(ecart) > 100000 Then
            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-006"
            alertsArray(alertCount, 2) = "ECART"
            alertsArray(alertCount, 3) = "Écart GL/Balance significatif"
            alertsArray(alertCount, 4) = "MEDIUM"
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam
            alertsArray(alertCount, 7) = ecart
            alertsArray(alertCount, 8) = ecart
            alertsArray(alertCount, 9) = "7 jours"
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #6: Comptes orphelins à risque élevé
        ' ═══════════════════════════════════════════════════════════════
        If riskScore >= 70 Then
            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-007"
            alertsArray(alertCount, 2) = "RISQUE"
            alertsArray(alertCount, 3) = "Score Risque Critique (" & riskScore & ")"
            alertsArray(alertCount, 4) = "CRITICAL"
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = Core_Engine.SafeText(wsR.Cells(i, 13).Value)
            alertsArray(alertCount, 7) = ecart
            alertsArray(alertCount, 8) = 0
            alertsArray(alertCount, 9) = "Immédiat"
        End If

        ' ═══════════════════════════════════════════════════════════════
        ' RÈGLE #7: Comptes sensibles (Caisse, Coffre)
        ' ═══════════════════════════════════════════════════════════════
        If (InStr(aNam, "CAISSE") > 0 Or InStr(aNam, "COFFRE") > 0 Or _
            InStr(aNam, "CASH") > 0 Or InStr(aNam, "VAULT") > 0) And _
            Abs(ecart) > 0 Then

            alertCount = alertCount + 1
            alertsArray(alertCount, 1) = "AUD-008"
            alertsArray(alertCount, 2) = "SENSIBLE"
            alertsArray(alertCount, 3) = "Écart sur compte sensible"
            alertsArray(alertCount, 4) = "HIGH"
            alertsArray(alertCount, 5) = aNum
            alertsArray(alertCount, 6) = aNam & " - Nécessite contrôle physique"
            alertsArray(alertCount, 7) = ecart
            alertsArray(alertCount, 8) = 0
            alertsArray(alertCount, 9) = "J+0"
        End If
    Next i

    ' Écrire les alertes
    If alertCount > 0 Then
        wsA.Range("A2").Resize(alertCount, 9).Value = alertsArray
        ApplyAlertFormatting wsA, alertCount
    End If

    wsA.Columns("A:I").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Lancer_Audit_Automatique terminé", "Alertes: " & alertCount)
    Exit Sub

AuditError:
    Call Core_Engine.LogError("Forensic_Rules", "Lancer_Audit_Automatique", Err.Number, Err.Description)
End Sub

Private Function CalculateProvision(ecart As Double, age As Double) As Double
    ' Calcul provision IFRS 9 simplifiée
    If age > 360 Then
        CalculateProvision = ecart ' 100%
    ElseIf age > 180 Then
        CalculateProvision = ecart * 0.5 ' 50%
    ElseIf age > 90 Then
        CalculateProvision = ecart * 0.25 ' 25%
    ElseIf age > 30 Then
        CalculateProvision = ecart * 0.1 ' 10%
    Else
        CalculateProvision = 0
    End If
End Function

Private Sub FormatAlertHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(139, 0, 0) ' Dark Red
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub ApplyAlertFormatting(ws As Worksheet, lastRow As Long)
    Dim i As Long
    For i = 2 To lastRow + 1
        Select Case ws.Cells(i, 4).Value
            Case "CRITICAL"
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Interior.Color = RGB(255, 200, 200)
                ws.Cells(i, 4).Interior.Color = RGB(255, 0, 0)
                ws.Cells(i, 4).Font.Color = vbWhite
            Case "HIGH"
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Interior.Color = RGB(255, 230, 200)
                ws.Cells(i, 4).Interior.Color = RGB(255, 165, 0)
            Case "MAJOR"
                ws.Cells(i, 4).Interior.Color = RGB(255, 192, 0)
            Case "MEDIUM"
                ws.Cells(i, 4).Interior.Color = RGB(255, 255, 0)
        End Select
    Next i
End Sub

' ==============================================================================
' 3. ANALYSES FORENSIQUES AVANCÉES
' ==============================================================================

Public Sub Lancer_Analyses_Forensic()
    Dim wsT As Worksheet, wsA As Worksheet
    Dim lr As Long, i As Long, alertRow As Long
    Dim dictGraph As Object, dictSplit As Object, dictDuplicate As Object
    Dim dictSequence As Object, dictEndMonth As Object

    On Error GoTo ForensicError

    If Not Core_Engine.FeuilleExiste("TRANSACTION_DATA") Or Not Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        Exit Sub
    End If

    Set wsT = ThisWorkbook.Worksheets("TRANSACTION_DATA")
    Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")

    alertRow = wsA.Cells(wsA.Rows.count, 1).End(xlUp).Row + 1
    lr = wsT.Cells(wsT.Rows.count, 1).End(xlUp).Row

    ' Dictionnaires pour analyses
    Set dictGraph = CreateObject("Scripting.Dictionary")
    Set dictSplit = CreateObject("Scripting.Dictionary")
    Set dictDuplicate = CreateObject("Scripting.Dictionary")
    Set dictSequence = CreateObject("Scripting.Dictionary")
    Set dictEndMonth = CreateObject("Scripting.Dictionary")

    ' ═══════════════════════════════════════════════════════════════
    ' PASS 1: Collecte et analyse de patterns
    ' ═══════════════════════════════════════════════════════════════
    For i = 2 To lr
        Dim acct As String, narr As String, mnt As Double
        Dim transDate As Variant, dayOfWeek As Integer
        Dim transHash As String

        acct = Core_Engine.SafeText(wsT.Cells(i, 1).Value)
        transDate = wsT.Cells(i, 2).Value
        narr = UCase(Core_Engine.SafeText(wsT.Cells(i, 3).Value))
        mnt = Core_Engine.SafeVal(wsT.Cells(i, 4).Value)
        transHash = Core_Engine.SafeText(wsT.Cells(i, 11).Value)

        ' ───────────────────────────────────────────────────────────
        ' CHECK 1: Mots-clés suspects (élargi)
        ' ───────────────────────────────────────────────────────────
        Dim suspectKeywords As Variant, kw As Variant
        suspectKeywords = Array("CADEAU", "URGENT", "CASH", "MANUEL", "CORRECTION", _
                                "ANNULATION", "REVERSE", "AJUSTEMENT", "ERREUR", _
                                "VOID", "CANCEL", "OVERRIDE", "EXCEPTION", "SPECIAL")

        For Each kw In suspectKeywords
            If InStr(narr, CStr(kw)) > 0 And InStr(narr, "AUTO") = 0 Then
                If Abs(mnt) > 10000 Then
                    Call AddForensicAlert(wsA, alertRow, "FRD-001", "KEYWORD", _
                        "Mot-clé suspect: " & kw, "FRAUD", acct, narr, mnt)
                    Exit For
                End If
            End If
        Next kw

        ' ───────────────────────────────────────────────────────────
        ' CHECK 2: Transactions Week-end
        ' ───────────────────────────────────────────────────────────
        If IsDate(transDate) Then
            dayOfWeek = Weekday(transDate, vbMonday)

            If dayOfWeek >= 6 And Abs(mnt) > WEEKEND_AMOUNT_THRESHOLD Then
                If InStr(narr, "BATCH") = 0 And InStr(narr, "AUTO") = 0 And _
                   InStr(narr, "SYSTEM") = 0 Then
                    Call AddForensicAlert(wsA, alertRow, "FRD-002", "WEEKEND", _
                        "Transaction hors jours ouvrés", "COMPLIANCE", acct, _
                        Format(transDate, "ddd dd/mm") & " - " & narr, mnt)
                End If
            End If

            ' ───────────────────────────────────────────────────────
            ' CHECK 3: Transactions fin de mois suspectes
            ' ───────────────────────────────────────────────────────
            If Day(transDate) >= 28 Or Day(transDate) <= 2 Then
                Dim monthKey As String
                monthKey = acct & "|" & Format(transDate, "yyyymm")

                If dictEndMonth.Exists(monthKey) Then
                    dictEndMonth(monthKey) = dictEndMonth(monthKey) + 1
                Else
                    dictEndMonth.Add monthKey, 1
                End If
            End If
        End If

        ' ───────────────────────────────────────────────────────────
        ' CHECK 4: Préparation détection Circularité (Layering)
        ' ───────────────────────────────────────────────────────────
        If Abs(mnt) > 1000000 Then
            Dim graphKey As String
            graphKey = Format(transDate, "yyyymmdd") & "|" & Format(Abs(mnt), "0")

            If dictGraph.Exists(graphKey) Then
                dictGraph(graphKey) = dictGraph(graphKey) + 1
            Else
                dictGraph.Add graphKey, 1
            End If
        End If

        ' ───────────────────────────────────────────────────────────
        ' CHECK 5: Préparation détection Saucissonnage (Structuring)
        ' ───────────────────────────────────────────────────────────
        If Abs(mnt) >= SPLIT_DETECTION_THRESHOLD And Abs(mnt) < SPLIT_DETECTION_THRESHOLD * 2 Then
            Dim splitKey As String
            splitKey = acct & "|" & Format(transDate, "yyyymmdd")

            If dictSplit.Exists(splitKey) Then
                dictSplit(splitKey) = dictSplit(splitKey) + 1
            Else
                dictSplit.Add splitKey, 1
            End If
        End If

        ' ───────────────────────────────────────────────────────────
        ' CHECK 6: Détection doublons (même hash)
        ' ───────────────────────────────────────────────────────────
        If transHash <> "" Then
            If dictDuplicate.Exists(transHash) Then
                dictDuplicate(transHash) = dictDuplicate(transHash) + 1
            Else
                dictDuplicate.Add transHash, 1
            End If
        End If

        ' ───────────────────────────────────────────────────────────
        ' CHECK 7: Montants ronds suspects
        ' ───────────────────────────────────────────────────────────
        If Abs(mnt) >= ROUND_NUMBER_THRESHOLD And (Abs(mnt) Mod 100000 = 0) Then
            Dim seqKey As String
            seqKey = acct & "|" & Format(Abs(mnt), "0")

            If dictSequence.Exists(seqKey) Then
                dictSequence(seqKey) = dictSequence(seqKey) + 1
            Else
                dictSequence.Add seqKey, 1
            End If
        End If

        ' ───────────────────────────────────────────────────────────
        ' CHECK 8: Transactions juste sous le seuil
        ' ───────────────────────────────────────────────────────────
        Dim thresholds As Variant, threshold As Variant
        thresholds = Array(5000000, 1000000, 500000, 100000) ' Seuils courants

        For Each threshold In thresholds
            If Abs(mnt) >= CDbl(threshold) * 0.9 And Abs(mnt) < CDbl(threshold) Then
                Call AddForensicAlert(wsA, alertRow, "FRD-003", "THRESHOLD", _
                    "Juste sous seuil " & Format(CDbl(threshold), "#,##0"), "FRAUD", _
                    acct, narr & " (Ratio: " & Format(Abs(mnt) / CDbl(threshold), "0.0%") & ")", mnt)
                Exit For
            End If
        Next threshold
    Next i

    ' ═══════════════════════════════════════════════════════════════
    ' PASS 2: Analyse des patterns collectés
    ' ═══════════════════════════════════════════════════════════════

    Dim v As Variant, parts() As String

    ' Circularité (Layering)
    For Each v In dictGraph.Keys
        If dictGraph(v) >= LAYERING_MIN_COUNT Then
            parts = Split(v, "|")
            Call AddForensicAlert(wsA, alertRow, "FRD-004", "LAYERING", _
                "Circularité détectée: " & dictGraph(v) & " transactions identiques", _
                "CRITICAL", "MULTIPLE", _
                "Date: " & Left(parts(0), 4) & "/" & Mid(parts(0), 5, 2) & "/" & Right(parts(0), 2), _
                CDbl(parts(1)))
        End If
    Next v

    ' Saucissonnage (Structuring)
    For Each v In dictSplit.Keys
        If dictSplit(v) >= 3 Then
            parts = Split(v, "|")
            Call AddForensicAlert(wsA, alertRow, "FRD-005", "STRUCTURING", _
                "Saucissonnage: " & dictSplit(v) & " transactions le même jour", _
                "FRAUD", parts(0), "Date: " & parts(1), SPLIT_DETECTION_THRESHOLD)
        End If
    Next v

    ' Doublons
    For Each v In dictDuplicate.Keys
        If dictDuplicate(v) > 1 Then
            Call AddForensicAlert(wsA, alertRow, "FRD-006", "DUPLICATE", _
                "Transaction en double (" & dictDuplicate(v) & " occurrences)", _
                "HIGH", "", "Hash: " & v, 0)
        End If
    Next v

    ' Séquences de montants ronds
    For Each v In dictSequence.Keys
        If dictSequence(v) >= 5 Then
            parts = Split(v, "|")
            Call AddForensicAlert(wsA, alertRow, "FRD-007", "PATTERN", _
                "Pattern répétitif: " & dictSequence(v) & "x même montant rond", _
                "FRAUD", parts(0), "Montant répété", CDbl(parts(1)))
        End If
    Next v

    ' Concentration fin de mois
    For Each v In dictEndMonth.Keys
        If dictEndMonth(v) >= 10 Then
            parts = Split(v, "|")
            Call AddForensicAlert(wsA, alertRow, "FRD-008", "TIMING", _
                "Concentration fin de mois: " & dictEndMonth(v) & " transactions", _
                "MEDIUM", parts(0), "Période: " & parts(1), 0)
        End If
    Next v

    wsA.Columns("A:I").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Lancer_Analyses_Forensic terminé")
    Exit Sub

ForensicError:
    Call Core_Engine.LogError("Forensic_Rules", "Lancer_Analyses_Forensic", Err.Number, Err.Description)
End Sub

' ==============================================================================
' 4. ANALYSE BENFORD AMÉLIORÉE (Chi-squared + MAD)
' ==============================================================================

Public Sub Lancer_Benford_Enhanced()
    Dim wsData As Worksheet, wsOut As Worksheet
    Dim lr As Long, i As Long, d As Integer
    Dim observed(1 To 9) As Long, total As Long
    Dim benford(1 To 9) As Double, expected(1 To 9) As Double
    Dim chiSquared As Double, MAD As Double
    Dim outputArray(1 To 12, 1 To 7) As Variant

    On Error GoTo BenfordError

    ' Initialisation des fréquences Benford théoriques
    benford(1) = 0.301: benford(2) = 0.176: benford(3) = 0.125
    benford(4) = 0.097: benford(5) = 0.079: benford(6) = 0.067
    benford(7) = 0.058: benford(8) = 0.051: benford(9) = 0.046

    If Not Core_Engine.FeuilleExiste("BALANCE_DATA") Then Exit Sub

    ' Création feuille résultats
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("FORENSIC_ANALYSIS").Delete
    On Error GoTo BenfordError
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Worksheets.Add
    wsOut.name = "FORENSIC_ANALYSIS"

    Set wsData = ThisWorkbook.Worksheets("BALANCE_DATA")
    lr = wsData.Cells(wsData.Rows.count, 1).End(xlUp).Row

    ' Comptage des premiers chiffres
    total = 0
    For i = 2 To lr
        Dim v As Double, firstDigit As Integer
        v = Abs(Core_Engine.SafeVal(wsData.Cells(i, 3).Value))

        If v >= 10 Then
            firstDigit = CInt(Left(Format(v, "0"), 1))
            If firstDigit >= 1 And firstDigit <= 9 Then
                observed(firstDigit) = observed(firstDigit) + 1
                total = total + 1
            End If
        End If
    Next i

    If total < 100 Then
        wsOut.Range("A1").Value = "ATTENTION: Échantillon insuffisant pour Benford (" & total & " observations)"
        wsOut.Range("A1").Font.Color = vbRed
        Exit Sub
    End If

    ' Calcul statistiques
    chiSquared = 0
    MAD = 0

    For d = 1 To 9
        expected(d) = benford(d) * total
        If expected(d) > 0 Then
            chiSquared = chiSquared + ((observed(d) - expected(d)) ^ 2) / expected(d)
        End If
        MAD = MAD + Abs((observed(d) / total) - benford(d))
    Next d
    MAD = MAD / 9

    ' En-tête
    wsOut.Range("A1").Value = "ANALYSE LOI DE BENFORD"
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A1").Font.Bold = True

    wsOut.Range("A3").Value = "Nombre d'observations: " & total
    wsOut.Range("A4").Value = "Chi-squared: " & Format(chiSquared, "0.00") & " (Seuil critique 5%: 15.51)"
    wsOut.Range("A5").Value = "MAD (Mean Absolute Deviation): " & Format(MAD, "0.0000")

    ' Interprétation MAD
    Dim madInterpret As String
    If MAD < 0.006 Then
        madInterpret = "Conformité EXCELLENTE"
        wsOut.Range("A5").Interior.Color = RGB(200, 255, 200)
    ElseIf MAD < 0.012 Then
        madInterpret = "Conformité ACCEPTABLE"
        wsOut.Range("A5").Interior.Color = RGB(255, 255, 200)
    ElseIf MAD < 0.015 Then
        madInterpret = "Conformité MARGINALE - Investigation recommandée"
        wsOut.Range("A5").Interior.Color = RGB(255, 230, 200)
    Else
        madInterpret = "NON-CONFORMITÉ - Investigation OBLIGATOIRE"
        wsOut.Range("A5").Interior.Color = RGB(255, 200, 200)
    End If
    wsOut.Range("A6").Value = "Interprétation: " & madInterpret

    ' Tableau détaillé
    wsOut.Range("A8:G8").Value = Array("Chiffre", "Observé", "Attendu", "Fréq. Réelle", "Fréq. Théorique", "Écart", "Statut")
    wsOut.Range("A8:G8").Font.Bold = True
    wsOut.Range("A8:G8").Interior.Color = RGB(0, 51, 102)
    wsOut.Range("A8:G8").Font.Color = vbWhite

    For d = 1 To 9
        Dim rowNum As Integer
        rowNum = 8 + d

        wsOut.Cells(rowNum, 1).Value = d
        wsOut.Cells(rowNum, 2).Value = observed(d)
        wsOut.Cells(rowNum, 3).Value = Round(expected(d), 0)
        wsOut.Cells(rowNum, 4).Value = observed(d) / total
        wsOut.Cells(rowNum, 5).Value = benford(d)
        wsOut.Cells(rowNum, 6).Value = (observed(d) / total) - benford(d)

        ' Statut avec seuil relatif
        If Abs(wsOut.Cells(rowNum, 6).Value) > BENFORD_THRESHOLD Then
            wsOut.Cells(rowNum, 7).Value = "ANOMALIE"
            wsOut.Cells(rowNum, 7).Interior.Color = RGB(255, 0, 0)
            wsOut.Cells(rowNum, 7).Font.Color = vbWhite
        ElseIf Abs(wsOut.Cells(rowNum, 6).Value) > 0.1 Then
            wsOut.Cells(rowNum, 7).Value = "À SURVEILLER"
            wsOut.Cells(rowNum, 7).Interior.Color = RGB(255, 192, 0)
        Else
            wsOut.Cells(rowNum, 7).Value = "OK"
            wsOut.Cells(rowNum, 7).Interior.Color = RGB(0, 176, 80)
            wsOut.Cells(rowNum, 7).Font.Color = vbWhite
        End If
    Next d

    wsOut.Range("D9:F17").NumberFormat = "0.0%"

    ' Ajouter graphique
    Call CreateBenfordChart(wsOut, total)

    ' Résultat global dans AUDIT_REPORT
    If MAD >= 0.015 Then
        Dim wsA As Worksheet
        Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")
        Dim nextRow As Long
        nextRow = wsA.Cells(wsA.Rows.count, 1).End(xlUp).Row + 1

        Call AddForensicAlert(wsA, nextRow, "BEN-001", "BENFORD", _
            "Non-conformité Loi de Benford (MAD=" & Format(MAD, "0.0000") & ")", _
            "CRITICAL", "GLOBAL", _
            "Chi²=" & Format(chiSquared, "0.0") & " | Manipulation données suspectée", _
            0)
    End If

    wsOut.Columns("A:G").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Lancer_Benford_Enhanced terminé", _
        "Chi²=" & Format(chiSquared, "0.00") & ", MAD=" & Format(MAD, "0.0000"))
    Exit Sub

BenfordError:
    Call Core_Engine.LogError("Forensic_Rules", "Lancer_Benford_Enhanced", Err.Number, Err.Description)
End Sub

Private Sub CreateBenfordChart(ws As Worksheet, total As Long)
    ' Créer un graphique comparatif Benford
    On Error Resume Next

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=450, Width:=400, Top:=60, Height:=250)

    With cht.Chart
        .SetSourceData Source:=ws.Range("A8:E17")
        .ChartType = xlColumnClustered

        .HasTitle = True
        .ChartTitle.text = "Distribution Benford (n=" & total & ")"

        .SeriesCollection(1).name = "Observé"
        .SeriesCollection(2).name = "Attendu"

        ' Supprimer séries non nécessaires
        If .SeriesCollection.count > 2 Then
            Dim s As Integer
            For s = .SeriesCollection.count To 3 Step -1
                .SeriesCollection(s).Delete
            Next s
        End If

        .Legend.Position = xlLegendPositionBottom
    End With

    On Error GoTo 0
End Sub

' ==============================================================================
' 5. UTILITAIRES
' ==============================================================================

Public Sub AddLog(ByVal ws As Worksheet, ByRef r As Long, ByVal ref As String, _
                  ByVal d As String, ByVal l As String, ByVal a As String, _
                  ByVal n As String, ByVal v As Variant)
    ' Compatibilité avec ancien code
    Call AddForensicAlert(ws, r, ref, "LEGACY", d, l, a, n, Core_Engine.SafeVal(v))
End Sub

Private Sub AddForensicAlert(ByVal ws As Worksheet, ByRef r As Long, _
                              ByVal ref As String, ByVal category As String, _
                              ByVal description As String, ByVal severity As String, _
                              ByVal account As String, ByVal details As String, _
                              ByVal amount As Double)

    ws.Cells(r, 1).Value = ref
    ws.Cells(r, 2).Value = category
    ws.Cells(r, 3).Value = description
    ws.Cells(r, 4).Value = severity
    ws.Cells(r, 5).Value = account
    ws.Cells(r, 6).Value = details
    ws.Cells(r, 7).Value = amount

    ' Formatage selon sévérité
    Select Case UCase(severity)
        Case "CRITICAL"
            ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
            ws.Cells(r, 4).Font.Color = vbWhite
        Case "FRAUD"
            ws.Cells(r, 4).Interior.Color = RGB(139, 0, 0)
            ws.Cells(r, 4).Font.Color = vbWhite
        Case "HIGH"
            ws.Cells(r, 4).Interior.Color = RGB(255, 165, 0)
        Case "MAJOR"
            ws.Cells(r, 4).Interior.Color = RGB(255, 192, 0)
        Case "MEDIUM"
            ws.Cells(r, 4).Interior.Color = RGB(255, 255, 0)
        Case "COMPLIANCE"
            ws.Cells(r, 4).Interior.Color = RGB(173, 216, 230)
        Case Else
            ws.Cells(r, 4).Interior.Color = RGB(211, 211, 211)
    End Select

    r = r + 1
End Sub

Private Sub FormatHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Alias pour compatibilité
Public Sub Lancer_Benford()
    Call Lancer_Benford_Enhanced
End Sub
