Attribute VB_Name = "SAFA_Tests"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SAFA_TESTS v10.0
' ==============================================================================
' Description: Smoke test de bout en bout, executable dans Excel.
'   1. Genere un jeu de donnees demo (Demo_Data) avec anomalies injectees
'   2. Execute le pipeline complet (Core_Engine.Lancer_Traitement_Complet)
'   3. Verifie que chaque feuille attendue existe et contient des donnees
'   4. Verifie que chaque anomalie injectee a bien ete detectee
'   5. Execute les analyses avancees et les rapports
'   Resultats: feuille TEST_RESULTS (PASS / FAIL par test)
' ==============================================================================

Private Const RESULT_SHEET As String = "TEST_RESULTS"

Private mWs As Worksheet
Private mRow As Long
Private mPass As Long
Private mFail As Long

' ==============================================================================
' POINT D'ENTREE
' ==============================================================================

Public Sub RunSmokeTestUI()
    Dim ok As Boolean
    If MsgBox("Le test automatique va REMPLACER les donnees importees par des donnees de demonstration," & vbCrLf & _
              "puis executer tout le pipeline (1 a 3 minutes)." & vbCrLf & vbCrLf & "Continuer ?", _
              vbYesNo + vbQuestion, "S.A.F.A - Tests") = vbNo Then Exit Sub

    ok = RunSmokeTest()

    On Error Resume Next
    ThisWorkbook.Sheets(RESULT_SHEET).Activate
    On Error GoTo 0

    MsgBox "Tests termines: " & mPass & " PASS / " & mFail & " FAIL" & vbCrLf & vbCrLf & _
           IIf(ok, "Tous les tests sont passes.", "Consultez la feuille TEST_RESULTS pour le detail."), _
           IIf(ok, vbInformation, vbExclamation), "S.A.F.A - Tests"
End Sub

Public Function RunSmokeTest() As Boolean
    Dim t0 As Double
    t0 = Timer

    Call InitResults

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- 1. Donnees demo ----
    Call RunStep("Generation donnees demo", "Demo_Data.GenerateDemoData")
    Call AssertSheetHasRows("BALANCE_RAW contient des lignes", "BALANCE_RAW", 10)
    Call AssertSheetHasRows("GLPROOF_RAW contient des lignes", "GLPROOF_RAW", 50)

    ' ---- 2. Pipeline complet ----
    Call RunStep("Pipeline complet (Lancer_Traitement_Complet)", "Core_Engine.Lancer_Traitement_Complet")
    Call AssertSheetHasRows("BALANCE_DATA genere", "BALANCE_DATA", 10)
    Call AssertSheetHasRows("GLPROOF_DATA genere", "GLPROOF_DATA", 10)
    Call AssertSheetHasRows("RECONCIL genere", "RECONCIL", 10)
    Call AssertSheetHasRows("TRANSACTION_DATA genere", "TRANSACTION_DATA", 50)
    Call AssertSheetHasRows("AUDIT_REPORT contient des alertes", "AUDIT_REPORT", 1)

    ' ---- 3. Rapprochement: comptes a ecart injectes ----
    Call AssertReconcilStatus("Ecarts injectes detectes (statut 'Ecart a analyser')", _
                              Demo_Data.DemoAccounts("ECART"), SAFA_Common.RECONCIL_STATUT_ECART)
    Call AssertReconcilPriority("Au moins un compte injecte en CRITICAL ou HIGH", Demo_Data.DemoAccounts("ECART"))

    ' ---- 4. Detecteurs forensiques (au moins une alerte sur un compte injecte) ----
    Call AssertAlertForAccounts("Structuration (FRD-005 / LAB-002) detectee", Demo_Data.DemoAccounts("STRUCTURING"))
    Call AssertAlertForAccounts("Seuil LAB/FT (FRD-003 / LAB-001) detecte", Demo_Data.DemoAccounts("LAB"))
    Call AssertAlertForAccounts("Transactions week-end (FRD-002) detectees", Demo_Data.DemoAccounts("WEEKEND"))
    Call AssertAlertForAccounts("Mots-cles suspects (FRD-001) detectes", Demo_Data.DemoAccounts("KEYWORD"))
    Call AssertAlertForAccounts("Doublons (FRD-006) detectes", Demo_Data.DemoAccounts("DUPLICATE"))
    Call AssertAlertForAccounts("Circularite / layering (FRD-004) detectee", Array("MULTIPLE"))
    Call AssertAlertForAccounts("Compte dormant reactive (IA-004) detecte", Demo_Data.DemoAccounts("DORMANT"))
    Call AssertAlertForAccounts("Anomalie Z-Score (IA-001/002) detectee", Demo_Data.DemoAccounts("ZSCORE"))
    Call AssertSheetHasRows("Analyse Benford executee (FORENSIC_ANALYSIS)", "FORENSIC_ANALYSIS", 9)
    Call AssertAlertRef("Non-conformite Benford (BEN-001) detectee (25% de soldes en 8/9)", "BEN-001")

    ' ---- 5. Conformite ----
    Call RunStep("Conformite reglementaire", "Regulatory_Compliance.Lancer_Verification_Conformite")
    Call AssertSheetHasRows("COMPLIANCE_CHECK contient des regles", "COMPLIANCE_CHECK", 3)
    Call AssertComplianceStatus("Suspens > 90 jours (COBAC-001) NON-CONFORME", "COBAC-001", "NON-CONFORME")

    ' ---- 5b. GL Monitoring (regles de la seance DAI du 09/09/2026) ----
    Call RunStep("GL Monitoring", "GL_Monitoring.Lancer_GL_Monitoring_Silencieux")
    Call AssertSheetHasRows("GL_MONITORING contient des constats", "GL_MONITORING", 5)
    Call AssertSheetExists("PROOFABLE_UNIVERSE cree", "PROOFABLE_UNIVERSE")
    Call AssertSheetExists("GL_RATING cree", "GL_RATING")
    Call AssertGLM("GLM-002 proof manquant (comptes Balance seule)", "GLM-002", Demo_Data.DemoAccounts("BALANCE_ONLY"))
    Call AssertGLM("GLM-003 sens anormal (banque a solde crediteur)", "GLM-003", Demo_Data.DemoAccounts("SENSE"))
    Call AssertGLM("GLM-004 proxy a solde non nul", "GLM-004", Demo_Data.DemoAccounts("PROXY"))
    Call AssertGLM("GLM-004 suspens a solde non nul", "GLM-004", Demo_Data.DemoAccounts("SUSPENS"))
    Call AssertGLM("GLM-005 debit sur compte de produit", "GLM-005", Demo_Data.DemoAccounts("REVENUE_DEBIT"))
    Call AssertGLM("GLM-006 charge constatee d'avance non amortie", "GLM-006", Demo_Data.DemoAccounts("PREPAID"))
    Call AssertGLM("GLM-009 ecart ATM over-aged", "GLM-009", Demo_Data.DemoAccounts("CASH_DIFF"))
    Call AssertGLM("GLM-010 caisse au-dela de la limite", "GLM-010", Demo_Data.DemoAccounts("CASH_LIMIT"))

    ' ---- 6. Rapports ----
    Call RunStepWb("Generation rapports (Report_Generator)", "Report_Generator.GenerateFullReport")
    Call AssertSheetHasRows("EXECUTIVE_SUMMARY genere", "EXECUTIVE_SUMMARY", 5)
    Call AssertSheetHasRows("DASHBOARD_RISQUE genere", "DASHBOARD_RISQUE", 5)
    Call AssertSheetHasRows("FORENSIC_REPORT genere", "FORENSIC_REPORT", 5)
    Call AssertSheetHasRows("COMPLIANCE_REPORT genere", "COMPLIANCE_REPORT", 5)
    Call AssertSheetHasRows("ECHANTILLON_TEST genere", "ECHANTILLON_TEST", 3)
    Call AssertCellNotEmpty("Top 10 du dashboard rempli (pas un placeholder)", "DASHBOARD_RISQUE", "B", "TOP 10 COMPTES A RISQUE", 2, 2)

    ' ---- 7. Analyses avancees ----
    Call RunStep("Analyse temporelle", "Temporal_Analysis.Lancer_Analyse_Temporelle")
    Call AssertSheetExists("TEMPORAL_ANALYSIS cree", "TEMPORAL_ANALYSIS")
    Call RunStep("Analyse reseau", "Network_Analysis.Lancer_Analyse_Reseau")
    Call AssertSheetExists("NETWORK_ANALYSIS cree", "NETWORK_ANALYSIS")
    Call RunStep("Diagnostic systeme", "Auto_Diagnostic.LancerDiagnosticComplet")
    Call AssertSheetExists("SYSTEM_DIAGNOSTIC cree", "SYSTEM_DIAGNOSTIC")

    ' ---- 8. Securite / crypto (optionnel: module Crypto_Provider) ----
    Call RunCryptoSelfTest

    ' ---- 9. Integrite journal d'audit ----
    Call RunAuditTrailCheck

    ' ---- Etat Excel propre ----
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    Call WriteSummary(Timer - t0)

    RunSmokeTest = (mFail = 0)
End Function

' ==============================================================================
' ETAPES
' ==============================================================================

Private Sub RunStep(name As String, macro As String)
    Dim t As Double
    t = Timer
    On Error Resume Next
    Err.Clear
    Application.Run macro
    If Err.Number <> 0 Then
        Call Record(name, "Execution sans erreur", "Erreur " & Err.Number & ": " & Err.Description, False)
    Else
        Call Record(name, "Execution sans erreur", "OK (" & Format(Timer - t, "0.0") & " s)", True)
    End If
    On Error GoTo 0
End Sub

Private Sub RunStepWb(name As String, macro As String)
    ' Variante pour les procedures prenant le classeur en parametre
    Dim t As Double
    t = Timer
    On Error Resume Next
    Err.Clear
    Application.Run macro, ThisWorkbook
    If Err.Number <> 0 Then
        Call Record(name, "Execution sans erreur", "Erreur " & Err.Number & ": " & Err.Description, False)
    Else
        Call Record(name, "Execution sans erreur", "OK (" & Format(Timer - t, "0.0") & " s)", True)
    End If
    On Error GoTo 0
End Sub

Private Sub RunCryptoSelfTest()
    Dim res As Variant
    On Error Resume Next
    Err.Clear
    res = Application.Run("Crypto_Provider.SelfTest")
    If Err.Number <> 0 Then
        Call Record("Crypto: auto-test", "OK", "Module Crypto_Provider absent (" & Err.Description & ")", False)
    Else
        Call Record("Crypto: auto-test (SHA-256, HMAC, PBKDF2, AES)", "OK", CStr(res), (CStr(res) = "OK"))
    End If
    On Error GoTo 0
End Sub

Private Sub RunAuditTrailCheck()
    Dim res As Variant
    On Error Resume Next
    Err.Clear
    res = Application.Run("Security_Module.VerifyAuditTrailIntegrity")
    If Err.Number <> 0 Then
        Call Record("Journal d'audit: integrite (chaine + hash)", "Valide", "Non verifiable (" & Err.Description & ")", False)
    Else
        Dim detail As String
        detail = Application.Run("Security_Module.LastIntegrityReport")
        Call Record("Journal d'audit: integrite (chaine + hash)", "Valide", IIf(detail = "", CStr(res), detail), CBool(res))
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' ASSERTIONS
' ==============================================================================

Private Sub AssertSheetExists(name As String, sheetName As String)
    Dim ok As Boolean
    ok = SAFA_Common.FeuilleExiste(sheetName)
    Call Record(name, "Feuille " & sheetName & " presente", IIf(ok, "Presente", "ABSENTE"), ok)
End Sub

Private Sub AssertSheetHasRows(name As String, sheetName As String, minRows As Long)
    Dim n As Long
    If Not SAFA_Common.FeuilleExiste(sheetName) Then
        Call Record(name, ">= " & minRows & " lignes", "Feuille absente", False)
        Exit Sub
    End If
    On Error Resume Next
    n = Application.CountA(ThisWorkbook.Sheets(sheetName).Columns(1)) - 1
    If n < 0 Then n = 0
    On Error GoTo 0
    Call Record(name, ">= " & minRows & " lignes", n & " lignes", n >= minRows)
End Sub

Private Sub AssertCellNotEmpty(name As String, sheetName As String, col As String, anchorText As String, rowOffset As Long, colIdx As Long)
    ' Cherche anchorText dans la colonne A, puis verifie que la cellule (ligne+rowOffset, colIdx) est non vide
    Dim ws As Worksheet, f As Range
    If Not SAFA_Common.FeuilleExiste(sheetName) Then
        Call Record(name, "Valeur presente", "Feuille absente", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set f = ws.Columns(1).Find(What:=anchorText, LookAt:=xlPart, MatchCase:=False)
    If f Is Nothing Then
        Call Record(name, "Valeur presente", "Ancre '" & anchorText & "' introuvable", False): Exit Sub
    End If
    Dim v As String
    v = SAFA_Common.SafeText(ws.Cells(f.Row + rowOffset, colIdx).Value)
    Call Record(name, "Valeur presente", IIf(v = "", "VIDE", v), v <> "")
End Sub

Private Sub AssertReconcilStatus(name As String, accounts As Variant, expectedStatus As String)
    Dim ws As Worksheet, lastRow As Long, i As Long, found As Long, total As Long
    Dim acct As String, k As Long
    If Not SAFA_Common.FeuilleExiste("RECONCIL") Then
        Call Record(name, "Statut '" & expectedStatus & "'", "RECONCIL absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("RECONCIL")
    lastRow = ws.Cells(ws.Rows.Count, SAFA_Common.RECONCIL_COL_COMPTE).End(xlUp).Row
    total = UBound(accounts) - LBound(accounts) + 1
    For k = LBound(accounts) To UBound(accounts)
        For i = 2 To lastRow
            If NormKey(ws.Cells(i, SAFA_Common.RECONCIL_COL_COMPTE).Value) = NormKey(accounts(k)) Then
                If SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.RECONCIL_COL_STATUT).Value) = expectedStatus Then found = found + 1
                Exit For
            End If
        Next i
    Next k
    Call Record(name, total & " comptes", found & " / " & total, found = total)
End Sub

Private Sub AssertReconcilPriority(name As String, accounts As Variant)
    Dim ws As Worksheet, lastRow As Long, i As Long, k As Long, hit As Boolean, p As String
    If Not SAFA_Common.FeuilleExiste("RECONCIL") Then
        Call Record(name, "CRITICAL/HIGH", "RECONCIL absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("RECONCIL")
    lastRow = ws.Cells(ws.Rows.Count, SAFA_Common.RECONCIL_COL_COMPTE).End(xlUp).Row
    For k = LBound(accounts) To UBound(accounts)
        For i = 2 To lastRow
            If NormKey(ws.Cells(i, SAFA_Common.RECONCIL_COL_COMPTE).Value) = NormKey(accounts(k)) Then
                p = UCase(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.RECONCIL_COL_PRIORITY).Value))
                If p = "CRITICAL" Or p = "HIGH" Then hit = True
                Exit For
            End If
        Next i
        If hit Then Exit For
    Next k
    Call Record(name, "CRITICAL ou HIGH", IIf(hit, "Trouve", "Aucun"), hit)
End Sub

Private Sub AssertAlertForAccounts(name As String, accounts As Variant)
    ' Robuste au layout: cherche le compte dans TOUTES les colonnes de AUDIT_REPORT
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, i As Long, j As Long, k As Long
    Dim hit As Boolean, cellKey As String
    If Not SAFA_Common.FeuilleExiste("AUDIT_REPORT") Then
        Call Record(name, "Alerte presente", "AUDIT_REPORT absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("AUDIT_REPORT")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 9
    For i = 2 To lastRow
        For j = 1 To lastCol
            cellKey = NormKey(ws.Cells(i, j).Value)
            If Len(cellKey) > 6 Then
                For k = LBound(accounts) To UBound(accounts)
                    If InStr(1, cellKey, NormKey(accounts(k))) > 0 Then hit = True: Exit For
                Next k
            End If
            If hit Then Exit For
        Next j
        If hit Then Exit For
    Next i
    Call Record(name, "Alerte sur compte injecte", IIf(hit, "Alerte trouvee (ligne " & i & ")", "AUCUNE alerte"), hit)
End Sub

Private Sub AssertGLM(name As String, rule As String, accounts As Variant)
    ' Verifie qu'un constat GL_MONITORING existe pour la regle ET l'un des comptes
    Dim ws As Worksheet, lastRow As Long, i As Long, k As Long, hit As Boolean
    If Not SAFA_Common.FeuilleExiste("GL_MONITORING") Then
        Call Record(name, "Constat " & rule, "GL_MONITORING absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("GL_MONITORING")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If UCase(SAFA_Common.SafeText(ws.Cells(i, 1).Value)) = UCase(rule) Then
            For k = LBound(accounts) To UBound(accounts)
                If NormKey(ws.Cells(i, 2).Value) = NormKey(accounts(k)) Then hit = True: Exit For
            Next k
        End If
        If hit Then Exit For
    Next i
    Call Record(name, "Constat " & rule & " sur compte injecte", IIf(hit, "Trouve (ligne " & i & ")", "AUCUN"), hit)
End Sub

Private Sub AssertAlertRef(name As String, refPrefix As String)
    ' Verifie qu'au moins une alerte AUDIT_REPORT a une Ref commencant par refPrefix
    Dim ws As Worksheet, lastRow As Long, i As Long, hit As Boolean
    If Not SAFA_Common.FeuilleExiste("AUDIT_REPORT") Then
        Call Record(name, "Alerte " & refPrefix, "AUDIT_REPORT absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("AUDIT_REPORT")
    lastRow = ws.Cells(ws.Rows.Count, SAFA_Common.AUDIT_COL_REF).End(xlUp).Row
    For i = 2 To lastRow
        If Left(UCase(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.AUDIT_COL_REF).Value)), Len(refPrefix)) = UCase(refPrefix) Then
            hit = True: Exit For
        End If
    Next i
    Call Record(name, "Alerte " & refPrefix, IIf(hit, "Trouvee", "AUCUNE"), hit)
End Sub

Private Sub AssertComplianceStatus(name As String, ruleId As String, expectedStatus As String)
    Dim ws As Worksheet, lastRow As Long, i As Long, st As String, found As Boolean
    If Not SAFA_Common.FeuilleExiste("COMPLIANCE_CHECK") Then
        Call Record(name, expectedStatus, "COMPLIANCE_CHECK absent", False): Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("COMPLIANCE_CHECK")
    lastRow = ws.Cells(ws.Rows.Count, SAFA_Common.COMPLIANCE_COL_ID).End(xlUp).Row
    For i = 2 To lastRow
        If UCase(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.COMPLIANCE_COL_ID).Value)) = UCase(ruleId) Then
            st = UCase(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.COMPLIANCE_COL_STATUT).Value))
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        Call Record(name, expectedStatus, "Regle " & ruleId & " absente", False)
    Else
        Call Record(name, expectedStatus, st, (st = UCase(expectedStatus)))
    End If
End Sub

Private Function NormKey(v As Variant) As String
    NormKey = UCase(Trim(Replace(Replace(SAFA_Common.SafeText(v), " ", ""), "'", "")))
End Function

' ==============================================================================
' RESULTATS
' ==============================================================================

Private Sub InitResults()
    Set mWs = SAFA_Common.GetOrCreateSheet(RESULT_SHEET, True)
    mRow = 1: mPass = 0: mFail = 0
    mWs.Range("A1:D1").Value = Array("Test", "Attendu", "Obtenu", "Statut")
    SAFA_Common.FormatHeader mWs.Range("A1:D1")
    mRow = 2
End Sub

Private Sub Record(name As String, expected As String, obtained As String, passed As Boolean)
    On Error Resume Next
    mWs.Cells(mRow, 1).Value = name
    mWs.Cells(mRow, 2).Value = expected
    mWs.Cells(mRow, 3).Value = obtained
    mWs.Cells(mRow, 4).Value = IIf(passed, "PASS", "FAIL")
    If passed Then
        mWs.Cells(mRow, 4).Interior.Color = RGB(200, 255, 200)
        mPass = mPass + 1
    Else
        mWs.Cells(mRow, 4).Interior.Color = RGB(255, 180, 180)
        mWs.Cells(mRow, 4).Font.Bold = True
        mFail = mFail + 1
    End If
    mRow = mRow + 1
    Debug.Print IIf(passed, "[PASS] ", "[FAIL] ") & name & " -> " & obtained
End Sub

Private Sub WriteSummary(duration As Double)
    On Error Resume Next
    mRow = mRow + 1
    mWs.Cells(mRow, 1).Value = "TOTAL"
    mWs.Cells(mRow, 2).Value = (mPass + mFail) & " tests"
    mWs.Cells(mRow, 3).Value = mPass & " PASS / " & mFail & " FAIL en " & Format(duration, "0.0") & " s"
    mWs.Cells(mRow, 4).Value = IIf(mFail = 0, "PASS", "FAIL")
    mWs.Range(mWs.Cells(mRow, 1), mWs.Cells(mRow, 4)).Font.Bold = True
    mWs.Cells(mRow, 4).Interior.Color = IIf(mFail = 0, RGB(120, 220, 140), RGB(255, 120, 120))
    mWs.Columns("A:D").AutoFit
    mWs.Columns("C").ColumnWidth = 60
    Call SAFA_Common.WriteAuditLog("TEST", "Smoke test: " & mPass & " PASS / " & mFail & " FAIL")
End Sub
