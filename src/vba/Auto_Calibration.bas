Attribute VB_Name = "Auto_Calibration"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: AUTO_CALIBRATION v10.2
' ==============================================================================
' Description: Calibration AUTOMATIQUE sur les donnees, sans hypothese a priori:
'   1. Numerotation Balance <-> GL Proof: plusieurs transformations candidates
'      sont evaluees (identique, injection SOL, suppression du code devise,
'      chiffres seuls, suffixe de n chiffres); celle qui rapproche le plus de
'      comptes SANS collision est retenue.
'   2. Signe des soldes Balance vs GL: si |GL + Balance| cadre mieux que
'      |GL - Balance| sur les comptes rapproches, la Balance est inversee.
'   3. Convention debit/credit: vote majoritaire sur les comptes dont la famille
'      est connue (classe ou libelle), separement pour le bilan et le resultat.
'   Resultats: variables globales SAFA_Common.g_* + feuille CALIBRATION + journal.
'   Mode force possible via settings.json (account_normalization, sign_convention).
' ==============================================================================

Private Const SHEET_CAL As String = "CALIBRATION"

' ==============================================================================
' 1 + 2. NUMEROTATION ET SIGNE BALANCE / GL
' ==============================================================================

Public Function CalibrerNumerotation(Optional silent As Boolean = True) As String
    On Error GoTo ErrHandler

    Dim dBal As Object, dGL As Object, glKeys As Object, balKeys As Object
    Dim modes As Variant, m As Variant, bestMode As String, bestRate As Double, bestMatched As Long
    Dim k As Variant, key As String, nBal As Long, nGL As Long, matched As Long, collisions As Long
    Dim agree As Long, flip As Long, bestAgree As Long, bestFlip As Long
    Dim sol As String, subCode As String, tol As Double, rate As Double
    Dim report As String, ws As Worksheet, r As Long
    Dim cfg As Config_Manager.GeneralConfig, forced As String

    Call ResetCalibration

    Set dBal = Core_Engine.ScanBalanceAccounts()
    Set dGL = Core_Engine.ScanGLProofAccounts()
    nBal = dBal.Count: nGL = dGL.Count

    sol = "799": tol = 100
    On Error Resume Next
    sol = SAFA_Common.SafeText(ThisWorkbook.Sheets("PARAM").Range("F2").Value)
    tol = SAFA_Common.SafeVal(ThisWorkbook.Sheets("PARAM").Range("B2").Value, 100)
    cfg = Config_Manager.GetGeneralConfig()
    On Error GoTo ErrHandler
    If sol = "" Then sol = "799"
    subCode = "0"
    forced = UCase(Trim(cfg.AccountNormalization))
    If forced = "" Then forced = "AUTO"

    Set ws = SAFA_Common.GetOrCreateSheet(SHEET_CAL, True)
    ws.Range("A1").Value = "CALIBRATION AUTOMATIQUE SUR LES DONNEES"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Executee le " & Format(Now, "dd/mm/yyyy hh:nn") & " - Balance: " & nBal & " comptes, GL Proof: " & nGL & " comptes, SOL " & sol & ", tolerance " & tol
    ws.Range("A3").Value = "Numerotation retenue": ws.Range("A4").Value = "Taux de rapprochement"
    ws.Range("A5").Value = "Inversion du signe Balance": ws.Range("A6").Value = "Convention bilan"
    ws.Range("A7").Value = "Convention resultat": ws.Range("A8").Value = "Mode impose (config)"
    ws.Range("A3:A8").Font.Bold = True
    ws.Range("B8").Value = forced

    ws.Range("A10:F10").Value = Array("Mode", "Description", "Comptes rapproches", "Taux", "Collisions", "Signe: |GL-Bal| ok / |GL+Bal| ok")
    SAFA_Common.FormatHeader ws.Range("A10:F10")
    r = 11

    If nBal = 0 Or nGL = 0 Then
        ws.Range("B3").Value = "SOL_INJECT": ws.Range("B4").Value = "n/a (donnees absentes)"
        SAFA_Common.g_BalanceTransform = "SOL_INJECT"
        CalibrerNumerotation = "SOL_INJECT"
        Exit Function
    End If

    modes = Array("NONE", "SOL_INJECT", "STRIP_CCY", "DIGITS", "SUFFIX12", "SUFFIX10", "SUFFIX8", "SUFFIX6")
    bestMode = "": bestRate = -1

    For Each m In modes
        ' Jeu de cles GL pour ce mode
        Set glKeys = CreateObject("Scripting.Dictionary")
        For Each k In dGL.Keys
            key = KeyFor(CStr(m), CStr(k), False, sol, subCode)
            If key <> "" Then
                If Not glKeys.Exists(key) Then glKeys.Add key, dGL(k)
            End If
        Next k
        ' Cles Balance, rapprochement et signe
        Set balKeys = CreateObject("Scripting.Dictionary")
        matched = 0: collisions = 0: agree = 0: flip = 0
        For Each k In dBal.Keys
            key = KeyFor(CStr(m), CStr(k), True, sol, subCode)
            If key <> "" Then
                If balKeys.Exists(key) Then
                    collisions = collisions + 1
                Else
                    balKeys.Add key, 1
                    If glKeys.Exists(key) Then
                        matched = matched + 1
                        If Abs(CDbl(glKeys(key)) - CDbl(dBal(k))) <= tol Then
                            agree = agree + 1
                        ElseIf Abs(CDbl(glKeys(key)) + CDbl(dBal(k))) <= tol Then
                            flip = flip + 1
                        End If
                    End If
                End If
            End If
        Next k
        rate = matched / Application.WorksheetFunction.Min(nBal, nGL)

        ws.Cells(r, 1).Value = CStr(m)
        ws.Cells(r, 2).Value = ModeDescription(CStr(m))
        ws.Cells(r, 3).Value = matched
        ws.Cells(r, 4).Value = rate: ws.Cells(r, 4).NumberFormat = "0.0%"
        ws.Cells(r, 5).Value = collisions
        ws.Cells(r, 6).Value = agree & " / " & flip
        r = r + 1

        ' Un mode avec collisions (deux comptes Balance -> une meme cle) est ecarte
        If collisions = 0 And rate > bestRate + 0.0001 Then
            bestRate = rate: bestMode = CStr(m): bestMatched = matched
            bestAgree = agree: bestFlip = flip
        End If
    Next m

    ' Mode impose par la configuration ?
    If forced <> "AUTO" Then
        If forced = "NONE" Or forced = "SOL_INJECT" Then bestMode = forced
    End If
    If bestMode = "" Then bestMode = "SOL_INJECT"

    ' Appliquer
    Select Case Left(bestMode, 6)
        Case "SUFFIX"
            SAFA_Common.g_BalanceTransform = "NONE"
            SAFA_Common.g_KeyMode = "SUFFIX"
            SAFA_Common.g_KeySuffix = CLng(Mid(bestMode, 7))
        Case Else
            Select Case bestMode
                Case "SOL_INJECT": SAFA_Common.g_BalanceTransform = "SOL_INJECT": SAFA_Common.g_KeyMode = ""
                Case "STRIP_CCY": SAFA_Common.g_BalanceTransform = "NONE": SAFA_Common.g_KeyMode = "STRIP_CCY"
                Case "DIGITS": SAFA_Common.g_BalanceTransform = "NONE": SAFA_Common.g_KeyMode = "DIGITS"
                Case Else: SAFA_Common.g_BalanceTransform = "NONE": SAFA_Common.g_KeyMode = ""
            End Select
    End Select
    SAFA_Common.g_FlipBalanceSign = (bestFlip > bestAgree And bestFlip >= 3)

    For r = 11 To 11 + UBound(modes)
        If ws.Cells(r, 1).Value = bestMode Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = RGB(214, 240, 220)
    Next r
    ws.Range("B3").Value = bestMode
    ws.Range("B4").Value = bestRate: ws.Range("B4").NumberFormat = "0.0%"
    ws.Range("B5").Value = IIf(SAFA_Common.g_FlipBalanceSign, "OUI (la Balance est inversee: |GL+Bal| cadre sur " & bestFlip & " comptes contre " & bestAgree & ")", "NON")
    If bestRate < 0.5 Then
        ws.Range("B4").Interior.Color = RGB(255, 170, 170)
        ws.Range("D3").Value = "ATTENTION: moins de 50% des comptes rapproches - verifier les fichiers sources (SOL, format des numeros)"
        ws.Range("D3").Font.Color = RGB(200, 0, 0): ws.Range("D3").Font.Bold = True
    End If
    ws.Columns("A:F").AutoFit

    report = "Numerotation: " & bestMode & " (" & Format(bestRate, "0.0%") & ", " & bestMatched & " comptes)" & _
             IIf(SAFA_Common.g_FlipBalanceSign, " - signe Balance inverse", "")
    Call SAFA_Common.WriteAuditLog("CALIBRATION", report)
    If Not silent Then MsgBox report, IIf(bestRate < 0.5, vbExclamation, vbInformation), "Calibration automatique"

    CalibrerNumerotation = bestMode
    Exit Function

ErrHandler:
    Call SAFA_Common.LogError("Auto_Calibration", "CalibrerNumerotation", Err.Number, Err.Description)
    If SAFA_Common.g_BalanceTransform = "" Then SAFA_Common.g_BalanceTransform = "SOL_INJECT"
    CalibrerNumerotation = SAFA_Common.g_BalanceTransform
End Function

Private Function KeyFor(mode As String, acct As String, isBalance As Boolean, sol As String, subCode As String) As String
    ' Cle de rapprochement d'un numero de compte pour un mode candidat
    Dim s As String
    s = UCase(Trim(Replace(Replace(Replace(acct, " ", ""), "'", ""), "-", "")))
    If s = "" Then Exit Function
    Select Case mode
        Case "NONE"
            KeyFor = s
        Case "SOL_INJECT"
            If isBalance Then
                If Len(s) >= 8 Then KeyFor = Left(s, 3) & sol & subCode & Mid(s, 8) Else KeyFor = s
            Else
                KeyFor = s
            End If
        Case "STRIP_CCY"
            Do While Len(s) > 0 And Not (Left(s, 1) Like "[0-9]")
                s = Mid(s, 2)
            Loop
            KeyFor = s
        Case "DIGITS"
            KeyFor = SAFA_Common.DigitsOnly(s)
        Case Else
            If Left(mode, 6) = "SUFFIX" Then
                s = SAFA_Common.DigitsOnly(s)
                If Len(s) > CLng(Mid(mode, 7)) Then s = Right(s, CLng(Mid(mode, 7)))
                KeyFor = s
            Else
                KeyFor = s
            End If
    End Select
End Function

Private Function ModeDescription(mode As String) As String
    Select Case mode
        Case "NONE": ModeDescription = "Numeros identiques dans la Balance et le GL Proof"
        Case "SOL_INJECT": ModeDescription = "Balance: 3 caracteres + SOL + sous-code + suffixe (extraction Finacle agence)"
        Case "STRIP_CCY": ModeDescription = "Suppression du code devise / lettres en tete des deux cotes"
        Case "DIGITS": ModeDescription = "Chiffres seulement, des deux cotes"
        Case Else: ModeDescription = "Chiffres seulement, compares sur les " & Mid(mode, 7) & " derniers"
    End Select
End Function

' ==============================================================================
' 3. CONVENTION DEBIT / CREDIT (bilan et resultat separement)
' ==============================================================================

Public Sub CalibrerSignes(dictBal As Object, Optional silent As Boolean = True)
    On Error GoTo ErrHandler

    Dim k As Variant, name As String, bal As Double, cls As GL_Monitoring.AccountClass
    Dim nBS As Long, okBS As Long, nPL As Long, okPL As Long, hasPAL As Boolean
    Dim cfgM As Config_Manager.GLMonitoringConfig, forced As String
    Dim ws As Worksheet

    cfgM = Config_Manager.GetGLMonitoringConfig()
    forced = UCase(Trim(cfgM.SignConvention))
    If forced = "" Then forced = "AUTO"

    ' Convention imposee ?
    If forced = "DEBIT_POSITIVE" Then
        SAFA_Common.g_DebitPositiveBS = 1: SAFA_Common.g_DebitPositivePL = 1
    ElseIf forced = "CREDIT_POSITIVE" Then
        SAFA_Common.g_DebitPositiveBS = -1: SAFA_Common.g_DebitPositivePL = -1
    Else
        ' Vote: sous l'hypothese "debit positif", un actif / une charge est > 0, un passif / produit < 0
        For Each k In dictBal.Keys
            name = dictBal(k)(0): bal = dictBal(k)(1)
            If bal <> 0 Then
                If InStr(UCase(k), UCase(cfgM.PLPrefix)) > 0 Then hasPAL = True
                cls = GL_Monitoring.ClassifyAccount(CStr(k), name, bal)
                Select Case cls.Family
                    Case "ACTIF"
                        If cls.Nature <> "TRANSIT" And cls.Nature <> "SUSPENS" And cls.Nature <> "PROXY" Then
                            nBS = nBS + 1: If bal > 0 Then okBS = okBS + 1
                        End If
                    Case "PASSIF", "CAPITAUX"
                        nBS = nBS + 1: If bal < 0 Then okBS = okBS + 1
                    Case "CHARGE", "PRODUIT"
                        ' Ne compter que les comptes dont la famille vient de la classe ou du libelle (pas du signe)
                        If InStr(cls.Reason, "classe") > 0 Or GL_Monitoring.ContainsAny(UCase(name), cfgM.KwRevenue) Or GL_Monitoring.ContainsAny(UCase(name), cfgM.KwExpense) Then
                            nPL = nPL + 1
                            If (cls.Family = "CHARGE" And bal > 0) Or (cls.Family = "PRODUIT" And bal < 0) Then okPL = okPL + 1
                        End If
                End Select
            End If
        Next k

        SAFA_Common.g_DebitPositiveBS = Verdict(okBS, nBS, 1)
        ' Resultat: a defaut de preuve, convention Finacle PAL (negatif = charge => credit positif), sinon debit positif
        SAFA_Common.g_DebitPositivePL = Verdict(okPL, nPL, IIf(hasPAL, -1, 1))
    End If

    ' Feuille CALIBRATION
    If SAFA_Common.FeuilleExiste(SHEET_CAL) Then
        Set ws = ThisWorkbook.Sheets(SHEET_CAL)
        ws.Range("B6").Value = ConventionLabel(SAFA_Common.g_DebitPositiveBS) & IIf(forced = "AUTO", " (vote: " & okBS & "/" & nBS & " comptes de bilan coherents avec 'debit positif')", " (impose)")
        ws.Range("B7").Value = ConventionLabel(SAFA_Common.g_DebitPositivePL) & IIf(forced = "AUTO", " (vote: " & okPL & "/" & nPL & " comptes de resultat coherents avec 'debit positif')", " (impose)")
        ws.Columns("B").AutoFit
    End If

    Call SAFA_Common.WriteAuditLog("CALIBRATION", "Conventions de signe: bilan " & ConventionLabel(SAFA_Common.g_DebitPositiveBS) & ", resultat " & ConventionLabel(SAFA_Common.g_DebitPositivePL))
    If Not silent Then MsgBox "Convention bilan: " & ConventionLabel(SAFA_Common.g_DebitPositiveBS) & vbCrLf & "Convention resultat: " & ConventionLabel(SAFA_Common.g_DebitPositivePL), vbInformation, "Calibration des signes"
    Exit Sub

ErrHandler:
    Call SAFA_Common.LogError("Auto_Calibration", "CalibrerSignes", Err.Number, Err.Description)
    If SAFA_Common.g_DebitPositiveBS = 0 Then SAFA_Common.g_DebitPositiveBS = 1
    If SAFA_Common.g_DebitPositivePL = 0 Then SAFA_Common.g_DebitPositivePL = 1
End Sub

Private Function Verdict(ok As Long, n As Long, fallback As Integer) As Integer
    ' 1 = debit positif, -1 = credit positif; majorite a 60 % requise, sinon repli
    If n < 5 Then Verdict = fallback: Exit Function
    If ok / n >= 0.6 Then
        Verdict = 1
    ElseIf ok / n <= 0.4 Then
        Verdict = -1
    Else
        Verdict = fallback
    End If
End Function

Public Function ConventionLabel(v As Integer) As String
    Select Case v
        Case 1: ConventionLabel = "DEBIT_POSITIVE"
        Case -1: ConventionLabel = "CREDIT_POSITIVE"
        Case Else: ConventionLabel = "INDETERMINEE"
    End Select
End Function

' ==============================================================================
' RESET
' ==============================================================================

Public Sub ResetCalibration()
    SAFA_Common.g_BalanceTransform = ""
    SAFA_Common.g_KeyMode = ""
    SAFA_Common.g_KeySuffix = 0
    SAFA_Common.g_FlipBalanceSign = False
    SAFA_Common.g_DebitPositiveBS = 0
    SAFA_Common.g_DebitPositivePL = 0
End Sub
