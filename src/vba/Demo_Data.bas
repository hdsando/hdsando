Attribute VB_Name = "Demo_Data"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: DEMO_DATA v10.0
' ==============================================================================
' Description: Generateur deterministe de donnees de demonstration, dans les
'   formats EXACTS attendus par Core_Engine (NettoyerBalance / NettoyerGLProof)
'   et Forensic_Rules (Generer_Base_Transactions), avec des anomalies injectees
'   sur des comptes connus afin que SAFA_Tests puisse verifier chaque detecteur.
'
' FORMAT BALANCE_RAW (export "Balance" Finacle):
'   ligne 1-2 : titres libres
'   ligne 3   : en-tete  A="Account Number" B="Account Name" C="Closing Balance" D="Currency"
'   lignes 4+ : donnees (compte brut sur 12 car.: 3 car. de classe + "0000" + 5 car. de suffixe)
'   Core_Engine normalise: Left(3) & SOL_ID & "0" & Mid(8)  ->  "471" & "799" & "0" & "00001"
'
' FORMAT GLPROOF_RAW (rapport "GL Proof" Finacle, machine d'etat):
'   A="ACCOUNT NUMBER"       B=":"  C=<compte normalise>
'   A="ACCOUNT NAME"         B=":"  C=<libelle>
'   A="BALANCE AS PER GL"           C=<montant>
'   A="BALANCE AS PER PROOF"        C=<montant>
'   A=<date>  B=<age en jours>  C=<narration>  D=<montant signe>   (n lignes)
'   A="DIFFERENCE"                  C=<montant>            (fin de bloc)
'
' ANOMALIES INJECTEES (voir DemoAccounts pour la liste des comptes):
'   ECART        comptes 1-15 : Balance <> GL (150 000 a 12 000 000 XAF)
'   SUSPENS      comptes 1-3  : suspens avec ecart et items > 90 jours  (COBAC-001)
'   TRANSIT      comptes 4-5  : transit avec ecart et items > 7 jours   (COBAC-002)
'   STRUCTURING  comptes 16-17: 4 ecritures 60-95k le meme jour        (FRD-005, LAB-002)
'   LAB          comptes 18-19: 4 750 000 (juste sous seuil, FRD-003) + 5 200 000 (LAB-001)
'   WEEKEND      comptes 20-22: ecritures > 50k un samedi / dimanche    (FRD-002)
'   KEYWORD      comptes 23-25: narrations CADEAU / URGENT MANUEL / CORRECTION (FRD-001)
'   DUPLICATE    comptes 26-27: ecriture strictement dupliquee          (FRD-006)
'   ROUNDTRIP    comptes 28-30: 2 500 000 le meme jour x6 (A->B->C->A)  (FRD-004)
'   DORMANT      compte 31    : 2 ecritures recentes, cumul > 1M        (IA-004)
'   ZSCORE       compte 32    : 9 ecritures ~100k + 1 de 3 000 000      (IA-001/002)
'   BALANCE_ONLY comptes 33-37: presents en Balance, absents du GL
'   GL_ONLY      comptes 38-42: presents au GL, absents de la Balance
'   PREPAID      compte 43    : charge constatee d'avance non amortie depuis > 30 j (GLM-006)
'   PROXY        compte 44    : compte proxy a solde non nul                     (GLM-004)
'   REVENUE_DEBIT compte 45   : debit sur compte de produit                      (GLM-005)
'   CASH_DIFF    compte 46    : ecart ATM non reconcilie > 180 j                 (GLM-009)
'   CASH_LIMIT   compte 47    : caisse au-dela de la limite de 5 000 000         (GLM-010)
'   SENSE        compte 48    : compte de banque (actif) a solde crediteur       (GLM-003)
'   BENFORD      25% des soldes commencent par 8 ou 9                   (BEN-001)
' ==============================================================================

Private Const SOL_ID As String = "799"
Private Const SHEET_BAL As String = "BALANCE_RAW"
Private Const SHEET_GL As String = "GLPROOF_RAW"

' ==============================================================================
' API PUBLIQUE
' ==============================================================================

Public Function DemoAccounts(kind As String) As Variant
    ' Retourne les numeros de compte normalises des anomalies injectees (String())
    Dim idx As Variant, i As Long
    Dim res() As String

    Select Case UCase(kind)
        Case "SUSPENS":       idx = Array(1, 2, 3)
        Case "TRANSIT":       idx = Array(4, 5)
        Case "ECART":         idx = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
        Case "STRUCTURING":   idx = Array(16, 17)
        Case "LAB":           idx = Array(18, 19)
        Case "WEEKEND":       idx = Array(20, 21, 22)
        Case "KEYWORD":       idx = Array(23, 24, 25)
        Case "DUPLICATE":     idx = Array(26, 27)
        Case "ROUNDTRIP":     idx = Array(28, 29, 30)
        Case "DORMANT":       idx = Array(31)
        Case "ZSCORE":        idx = Array(32)
        Case "BALANCE_ONLY":  idx = Array(33, 34, 35, 36, 37)
        Case "GL_ONLY":       idx = Array(38, 39, 40, 41, 42)
        Case "PREPAID":       idx = Array(43)
        Case "PROXY":         idx = Array(44)
        Case "REVENUE_DEBIT": idx = Array(45)
        Case "CASH_DIFF":     idx = Array(46)
        Case "CASH_LIMIT":    idx = Array(47)
        Case "SENSE":         idx = Array(48)
        Case Else:            idx = Array(0)
    End Select

    ReDim res(0 To UBound(idx))
    For i = 0 To UBound(idx)
        res(i) = CleanAccount(CLng(idx(i)))
    Next i
    DemoAccounts = res
End Function

Public Sub ClearDemoData()
    On Error Resume Next
    If SAFA_Common.FeuilleExiste(SHEET_BAL) Then ThisWorkbook.Sheets(SHEET_BAL).Cells.Clear
    If SAFA_Common.FeuilleExiste(SHEET_GL) Then ThisWorkbook.Sheets(SHEET_GL).Cells.Clear
    On Error GoTo 0
End Sub

Public Sub GenerateDemoData(Optional nbComptes As Long = 300, Optional nbEcritures As Long = 4000, Optional seed As Long = 42)
    On Error GoTo ErrHandler

    Dim wsB As Worksheet, wsG As Worksheet
    Dim n As Long, r As Long, k As Long
    Dim closing() As Double, glBal() As Double
    Dim balArr() As Variant, glArr() As Variant
    Dim glRows As Long, txTotal As Long

    If nbComptes < 60 Then nbComptes = 60
    If nbEcritures < nbComptes * 3 Then nbEcritures = nbComptes * 3

    ' Generateur deterministe (memes donnees a chaque execution pour un seed donne)
    Rnd -1
    Randomize seed

    Application.ScreenUpdating = False

    Call EnsureParamSheet

    ' ---------------------------------------------------------------
    ' 1. Soldes de cloture (Balance) - 25% violent Benford (1er chiffre 8/9)
    ' ---------------------------------------------------------------
    ReDim closing(1 To nbComptes)
    ReDim glBal(1 To nbComptes)
    For n = 1 To nbComptes
        If (n Mod 4) = 0 Then
            closing(n) = (8 + Rnd * 1.99) * (10 ^ (4 + Int(Rnd * 4)))
        Else
            closing(n) = 10 ^ (4 + Rnd * 4)
        End If
        closing(n) = Round(closing(n), 0)
        ' Sens normal: capitaux (1), produits (7) et fournisseurs (401) au credit (negatif)
        If Left(PrefixFor(n), 1) = "1" Or Left(PrefixFor(n), 1) = "7" Or PrefixFor(n) = "401" Then closing(n) = -closing(n)
        ' Caisses dans la limite (5 000 000), sauf le compte 47 injecte
        If (PrefixFor(n) = "531" Or PrefixFor(n) = "571") And n <> 47 Then
            If Abs(closing(n)) > 4500000 Then closing(n) = 4500000 - (n Mod 100) * 1000
        End If
        glBal(n) = closing(n)
    Next n

    ' Comptes injectes pour le GL Monitoring (seance DAI 09/09/2026)
    closing(43) = 4000000: glBal(43) = closing(43)          ' prepaid: solde attendu apres 8 mensualites
    closing(44) = 1250000: glBal(44) = closing(44)          ' proxy non nul
    closing(46) = 950000: glBal(46) = closing(46)           ' ecart ATM non regularise
    closing(47) = 8500000: glBal(47) = closing(47)          ' caisse au-dela de la limite
    closing(48) = -Abs(closing(48)): glBal(48) = closing(48) ' compte de banque (actif) crediteur

    ' Ecarts injectes (comptes 1..15)
    glBal(1) = closing(1) + 12000000
    glBal(2) = closing(2) - 3500000
    glBal(3) = closing(3) + 650000
    glBal(4) = closing(4) - 250000
    glBal(5) = closing(5) + 180000
    For n = 6 To 15
        glBal(n) = closing(n) + Sgn(Rnd - 0.5) * Round(150000 + Rnd * 7850000, 0)
        If glBal(n) = closing(n) Then glBal(n) = closing(n) + 200000
    Next n

    ' ---------------------------------------------------------------
    ' 2. BALANCE_RAW
    ' ---------------------------------------------------------------
    Set wsB = SAFA_Common.GetOrCreateSheet(SHEET_BAL, True)
    wsB.Range("A1").Value = "BALANCE GENERALE - SOL " & SOL_ID & " - DONNEES DE DEMONSTRATION (fictives)"
    wsB.Range("A2").Value = "Edition du " & Format(Date, "dd/mm/yyyy")
    wsB.Range("A3:D3").Value = Array("Account Number", "Account Name", "Closing Balance", "Currency")
    wsB.Range("A3:D3").Font.Bold = True

    ReDim balArr(1 To nbComptes, 1 To 4)
    r = 0
    For n = 1 To nbComptes
        If Not IsGlOnly(n) Then
            r = r + 1
            balArr(r, 1) = "'" & RawAccount(n)
            balArr(r, 2) = NameFor(n)
            balArr(r, 3) = closing(n)
            balArr(r, 4) = "XAF"
        End If
    Next n
    wsB.Range("A4").Resize(r, 4).Value = balArr
    wsB.Cells(4 + r, 1).Value = "Total"
    wsB.Cells(4 + r, 3).Formula = "=SUM(C4:C" & (3 + r) & ")"
    wsB.Columns("A:D").AutoFit

    ' ---------------------------------------------------------------
    ' 3. GLPROOF_RAW (rapport Finacle)
    ' ---------------------------------------------------------------
    Set wsG = SAFA_Common.GetOrCreateSheet(SHEET_GL, True)
    glRows = nbComptes * 8 + nbEcritures + 500
    ReDim glArr(1 To glRows, 1 To 4)
    r = 0
    Call AddRow(glArr, r, "GL PROOF REPORT - SOL " & SOL_ID & " - DONNEES DE DEMONSTRATION (fictives)", "", "", "")
    Call AddRow(glArr, r, "Report date: " & Format(Date, "dd/mm/yyyy"), "", "", "")
    Call AddRow(glArr, r, "", "", "", "")

    Dim avgTx As Double
    avgTx = nbEcritures / nbComptes

    For n = 1 To nbComptes
        If Not IsBalanceOnly(n) Then
            Dim proof As Double
            Call AddRow(glArr, r, "ACCOUNT NUMBER", ":", "'" & CleanAccount(n), "")
            Call AddRow(glArr, r, "ACCOUNT NAME", ":", NameFor(n), "")
            Call AddRow(glArr, r, "BALANCE AS PER GL", "", glBal(n), "")

            ' Les ecritures sont ajoutees apres la ligne PROOF; on reserve la ligne
            Dim proofRow As Long
            r = r + 1
            proofRow = r
            glArr(r, 1) = "BALANCE AS PER PROOF"

            proof = WriteTransactions(glArr, r, n, avgTx, txTotal)
            glArr(proofRow, 3) = proof

            Call AddRow(glArr, r, "DIFFERENCE", "", glBal(n) - proof, "")
            Call AddRow(glArr, r, "", "", "", "")
        End If
        If r > glRows - 60 Then Exit For
    Next n

    wsG.Range("A1").Resize(r, 4).Value = glArr
    wsG.Columns("A").NumberFormat = "dd/mm/yyyy"
    wsG.Columns("D").NumberFormat = "#,##0"
    wsG.Columns("A:D").AutoFit

    Application.ScreenUpdating = True

    Call SAFA_Common.WriteAuditLog("DEMO", "Donnees de demonstration generees", _
                                   nbComptes & " comptes, " & txTotal & " ecritures, seed=" & seed)
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur GenerateDemoData: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

' ==============================================================================
' GENERATION DES ECRITURES PAR COMPTE (anomalies injectees selon l'index)
' Retourne la somme des ecritures (balance "as per proof")
' ==============================================================================

Private Function WriteTransactions(ByRef arr() As Variant, ByRef r As Long, n As Long, avgTx As Double, ByRef txTotal As Long) As Double
    Dim i As Long, cnt As Long, age As Long, amt As Double, total As Double
    Dim d As Date

    Select Case n
        Case 1, 2, 3
            ' SUSPENS: items anciens (> 90 j) + quelques recents
            For i = 1 To 6
                age = 95 + Int(Rnd * 305)
                If n = 1 And i = 1 Then age = 400   ' garantit Age Max > 365 j -> score CRITICAL sur le compte 1
                amt = Sgn(Rnd - 0.5) * RandAmount(200000, 3000000)
                total = total + AddTx(arr, r, age, "SUSPENS A REGULARISER REF " & Format(1000 + i * n, "0000"), amt, txTotal)
            Next i
            For i = 1 To 4
                total = total + AddTx(arr, r, 1 + Int(Rnd * 30), RandNarration(), Sgn(Rnd - 0.5) * RandAmount(20000, 500000), txTotal)
            Next i

        Case 4, 5
            ' TRANSIT: items non apures entre 10 et 30 jours
            For i = 1 To 5
                total = total + AddTx(arr, r, 10 + Int(Rnd * 21), "TRANSIT COMPENSATION LOT " & Format(i, "00"), Sgn(Rnd - 0.5) * RandAmount(100000, 2000000), txTotal)
            Next i

        Case 16, 17
            ' STRUCTURING: 4 ecritures 60k-95k le meme jour (+ activite normale)
            For i = 1 To 4
                total = total + AddTx(arr, r, 3, "DEPOT ESPECES FRACTIONNE " & i, RandAmount(60000, 95000), txTotal)
            Next i
            total = total + NormalActivity(arr, r, 6, txTotal)

        Case 18, 19
            ' LAB/FT: juste sous le seuil (FRD-003) + au-dessus du seuil de declaration (LAB-001)
            total = total + AddTx(arr, r, 5, "VIREMENT RECU CLIENT IMPORT-EXPORT", 4750000, txTotal)
            total = total + AddTx(arr, r, 8, "VIREMENT INTERNATIONAL RECU", 5200000, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)

        Case 20, 21, 22
            ' WEEKEND: samedi et dimanche derniers, montants > 50 000
            d = LastWeekday(6)
            total = total + AddTxDate(arr, r, d, "VIREMENT CLIENT " & CleanAccount(n + 50), RandAmount(120000, 900000), txTotal)
            d = LastWeekday(7)
            total = total + AddTxDate(arr, r, d, "RETRAIT GUICHET EXCEPTIONNEL", -RandAmount(60000, 400000), txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)

        Case 23
            total = total + AddTx(arr, r, 4, "PAIEMENT CADEAU DIRECTION GENERALE", -250000, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)
        Case 24
            total = total + AddTx(arr, r, 6, "VIREMENT URGENT MANUEL HORS PROCEDURE", -180000, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)
        Case 25
            total = total + AddTx(arr, r, 9, "CORRECTION SAISIE OPERATEUR", 95000, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)

        Case 26, 27
            ' DUPLICATE: ecriture strictement identique deux fois
            total = total + AddTx(arr, r, 4, "REGLEMENT FACTURE 2024-0117", -437500, txTotal)
            total = total + AddTx(arr, r, 4, "REGLEMENT FACTURE 2024-0117", -437500, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal)

        Case 28, 29, 30
            ' ROUNDTRIP: 2 500 000 le meme jour sur 3 comptes (6 ecritures identiques en montant)
            total = total + AddTx(arr, r, 6, "VIREMENT RECU DE " & CleanAccount(IIf(n = 28, 30, n - 1)), 2500000, txTotal)
            total = total + AddTx(arr, r, 6, "VIREMENT EMIS VERS " & CleanAccount(IIf(n = 30, 28, n + 1)), -2500000, txTotal)
            total = total + NormalActivity(arr, r, 5, txTotal)

        Case 31
            ' DORMANT: 2 ecritures recentes, cumul > 1 000 000
            total = total + AddTx(arr, r, 3, "VIREMENT RECU REACTIVATION", 900000, txTotal)
            total = total + AddTx(arr, r, 5, "VIREMENT RECU COMPLEMENT", 800000, txTotal)

        Case 32
            ' ZSCORE: 9 ecritures ~100k + 1 outlier 3 000 000
            For i = 1 To 9
                total = total + AddTx(arr, r, 2 + i * 3, RandNarration(), RandAmount(80000, 120000), txTotal)
            Next i
            total = total + AddTx(arr, r, 2, "VIREMENT RECU EXCEPTIONNEL", 3000000, txTotal)

        Case 43
            ' PREPAID: amortissements mensuels arretes il y a plus de 30 jours
            For i = 1 To 3
                total = total + AddTx(arr, r, 45 + i * 25, "AMORTISSEMENT MENSUEL LOYER", -1000000, txTotal)
            Next i

        Case 44
            ' PROXY: transaction echouee non re-imputee
            total = total + AddTx(arr, r, 3, "PROXY ECHEC SYSTEME LOT 17", 1250000, txTotal)

        Case 45
            ' PRODUIT: un debit (extourne) parmi des credits
            total = total + AddTx(arr, r, 4, "EXTOURNE COMMISSION CLIENT", -150000, txTotal)
            total = total + NormalActivity(arr, r, 6, txTotal, 1)

        Case 46
            ' ECART ATM: items anciens non reconcilies
            For i = 1 To 4
                total = total + AddTx(arr, r, 100 + i * 75, "ECART ATM GAB 0" & i & " NON RECONCILIE", RandAmount(50000, 400000), txTotal)
            Next i

        Case 47, 48
            total = total + NormalActivity(arr, r, 5, txTotal)

        Case Else
            cnt = 3 + Int(avgTx * (0.5 + Rnd))
            ' Comptes de produits: credits seulement; comptes de charges: debits seulement
            If PrefixFor(n) = "701" Then
                total = total + NormalActivity(arr, r, cnt, txTotal, 1)
            ElseIf PrefixFor(n) = "601" Then
                total = total + NormalActivity(arr, r, cnt, txTotal, -1)
            Else
                total = total + NormalActivity(arr, r, cnt, txTotal)
            End If
    End Select

    WriteTransactions = total
End Function

Private Function NormalActivity(ByRef arr() As Variant, ByRef r As Long, cnt As Long, ByRef txTotal As Long, Optional signMode As Long = 0) As Double
    ' signMode: 0 = aleatoire, 1 = credits uniquement (positif), -1 = debits uniquement (negatif)
    Dim i As Long, total As Double, amt As Double
    For i = 1 To cnt
        amt = RandAmount(10000, 4000000)
        ' Eviter les bandes "juste sous seuil" pour que FRD-003 ne se declenche que sur les comptes injectes
        If amt >= 90000 And amt < 100000 Then amt = 85000
        If amt >= 450000 And amt < 500000 Then amt = 420000
        If amt >= 900000 And amt < 1000000 Then amt = 850000
        If amt >= 4500000 Then amt = 4200000
        If signMode = 0 Then
            If Rnd < 0.5 Then amt = -amt
        ElseIf signMode < 0 Then
            amt = -amt
        End If
        total = total + AddTx(arr, r, 1 + Int(Rnd * 60), RandNarration(), amt, txTotal)
    Next i
    NormalActivity = total
End Function

Private Function AddTx(ByRef arr() As Variant, ByRef r As Long, age As Long, narration As String, amount As Double, ByRef txTotal As Long) As Double
    Dim d As Date
    If age < 1 Then age = 1
    d = Date - age
    ' Eviter un week-end involontaire (les week-ends sont injectes explicitement)
    Do While Weekday(d, vbMonday) >= 6
        d = d - 1
        age = age + 1
    Loop
    r = r + 1
    arr(r, 1) = d
    arr(r, 2) = age
    arr(r, 3) = narration
    arr(r, 4) = Round(amount, 0)
    txTotal = txTotal + 1
    AddTx = Round(amount, 0)
End Function

Private Function AddTxDate(ByRef arr() As Variant, ByRef r As Long, d As Date, narration As String, amount As Double, ByRef txTotal As Long) As Double
    Dim age As Long
    age = CLng(Date - d)
    If age < 1 Then age = 1
    r = r + 1
    arr(r, 1) = d
    arr(r, 2) = age
    arr(r, 3) = narration
    arr(r, 4) = Round(amount, 0)
    txTotal = txTotal + 1
    AddTxDate = Round(amount, 0)
End Function

' ==============================================================================
' HELPERS
' ==============================================================================

Private Sub AddRow(ByRef arr() As Variant, ByRef r As Long, a As Variant, b As Variant, c As Variant, d As Variant)
    r = r + 1
    arr(r, 1) = a
    arr(r, 2) = b
    arr(r, 3) = c
    arr(r, 4) = d
End Sub

Private Sub EnsureParamSheet()
    Dim ws As Worksheet
    Set ws = SAFA_Common.GetOrCreateSheet("PARAM", False)
    ws.Range("A1").Value = "Parametre"
    ws.Range("B1").Value = "Valeur"
    ws.Range("A2").Value = "Tolerance"
    ws.Range("B2").Value = 100
    ws.Range("E1").Value = "SOL ID"
    ws.Range("F1").Value = "SOL ID"
    ws.Range("F2").Value = "'" & SOL_ID
    ' Exclusions Non-Proofable: comptes techniques uniquement (jamais SUSPENS/TRANSIT)
    ws.Range("D1").Value = "Patterns Exclusion"
    ws.Range("D2:D50").ClearContents
    ws.Range("D2").Value = "ISO"
    ws.Range("D3").Value = "NOSTRO"
    ws.Range("D4").Value = "VOSTRO"
End Sub

Private Function IsGlOnly(n As Long) As Boolean
    IsGlOnly = (n >= 38 And n <= 42)
End Function

Private Function IsBalanceOnly(n As Long) As Boolean
    IsBalanceOnly = (n >= 33 And n <= 37)
End Function

Private Function PrefixFor(n As Long) As String
    Select Case n
        Case 1 To 3: PrefixFor = "471"
        Case 4 To 5: PrefixFor = "472"
        Case 6 To 15: PrefixFor = Choose(((n - 6) Mod 5) + 1, "411", "401", "521", "601", "512")
        Case 16, 17: PrefixFor = "531"
        Case 18, 19: PrefixFor = "411"
        Case 20 To 22: PrefixFor = "521"
        Case 23 To 25: PrefixFor = "601"
        Case 26, 27: PrefixFor = "401"
        Case 28 To 30: PrefixFor = "512"
        Case 31, 32: PrefixFor = "521"
        Case 33 To 37: PrefixFor = "411"
        Case 38 To 42: PrefixFor = "401"
        Case 43: PrefixFor = "476"   ' charges constatees d'avance
        Case 44: PrefixFor = "472"   ' proxy
        Case 45: PrefixFor = "701"   ' produit avec debit
        Case 46: PrefixFor = "571"   ' ecart caisse / ATM
        Case 47: PrefixFor = "571"   ' caisse au-dela de la limite
        Case 48: PrefixFor = "521"   ' banque a solde crediteur (sens anormal)
        Case Else: PrefixFor = Choose(((n - 49) Mod 11) + 1, "101", "211", "311", "401", "411", "521", "571", "601", "701", "512", "531")
    End Select
End Function

Private Function RawAccount(n As Long) As String
    RawAccount = PrefixFor(n) & "0000" & Format(n, "00000")
End Function

Private Function CleanAccount(n As Long) As String
    CleanAccount = PrefixFor(n) & SOL_ID & "0" & Format(n, "00000")
End Function

Private Function NameFor(n As Long) As String
    Dim base As String
    Select Case n
        Case 43: NameFor = "CHARGES CONSTATEES D AVANCE LOYER SIEGE": Exit Function
        Case 44: NameFor = "PROXY GAB TRANSACTIONS ECHOUEES": Exit Function
        Case 45: NameFor = "PRODUITS COMMISSIONS TRANSFERTS": Exit Function
        Case 46: NameFor = "ECART CAISSE ATM AGENCE": Exit Function
        Case 47: NameFor = "CAISSE PRINCIPALE AGENCE": Exit Function
        Case 48: NameFor = "BANQUE BEAC COMPTE COURANT": Exit Function
    End Select
    Select Case PrefixFor(n)
        Case "101": base = "CAPITAL SOCIAL"
        Case "211": base = "IMMOBILISATIONS INCORPORELLES"
        Case "311": base = "STOCKS FOURNITURES"
        Case "401": base = "FOURNISSEURS LOCAUX"
        Case "411": base = "CLIENTS ENTREPRISES"
        Case "471": base = "SUSPENS VIREMENTS RECUS"
        Case "472": base = "TRANSIT COMPENSATION"
        Case "512": base = "BANQUE CORRESPONDANT"
        Case "521": base = "BANQUE BEAC"
        Case "531": base = "CAISSE AGENCE"
        Case "571": base = "CAISSE PRINCIPALE"
        Case "601": base = "CHARGES EXTERNES"
        Case "701": base = "PRODUITS BANCAIRES"
        Case Else: base = "COMPTE DIVERS"
    End Select
    NameFor = base & " " & Format(n, "000")
End Function

Private Function RandAmount(minV As Double, maxV As Double) As Double
    ' Log-uniforme entre minV et maxV (distribution realiste, conforme Benford)
    RandAmount = Round(Exp(Log(minV) + Rnd * (Log(maxV) - Log(minV))), 0)
End Function

Private Function RandNarration() As String
    Dim list As Variant
    list = Array("VIREMENT SALAIRE", "REGLEMENT FACTURE", "FRAIS TENUE DE COMPTE", "DEPOT ESPECES", _
                 "RETRAIT GUICHET", "COMMISSION TRANSFERT", "REMISE CHEQUE", "PRELEVEMENT ASSURANCE", _
                 "LOYER AGENCE", "ACHAT FOURNITURES", "VIREMENT CLIENT", "REMBOURSEMENT PRET")
    RandNarration = list(Int(Rnd * (UBound(list) + 1))) & " REF " & Format(Int(Rnd * 900000) + 100000, "000000")
End Function

Private Function LastWeekday(vbDay As Long) As Date
    ' vbDay: 6 = samedi, 7 = dimanche (base vbMonday)
    Dim d As Date
    d = Date - 1
    Do While Weekday(d, vbMonday) <> vbDay
        d = d - 1
    Loop
    LastWeekday = d
End Function
