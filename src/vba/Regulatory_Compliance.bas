Attribute VB_Name = "Regulatory_Compliance"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: REGULATORY_COMPLIANCE v10.0
' ==============================================================================
' Description: Vérifications de conformité réglementaire
' Réglementations couvertes:
'   - COBAC (Commission Bancaire d'Afrique Centrale)
'   - BEAC (Banque des États de l'Afrique Centrale)
'   - OHADA/SYSCOHADA (Plan comptable)
'   - Règles prudentielles bancaires
'   - Lutte Anti-Blanchiment (LAB/FT)
' ==============================================================================

' --- CONSTANTES RÉGLEMENTAIRES ---
Private Const COBAC_SUSPENS_LIMIT_DAYS As Long = 90
Private Const COBAC_TRANSIT_LIMIT_DAYS As Long = 7
Private Const BEAC_RESERVE_RATIO As Double = 0.07  ' 7% réserves obligatoires
Private Const LAB_THRESHOLD_XAF As Double = 5000000  ' Seuil déclaration LAB
Private Const LARGE_EXPOSURE_RATIO As Double = 0.25  ' 25% des fonds propres

' --- TYPES ---
Private Type ComplianceCheck
    RuleID As String
    RuleName As String
    Regulation As String
    Status As String
    Details As String
    Impact As String
    Recommendation As String
End Type

' ==============================================================================
' 1. POINT D'ENTRÉE PRINCIPAL
' ==============================================================================

Public Sub Lancer_Verification_Conformite()
    Dim wsComp As Worksheet
    Dim checkRow As Long

    On Error GoTo CompError

    ' Créer feuille de conformité
    Set wsComp = Core_Engine.GetOrCreateSheet("COMPLIANCE_CHECK", True)

    ' En-têtes
    wsComp.Range("A1:H1").Value = Array("Règle ID", "Règle", "Réglementation", _
                                         "Statut", "Détails", "Impact", "Recommandation", "Priorité")
    FormatComplianceHeader wsComp.Range("A1:H1")

    checkRow = 2

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 1: RÈGLES COBAC
    ' ═══════════════════════════════════════════════════════════════
    Call Check_COBAC_Suspens(wsComp, checkRow)
    Call Check_COBAC_Transit(wsComp, checkRow)
    Call Check_COBAC_Provisions(wsComp, checkRow)

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 2: RÈGLES OHADA/SYSCOHADA
    ' ═══════════════════════════════════════════════════════════════
    Call Check_OHADA_Balance(wsComp, checkRow)
    Call Check_OHADA_ClasseComptes(wsComp, checkRow)

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 3: RÈGLES LAB/FT (Lutte Anti-Blanchiment)
    ' ═══════════════════════════════════════════════════════════════
    Call Check_LAB_Seuils(wsComp, checkRow)
    Call Check_LAB_Structuration(wsComp, checkRow)

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 4: RÈGLES PRUDENTIELLES
    ' ═══════════════════════════════════════════════════════════════
    Call Check_Prudentiel_Concentrations(wsComp, checkRow)

    ' Finalisation
    wsComp.Columns("A:H").AutoFit

    ' Résumé
    Call CreateComplianceSummary(wsComp, checkRow)

    Call Core_Engine.WriteToAuditLog("PROCESS", "Lancer_Verification_Conformite terminé", _
                                      "Vérifications: " & (checkRow - 2))
    Exit Sub

CompError:
    Call Core_Engine.LogError("Regulatory_Compliance", "Lancer_Verification_Conformite", _
                               Err.Number, Err.Description)
End Sub

' ==============================================================================
' 2. RÈGLES COBAC
' ==============================================================================

Private Sub Check_COBAC_Suspens(ws As Worksheet, ByRef r As Long)
    ' COBAC R-2001/01: Les comptes de suspens doivent être apurés sous 90 jours
    Dim wsR As Worksheet
    Dim i As Long, lr As Long, violations As Long
    Dim totalAmount As Double

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    violations = 0
    totalAmount = 0

    For i = 2 To lr
        Dim age As Double, status As String, name As String

        age = Core_Engine.SafeVal(wsR.Cells(i, 8).Value)
        status = Core_Engine.SafeText(wsR.Cells(i, 6).Value)
        name = UCase(Core_Engine.SafeText(wsR.Cells(i, 2).Value))

        ' Vérifier comptes de suspens > 90 jours
        If (InStr(name, "SUSPENS") > 0 Or InStr(name, "SUSPENSE") > 0) And _
           age > COBAC_SUSPENS_LIMIT_DAYS And status = "Ecart à analyser" Then
            violations = violations + 1
            totalAmount = totalAmount + Abs(Core_Engine.SafeVal(wsR.Cells(i, 5).Value))
        End If
    Next i

    ' Enregistrer résultat
    ws.Cells(r, 1).Value = "COBAC-001"
    ws.Cells(r, 2).Value = "Apurement Suspens 90j"
    ws.Cells(r, 3).Value = "COBAC R-2001/01"

    If violations = 0 Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Aucun suspens > 90 jours"
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "NON-CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = violations & " compte(s) en violation, Total: " & Format(totalAmount, "#,##0") & " XAF"
        ws.Cells(r, 6).Value = "Sanction COBAC possible"
        ws.Cells(r, 7).Value = "Apurer immédiatement ou provisionner"
        ws.Cells(r, 8).Value = "CRITICAL"
        ws.Cells(r, 8).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 8).Font.Color = vbWhite
    End If

    r = r + 1
End Sub

Private Sub Check_COBAC_Transit(ws As Worksheet, ByRef r As Long)
    ' COBAC: Les comptes de transit doivent être apurés quotidiennement (tolérance J+7)
    Dim wsR As Worksheet
    Dim i As Long, lr As Long, violations As Long
    Dim totalAmount As Double

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    violations = 0
    totalAmount = 0

    For i = 2 To lr
        Dim age As Double, balance As Double, name As String

        age = Core_Engine.SafeVal(wsR.Cells(i, 8).Value)
        balance = Core_Engine.SafeVal(wsR.Cells(i, 4).Value)
        name = UCase(Core_Engine.SafeText(wsR.Cells(i, 2).Value))

        If (InStr(name, "TRANSIT") > 0 Or InStr(name, "PASSAGE") > 0 Or _
            InStr(name, "CLEARING") > 0) And age > COBAC_TRANSIT_LIMIT_DAYS And _
            Abs(balance) > 1000 Then
            violations = violations + 1
            totalAmount = totalAmount + Abs(balance)
        End If
    Next i

    ws.Cells(r, 1).Value = "COBAC-002"
    ws.Cells(r, 2).Value = "Apurement Transit J+7"
    ws.Cells(r, 3).Value = "COBAC - Règles Comptables"

    If violations = 0 Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Comptes de transit correctement apurés"
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "NON-CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(255, 165, 0)
        ws.Cells(r, 5).Value = violations & " compte(s), Montant: " & Format(totalAmount, "#,##0") & " XAF"
        ws.Cells(r, 6).Value = "Risque opérationnel"
        ws.Cells(r, 7).Value = "Investiguer et régulariser"
        ws.Cells(r, 8).Value = "HIGH"
        ws.Cells(r, 8).Interior.Color = RGB(255, 165, 0)
    End If

    r = r + 1
End Sub

Private Sub Check_COBAC_Provisions(ws As Worksheet, ByRef r As Long)
    ' Vérification adéquation des provisions (simplifiée)
    Dim wsR As Worksheet
    Dim totalEcart As Double, totalProvision As Double
    Dim ratio As Double

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub

    totalEcart = Application.WorksheetFunction.SumIf(wsR.Range("F:F"), "Ecart à analyser", wsR.Range("E:E"))
    totalProvision = Application.WorksheetFunction.Sum(wsR.Range("O:O"))
    On Error GoTo 0

    If Abs(totalEcart) > 0 Then
        ratio = totalProvision / Abs(totalEcart)
    Else
        ratio = 1
    End If

    ws.Cells(r, 1).Value = "COBAC-003"
    ws.Cells(r, 2).Value = "Couverture Provisions"
    ws.Cells(r, 3).Value = "COBAC/IFRS9"

    If ratio >= 0.5 Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Taux couverture: " & Format(ratio, "0%")
        ws.Cells(r, 8).Value = "LOW"
    ElseIf ratio >= 0.25 Then
        ws.Cells(r, 4).Value = "ATTENTION"
        ws.Cells(r, 4).Interior.Color = RGB(255, 255, 0)
        ws.Cells(r, 5).Value = "Taux couverture: " & Format(ratio, "0%") & " - Insuffisant"
        ws.Cells(r, 6).Value = "Sous-provisionnement"
        ws.Cells(r, 7).Value = "Revoir calcul provisions IFRS9"
        ws.Cells(r, 8).Value = "MEDIUM"
    Else
        ws.Cells(r, 4).Value = "NON-CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Taux couverture: " & Format(ratio, "0%") & " - Critique"
        ws.Cells(r, 6).Value = "Impact P&L significatif"
        ws.Cells(r, 7).Value = "Constituer provisions complémentaires"
        ws.Cells(r, 8).Value = "CRITICAL"
        ws.Cells(r, 8).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 8).Font.Color = vbWhite
    End If

    r = r + 1
End Sub

' ==============================================================================
' 3. RÈGLES OHADA/SYSCOHADA
' ==============================================================================

Private Sub Check_OHADA_Balance(ws As Worksheet, ByRef r As Long)
    ' Vérification équilibre comptable (Actif = Passif)
    Dim wsR As Worksheet
    Dim totalActif As Double, totalPassif As Double
    Dim i As Long, lr As Long

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row

    For i = 2 To lr
        Dim acct As String, balance As Double
        acct = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        balance = Core_Engine.SafeVal(wsR.Cells(i, 4).Value)

        ' Classes 1-5 = Bilan
        Select Case Left(acct, 1)
            Case "1", "2", "3" ' Actif (simplifié)
                If balance > 0 Then totalActif = totalActif + balance
            Case "4", "5" ' Passif (simplifié)
                If balance < 0 Then totalPassif = totalPassif + Abs(balance)
        End Select
    Next i

    ws.Cells(r, 1).Value = "OHADA-001"
    ws.Cells(r, 2).Value = "Équilibre Bilan"
    ws.Cells(r, 3).Value = "SYSCOHADA Art. 7"

    Dim ecartBilan As Double
    ecartBilan = totalActif - totalPassif

    If Abs(ecartBilan) < 1000 Then ' Tolérance d'arrondi
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Bilan équilibré (Écart: " & Format(ecartBilan, "#,##0") & ")"
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "ANOMALIE"
        ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Déséquilibre: " & Format(ecartBilan, "#,##0") & " XAF"
        ws.Cells(r, 6).Value = "États financiers incorrects"
        ws.Cells(r, 7).Value = "Identifier et corriger l'écart"
        ws.Cells(r, 8).Value = "CRITICAL"
        ws.Cells(r, 8).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 8).Font.Color = vbWhite
    End If

    r = r + 1
End Sub

Private Sub Check_OHADA_ClasseComptes(ws As Worksheet, ByRef r As Long)
    ' Vérification cohérence des classes comptables
    Dim wsR As Worksheet
    Dim anomalies As Long
    Dim i As Long, lr As Long

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    anomalies = 0

    For i = 2 To lr
        Dim acct As String, balance As Double, name As String
        acct = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        balance = Core_Engine.SafeVal(wsR.Cells(i, 4).Value)
        name = UCase(Core_Engine.SafeText(wsR.Cells(i, 2).Value))

        ' Vérification sens logique selon classe
        Select Case Left(acct, 1)
            Case "2" ' Immobilisations - normalement débitrices
                If balance < -100000 And InStr(name, "AMORT") = 0 And InStr(name, "PROV") = 0 Then
                    anomalies = anomalies + 1
                End If
            Case "4" ' Tiers
                ' Les clients (41) sont normalement débiteurs
                If Left(acct, 2) = "41" And balance < -100000 Then
                    anomalies = anomalies + 1
                End If
                ' Les fournisseurs (40) sont normalement créditeurs
                If Left(acct, 2) = "40" And balance > 100000 Then
                    anomalies = anomalies + 1
                End If
            Case "5" ' Trésorerie
                ' Banque/Caisse généralement débiteurs
                If (InStr(name, "BANQUE") > 0 Or InStr(name, "CAISSE") > 0) And _
                   balance < -1000000 Then
                    anomalies = anomalies + 1
                End If
        End Select
    Next i

    ws.Cells(r, 1).Value = "OHADA-002"
    ws.Cells(r, 2).Value = "Cohérence Classes Comptes"
    ws.Cells(r, 3).Value = "SYSCOHADA"

    If anomalies = 0 Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Sens des comptes cohérents"
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "ATTENTION"
        ws.Cells(r, 4).Interior.Color = RGB(255, 255, 0)
        ws.Cells(r, 5).Value = anomalies & " compte(s) avec sens anormal"
        ws.Cells(r, 6).Value = "Possible erreur d'imputation"
        ws.Cells(r, 7).Value = "Vérifier les comptes concernés"
        ws.Cells(r, 8).Value = "MEDIUM"
    End If

    r = r + 1
End Sub

' ==============================================================================
' 4. RÈGLES LAB/FT (Lutte Anti-Blanchiment)
' ==============================================================================

Private Sub Check_LAB_Seuils(ws As Worksheet, ByRef r As Long)
    ' Vérification des transactions au-dessus du seuil de déclaration
    Dim wsT As Worksheet
    Dim i As Long, lr As Long, largeTransactions As Long
    Dim totalAmount As Double

    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets("TRANSACTION_DATA")
    If wsT Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsT.Cells(wsT.Rows.count, 1).End(xlUp).Row
    largeTransactions = 0
    totalAmount = 0

    For i = 2 To lr
        Dim mnt As Double
        mnt = Abs(Core_Engine.SafeVal(wsT.Cells(i, 4).Value))

        If mnt >= LAB_THRESHOLD_XAF Then
            largeTransactions = largeTransactions + 1
            totalAmount = totalAmount + mnt
        End If
    Next i

    ws.Cells(r, 1).Value = "LAB-001"
    ws.Cells(r, 2).Value = "Transactions > Seuil LAB"
    ws.Cells(r, 3).Value = "Règlement COBAC LAB/FT"

    ws.Cells(r, 4).Value = "INFORMATION"
    ws.Cells(r, 4).Interior.Color = RGB(173, 216, 230)
    ws.Cells(r, 5).Value = largeTransactions & " transaction(s) >= " & Format(LAB_THRESHOLD_XAF, "#,##0") & " XAF"
    ws.Cells(r, 6).Value = "Montant total: " & Format(totalAmount, "#,##0") & " XAF"
    ws.Cells(r, 7).Value = "Vérifier déclarations de soupçon si nécessaire"

    If largeTransactions > 50 Then
        ws.Cells(r, 8).Value = "MEDIUM"
    Else
        ws.Cells(r, 8).Value = "LOW"
    End If

    r = r + 1
End Sub

Private Sub Check_LAB_Structuration(ws As Worksheet, ByRef r As Long)
    ' Vérification des patterns de structuration (fractionnement)
    Dim wsA As Worksheet
    Dim structuringAlerts As Long

    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")
    If wsA Is Nothing Then Exit Sub

    structuringAlerts = Application.CountIf(wsA.Range("B:B"), "STRUCTURING")
    On Error GoTo 0

    ws.Cells(r, 1).Value = "LAB-002"
    ws.Cells(r, 2).Value = "Détection Structuration"
    ws.Cells(r, 3).Value = "Règlement COBAC LAB/FT"

    If structuringAlerts = 0 Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Aucun pattern de structuration détecté"
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "ALERTE"
        ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = structuringAlerts & " cas de structuration potentielle"
        ws.Cells(r, 6).Value = "Possible contournement seuils LAB"
        ws.Cells(r, 7).Value = "Analyser et signaler si confirmé"
        ws.Cells(r, 8).Value = "CRITICAL"
        ws.Cells(r, 8).Interior.Color = RGB(255, 0, 0)
        ws.Cells(r, 8).Font.Color = vbWhite
    End If

    r = r + 1
End Sub

' ==============================================================================
' 5. RÈGLES PRUDENTIELLES
' ==============================================================================

Private Sub Check_Prudentiel_Concentrations(ws As Worksheet, ByRef r As Long)
    ' Vérification des concentrations de risques
    Dim wsR As Worksheet
    Dim totalEcart As Double, maxEcart As Double
    Dim maxAccount As String
    Dim i As Long, lr As Long

    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    If wsR Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    maxEcart = 0

    For i = 2 To lr
        Dim ecart As Double
        ecart = Abs(Core_Engine.SafeVal(wsR.Cells(i, 5).Value))
        totalEcart = totalEcart + ecart

        If ecart > maxEcart Then
            maxEcart = ecart
            maxAccount = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        End If
    Next i

    ws.Cells(r, 1).Value = "PRUD-001"
    ws.Cells(r, 2).Value = "Concentration Écarts"
    ws.Cells(r, 3).Value = "Règles Prudentielles"

    Dim concentration As Double
    If totalEcart > 0 Then
        concentration = maxEcart / totalEcart
    Else
        concentration = 0
    End If

    If concentration < LARGE_EXPOSURE_RATIO Then
        ws.Cells(r, 4).Value = "CONFORME"
        ws.Cells(r, 4).Interior.Color = RGB(0, 176, 80)
        ws.Cells(r, 4).Font.Color = vbWhite
        ws.Cells(r, 5).Value = "Concentration max: " & Format(concentration, "0%")
        ws.Cells(r, 8).Value = "LOW"
    Else
        ws.Cells(r, 4).Value = "ATTENTION"
        ws.Cells(r, 4).Interior.Color = RGB(255, 165, 0)
        ws.Cells(r, 5).Value = "Compte " & maxAccount & " = " & Format(concentration, "0%") & " du total"
        ws.Cells(r, 6).Value = "Risque de concentration"
        ws.Cells(r, 7).Value = "Analyser en priorité ce compte"
        ws.Cells(r, 8).Value = "HIGH"
        ws.Cells(r, 8).Interior.Color = RGB(255, 165, 0)
    End If

    r = r + 1
End Sub

' ==============================================================================
' 6. RÉSUMÉ CONFORMITÉ
' ==============================================================================

Private Sub CreateComplianceSummary(ws As Worksheet, lastRow As Long)
    Dim conforme As Long, nonConforme As Long, attention As Long
    Dim i As Long

    ' Compter les statuts
    For i = 2 To lastRow - 1
        Select Case ws.Cells(i, 4).Value
            Case "CONFORME"
                conforme = conforme + 1
            Case "NON-CONFORME", "ALERTE", "ANOMALIE"
                nonConforme = nonConforme + 1
            Case "ATTENTION"
                attention = attention + 1
        End Select
    Next i

    ' Ajouter résumé en haut
    ws.Rows("1:1").Insert Shift:=xlDown
    ws.Rows("1:1").Insert Shift:=xlDown
    ws.Rows("1:1").Insert Shift:=xlDown

    ws.Range("A1").Value = "SYNTHÈSE CONFORMITÉ RÉGLEMENTAIRE"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "Date: " & Format(Now, "dd/mm/yyyy hh:nn")

    ws.Range("A3:D3").Value = Array("CONFORME", "NON-CONFORME", "ATTENTION", "TOTAL")
    ws.Range("A4:D4").Value = Array(conforme, nonConforme, attention, conforme + nonConforme + attention)

    ws.Range("A3").Interior.Color = RGB(0, 176, 80)
    ws.Range("B3").Interior.Color = RGB(255, 0, 0)
    ws.Range("C3").Interior.Color = RGB(255, 255, 0)

    ws.Range("A3:D3").Font.Bold = True
    ws.Range("A4:D4").Font.Bold = True

    ' Score global de conformité
    Dim score As Double
    If conforme + nonConforme + attention > 0 Then
        score = conforme / (conforme + nonConforme + attention)
    Else
        score = 1
    End If

    ws.Range("F3").Value = "Score Conformité:"
    ws.Range("G3").Value = Format(score, "0%")
    ws.Range("G3").Font.Bold = True
    ws.Range("G3").Font.Size = 16

    If score >= 0.8 Then
        ws.Range("G3").Font.Color = RGB(0, 150, 0)
    ElseIf score >= 0.5 Then
        ws.Range("G3").Font.Color = RGB(255, 165, 0)
    Else
        ws.Range("G3").Font.Color = RGB(255, 0, 0)
    End If
End Sub

' ==============================================================================
' 7. UTILITAIRES
' ==============================================================================

Private Sub FormatComplianceHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
End Sub
