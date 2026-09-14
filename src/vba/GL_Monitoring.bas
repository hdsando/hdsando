Attribute VB_Name = "GL_Monitoring"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: GL_MONITORING v10.1
' ==============================================================================
' Description: Controles "GL Monitoring" issus de la Knowledge Sharing Session
'   DAI du 09/09/2026 (integrite du grand livre et revue des justificatifs) et du
'   support Groupe "GL Integrity & Proof Review". Le module travaille sur les
'   feuilles produites par le pipeline (BALANCE_DATA, GLPROOF_DATA,
'   TRANSACTION_DATA) et sur deux feuilles optionnelles:
'     PROOFABLE_LIST   : liste des comptes proofables tenue par le Controle Interne
'     BALANCE_PREV_RAW : balance de la periode precedente (variations)
'
' REGLES (ref. section 5 du compte rendu):
'   GLM-001 Univers proofable: classification de chaque compte (proofable / P&L /
'           engagement / systeme) et exhaustivite par rapport a la liste CI
'   GLM-002 Proof manquant: compte proofable sans bloc GL Proof
'   GLM-003 Sens de solde anormal: actif crediteur, passif debiteur, produit
'           debiteur, charge creditrice
'   GLM-004 Transit / proxy / suspens non nuls (attendu: zero sous 24 h)
'   GLM-005 Debit sur compte de produit (admis le jour meme, sinon approbation)
'   GLM-006 Charges constatees d'avance: amortissement mensuel non constate
'   GLM-007 Variation des charges / produits recurrents (20 % / 50 %)
'   GLM-008 Contreparties des comptes d'attente (schema caisse -> attente -> tiers)
'   GLM-009 Items over-aged (ecarts caisse / ATM, 90 / 180 / 360 jours)
'   GLM-010 Limites de caisse et de coffre
'   GLM-011 Imputations inter-agences (INTERSOL) a documenter
'   GLM-012 Saisie manuelle sur compte automatise / systeme
'   --- Support Groupe "GL Integrity & Proof Review" (diapos 16-25) ---
'   GLM-013 Cheques de direction perimes a transferer en non reclames
'   GLM-014 Soldes inhabituels / exceptionnels (vs periode precedente)
'   GLM-015 Mouvements de depenses significatifs (regularite, autorisation)
'   GLM-016 Remboursements de pertes liees a la fraude et narrations sensibles
'   GLM-017 Travaux en cours (WIP) anciens: comptabilisation / capitalisation
'   GLM-018 Charges stockees dans les comptes d'attente / transit (interdit)
'   Feuilles: GL_MONITORING (constats), PROOFABLE_UNIVERSE (univers), GL_RATING (note)
' ==============================================================================

Public Type AccountClass
    Family As String       ' ACTIF / PASSIF / CAPITAUX / CHARGE / PRODUIT / ENGAGEMENT / TIERS / INCONNU
    Nature As String       ' STANDARD / TRANSIT / PROXY / SUSPENS / PREPAID / CASH / DIFFERENCE / PL / SYSTEM / INTERSOL
    Proofable As Boolean
    Reason As String
End Type

Private mCfg As Config_Manager.GLMonitoringConfig
Private mCfgLoaded As Boolean
Private mWsOut As Worksheet
Private mRow As Long
Private mWsAudit As Worksheet
Private mAuditRow As Long
Private mPoints As Object
Private mCounts As Object

' ==============================================================================
' POINTS D'ENTREE
' ==============================================================================

Public Sub Lancer_GL_Monitoring_Silencieux()
    Call Lancer_GL_Monitoring(False)
End Sub

Public Sub Lancer_GL_Monitoring(Optional showMsg As Boolean = True)
    On Error GoTo ErrHandler

    Dim wsBal As Worksheet, wsGL As Worksheet, wsTx As Worksheet
    Dim dictBal As Object, dictGL As Object, dictCI As Object, dictPrev As Object
    Dim lr As Long, i As Long, k As String
    Dim t0 As Double
    t0 = Timer

    Call LoadCfg
    Set mPoints = CreateObject("Scripting.Dictionary")
    Set mCounts = CreateObject("Scripting.Dictionary")

    If Not SAFA_Common.FeuilleExiste("BALANCE_DATA") Then
        If showMsg Then MsgBox "BALANCE_DATA absente: lancez d'abord l'analyse complete.", vbExclamation, "GL Monitoring"
        Exit Sub
    End If
    Set wsBal = ThisWorkbook.Sheets("BALANCE_DATA")
    If SAFA_Common.FeuilleExiste("GLPROOF_DATA") Then Set wsGL = ThisWorkbook.Sheets("GLPROOF_DATA")
    If SAFA_Common.FeuilleExiste("TRANSACTION_DATA") Then Set wsTx = ThisWorkbook.Sheets("TRANSACTION_DATA")

    ' Feuille de constats
    Set mWsOut = SAFA_Common.GetOrCreateSheet("GL_MONITORING", True)
    mWsOut.Range("A1:K1").Value = Array("Regle", "Compte", "Libelle", "Famille", "Nature", "Solde", _
                                        "Age max (j)", "Constat", "Gravite", "Action attendue", "Points")
    SAFA_Common.FormatHeader mWsOut.Range("A1:K1")
    mRow = 2

    ' AUDIT_REPORT (alertes HIGH / CRITICAL)
    Set mWsAudit = Nothing
    If SAFA_Common.FeuilleExiste("AUDIT_REPORT") Then
        Set mWsAudit = ThisWorkbook.Sheets("AUDIT_REPORT")
        mAuditRow = mWsAudit.Cells(mWsAudit.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ' --- Balance: cle -> Array(libelle, solde) ---
    Set dictBal = CreateObject("Scripting.Dictionary")
    lr = wsBal.Cells(wsBal.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsBal.Cells(i, 1).Value))
        If k <> "" And Not dictBal.Exists(k) Then
            dictBal.Add k, Array(SAFA_Common.SafeText(wsBal.Cells(i, 2).Value), SAFA_Common.SafeVal(wsBal.Cells(i, 3).Value))
        End If
    Next i

    ' --- GL Proof: cle -> Array(libelle, solde GL, nb trans, age max, derniere date) ---
    Set dictGL = CreateObject("Scripting.Dictionary")
    If Not wsGL Is Nothing Then
        lr = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lr
            k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsGL.Cells(i, 1).Value))
            If k <> "" And Not dictGL.Exists(k) Then
                dictGL.Add k, Array(SAFA_Common.SafeText(wsGL.Cells(i, 2).Value), SAFA_Common.SafeVal(wsGL.Cells(i, 4).Value), _
                                    SAFA_Common.SafeVal(wsGL.Cells(i, 6).Value), SAFA_Common.SafeVal(wsGL.Cells(i, 7).Value), _
                                    SAFA_Common.SafeDate(wsGL.Cells(i, 10).Value))
            End If
        Next i
    End If

    Set dictCI = LoadKeyList("PROOFABLE_LIST")
    Set dictPrev = LoadPreviousBalance()

    ' Conventions debit/credit detectees sur les donnees (ou imposees par la configuration)
    Call Auto_Calibration.CalibrerSignes(dictBal, True)

    ' --- Regles ---
    Call ConstruireUnivers(dictBal, dictGL, dictCI)
    Call Regle_Sens(dictBal)
    Call Regle_TransitSuspens(dictBal, dictGL)
    Call Regle_DebitProduits(dictBal, wsTx)
    Call Regle_Prepaid(dictBal, dictGL)
    Call Regle_Variation(dictBal, dictPrev)
    Call Regle_OverAged(dictBal, dictGL)
    Call Regle_Limites(dictBal)
    Call Regle_Intersol(dictBal)
    Call Regle_SaisiesManuelles(dictBal, wsTx)
    Call Regle_ChequesDirection(dictBal, dictGL)
    Call Regle_SoldesInhabituels(dictBal, dictPrev)
    Call Regle_MouvementsDepenses(dictBal, wsTx)
    Call Regle_WIP(dictBal, dictGL)
    Call Regle_ChargesEnAttente(dictBal, wsTx)
    Call Regle_Contreparties

    ' --- Mise en forme et rating ---
    mWsOut.Columns("A:K").AutoFit
    mWsOut.Columns("H").ColumnWidth = 70
    mWsOut.Columns("J").ColumnWidth = 55
    mWsOut.Columns("H:J").WrapText = True
    mWsOut.Range("F2:F" & mRow).NumberFormat = "#,##0"
    If mRow > 2 Then mWsOut.Range("A1:K" & (mRow - 1)).AutoFilter

    Call EcrireRating

    Call SAFA_Common.WriteAuditLog("PROCESS", "GL Monitoring termine", (mRow - 2) & " constats en " & Format(Timer - t0, "0.0") & " s")

    If showMsg Then
        MsgBox "GL Monitoring termine: " & (mRow - 2) & " constat(s)." & vbCrLf & vbCrLf & _
               "Feuilles: GL_MONITORING (constats), PROOFABLE_UNIVERSE (univers), GL_RATING (note).", _
               vbInformation, "GL Monitoring"
    End If
    Exit Sub

ErrHandler:
    Call SAFA_Common.LogError("GL_Monitoring", "Lancer_GL_Monitoring", Err.Number, Err.Description)
    If showMsg Then MsgBox "Erreur GL Monitoring: " & Err.Description, vbCritical, "GL Monitoring"
End Sub

' ==============================================================================
' CLASSIFICATION DES COMPTES (nomenclature Finacle UBA + repli OHADA)
' ==============================================================================

Public Function ClassifyAccount(acct As String, name As String, balance As Double) As AccountClass
    Dim r As AccountClass
    Dim a As String, n As String, core As String, both As String
    Dim ccy As Variant, isUba As Boolean, cls As String

    Call LoadCfg
    a = UCase(Trim(acct)): n = UCase(Trim(name)): both = a & " " & n
    r.Family = "INCONNU": r.Nature = "STANDARD": r.Proofable = True: r.Reason = ""

    ' Prefixe devise (comptes internes Finacle: XAF..., USD..., EUR..., GBP...)
    core = a
    For Each ccy In Split(mCfg.CurrencyPrefixes, "|")
        If Len(ccy) > 0 Then
            If Left(a, Len(ccy)) = UCase(ccy) Then isUba = True: core = Mid(a, Len(ccy) + 1): Exit For
        End If
    Next ccy
    If Len(core) > 0 Then
        If core Like "[0-9]*" Then cls = Left(core, 1)
    End If

    ' 1. Comptes systeme (ecritures exclusivement automatiques): non proofables
    If ContainsAny(both, mCfg.SystemPrefixes) Then
        r.Nature = "SYSTEM": r.Proofable = False
        r.Reason = "Compte systeme (IENC / RFI / position / Finnone): ecritures automatiques"
        Call FamilyByName(r, n)
        ClassifyAccount = r
        Exit Function
    End If

    ' 2. Comptes de resultat: prefixe PAL (Finacle) ou classes 6 / 7 (OHADA)
    If (Len(mCfg.PLPrefix) > 0 And InStr(a, UCase(mCfg.PLPrefix)) > 0) Then
        r.Nature = "PL": r.Proofable = False
        ' Famille par le signe selon la convention detectee pour le resultat
        ' (Finacle PAL par defaut: montant negatif = charge, sans signe = produit)
        If PLDebitPositive() Then
            r.Family = IIf(balance < 0, "PRODUIT", "CHARGE")
        Else
            r.Family = IIf(balance < 0, "CHARGE", "PRODUIT")
        End If
        If ContainsAny(n, mCfg.KwRevenue) Then r.Family = "PRODUIT"
        If ContainsAny(n, mCfg.KwExpense) Then r.Family = "CHARGE"
        r.Reason = "Compte de resultat (PAL): non proofable"
        ClassifyAccount = r
        Exit Function
    End If
    If cls <> "" And ContainsAny(cls, mCfg.PLClasses) Then
        r.Nature = "PL": r.Proofable = False
        r.Family = IIf(cls = "6", "CHARGE", "PRODUIT")
        r.Reason = "Compte de resultat (classe " & cls & "): non proofable"
        ClassifyAccount = r
        Exit Function
    End If

    ' 3. Engagements / hors bilan: non proofables
    If cls <> "" And ContainsAny(cls, mCfg.NonProofableClasses) Then
        r.Family = "ENGAGEMENT": r.Proofable = False
        r.Reason = "Engagement / hors bilan (classe " & cls & "): non proofable"
        ClassifyAccount = r
        Exit Function
    End If

    ' 4. Nature par mots-cles (ordre: proxy, transit, suspens, prepaid, ecart, caisse)
    If ContainsAny(both, mCfg.InterbranchPrefix) Then
        r.Nature = "INTERSOL"
    ElseIf ContainsAny(n, mCfg.KwMgrCheque) Then
        r.Nature = "MGR_CHEQUE"
    ElseIf ContainsAny(n, mCfg.KwWip) Then
        r.Nature = "WIP"
    ElseIf ContainsAny(a, mCfg.PrepaidGLCodes) Then
        r.Nature = "PREPAID"
    ElseIf ContainsAny(both, mCfg.KwProxy) Then
        r.Nature = "PROXY"
    ElseIf ContainsAny(both, mCfg.KwTransit) Then
        r.Nature = "TRANSIT"
    ElseIf ContainsAny(both, mCfg.KwSuspense) Then
        r.Nature = "SUSPENS"
    ElseIf ContainsAny(both, mCfg.KwPrepaid) Then
        r.Nature = "PREPAID"
    ElseIf ContainsAny(both, mCfg.KwDifference) Then
        r.Nature = "DIFFERENCE"
    ElseIf ContainsAny(both, mCfg.KwCash) Or ContainsAny(both, mCfg.KwVault) Then
        r.Nature = "CASH"
    End If

    ' 5. Famille
    Select Case r.Nature
        Case "PREPAID", "CASH", "WIP": r.Family = "ACTIF"
        Case "MGR_CHEQUE": r.Family = "PASSIF"
        Case "TRANSIT", "PROXY", "SUSPENS", "INTERSOL": r.Family = "TIERS"
        Case "DIFFERENCE"
            ' Excedents de caisse (overage) = passif ; manquants (shortage) = actif
            If ContainsAny(n, mCfg.KwOverage) Then
                r.Family = "PASSIF"
            ElseIf ContainsAny(n, mCfg.KwShortage) Then
                r.Family = "ACTIF"
            Else
                r.Family = "TIERS"
            End If
        Case Else
            If cls <> "" Then
                Select Case cls
                    Case "1": r.Family = "CAPITAUX"
                    Case "2", "3", "5": r.Family = "ACTIF"
                    Case "4"
                        Select Case Mid(core, 2, 1)
                            Case "0": r.Family = "PASSIF"
                            Case "1": r.Family = "ACTIF"
                            Case Else: r.Family = "TIERS"
                        End Select
                    Case Else: r.Family = "INCONNU"
                End Select
            End If
            Call FamilyByName(r, n)
    End Select

    r.Reason = "Compte interne proofable"
    ClassifyAccount = r
End Function

Private Sub FamilyByName(ByRef r As AccountClass, n As String)
    ' Affine la famille par les libelles (utile pour la nomenclature Finacle sans classe numerique)
    If r.Family = "INCONNU" Or r.Family = "TIERS" Then
        If ContainsAny(n, mCfg.KwAsset) Then
            r.Family = "ACTIF"
        ElseIf ContainsAny(n, mCfg.KwLiability) Then
            r.Family = "PASSIF"
        End If
    End If
End Sub

' ==============================================================================
' GLM-001 / GLM-002 : UNIVERS PROOFABLE ET EXHAUSTIVITE
' ==============================================================================

Private Sub ConstruireUnivers(dictBal As Object, dictGL As Object, dictCI As Object)
    Dim ws As Worksheet, r As Long, k As Variant
    Dim cls As AccountClass, name As String, bal As Double
    Dim nProof As Long, nNonProof As Long, nMissing As Long

    Set ws = SAFA_Common.GetOrCreateSheet("PROOFABLE_UNIVERSE", True)
    ws.Range("A1:I1").Value = Array("Compte", "Libelle", "Solde", "Famille", "Nature", "Proofable", "Motif", "Dans liste CI", "Proof present")
    SAFA_Common.FormatHeader ws.Range("A1:I1")
    r = 2

    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        ws.Cells(r, 1).Value = "'" & k
        ws.Cells(r, 2).Value = name
        ws.Cells(r, 3).Value = bal
        ws.Cells(r, 4).Value = cls.Family
        ws.Cells(r, 5).Value = cls.Nature
        ws.Cells(r, 6).Value = IIf(cls.Proofable, "OUI", "NON")
        ws.Cells(r, 7).Value = cls.Reason
        ws.Cells(r, 8).Value = IIf(dictCI.Count = 0, "n/a", IIf(dictCI.Exists(k), "OUI", "NON"))
        ws.Cells(r, 9).Value = IIf(dictGL.Exists(k), "OUI", "NON")
        If cls.Proofable Then
            nProof = nProof + 1
            ws.Cells(r, 6).Interior.Color = RGB(214, 240, 220)
            ' GLM-001: proofable mais absent de la liste CI
            If dictCI.Count > 0 And Not dictCI.Exists(k) Then
                Call Ecrire("GLM-001", CStr(k), name, cls, bal, 0, _
                    "Compte proofable absent de la liste des comptes proofables du Controle Interne (exhaustivite)", _
                    "MEDIUM", "Faire ajouter le compte a la liste CI et exiger un proof", mCfg.RatingPointsUniverse)
            End If
            ' GLM-002: proof manquant
            If Not dictGL.Exists(k) Then
                nMissing = nMissing + 1
                ws.Cells(r, 9).Interior.Color = RGB(252, 226, 226)
                Call Ecrire("GLM-002", CStr(k), name, cls, bal, 0, _
                    IIf(bal = 0, "Compte proofable sans bloc GL Proof (solde nul: proof 'neant' attendu)", _
                                 "Compte proofable sans bloc GL Proof: PROOF MANQUANT"), _
                    IIf(bal = 0, "LOW", IIf(Abs(bal) > 1000000, "CRITICAL", "HIGH")), _
                    "Exiger le proof du compte aupres des Operations / FINCON", IIf(bal = 0, 0, mCfg.RatingPointsProofMissing))
            End If
        Else
            nNonProof = nNonProof + 1
            ws.Cells(r, 6).Interior.Color = RGB(235, 235, 235)
        End If
        r = r + 1
    Next k

    ' Comptes de la liste CI absents de la balance
    For Each k In dictCI.Keys
        If Not dictBal.Exists(k) Then
            Dim c2 As AccountClass
            c2.Family = "": c2.Nature = ""
            Call Ecrire("GLM-001", CStr(k), "(liste CI)", c2, 0, 0, _
                "Compte de la liste CI absent de la balance extraite (liste obsolete ou perimetre different)", _
                "LOW", "Mettre a jour la liste des comptes proofables", 0)
        End If
    Next k

    ws.Columns("A:I").AutoFit
    ws.Range("C2:C" & r).NumberFormat = "#,##0"
    If r > 2 Then ws.Range("A1:I" & (r - 1)).AutoFilter

    If dictCI.Count = 0 Then
        Dim c3 As AccountClass
        Call Ecrire("GLM-001", "-", "-", c3, 0, 0, _
            "Liste CI des comptes proofables absente (feuille PROOFABLE_LIST): exhaustivite non verifiee. " & _
            nProof & " comptes proofables identifies par SAFA, " & nNonProof & " non proofables, " & nMissing & " sans proof", _
            "INFO", "Importer la liste CI (bouton 'Liste proofables') puis relancer", 0)
    End If
End Sub

' ==============================================================================
' GLM-003 : SENS DES SOLDES
' ==============================================================================

Private Sub Regle_Sens(dictBal As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double, sgn As Double
    Dim constat As String

    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        If bal <> 0 Then
            cls = ClassifyAccount(CStr(k), name, bal)
            ' sgn > 0 = debiteur, selon la convention detectee pour la famille (bilan ou resultat)
            If cls.Family = "CHARGE" Or cls.Family = "PRODUIT" Then
                sgn = IIf(PLDebitPositive(), bal, -bal)
            Else
                sgn = IIf(BSDebitPositive(), bal, -bal)
            End If
            constat = ""
            If cls.Nature = "PL" And InStr(UCase(k), UCase(mCfg.PLPrefix)) > 0 Then
                ' Comptes PAL: la famille est deduite du signe; on ne verifie que par le libelle
                If ContainsAny(name, mCfg.KwRevenue) And sgn > 0 Then constat = "Compte de produit a solde DEBITEUR"
                If ContainsAny(name, mCfg.KwExpense) And sgn < 0 Then constat = "Compte de charge a solde CREDITEUR"
            Else
                Select Case cls.Family
                    Case "ACTIF": If sgn < 0 Then constat = "Compte d'actif a solde CREDITEUR"
                    Case "PASSIF", "CAPITAUX": If sgn > 0 Then constat = "Compte de passif a solde DEBITEUR"
                    Case "CHARGE": If sgn < 0 Then constat = "Compte de charge a solde CREDITEUR"
                    Case "PRODUIT": If sgn > 0 Then constat = "Compte de produit a solde DEBITEUR"
                End Select
            End If
            If constat <> "" Then
                Call Ecrire("GLM-003", CStr(k), name, cls, bal, 0, constat & " (" & Format(bal, "#,##0") & ")", _
                    IIf(Abs(bal) > 1000000, "HIGH", "MEDIUM"), _
                    "Verifier l'origine des transactions (journal, pieces), l'autorisation et la correction", mCfg.RatingPointsSense)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-004 : TRANSIT / PROXY / SUSPENS NON NULS
' ==============================================================================

Private Sub Regle_TransitSuspens(dictBal As Object, dictGL As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double
    Dim ageMax As Double, pts As Double, grav As String, tranches As Long

    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        If (cls.Nature = "TRANSIT" Or cls.Nature = "PROXY" Or cls.Nature = "SUSPENS") And bal <> 0 Then
            ageMax = 0
            If dictGL.Exists(k) Then ageMax = dictGL(k)(3)
            If ageMax > 30 Or Abs(bal) > 5000000 Then
                grav = "CRITICAL"
            ElseIf ageMax > mCfg.SuspenseRatingDays Then
                grav = "HIGH"
            Else
                grav = "MEDIUM"
            End If
            tranches = 1
            If mCfg.RatingTrancheDays > 0 Then tranches = Application.WorksheetFunction.Max(1, -Int(-ageMax / mCfg.RatingTrancheDays))
            pts = Application.WorksheetFunction.Min(10, mCfg.RatingPointsPerTranche * tranches)
            Call Ecrire("GLM-004", CStr(k), name, cls, bal, ageMax, _
                "Compte " & LCase(cls.Nature) & " a solde non nul (attendu: zero sous " & mCfg.TransitZeroDays & " jour)" & _
                IIf(ageMax > 0, " - item le plus ancien: " & Format(ageMax, "0") & " j", ""), _
                grav, "Examiner les mouvements du jour ET les contreparties (compte cible), regulariser; ne jamais y stocker de charges", pts)
        End If
    Next k
End Sub

' ==============================================================================
' GLM-005 : DEBITS SUR COMPTES DE PRODUITS
' ==============================================================================

Private Sub Regle_DebitProduits(dictBal As Object, wsTx As Worksheet)
    Dim lr As Long, i As Long, k As String, mnt As Double, d As Variant
    Dim dictCredits As Object, cls As AccountClass, name As String
    Dim dictIsRevenue As Object, key2 As String, nAlert As Long

    If wsTx Is Nothing Then
        Dim c0 As AccountClass
        Call Ecrire("GLM-005", "-", "-", c0, 0, 0, "TRANSACTION_DATA absente: debits sur comptes de produits non evalues", "INFO", "Relancer l'analyse complete", 0)
        Exit Sub
    End If

    Set dictCredits = CreateObject("Scripting.Dictionary")
    Set dictIsRevenue = CreateObject("Scripting.Dictionary")
    lr = wsTx.Cells(wsTx.Rows.Count, 1).End(xlUp).Row

    ' Passe 1: credits par compte|date|montant (pour detecter la correction le jour meme)
    For i = 2 To lr
        k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsTx.Cells(i, 1).Value))
        mnt = SAFA_Common.SafeVal(wsTx.Cells(i, 4).Value)
        d = wsTx.Cells(i, 2).Value
        If k <> "" And mnt > 0 And IsDate(d) Then
            key2 = k & "|" & Format(CDate(d), "yyyymmdd") & "|" & Format(Abs(mnt), "0")
            dictCredits(key2) = 1
        End If
    Next i

    ' Passe 2: debits sur comptes de produits
    For i = 2 To lr
        k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsTx.Cells(i, 1).Value))
        mnt = SAFA_Common.SafeVal(wsTx.Cells(i, 4).Value)
        d = wsTx.Cells(i, 2).Value
        If k <> "" And mnt < 0 Then
            If Not dictIsRevenue.Exists(k) Then
                name = ""
                If dictBal.Exists(k) Then name = dictBal(k)(0)
                cls = ClassifyAccount(k, name, IIf(dictBal.Exists(k), dictBal(k)(1), 0))
                dictIsRevenue(k) = (cls.Family = "PRODUIT")
            End If
            If dictIsRevenue(k) Then
                name = "": If dictBal.Exists(k) Then name = dictBal(k)(0)
                cls = ClassifyAccount(k, name, 0)
                key2 = k & "|" & IIf(IsDate(d), Format(CDate(d), "yyyymmdd"), "") & "|" & Format(Abs(mnt), "0")
                If dictCredits.Exists(key2) Then
                    Call Ecrire("GLM-005", k, name, cls, mnt, 0, _
                        "Debit sur compte de produit corrige le jour meme (" & Format(mnt, "#,##0") & " le " & Format(d, "dd/mm/yyyy") & "): " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                        "LOW", "Verifier la piece de correction", 0)
                Else
                    nAlert = nAlert + 1
                    Call Ecrire("GLM-005", k, name, cls, mnt, 0, _
                        "DEBIT sur compte de produit (" & Format(mnt, "#,##0") & " le " & Format(d, "dd/mm/yyyy") & "): " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                        "HIGH", "Extourne / annulation de produit: verifier l'approbation requise (Groupe pour les interets sur comptes dormants)", mCfg.RatingPointsRevenueDebit)
                End If
            End If
        End If
    Next i
End Sub

' ==============================================================================
' GLM-006 : CHARGES CONSTATEES D'AVANCE
' ==============================================================================

Private Sub Regle_Prepaid(dictBal As Object, dictGL As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double
    Dim lastDate As Date, nbTrans As Double, jours As Long

    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        If bal <> 0 Then
            cls = ClassifyAccount(CStr(k), name, bal)
            If cls.Nature = "PREPAID" Then
                lastDate = 0: nbTrans = 0
                If dictGL.Exists(k) Then lastDate = dictGL(k)(4): nbTrans = dictGL(k)(2)
                If lastDate = 0 Or nbTrans = 0 Then
                    Call Ecrire("GLM-006", CStr(k), name, cls, bal, 0, _
                        "Charge constatee d'avance sans mouvement au GL Proof: amortissement mensuel non constate", _
                        "MEDIUM", "Verifier le tableau d'amortissement lineaire (solde attendu = paye - mensualites ecoulees), pieces, duree approuvee, instruction permanente", mCfg.RatingPointsPrepaid)
                Else
                    jours = DateDiff("d", lastDate, Date)
                    If jours > mCfg.PrepaidRegularizationDays Then
                        Call Ecrire("GLM-006", CStr(k), name, cls, bal, jours, _
                            "Dernier mouvement il y a " & jours & " jours (> " & mCfg.PrepaidRegularizationDays & "): amortissement mensuel non constate", _
                            "MEDIUM", "Regulariser sous 30 jours; verifier le tableau d'amortissement et les autorisations", mCfg.RatingPointsPrepaid)
                    End If
                End If
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-007 : VARIATIONS DES CHARGES / PRODUITS RECURRENTS
' ==============================================================================

Private Sub Regle_Variation(dictBal As Object, dictPrev As Object)
    Dim k As Variant, cls As AccountClass, name As String, cur As Double, prev As Double, v As Double
    Dim c0 As AccountClass

    If dictPrev.Count = 0 Then
        Call Ecrire("GLM-007", "-", "-", c0, 0, 0, _
            "Balance de la periode precedente absente (feuille BALANCE_PREV_RAW): variations des charges recurrentes non evaluees", _
            "INFO", "Importer la balance N-1 (bouton 'Balance N-1') puis relancer", 0)
        Exit Sub
    End If

    For Each k In dictBal.Keys
        name = dictBal(k)(0): cur = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, cur)
        If cls.Family = "CHARGE" Or cls.Family = "PRODUIT" Then
            If dictPrev.Exists(k) Then
                prev = dictPrev(k)
                If Abs(prev) >= mCfg.VariationMinAmount Then
                    v = (cur - prev) / Abs(prev)
                    If Abs(v) >= mCfg.ExpenseVariationCritical Then
                        Call Ecrire("GLM-007", CStr(k), name, cls, cur, 0, _
                            "Variation de " & Format(v, "+0%;-0%") & " vs periode precedente (" & Format(prev, "#,##0") & " -> " & Format(cur, "#,##0") & ")", _
                            "HIGH", "Demander justificatifs, autorisation et approbation (variation > " & Format(mCfg.ExpenseVariationCritical, "0%") & ")", 2)
                    ElseIf Abs(v) >= mCfg.ExpenseVariationWarn Then
                        Call Ecrire("GLM-007", CStr(k), name, cls, cur, 0, _
                            "Variation de " & Format(v, "+0%;-0%") & " vs periode precedente (" & Format(prev, "#,##0") & " -> " & Format(cur, "#,##0") & ")", _
                            "MEDIUM", "Demander une explication (variation > " & Format(mCfg.ExpenseVariationWarn, "0%") & ")", 1)
                    End If
                End If
            ElseIf Abs(cur) >= mCfg.VariationMinAmount Then
                Call Ecrire("GLM-007", CStr(k), name, cls, cur, 0, "Compte de resultat absent de la balance precedente (nouveau)", "LOW", "Verifier la creation du compte", 0)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-009 : ITEMS OVER-AGED (ecarts caisse / ATM et autres comptes proofables)
' ==============================================================================

Private Sub Regle_OverAged(dictBal As Object, dictGL As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double, ageMax As Double
    Dim grav As String, pts As Double

    For Each k In dictGL.Keys
        ageMax = dictGL(k)(3)
        If ageMax >= mCfg.OverAgedTier1 Then
            name = dictGL(k)(0): bal = dictGL(k)(1)
            If dictBal.Exists(k) Then name = dictBal(k)(0): bal = dictBal(k)(1)
            cls = ClassifyAccount(CStr(k), name, bal)
            If cls.Proofable And cls.Nature <> "TRANSIT" And cls.Nature <> "PROXY" And cls.Nature <> "SUSPENS" Then
                If ageMax >= mCfg.OverAgedTier3 Then
                    grav = "HIGH": pts = mCfg.RatingPointsOverAged * 1.5
                ElseIf ageMax >= mCfg.OverAgedTier2 Then
                    grav = "MEDIUM": pts = mCfg.RatingPointsOverAged
                Else
                    grav = "LOW": pts = mCfg.RatingPointsOverAged / 2
                End If
                Call Ecrire("GLM-009", CStr(k), name, cls, bal, ageMax, _
                    IIf(cls.Nature = "DIFFERENCE", "Ecart de caisse / ATM", "Items") & " non regularise(s) depuis " & Format(ageMax, "0") & " jours (over-aged)", _
                    grav, "Detailler operation par operation, identifier les transactions et les responsables, provisionner si necessaire", pts)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-010 : LIMITES DE CAISSE ET DE COFFRE
' ==============================================================================

Private Sub Regle_Limites(dictBal As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double, isVault As Boolean

    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        If cls.Nature = "CASH" Then
            isVault = ContainsAny(UCase(name), mCfg.KwVault)
            If isVault And mCfg.VaultLimit > 0 And Abs(bal) > mCfg.VaultLimit Then
                Call Ecrire("GLM-010", CStr(k), name, cls, bal, 0, _
                    "Solde de coffre (" & Format(bal, "#,##0") & ") au-dela du montant assure (" & Format(mCfg.VaultLimit, "#,##0") & ")", _
                    "HIGH", "Transferer l'excedent a la Banque centrale; verifier la police d'assurance", mCfg.RatingPointsCashLimit)
            ElseIf Not isVault And mCfg.CashLimit > 0 And Abs(bal) > mCfg.CashLimit Then
                Call Ecrire("GLM-010", CStr(k), name, cls, bal, 0, _
                    "Solde de caisse (" & Format(bal, "#,##0") & ") au-dela de la limite (" & Format(mCfg.CashLimit, "#,##0") & ")", _
                    "HIGH", "Verifier la limite approuvee et le montant assure; excedent a transferer", mCfg.RatingPointsCashLimit)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-011 : INTERSOL / IMPUTATIONS INTER-AGENCES
' ==============================================================================

Private Sub Regle_Intersol(dictBal As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double
    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        If cls.Nature = "INTERSOL" And bal <> 0 Then
            Call Ecrire("GLM-011", CStr(k), name, cls, bal, 0, _
                "Compte inter-agences (INTERSOL) a solde non nul", "LOW", _
                "Documenter chaque charge imputee par le siege et verifier sa coherence avec l'activite reelle de l'agence", 0)
        End If
    Next k
End Sub

' ==============================================================================
' GLM-012 : SAISIES MANUELLES SUR COMPTES AUTOMATISES
' ==============================================================================

Private Sub Regle_SaisiesManuelles(dictBal As Object, wsTx As Worksheet)
    Dim c0 As AccountClass, colUser As Long, j As Long, h As String
    Dim lr As Long, i As Long, k As String, user As String, cls As AccountClass, name As String

    If wsTx Is Nothing Then Exit Sub
    colUser = 0
    For j = 1 To 40
        h = UCase(SAFA_Common.SafeText(wsTx.Cells(1, j).Value))
        If h <> "" Then
            If InStr(h, "USER") > 0 Or InStr(h, "POST") > 0 Or InStr(h, "MAKER") > 0 Or InStr(h, "UTILISATEUR") > 0 Or InStr(h, "SAISI") > 0 Then colUser = j: Exit For
        End If
    Next j

    If colUser = 0 Then
        Call Ecrire("GLM-012", "-", "-", c0, 0, 0, _
            "Identifiant du posteur absent de TRANSACTION_DATA: saisies manuelles sur comptes systeme / produits non evaluees (dans le journal Finacle: CDCI = automatique, identifiant du caissier = manuel)", _
            "INFO", "Extraire le journal avec la colonne utilisateur pour activer ce controle", 0)
        Exit Sub
    End If

    lr = wsTx.Cells(wsTx.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        user = UCase(SAFA_Common.SafeText(wsTx.Cells(i, colUser).Value))
        If user <> "" And Not ContainsAny(user, mCfg.AutoUserIds) Then
            k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsTx.Cells(i, 1).Value))
            name = "": If dictBal.Exists(k) Then name = dictBal(k)(0)
            cls = ClassifyAccount(k, name, 0)
            If cls.Nature = "SYSTEM" Or cls.Family = "PRODUIT" Then
                Call Ecrire("GLM-012", k, name, cls, SAFA_Common.SafeVal(wsTx.Cells(i, 4).Value), 0, _
                    "Saisie MANUELLE (" & user & ") sur compte " & IIf(cls.Nature = "SYSTEM", "systeme", "de produit") & " le " & Format(wsTx.Cells(i, 2).Value, "dd/mm/yyyy") & ": " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                    "CRITICAL", "Approfondir: motif de la saisie et autorisations prealables", mCfg.RatingPointsRevenueDebit)
            End If
        End If
    Next i
End Sub

' ==============================================================================
' GLM-008 : CONTREPARTIES DES COMPTES D'ATTENTE (information / rappel)
' ==============================================================================

Private Sub Regle_Contreparties()
    Dim c0 As AccountClass
    Call Ecrire("GLM-008", "-", "-", c0, 0, 0, _
        "Schema de fraude connu: debit caisse -> credit compte d'attente, puis debit attente -> credit compte d'un employe ou d'un proche; le compte d'attente ressort a zero. " & _
        "Le controle des contreparties n'est pas automatisable sans extrait du journal avec compte de contrepartie et code departemental.", _
        "INFO", "Dans Finacle, lire les contreparties de chaque mouvement des comptes d'attente / transit / proxy listes en GLM-004 et verifier le code departemental", 0)
End Sub

' ==============================================================================
' GLM-013 : CHEQUES DE DIRECTION PERIMES
' ==============================================================================

Private Sub Regle_ChequesDirection(dictBal As Object, dictGL As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double, ageMax As Double
    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        If cls.Nature = "MGR_CHEQUE" And bal <> 0 Then
            ageMax = 0
            If dictGL.Exists(k) Then ageMax = dictGL(k)(3)
            If ageMax > mCfg.MgrChequeStaleDays Then
                Call Ecrire("GLM-013", CStr(k), name, cls, bal, ageMax, _
                    "Cheques de direction en circulation depuis " & Format(ageMax, "0") & " jours (> " & mCfg.MgrChequeStaleDays & "): cheques perimes", _
                    IIf(Abs(bal) > 5000000, "HIGH", "MEDIUM"), _
                    "Transferer les cheques perimes vers le compte des elements non reclames; verifier le detail par cheque", mCfg.RatingPointsOther)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-014 : SOLDES INHABITUELS / EXCEPTIONNELS (tous comptes internes)
' ==============================================================================

Private Sub Regle_SoldesInhabituels(dictBal As Object, dictPrev As Object)
    Dim k As Variant, cls As AccountClass, name As String, cur As Double, prev As Double
    If dictPrev.Count = 0 Then Exit Sub   ' deja signale en GLM-007
    For Each k In dictBal.Keys
        name = dictBal(k)(0): cur = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, cur)
        If cls.Proofable And cls.Nature <> "PL" And Abs(cur) >= mCfg.UnusualBalanceMin Then
            If dictPrev.Exists(k) Then
                prev = dictPrev(k)
                If Abs(cur) >= mCfg.UnusualBalanceFactor * Abs(prev) Then
                    Call Ecrire("GLM-014", CStr(k), name, cls, cur, 0, _
                        "Solde inhabituel: " & Format(prev, "#,##0") & " -> " & Format(cur, "#,##0") & " (x" & IIf(prev = 0, "n/a", Format(Abs(cur) / Abs(prev), "0.0")) & ")", _
                        IIf(Abs(cur) > 10000000, "HIGH", "MEDIUM"), _
                        "Rapprocher avec le journal et les pieces: autorisation, approbation, authenticite, existence", mCfg.RatingPointsOther)
                End If
            Else
                Call Ecrire("GLM-014", CStr(k), name, cls, cur, 0, _
                    "Compte absent de la periode precedente avec un solde significatif", "LOW", _
                    "Verifier l'ouverture du compte et la justification du solde", 0)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-015 / GLM-016 : MOUVEMENTS DE DEPENSES ET NARRATIONS SENSIBLES
' ==============================================================================

Private Sub Regle_MouvementsDepenses(dictBal As Object, wsTx As Worksheet)
    Dim lr As Long, i As Long, k As String, mnt As Double, narr As String, name As String
    Dim cls As AccountClass, dictFam As Object
    If wsTx Is Nothing Then Exit Sub
    Set dictFam = CreateObject("Scripting.Dictionary")
    lr = wsTx.Cells(wsTx.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsTx.Cells(i, 1).Value))
        mnt = SAFA_Common.SafeVal(wsTx.Cells(i, 4).Value)
        narr = UCase(SAFA_Common.SafeText(wsTx.Cells(i, 3).Value))
        If k <> "" And mnt <> 0 Then
            name = "": If dictBal.Exists(k) Then name = dictBal(k)(0)
            If Not dictFam.Exists(k) Then
                cls = ClassifyAccount(k, name, IIf(dictBal.Exists(k), dictBal(k)(1), 0))
                dictFam(k) = cls.Family
            End If
            ' GLM-016: narration sensible (perte, fraude, detournement) sur tout compte
            If ContainsAny(narr, mCfg.KwFraudLoss) Then
                cls = ClassifyAccount(k, name, 0)
                Call Ecrire("GLM-016", k, name, cls, mnt, 0, _
                    "Ecriture sensible (" & Format(mnt, "#,##0") & " le " & Format(wsTx.Cells(i, 2).Value, "dd/mm/yyyy") & "): " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                    "HIGH", "Remboursement de perte / fraude: verifier l'approbation requise et le dossier", mCfg.RatingPointsOther)
            ' GLM-015: depense significative
            ElseIf dictFam(k) = "CHARGE" And Abs(mnt) >= mCfg.ExpenseMovementThreshold Then
                cls = ClassifyAccount(k, name, 0)
                Call Ecrire("GLM-015", k, name, cls, mnt, 0, _
                    "Mouvement de depense significatif (" & Format(mnt, "#,##0") & " le " & Format(wsTx.Cells(i, 2).Value, "dd/mm/yyyy") & "): " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                    IIf(Abs(mnt) >= 2 * mCfg.ExpenseMovementThreshold, "HIGH", "MEDIUM"), _
                    "Verifier la regularite, l'autorisation et la delegation de pouvoir de l'approbateur (limite locale / Groupe)", mCfg.RatingPointsOther)
            End If
        End If
    Next i
End Sub

' ==============================================================================
' GLM-017 : TRAVAUX EN COURS (WIP)
' ==============================================================================

Private Sub Regle_WIP(dictBal As Object, dictGL As Object)
    Dim k As Variant, cls As AccountClass, name As String, bal As Double, ageMax As Double, nbTrans As Double
    For Each k In dictBal.Keys
        name = dictBal(k)(0): bal = dictBal(k)(1)
        cls = ClassifyAccount(CStr(k), name, bal)
        If cls.Nature = "WIP" And bal <> 0 Then
            ageMax = 0: nbTrans = 0
            If dictGL.Exists(k) Then ageMax = dictGL(k)(3): nbTrans = dictGL(k)(2)
            If ageMax > mCfg.WipStaleDays Or nbTrans = 0 Then
                Call Ecrire("GLM-017", CStr(k), name, cls, bal, ageMax, _
                    "Travaux en cours " & IIf(nbTrans = 0, "sans detail au GL Proof", "avec des items de " & Format(ageMax, "0") & " jours") & ": comptabilisation ou capitalisation a verifier", _
                    IIf(Abs(bal) > mCfg.ExpenseMovementThreshold, "HIGH", "MEDIUM"), _
                    "Verifier l'approbation du projet, la delegation de pouvoir, l'avancement, la concordance approuve/comptabilise et les pieces (devis, factures, recus)", mCfg.RatingPointsOther)
            End If
        End If
    Next k
End Sub

' ==============================================================================
' GLM-018 : CHARGES STOCKEES DANS LES COMPTES D'ATTENTE / TRANSIT
' ==============================================================================

Private Sub Regle_ChargesEnAttente(dictBal As Object, wsTx As Worksheet)
    Dim lr As Long, i As Long, k As String, mnt As Double, narr As String, name As String
    Dim cls As AccountClass, dictNat As Object
    If wsTx Is Nothing Then Exit Sub
    Set dictNat = CreateObject("Scripting.Dictionary")
    lr = wsTx.Cells(wsTx.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(wsTx.Cells(i, 1).Value))
        mnt = SAFA_Common.SafeVal(wsTx.Cells(i, 4).Value)
        narr = UCase(SAFA_Common.SafeText(wsTx.Cells(i, 3).Value))
        If k <> "" And mnt <> 0 Then
            name = "": If dictBal.Exists(k) Then name = dictBal(k)(0)
            If Not dictNat.Exists(k) Then
                cls = ClassifyAccount(k, name, IIf(dictBal.Exists(k), dictBal(k)(1), 0))
                dictNat(k) = cls.Nature
            End If
            If (dictNat(k) = "SUSPENS" Or dictNat(k) = "TRANSIT" Or dictNat(k) = "PROXY") And ContainsAny(narr, mCfg.KwExpenseNarration) Then
                cls = ClassifyAccount(k, name, 0)
                Call Ecrire("GLM-018", k, name, cls, mnt, 0, _
                    "Charge apparemment stockee dans un compte " & LCase(dictNat(k)) & " (" & Format(mnt, "#,##0") & " le " & Format(wsTx.Cells(i, 2).Value, "dd/mm/yyyy") & "): " & SAFA_Common.SafeText(wsTx.Cells(i, 3).Value, 80), _
                    "CRITICAL", "Interdit: ne jamais stocker de charges en compte d'attente (preservation du resultat). Imputer en charge, verifier l'approbation", mCfg.RatingPointsOther * 2)
            End If
        End If
    Next i
End Sub

' ==============================================================================
' RATING
' ==============================================================================

Private Sub EcrireRating()
    Dim ws As Worksheet, r As Long, rule As Variant, total As Double, score As Double
    Dim labels As Object, sol As String

    Set labels = CreateObject("Scripting.Dictionary")
    labels("GLM-001") = "Exhaustivite de l'univers proofable"
    labels("GLM-002") = "Proofs manquants"
    labels("GLM-003") = "Sens des soldes"
    labels("GLM-004") = "Transit / proxy / suspens non nuls"
    labels("GLM-005") = "Debits sur comptes de produits"
    labels("GLM-006") = "Charges constatees d'avance"
    labels("GLM-007") = "Variations charges / produits"
    labels("GLM-008") = "Contreparties comptes d'attente"
    labels("GLM-009") = "Items over-aged"
    labels("GLM-010") = "Limites caisse / coffre"
    labels("GLM-011") = "Imputations inter-agences"
    labels("GLM-012") = "Saisies manuelles comptes automatises"
    labels("GLM-013") = "Cheques de direction perimes"
    labels("GLM-014") = "Soldes inhabituels / exceptionnels"
    labels("GLM-015") = "Mouvements de depenses significatifs"
    labels("GLM-016") = "Pertes / fraude: approbations"
    labels("GLM-017") = "Travaux en cours (WIP)"
    labels("GLM-018") = "Charges stockees en comptes d'attente"

    sol = "799"
    On Error Resume Next
    sol = SAFA_Common.SafeText(ThisWorkbook.Sheets("PARAM").Range("F2").Value)
    On Error GoTo 0

    Set ws = SAFA_Common.GetOrCreateSheet("GL_RATING", True)
    ws.Range("A1").Value = "GL RATING - NOTE MENSUELLE DU GRAND LIVRE"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Perimetre: SOL " & sol & " - Date: " & Format(Date, "dd/mm/yyyy") & " - Grille: " & _
                           mCfg.RatingPointsPerTranche & " pt par tranche de " & mCfg.RatingTrancheDays & " j (transit/suspens), a verifier avec le template CI en vigueur (grille revisee du 28/02/2026)"
    ws.Range("A2").Font.Italic = True

    ws.Range("A4:D4").Value = Array("Regle", "Controle", "Nb constats", "Points deduits")
    SAFA_Common.FormatHeader ws.Range("A4:D4")
    r = 5
    For Each rule In labels.Keys
        ws.Cells(r, 1).Value = rule
        ws.Cells(r, 2).Value = labels(rule)
        ws.Cells(r, 3).Value = IIf(mCounts.Exists(rule), mCounts(rule), 0)
        ws.Cells(r, 4).Value = IIf(mPoints.Exists(rule), Application.WorksheetFunction.Min(30, mPoints(rule)), 0)
        total = total + ws.Cells(r, 4).Value
        If ws.Cells(r, 4).Value > 0 Then ws.Cells(r, 4).Interior.Color = RGB(252, 226, 226)
        r = r + 1
    Next rule

    score = Application.WorksheetFunction.Max(0, 100 - total)
    r = r + 1
    ws.Cells(r, 2).Value = "TOTAL POINTS DEDUITS": ws.Cells(r, 4).Value = total
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "NOTE GL (sur 100)": ws.Cells(r, 4).Value = score
    ws.Cells(r, 4).Font.Size = 16: ws.Cells(r, 4).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "Appreciation"
    If score >= 90 Then
        ws.Cells(r, 4).Value = "SATISFAISANT": ws.Cells(r, 4).Interior.Color = RGB(200, 255, 200)
    ElseIf score >= 75 Then
        ws.Cells(r, 4).Value = "ACCEPTABLE": ws.Cells(r, 4).Interior.Color = RGB(255, 245, 190)
    ElseIf score >= 60 Then
        ws.Cells(r, 4).Value = "A AMELIORER": ws.Cells(r, 4).Interior.Color = RGB(255, 220, 170)
    Else
        ws.Cells(r, 4).Value = "INSATISFAISANT": ws.Cells(r, 4).Interior.Color = RGB(255, 170, 170)
    End If
    r = r + 2
    ws.Cells(r, 1).Value = "Les points sont deduits en une seule fois par le consolidateur (pas par agence) pour eviter les erreurs de moyenne. Les seuils d'appreciation sont indicatifs et parametrables."
    ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Size = 9
    ws.Columns("A:D").AutoFit
    ws.Columns("A").ColumnWidth = 12
End Sub

' ==============================================================================
' HELPERS
' ==============================================================================

Private Sub Ecrire(rule As String, acct As String, name As String, cls As AccountClass, bal As Double, ageMax As Double, _
                   constat As String, grav As String, action As String, pts As Double)
    With mWsOut
        .Cells(mRow, 1).Value = rule
        .Cells(mRow, 2).Value = IIf(acct = "-", "-", "'" & acct)
        .Cells(mRow, 3).Value = name
        .Cells(mRow, 4).Value = cls.Family
        .Cells(mRow, 5).Value = cls.Nature
        .Cells(mRow, 6).Value = bal
        .Cells(mRow, 7).Value = IIf(ageMax > 0, ageMax, "")
        .Cells(mRow, 8).Value = constat
        .Cells(mRow, 9).Value = grav
        .Cells(mRow, 10).Value = action
        .Cells(mRow, 11).Value = pts
        Select Case grav
            Case "CRITICAL": .Cells(mRow, 9).Interior.Color = RGB(255, 0, 0): .Cells(mRow, 9).Font.Color = vbWhite
            Case "HIGH": .Cells(mRow, 9).Interior.Color = RGB(255, 165, 0)
            Case "MEDIUM": .Cells(mRow, 9).Interior.Color = RGB(255, 255, 0)
            Case "LOW": .Cells(mRow, 9).Interior.Color = RGB(200, 255, 200)
            Case Else: .Cells(mRow, 9).Interior.Color = RGB(220, 230, 240)
        End Select
    End With
    mRow = mRow + 1

    If grav <> "INFO" Then
        mCounts(rule) = IIf(mCounts.Exists(rule), mCounts(rule), 0) + 1
        mPoints(rule) = IIf(mPoints.Exists(rule), mPoints(rule), 0) + pts
    End If

    ' Alertes HIGH / CRITICAL dans AUDIT_REPORT (layout canonique)
    If (grav = "HIGH" Or grav = "CRITICAL") And Not mWsAudit Is Nothing Then
        With mWsAudit
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_REF).Value = rule
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_CATEGORIE).Value = "GL_MONITORING"
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_RISQUE).Value = Left(constat, 120)
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_NIVEAU).Value = grav
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_COMPTE).Value = "'" & acct
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_DESCRIPTION).Value = action
            .Cells(mAuditRow, SAFA_Common.AUDIT_COL_VALEUR).Value = bal
            If grav = "CRITICAL" Then
                .Cells(mAuditRow, SAFA_Common.AUDIT_COL_NIVEAU).Interior.Color = RGB(255, 0, 0)
                .Cells(mAuditRow, SAFA_Common.AUDIT_COL_NIVEAU).Font.Color = vbWhite
            Else
                .Cells(mAuditRow, SAFA_Common.AUDIT_COL_NIVEAU).Interior.Color = RGB(255, 165, 0)
            End If
        End With
        mAuditRow = mAuditRow + 1
    End If
End Sub

Private Sub LoadCfg()
    On Error Resume Next
    mCfg = Config_Manager.GetGLMonitoringConfig()
    On Error GoTo 0
    If mCfg.RatingTrancheDays <= 0 Then mCfg.RatingTrancheDays = 7
    If mCfg.RatingPointsPerTranche <= 0 Then mCfg.RatingPointsPerTranche = 2.5
    If mCfg.PrepaidRegularizationDays <= 0 Then mCfg.PrepaidRegularizationDays = 30
    If mCfg.OverAgedTier1 <= 0 Then mCfg.OverAgedTier1 = 90: mCfg.OverAgedTier2 = 180: mCfg.OverAgedTier3 = 360
    If mCfg.CurrencyPrefixes = "" Then mCfg.CurrencyPrefixes = "XAF|XOF|USD|EUR|GBP"
    If mCfg.KwTransit = "" Then mCfg.KwTransit = "TRANSIT"
    If mCfg.KwProxy = "" Then mCfg.KwProxy = "PROXY"
    If mCfg.KwSuspense = "" Then mCfg.KwSuspense = "SUSPENS|ATTENTE|UNCLAIMED|NON RECLAM|SUNDRY"
    If mCfg.KwPrepaid = "" Then mCfg.KwPrepaid = "PREPAID|D AVANCE|D'AVANCE"
    If mCfg.KwCash = "" Then mCfg.KwCash = "CAISSE|CASH|TILL|ATM|GAB"
    If mCfg.KwVault = "" Then mCfg.KwVault = "COFFRE|VAULT"
    If mCfg.KwDifference = "" Then mCfg.KwDifference = "ECART|OVERAGE|SHORTAGE|DIFFERENCE"
    If mCfg.PLPrefix = "" Then mCfg.PLPrefix = "PAL"
    If mCfg.PLClasses = "" Then mCfg.PLClasses = "6|7"
    If mCfg.NonProofableClasses = "" Then mCfg.NonProofableClasses = "8|9"
    If mCfg.SystemPrefixes = "" Then mCfg.SystemPrefixes = "IENC|RFI|FINNONE|POSITION FCY|POSITION LCY"
    If mCfg.InterbranchPrefix = "" Then mCfg.InterbranchPrefix = "INTERSOL"
    If mCfg.AutoUserIds = "" Then mCfg.AutoUserIds = "CDCI|SYSTEM|BATCH|AUTO"
    If mCfg.MgrChequeStaleDays <= 0 Then mCfg.MgrChequeStaleDays = 180
    If mCfg.WipStaleDays <= 0 Then mCfg.WipStaleDays = 180
    If mCfg.ExpenseMovementThreshold <= 0 Then mCfg.ExpenseMovementThreshold = 5000000
    If mCfg.UnusualBalanceFactor <= 1 Then mCfg.UnusualBalanceFactor = 3
    If mCfg.UnusualBalanceMin <= 0 Then mCfg.UnusualBalanceMin = 1000000
    If mCfg.RatingPointsOther <= 0 Then mCfg.RatingPointsOther = 2
    If mCfg.PrepaidGLCodes = "" Then mCfg.PrepaidGLCodes = "16090|16091"
    If mCfg.KwMgrCheque = "" Then mCfg.KwMgrCheque = "CHEQUE DE DIRECTION|MANAGER CHEQUE|MANAGERS CHEQUE|BANK DRAFT"
    If mCfg.KwWip = "" Then mCfg.KwWip = "WIP|TRAVAUX EN COURS|WORK IN PROGRESS|IMMOBILISATION EN COURS"
    If mCfg.KwExpenseNarration = "" Then mCfg.KwExpenseNarration = "LOYER|RENT|SALAIRE|SALARY|FACTURE|INVOICE|FRAIS|ENTRETIEN|MAINTENANCE|FOURNITURE|CARBURANT|ACHAT|HONORAIRE"
    If mCfg.KwFraudLoss = "" Then mCfg.KwFraudLoss = "FRAUDE|FRAUD|PERTE|LOSS|DETOURNEMENT"
    If mCfg.KwOverage = "" Then mCfg.KwOverage = "OVERAGE|EXCEDENT"
    If mCfg.KwShortage = "" Then mCfg.KwShortage = "SHORTAGE|MANQUANT|DEFICIT"
    mCfgLoaded = True
End Sub

Private Function BSDebitPositive() As Boolean
    ' Convention du bilan: detectee (Auto_Calibration) sinon debit positif
    BSDebitPositive = (SAFA_Common.g_DebitPositiveBS >= 0)
End Function

Private Function PLDebitPositive() As Boolean
    ' Convention du resultat: detectee sinon convention Finacle PAL (credit positif)
    If SAFA_Common.g_DebitPositivePL = 0 Then PLDebitPositive = False Else PLDebitPositive = (SAFA_Common.g_DebitPositivePL > 0)
End Function

Public Function ContainsAny(text As String, pipeList As String) As Boolean
    Dim t As Variant
    If pipeList = "" Then Exit Function
    For Each t In Split(pipeList, "|")
        If Len(Trim(t)) > 0 Then
            If InStr(1, text, UCase(Trim(t)), vbTextCompare) > 0 Then ContainsAny = True: Exit Function
        End If
    Next t
End Function

Private Function LoadKeyList(sheetName As String) As Object
    ' Dictionnaire des comptes (colonne A) d'une feuille optionnelle
    Dim d As Object, ws As Worksheet, lr As Long, i As Long, v As String, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not SAFA_Common.FeuilleExiste(sheetName) Then Set LoadKeyList = d: Exit Function
    Set ws = ThisWorkbook.Sheets(sheetName)
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lr
        v = SAFA_Common.SafeText(ws.Cells(i, 1).Value)
        If Len(v) > 5 And InStr(1, v, "ACCOUNT", vbTextCompare) = 0 And InStr(1, v, "COMPTE", vbTextCompare) = 0 Then
            k = SAFA_Common.NormalizeAccountKey(v)
            If Not d.Exists(k) Then d.Add k, i
        End If
    Next i
    Set LoadKeyList = d
End Function

Private Function LoadPreviousBalance() As Object
    ' Balance N-1 (feuille BALANCE_PREV_RAW, meme format que BALANCE_RAW): cle -> solde
    Dim d As Object, ws As Worksheet, rng As Range, hRow As Long, cAcct As Long, cBal As Long
    Dim lr As Long, i As Long, acct As String, solId As String, k As String, term As Variant, j As Long, h As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not SAFA_Common.FeuilleExiste("BALANCE_PREV_RAW") Then Set LoadPreviousBalance = d: Exit Function
    Set ws = ThisWorkbook.Sheets("BALANCE_PREV_RAW")

    Set rng = Nothing
    For Each term In Array("Account Number", "Acct Num", "ACCOUNT", "Compte")
        Set rng = ws.Range("A1:Z50").Find(What:=CStr(term), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not rng Is Nothing Then Exit For
    Next term
    If rng Is Nothing Then Set LoadPreviousBalance = d: Exit Function
    hRow = rng.Row: cAcct = rng.Column

    cBal = 0
    For j = 1 To 30
        h = UCase(SAFA_Common.SafeText(ws.Cells(hRow, j).Value))
        If InStr(h, "CLOSING") > 0 Or InStr(h, "SOLDE") > 0 Or InStr(h, "BALANCE") > 0 Or InStr(h, "CLR_BAL") > 0 Then cBal = j: Exit For
    Next j
    If cBal = 0 Then Set LoadPreviousBalance = d: Exit Function

    solId = "799"
    On Error Resume Next
    solId = SAFA_Common.SafeText(ThisWorkbook.Sheets("PARAM").Range("F2").Value)
    On Error GoTo 0
    If solId = "" Then solId = "799"

    lr = ws.Cells(ws.Rows.Count, cAcct).End(xlUp).Row
    For i = hRow + 1 To lr
        acct = SAFA_Common.SafeText(ws.Cells(i, cAcct).Value)
        If Len(acct) > 5 And Left(acct, 1) <> "-" And InStr(1, acct, "Total", vbTextCompare) = 0 Then
            k = SAFA_Common.NormalizeAccountKey(Core_Engine.NormalizeBalanceAccount(acct, solId, "0"))
            If Not d.Exists(k) Then d.Add k, SAFA_Common.SafeVal(ws.Cells(i, cBal).Value)
        End If
    Next i
    Set LoadPreviousBalance = d
End Function
