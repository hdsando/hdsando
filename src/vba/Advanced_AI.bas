Attribute VB_Name = "Advanced_AI"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: ADVANCED_AI v10.0 (Enhanced)
' ==============================================================================
' Description: Intelligence Artificielle et Analyses Statistiques Avancées
' Fonctionnalités:
'   - Z-Score amélioré avec détection multi-niveau
'   - Analyse de vélocité avec tendances
'   - Clustering comportemental (K-means simplifié)
'   - Détection de dormance/réveil de comptes
'   - Scoring prédictif de risque
'   - Dashboard exécutif automatisé
' ==============================================================================

' --- CONSTANTES ---
Private Const Z_SCORE_THRESHOLD As Double = 2.5
Private Const Z_SCORE_CRITICAL As Double = 3.5
Private Const VELOCITY_THRESHOLD As Double = 0.2
Private Const MIN_TRANSACTIONS_FOR_STATS As Long = 5
Private Const DORMANCY_DAYS As Long = 90

' --- TYPES ---
Private Type AccountStats
    AccountNumber As String
    Mean As Double
    StdDev As Double
    TransactionCount As Long
    TotalVolume As Double
    LastActivity As Date
    IsDormant As Boolean
End Type

Private Type ClusterInfo
    ClusterID As Integer
    ClusterName As String
    MeanAmount As Double
    MeanFrequency As Double
    AccountCount As Long
End Type

' ==============================================================================
' 1. ANALYSE IA PAR COMPTE (Z-SCORE AMÉLIORÉ)
' ==============================================================================

Public Sub Lancer_IA_Par_Compte()
    Dim wsT As Worksheet, wsA As Worksheet
    Dim lr As Long, r As Long, i As Long
    Dim dictStats As Object
    Dim key As String, mnt As Double

    On Error GoTo IAError

    ' Vérification dépendances
    If Not Core_Engine.FeuilleExiste("TRANSACTION_DATA") Or _
       Not Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        Exit Sub
    End If

    Set wsT = ThisWorkbook.Worksheets("TRANSACTION_DATA")
    Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")

    If Application.CountA(wsT.Range("A:A")) < 2 Then Exit Sub

    r = wsA.Cells(wsA.Rows.count, 1).End(xlUp).Row + 1
    lr = wsT.Cells(wsT.Rows.count, 1).End(xlUp).Row

    ' Dictionnaires pour statistiques
    Set dictStats = CreateObject("Scripting.Dictionary")

    Dim dictSum As Object, dictSumSq As Object, dictCount As Object
    Dim dictMin As Object, dictMax As Object, dictLastDate As Object

    Set dictSum = CreateObject("Scripting.Dictionary")
    Set dictSumSq = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictMin = CreateObject("Scripting.Dictionary")
    Set dictMax = CreateObject("Scripting.Dictionary")
    Set dictLastDate = CreateObject("Scripting.Dictionary")

    ' ═══════════════════════════════════════════════════════════════
    ' PASS 1: Calcul des statistiques par compte
    ' ═══════════════════════════════════════════════════════════════
    For i = 2 To lr
        key = Core_Engine.SafeText(wsT.Cells(i, 1).Value)
        mnt = Core_Engine.SafeVal(wsT.Cells(i, 4).Value)

        Dim transDate As Variant
        transDate = wsT.Cells(i, 2).Value

        If key <> "" Then
            If Not dictSum.Exists(key) Then
                dictSum.Add key, mnt
                dictSumSq.Add key, mnt * mnt
                dictCount.Add key, 1
                dictMin.Add key, mnt
                dictMax.Add key, mnt
                If IsDate(transDate) Then dictLastDate.Add key, CDate(transDate) Else dictLastDate.Add key, Date
            Else
                dictSum(key) = dictSum(key) + mnt
                dictSumSq(key) = dictSumSq(key) + (mnt * mnt)
                dictCount(key) = dictCount(key) + 1
                If mnt < dictMin(key) Then dictMin(key) = mnt
                If mnt > dictMax(key) Then dictMax(key) = mnt
                If IsDate(transDate) Then
                    If CDate(transDate) > dictLastDate(key) Then dictLastDate(key) = CDate(transDate)
                End If
            End If
        End If
    Next i

    ' ═══════════════════════════════════════════════════════════════
    ' PASS 2: Calcul Z-Score et détection anomalies
    ' ═══════════════════════════════════════════════════════════════
    For i = 2 To lr
        key = Core_Engine.SafeText(wsT.Cells(i, 1).Value)

        If key <> "" And dictCount.Exists(key) Then
            If dictCount(key) >= MIN_TRANSACTIONS_FOR_STATS Then
                Dim mean As Double, variance As Double, stdDev As Double, z As Double
                Dim n As Long

                n = dictCount(key)
                mean = dictSum(key) / n
                variance = (dictSumSq(key) / n) - (mean * mean)

                If variance > 0 Then
                    stdDev = Sqr(variance)
                    mnt = Core_Engine.SafeVal(wsT.Cells(i, 4).Value)
                    z = (mnt - mean) / stdDev

                    ' Seuils de détection multi-niveau
                    If Abs(z) > Z_SCORE_CRITICAL And Abs(mnt) > 1000000 Then
                        Call AddAIAlert(wsA, r, "IA-001", "Z-SCORE CRITIQUE", _
                            "Transaction hautement atypique (Z=" & Format(z, "0.00") & ")", _
                            "CRITICAL", key, _
                            "Moy=" & Format(mean, "#,##0") & " | Ecart-type=" & Format(stdDev, "#,##0"), _
                            mnt)

                    ElseIf Abs(z) > Z_SCORE_THRESHOLD And Abs(mnt) > 100000 Then
                        Call AddAIAlert(wsA, r, "IA-002", "Z-SCORE", _
                            "Transaction atypique (Z=" & Format(z, "0.00") & ")", _
                            "HIGH", key, _
                            "Moy=" & Format(mean, "#,##0") & " | n=" & n, _
                            mnt)
                    End If

                    ' Détection Min/Max historiques
                    If mnt = dictMax(key) And Abs(mnt) > 500000 Then
                        Call AddAIAlert(wsA, r, "IA-003", "RECORD", _
                            "Montant maximum historique", _
                            "MEDIUM", key, _
                            "Précédent max: " & Format(dictMax(key), "#,##0"), _
                            mnt)
                    End If
                End If
            End If
        End If
    Next i

    ' ═══════════════════════════════════════════════════════════════
    ' PASS 3: Détection comptes dormants qui se réveillent
    ' ═══════════════════════════════════════════════════════════════
    Dim vKey As Variant
    For Each vKey In dictLastDate.Keys
        If dictLastDate(vKey) > Date - 30 Then ' Activité récente
            ' Vérifier s'il y avait une période de dormance avant
            ' (Logique simplifiée - en production, analyser l'historique complet)
            If dictCount(vKey) <= 3 And Abs(dictSum(vKey)) > 1000000 Then
                Call AddAIAlert(wsA, r, "IA-004", "DORMANCY", _
                    "Compte avec peu d'historique et gros volumes", _
                    "HIGH", CStr(vKey), _
                    "Transactions: " & dictCount(vKey) & " | Dernière: " & Format(dictLastDate(vKey), "dd/mm/yyyy"), _
                    dictSum(vKey))
            End If
        End If
    Next vKey

    wsA.Columns("A:I").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Lancer_IA_Par_Compte terminé")
    Exit Sub

IAError:
    Call Core_Engine.LogError("Advanced_AI", "Lancer_IA_Par_Compte", Err.Number, Err.Description)
End Sub

' ==============================================================================
' 2. GESTION HISTORIQUE ET VÉLOCITÉ
' ==============================================================================

Public Sub Gerer_Historique()
    Dim wsH As Worksheet, wsR As Worksheet, wsA As Worksheet
    Dim lr As Long, i As Long, r As Long, histRow As Long
    Dim dictOldEcart As Object
    Dim acct As String, newEcart As Double, oldEcart As Double
    Dim velocity As Double

    On Error GoTo HistError

    If Not Core_Engine.FeuilleExiste("RECONCIL") Or _
       Not Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        Exit Sub
    End If

    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")

    ' Créer ou récupérer HISTORY_LOG
    On Error Resume Next
    Set wsH = ThisWorkbook.Worksheets("HISTORY_LOG")
    If wsH Is Nothing Then
        Set wsH = ThisWorkbook.Worksheets.Add
        wsH.name = "HISTORY_LOG"
        wsH.Range("A1:F1").Value = Array("Date", "Compte", "Ecart", "Tendance", "Score Risque", "Commentaire")
        wsH.Range("A1:F1").Font.Bold = True
    End If
    On Error GoTo HistError

    ' Charger historique précédent dans dictionnaire
    Set dictOldEcart = CreateObject("Scripting.Dictionary")

    lr = wsH.Cells(wsH.Rows.count, 1).End(xlUp).Row
    If lr > 1 Then
        ' Récupérer la dernière valeur pour chaque compte
        For i = lr To 2 Step -1
            acct = Core_Engine.SafeText(wsH.Cells(i, 2).Value)
            If acct <> "" And Not dictOldEcart.Exists(acct) Then
                dictOldEcart.Add acct, Core_Engine.SafeVal(wsH.Cells(i, 3).Value)
            End If
        Next i
    End If

    histRow = lr + 1
    r = wsA.Cells(wsA.Rows.count, 1).End(xlUp).Row + 1

    ' Analyser écarts actuels et comparer avec historique
    For i = 2 To wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
        acct = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        newEcart = Core_Engine.SafeVal(wsR.Cells(i, 5).Value)

        ' Enregistrer dans historique
        wsH.Cells(histRow, 1).Value = Date
        wsH.Cells(histRow, 2).Value = "'" & acct
        wsH.Cells(histRow, 3).Value = newEcart

        ' Calcul vélocité si historique existe
        If dictOldEcart.Exists(acct) Then
            oldEcart = dictOldEcart(acct)

            If Abs(oldEcart) > 0 Then
                velocity = (Abs(newEcart) - Abs(oldEcart)) / Abs(oldEcart)

                ' Déterminer tendance
                If velocity > 0.5 Then
                    wsH.Cells(histRow, 4).Value = "EXPLOSION ↑↑"
                    wsH.Cells(histRow, 4).Interior.Color = RGB(255, 0, 0)
                ElseIf velocity > VELOCITY_THRESHOLD Then
                    wsH.Cells(histRow, 4).Value = "HAUSSE ↑"
                    wsH.Cells(histRow, 4).Interior.Color = RGB(255, 165, 0)
                ElseIf velocity < -0.5 Then
                    wsH.Cells(histRow, 4).Value = "CHUTE ↓↓"
                    wsH.Cells(histRow, 4).Interior.Color = RGB(0, 176, 80)
                ElseIf velocity < -VELOCITY_THRESHOLD Then
                    wsH.Cells(histRow, 4).Value = "BAISSE ↓"
                    wsH.Cells(histRow, 4).Interior.Color = RGB(144, 238, 144)
                Else
                    wsH.Cells(histRow, 4).Value = "STABLE →"
                End If

                ' Alertes sur vélocité anormale
                If velocity > 0.5 And Abs(newEcart) > 100000 Then
                    Call AddAIAlert(wsA, r, "VEL-001", "VELOCITY CRITIQUE", _
                        "Explosion écart +" & Format(velocity, "0%"), _
                        "CRITICAL", acct, _
                        "Ancien: " & Format(oldEcart, "#,##0") & " | Nouveau: " & Format(newEcart, "#,##0"), _
                        newEcart)
                ElseIf velocity > VELOCITY_THRESHOLD And Abs(newEcart) > 100000 Then
                    Call AddAIAlert(wsA, r, "VEL-002", "VELOCITY", _
                        "Hausse significative écart +" & Format(velocity, "0%"), _
                        "HIGH", acct, _
                        "Tendance haussière détectée", _
                        newEcart)
                End If
            End If
        Else
            wsH.Cells(histRow, 4).Value = "NOUVEAU"
            wsH.Cells(histRow, 4).Interior.Color = RGB(173, 216, 230)
        End If

        wsH.Cells(histRow, 5).Value = Core_Engine.SafeVal(wsR.Cells(i, 12).Value) ' Risk Score
        histRow = histRow + 1
    Next i

    wsH.Visible = xlSheetHidden
    wsH.Columns("A:F").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Gerer_Historique terminé", "Entrées: " & (histRow - lr - 1))
    Exit Sub

HistError:
    Call Core_Engine.LogError("Advanced_AI", "Gerer_Historique", Err.Number, Err.Description)
End Sub

' ==============================================================================
' 3. CLUSTERING COMPORTEMENTAL
' ==============================================================================

Public Sub Analyser_Comportement_Comptes()
    Dim wsR As Worksheet, wsCluster As Worksheet
    Dim lr As Long, i As Long, outRow As Long
    Dim dictVolume As Object, dictTrans As Object
    Dim clusters(1 To 4) As ClusterInfo
    Dim outputArray() As Variant

    On Error GoTo ClusterError

    If Not Core_Engine.FeuilleExiste("RECONCIL") Then Exit Sub

    Set wsR = ThisWorkbook.Worksheets("RECONCIL")

    ' Créer feuille clusters
    Set wsCluster = Core_Engine.GetOrCreateSheet("CLUSTER_ANALYSIS", True)

    ' Initialiser clusters
    clusters(1).ClusterID = 1: clusters(1).ClusterName = "DORMANT"
    clusters(2).ClusterID = 2: clusters(2).ClusterName = "FAIBLE ACTIVITE"
    clusters(3).ClusterID = 3: clusters(3).ClusterName = "ACTIVITE NORMALE"
    clusters(4).ClusterID = 4: clusters(4).ClusterName = "HAUTE ACTIVITE"

    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row
    ReDim outputArray(1 To lr, 1 To 6)
    outRow = 0

    ' Classifier chaque compte
    For i = 2 To lr
        Dim acct As String, volume As Double, nbTrans As Long
        Dim clusterID As Integer

        acct = Core_Engine.SafeText(wsR.Cells(i, 1).Value)
        volume = Abs(Core_Engine.SafeVal(wsR.Cells(i, 5).Value)) ' Écart comme proxy volume
        nbTrans = Core_Engine.SafeVal(wsR.Cells(i, 7).Value)

        ' Classification simple basée sur volume et transactions
        If nbTrans = 0 Then
            clusterID = 1 ' Dormant
        ElseIf nbTrans <= 5 And volume < 100000 Then
            clusterID = 2 ' Faible activité
        ElseIf nbTrans <= 50 Or volume < 1000000 Then
            clusterID = 3 ' Normal
        Else
            clusterID = 4 ' Haute activité
        End If

        outRow = outRow + 1
        outputArray(outRow, 1) = acct
        outputArray(outRow, 2) = Core_Engine.SafeText(wsR.Cells(i, 2).Value)
        outputArray(outRow, 3) = clusters(clusterID).ClusterName
        outputArray(outRow, 4) = nbTrans
        outputArray(outRow, 5) = volume
        outputArray(outRow, 6) = Core_Engine.SafeVal(wsR.Cells(i, 12).Value)

        clusters(clusterID).AccountCount = clusters(clusterID).AccountCount + 1
    Next i

    ' En-têtes
    wsCluster.Range("A1:F1").Value = Array("Compte", "Libellé", "Cluster", "Nb Trans", "Volume", "Risk Score")
    FormatHeader wsCluster.Range("A1:F1")

    ' Données
    If outRow > 0 Then
        wsCluster.Range("A2").Resize(outRow, 6).Value = outputArray
    End If

    ' Résumé des clusters
    wsCluster.Range("H1").Value = "RÉSUMÉ CLUSTERS"
    wsCluster.Range("H1").Font.Bold = True
    wsCluster.Range("H2:J2").Value = Array("Cluster", "Nb Comptes", "% Total")

    For i = 1 To 4
        wsCluster.Cells(i + 2, 8).Value = clusters(i).ClusterName
        wsCluster.Cells(i + 2, 9).Value = clusters(i).AccountCount
        wsCluster.Cells(i + 2, 10).Value = clusters(i).AccountCount / outRow
    Next i

    wsCluster.Range("J3:J6").NumberFormat = "0.0%"
    wsCluster.Columns("A:J").AutoFit

    Call Core_Engine.WriteToAuditLog("PROCESS", "Analyser_Comportement_Comptes terminé")
    Exit Sub

ClusterError:
    Call Core_Engine.LogError("Advanced_AI", "Analyser_Comportement_Comptes", Err.Number, Err.Description)
End Sub

' ==============================================================================
' 4. FINALISATION RAPPORT ET DASHBOARD
' ==============================================================================

Public Sub Finaliser_Rapport()
    Dim wsR As Worksheet, wsD As Worksheet, wsS As Worksheet, wsExec As Worksheet
    Dim pc As PivotCache, pt As PivotTable
    Dim lr As Long, i As Long, cnt As Integer

    On Error GoTo RapportError

    If Not Core_Engine.FeuilleExiste("RECONCIL") Then Exit Sub

    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    lr = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    ' avec Report_Generator qui crée les mêmes feuilles
    Set wsD = Core_Engine.GetOrCreateSheet("DASHBOARD_RISQUE", True)
    Set wsS = Core_Engine.GetOrCreateSheet("ECHANTILLON_TEST", True)
    Set wsExec = Core_Engine.GetOrCreateSheet("EXECUTIVE_SUMMARY", True)

    ' ═══════════════════════════════════════════════════════════════
    ' 1. CALCUL PROVISIONS IFRS 9
    ' ═══════════════════════════════════════════════════════════════
    If lr > 1 Then
        ' Ajouter colonne provision si pas présente
        If wsR.Cells(1, 15).Value = "" Then
            wsR.Cells(1, 15).Value = "Prov. IFRS9"
        End If

        For i = 2 To lr
            Dim age As Double, ecart As Double, provision As Double

            age = Core_Engine.SafeVal(wsR.Cells(i, 8).Value)
            ecart = Core_Engine.SafeVal(wsR.Cells(i, 5).Value)

            ' Buckets IFRS 9
            If age > 360 Then
                provision = Abs(ecart) ' Stage 3 - 100%
            ElseIf age > 180 Then
                provision = Abs(ecart) * 0.5 ' Stage 2 - 50%
            ElseIf age > 90 Then
                provision = Abs(ecart) * 0.25 ' Stage 2 - 25%
            ElseIf age > 30 Then
                provision = Abs(ecart) * 0.1 ' Stage 1 - 10%
            Else
                provision = Abs(ecart) * 0.01 ' Stage 1 - 1%
            End If

            wsR.Cells(i, 15).Value = provision
        Next i
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' 2. DASHBOARD AVEC TCD
    ' ═══════════════════════════════════════════════════════════════
    Call CreateDashboard(wsD, wsR, lr)

    ' ═══════════════════════════════════════════════════════════════
    ' 3. ÉCHANTILLON DE TEST
    ' ═══════════════════════════════════════════════════════════════
    Call CreateTestSample(wsS, wsR, lr)

    ' ═══════════════════════════════════════════════════════════════
    ' 4. EXECUTIVE SUMMARY
    ' ═══════════════════════════════════════════════════════════════
    Call CreateExecutiveSummary(wsExec)

    Call Core_Engine.WriteToAuditLog("PROCESS", "Finaliser_Rapport terminé")
    Exit Sub

RapportError:
    Call Core_Engine.LogError("Advanced_AI", "Finaliser_Rapport", Err.Number, Err.Description)
End Sub

Private Sub CreateDashboard(wsD As Worksheet, wsR As Worksheet, lr As Long)
    Dim pc As PivotCache, pt As PivotTable

    On Error GoTo DashError

    ' Titre
    wsD.Range("B2").Value = "TABLEAU DE BORD RISQUES AUDIT"
    wsD.Range("B2").Font.Size = 18
    wsD.Range("B2").Font.Bold = True
    wsD.Range("B2").Font.Color = RGB(0, 51, 102)

    wsD.Range("B3").Value = "Généré le: " & Format(Now, "dd/mm/yyyy hh:nn")
    wsD.Range("B3").Font.Italic = True

    ' KPIs
    Dim nbTotal As Long, nbEcarts As Long, nbCritical As Long
    Dim totalEcart As Double, totalProvision As Double

    nbTotal = lr - 1
    nbEcarts = Application.CountIf(wsR.Range("F:F"), "Ecart à analyser")
    nbCritical = Application.CountIf(wsR.Range("N:N"), "CRITICAL")
    totalEcart = Application.SumIf(wsR.Range("F:F"), "Ecart à analyser", wsR.Range("E:E"))
    totalProvision = Application.WorksheetFunction.Sum(wsR.Range("O:O"))

    ' Zone KPIs
    wsD.Range("B5").Value = "INDICATEURS CLÉS"
    wsD.Range("B5").Font.Bold = True
    wsD.Range("B5:E5").Interior.Color = RGB(0, 51, 102)
    wsD.Range("B5:E5").Font.Color = vbWhite

    wsD.Range("B6:E6").Value = Array("Comptes analysés", "Écarts détectés", "Risques CRITIQUES", "Provision IFRS9")
    wsD.Range("B7:E7").Value = Array(nbTotal, nbEcarts, nbCritical, totalProvision)
    wsD.Range("B7:E7").Font.Size = 16
    wsD.Range("B7:E7").Font.Bold = True
    wsD.Range("E7").NumberFormat = "#,##0"

    ' Formatage conditionnel KPIs
    If nbCritical > 0 Then
        wsD.Range("D7").Interior.Color = RGB(255, 0, 0)
        wsD.Range("D7").Font.Color = vbWhite
    End If

    ' TCD si données suffisantes
    If lr > 2 Then
        Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, wsR.Range("A1:O" & lr))
        Set pt = pc.CreatePivotTable(wsD.Range("B10"), "TCD_Risques")

        With pt
            .PivotFields("Statut").Orientation = xlPageField
            On Error Resume Next
            .PivotFields("Statut").CurrentPage = "Ecart à analyser"
            On Error GoTo DashError

            .PivotFields("Priority").Orientation = xlRowField
            .PivotFields("Type").Orientation = xlRowField

            .AddDataField .PivotFields("Ecart"), "Total Écarts", xlSum
            .AddDataField .PivotFields("AccountNumber"), "Nb Comptes", xlCount

            .DataFields("Total Écarts").NumberFormat = "#,##0"

            .PivotFields("Priority").AutoSort xlDescending, "Total Écarts"
        End With
    End If

    wsD.Columns("B:F").AutoFit
    Exit Sub

DashError:
    ' Continuer même si TCD échoue
    Resume Next
End Sub

Private Sub CreateTestSample(wsS As Worksheet, wsR As Worksheet, lr As Long)
    Dim i As Long, sampleRow As Long

    ' En-têtes
    wsS.Range("A1:G1").Value = Array("Priorité", "Compte", "Libellé", "Écart", "Age Max", "Risk Score", "Action Requise")
    FormatHeader wsS.Range("A1:G1")

    ' Trier par Risk Score décroissant
    wsR.Range("A1:O" & lr).Sort Key1:=wsR.Range("L2"), Order1:=xlDescending, Header:=xlYes

    ' Sélectionner top 20
    sampleRow = 2
    For i = 2 To lr
        If wsR.Cells(i, 6).Value = "Ecart à analyser" And sampleRow <= 21 Then
            wsS.Cells(sampleRow, 1).Value = wsR.Cells(i, 14).Value ' Priority
            wsS.Cells(sampleRow, 2).Value = wsR.Cells(i, 1).Value
            wsS.Cells(sampleRow, 3).Value = wsR.Cells(i, 2).Value
            wsS.Cells(sampleRow, 4).Value = wsR.Cells(i, 5).Value
            wsS.Cells(sampleRow, 5).Value = wsR.Cells(i, 8).Value
            wsS.Cells(sampleRow, 6).Value = wsR.Cells(i, 12).Value

            ' Recommandation automatique
            Select Case wsR.Cells(i, 14).Value
                Case "CRITICAL"
                    wsS.Cells(sampleRow, 7).Value = "Investigation immédiate requise"
                    wsS.Cells(sampleRow, 1).Interior.Color = RGB(255, 0, 0)
                    wsS.Cells(sampleRow, 1).Font.Color = vbWhite
                Case "HIGH"
                    wsS.Cells(sampleRow, 7).Value = "Analyser sous 48h"
                    wsS.Cells(sampleRow, 1).Interior.Color = RGB(255, 165, 0)
                Case "MEDIUM"
                    wsS.Cells(sampleRow, 7).Value = "Analyser cette semaine"
                    wsS.Cells(sampleRow, 1).Interior.Color = RGB(255, 255, 0)
                Case Else
                    wsS.Cells(sampleRow, 7).Value = "Suivi standard"
            End Select

            sampleRow = sampleRow + 1
        End If
    Next i

    wsS.Range("D2:D" & sampleRow).NumberFormat = "#,##0"
    wsS.Columns("A:G").AutoFit
End Sub

Private Sub CreateExecutiveSummary(wsExec As Worksheet)
    Dim wsA As Worksheet, wsR As Worksheet

    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets("AUDIT_REPORT")
    Set wsR = ThisWorkbook.Worksheets("RECONCIL")
    On Error GoTo 0

    ' Titre
    wsExec.Range("B2").Value = "SYNTHÈSE EXÉCUTIVE - AUDIT S.A.F.A"
    wsExec.Range("B2").Font.Size = 20
    wsExec.Range("B2").Font.Bold = True

    wsExec.Range("B3").Value = "Date: " & Format(Date, "dd mmmm yyyy")
    wsExec.Range("B4").Value = "Analyste: " & Environ("USERNAME")

    ' Section 1: Vue d'ensemble
    wsExec.Range("B6").Value = "1. VUE D'ENSEMBLE"
    wsExec.Range("B6").Font.Bold = True
    wsExec.Range("B6").Font.Size = 14

    Dim nbComptes As Long, nbAlertes As Long
    If Not wsR Is Nothing Then nbComptes = wsR.Cells(wsR.Rows.count, 1).End(xlUp).Row - 1
    If Not wsA Is Nothing Then nbAlertes = wsA.Cells(wsA.Rows.count, 1).End(xlUp).Row - 1

    wsExec.Range("B7").Value = "Comptes analysés: " & nbComptes
    wsExec.Range("B8").Value = "Alertes générées: " & nbAlertes

    ' Section 2: Risques majeurs
    wsExec.Range("B10").Value = "2. RISQUES MAJEURS"
    wsExec.Range("B10").Font.Bold = True
    wsExec.Range("B10").Font.Size = 14

    If Not wsA Is Nothing Then
        Dim criticalCount As Long, fraudCount As Long
        criticalCount = Application.CountIf(wsA.Range("D:D"), "CRITICAL")
        fraudCount = Application.CountIf(wsA.Range("D:D"), "FRAUD")

        wsExec.Range("B11").Value = "• Alertes CRITIQUES: " & criticalCount
        wsExec.Range("B12").Value = "• Suspicions de FRAUDE: " & fraudCount

        If criticalCount > 0 Or fraudCount > 0 Then
            wsExec.Range("B11:B12").Font.Color = RGB(255, 0, 0)
        End If
    End If

    ' Section 3: Recommandations
    wsExec.Range("B14").Value = "3. RECOMMANDATIONS"
    wsExec.Range("B14").Font.Bold = True
    wsExec.Range("B14").Font.Size = 14

    wsExec.Range("B15").Value = "• Prioriser l'investigation des comptes marqués CRITICAL"
    wsExec.Range("B16").Value = "• Vérifier la conformité Benford si anomalie détectée"
    wsExec.Range("B17").Value = "• Documenter toutes les régularisations"
    wsExec.Range("B18").Value = "• Planifier revue des comptes de passage"

    wsExec.Columns("B").AutoFit
End Sub

' ==============================================================================
' 5. UTILITAIRES
' ==============================================================================

Private Sub AddAIAlert(ByVal ws As Worksheet, ByRef r As Long, _
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

    ' Formatage
    Select Case UCase(severity)
        Case "CRITICAL"
            ws.Cells(r, 4).Interior.Color = RGB(255, 0, 0)
            ws.Cells(r, 4).Font.Color = vbWhite
        Case "HIGH"
            ws.Cells(r, 4).Interior.Color = RGB(255, 165, 0)
        Case "MEDIUM"
            ws.Cells(r, 4).Interior.Color = RGB(255, 255, 0)
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
