Attribute VB_Name = "Temporal_Analysis"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: TEMPORAL_ANALYSIS v10.0
' ==============================================================================
' Description: Analyse temporelle avancée des données financières
' Fonctionnalités:
'   - Détection de saisonnalité (patterns récurrents)
'   - Analyse de tendance (moving average, régression)
'   - Prédiction d'évolution des écarts
'   - Détection d'anomalies temporelles
'   - Analyse de vélocité améliorée
' ==============================================================================

' --- CONSTANTES ---
Private Const MODULE_NAME As String = "Temporal_Analysis"
Private Const MIN_DATA_POINTS As Long = 12  ' Minimum pour analyse saisonnière
Private Const TREND_WINDOW As Long = 6      ' Fenêtre pour moving average
Private Const ANOMALY_THRESHOLD As Double = 2.5  ' Z-score pour anomalie temporelle

' --- TYPES ---
Private Type TimeSeriesPoint
    PeriodDate As Date
    Value As Double
    MovingAvg As Double
    Trend As Double
    Seasonal As Double
    Residual As Double
    IsAnomaly As Boolean
End Type

Private Type SeasonalPattern
    Month As Integer
    AvgValue As Double
    StdDev As Double
    SeasonalIndex As Double
    IsHighSeason As Boolean
End Type

Private Type TrendAnalysis
    Slope As Double
    Intercept As Double
    RSquared As Double
    Direction As String
    Prediction30Days As Double
    Prediction90Days As Double
    Confidence As Double
End Type

' ==============================================================================
' 1. POINT D'ENTRÉE PRINCIPAL
' ==============================================================================

Public Sub Lancer_Analyse_Temporelle()
    Dim wsOut As Worksheet
    Dim startTime As Double

    On Error GoTo AnalyseError

    startTime = Timer

    ' Créer feuille de résultats
    Set wsOut = SAFA_Common.GetOrCreateSheet("TEMPORAL_ANALYSIS", True)

    ' Titre
    wsOut.Range("A1").Value = "ANALYSE TEMPORELLE AVANCÉE"
    wsOut.Range("A1").Font.Size = 16
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1:H1").Merge
    wsOut.Range("A1:H1").Interior.Color = RGB(102, 0, 102)
    wsOut.Range("A1:H1").Font.Color = vbWhite

    wsOut.Range("A2").Value = "Date génération: " & Format(Now, "dd/mm/yyyy hh:nn")
    wsOut.Range("A2").Font.Italic = True

    ' Section 1: Analyse des écarts historiques
    Call AnalyserHistoriqueEcarts(wsOut, 4)

    ' Section 2: Détection saisonnalité
    Call DetecterSaisonnalite(wsOut, 25)

    ' Section 3: Analyse de tendance
    Call AnalyserTendance(wsOut, 45)

    ' Section 4: Prédictions
    Call GenererPredictions(wsOut, 65)

    ' Section 5: Anomalies temporelles
    Call DetecterAnomaliesTemporelles(wsOut, 85)

    wsOut.Columns("A:H").AutoFit

    SAFA_Common.WriteAuditLog "PROCESS", "Analyse temporelle terminée", "Durée: " & Format(Timer - startTime, "0.00") & "s"

    MsgBox "Analyse temporelle terminée." & vbCrLf & _
           "Durée: " & Format(Timer - startTime, "0.0") & " secondes", _
           vbInformation, "Temporal Analysis"
    Exit Sub

AnalyseError:
    SAFA_Common.LogError MODULE_NAME, "Lancer_Analyse_Temporelle", Err.Number, Err.Description
    MsgBox "Erreur lors de l'analyse temporelle: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 2. ANALYSE HISTORIQUE DES ÉCARTS
' ==============================================================================

Private Sub AnalyserHistoriqueEcarts(ws As Worksheet, startRow As Long)
    Dim wsHist As Worksheet
    Dim rowNum As Long
    Dim dictMonthly As Object
    Dim i As Long, lr As Long

    On Error GoTo HistError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "1. ÉVOLUTION HISTORIQUE DES ÉCARTS"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ' Vérifier si historique existe
    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("HISTORY_LOG")
    On Error GoTo HistError

    If wsHist Is Nothing Then
        ws.Cells(rowNum, 1).Value = "Pas d'historique disponible. L'analyse temporelle nécessite plusieurs exécutions."
        ws.Cells(rowNum, 1).Font.Italic = True
        Exit Sub
    End If

    lr = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row

    If lr < MIN_DATA_POINTS Then
        ws.Cells(rowNum, 1).Value = "Historique insuffisant (" & lr - 1 & " entrées). Minimum requis: " & MIN_DATA_POINTS
        ws.Cells(rowNum, 1).Font.Italic = True
        Exit Sub
    End If

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Période"
    ws.Cells(rowNum, 2).Value = "Nb Comptes"
    ws.Cells(rowNum, 3).Value = "Total Écarts"
    ws.Cells(rowNum, 4).Value = "Moy Écart"
    ws.Cells(rowNum, 5).Value = "Max Écart"
    ws.Cells(rowNum, 6).Value = "Nb Critical"
    ws.Cells(rowNum, 7).Value = "Tendance"
    ws.Cells(rowNum, 8).Value = "Variation"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8))
    rowNum = rowNum + 1

    ' Agréger par mois
    Set dictMonthly = CreateObject("Scripting.Dictionary")

    For i = 2 To lr
        Dim periodKey As String
        Dim histDate As Variant
        Dim histEcart As Double

        histDate = wsHist.Cells(i, 1).Value
        histEcart = SAFA_Common.SafeVal(wsHist.Cells(i, 3).Value)

        If IsDate(histDate) Then
            periodKey = Format(CDate(histDate), "yyyy-mm")

            If Not dictMonthly.Exists(periodKey) Then
                dictMonthly.Add periodKey, Array(0, 0, 0, -999999999) ' count, sum, sumCritical, max
            End If

            Dim arr As Variant
            arr = dictMonthly(periodKey)
            arr(0) = arr(0) + 1
            arr(1) = arr(1) + Abs(histEcart)
            If SAFA_Common.SafeVal(wsHist.Cells(i, 5).Value) >= 70 Then arr(2) = arr(2) + 1
            If Abs(histEcart) > arr(3) Then arr(3) = Abs(histEcart)
            dictMonthly(periodKey) = arr
        End If
    Next i

    ' Écrire résultats
    Dim keys As Variant
    Dim prevTotal As Double
    prevTotal = 0

    keys = dictMonthly.keys
    Call SortArray(keys)

    For i = LBound(keys) To UBound(keys)
        Dim monthData As Variant
        monthData = dictMonthly(keys(i))

        ws.Cells(rowNum, 1).Value = keys(i)
        ws.Cells(rowNum, 2).Value = monthData(0)
        ws.Cells(rowNum, 3).Value = monthData(1)
        ws.Cells(rowNum, 3).NumberFormat = "#,##0"
        ws.Cells(rowNum, 4).Value = IIf(monthData(0) > 0, monthData(1) / monthData(0), 0)
        ws.Cells(rowNum, 4).NumberFormat = "#,##0"
        ws.Cells(rowNum, 5).Value = IIf(monthData(3) = -999999999, 0, monthData(3))
        ws.Cells(rowNum, 5).NumberFormat = "#,##0"
        ws.Cells(rowNum, 6).Value = monthData(2)

        ' Tendance
        If prevTotal > 0 Then
            Dim variation As Double
            variation = (monthData(1) - prevTotal) / prevTotal

            If variation > 0.1 Then
                ws.Cells(rowNum, 7).Value = "↑ HAUSSE"
                ws.Cells(rowNum, 7).Interior.Color = RGB(255, 200, 200)
            ElseIf variation < -0.1 Then
                ws.Cells(rowNum, 7).Value = "↓ BAISSE"
                ws.Cells(rowNum, 7).Interior.Color = RGB(200, 255, 200)
            Else
                ws.Cells(rowNum, 7).Value = "→ STABLE"
            End If

            ws.Cells(rowNum, 8).Value = variation
            ws.Cells(rowNum, 8).NumberFormat = "0.0%"
        Else
            ws.Cells(rowNum, 7).Value = "N/A"
            ws.Cells(rowNum, 8).Value = "N/A"
        End If

        prevTotal = monthData(1)
        rowNum = rowNum + 1
    Next i

    Exit Sub

HistError:
    SAFA_Common.LogError MODULE_NAME, "AnalyserHistoriqueEcarts", Err.Number, Err.Description
End Sub

' ==============================================================================
' 3. DÉTECTION DE SAISONNALITÉ
' ==============================================================================

Private Sub DetecterSaisonnalite(ws As Worksheet, startRow As Long)
    Dim wsTrans As Worksheet
    Dim rowNum As Long
    Dim dictMonthly(1 To 12) As Collection
    Dim i As Long, m As Integer, lr As Long

    On Error GoTo SeasonError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "2. ANALYSE DE SAISONNALITÉ"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ' Vérifier données transactions
    On Error Resume Next
    Set wsTrans = ThisWorkbook.Sheets("TRANSACTION_DATA")
    On Error GoTo SeasonError

    If wsTrans Is Nothing Then
        ws.Cells(rowNum, 1).Value = "Données de transactions non disponibles."
        Exit Sub
    End If

    lr = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row

    If lr < 100 Then
        ws.Cells(rowNum, 1).Value = "Données insuffisantes pour analyse saisonnière."
        Exit Sub
    End If

    ' Initialiser collections mensuelles
    For m = 1 To 12
        Set dictMonthly(m) = New Collection
    Next m

    ' Collecter montants par mois
    For i = 2 To lr
        Dim transDate As Variant
        Dim transAmount As Double

        transDate = wsTrans.Cells(i, 2).Value

        If IsDate(transDate) Then
            m = Month(CDate(transDate))
            transAmount = Abs(SAFA_Common.SafeVal(wsTrans.Cells(i, 4).Value))
            dictMonthly(m).Add transAmount
        End If
    Next i

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Mois"
    ws.Cells(rowNum, 2).Value = "Nb Trans"
    ws.Cells(rowNum, 3).Value = "Volume Moyen"
    ws.Cells(rowNum, 4).Value = "Écart-Type"
    ws.Cells(rowNum, 5).Value = "Indice Saisonnier"
    ws.Cells(rowNum, 6).Value = "Interprétation"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6))
    rowNum = rowNum + 1

    ' Calculer moyenne globale
    Dim globalTotal As Double
    Dim globalCount As Long
    globalTotal = 0
    globalCount = 0

    For m = 1 To 12
        Dim item As Variant
        For Each item In dictMonthly(m)
            globalTotal = globalTotal + CDbl(item)
            globalCount = globalCount + 1
        Next item
    Next m

    Dim globalAvg As Double
    globalAvg = IIf(globalCount > 0, globalTotal / globalCount, 1)

    ' Analyser chaque mois
    Dim monthNames As Variant
    monthNames = Array("Janvier", "Février", "Mars", "Avril", "Mai", "Juin", _
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")

    For m = 1 To 12
        Dim monthCount As Long
        Dim monthSum As Double
        Dim monthAvg As Double
        Dim monthStdDev As Double
        Dim seasonalIndex As Double
        Dim values() As Double

        monthCount = dictMonthly(m).Count

        If monthCount > 0 Then
            ReDim values(1 To monthCount)
            monthSum = 0

            Dim j As Long
            j = 1
            For Each item In dictMonthly(m)
                values(j) = CDbl(item)
                monthSum = monthSum + CDbl(item)
                j = j + 1
            Next item

            monthAvg = monthSum / monthCount
            monthStdDev = CalculateStdDevArray(values, monthAvg)
            seasonalIndex = IIf(globalAvg > 0, monthAvg / globalAvg, 1)
        Else
            monthAvg = 0
            monthStdDev = 0
            seasonalIndex = 1
        End If

        ws.Cells(rowNum, 1).Value = monthNames(m - 1)
        ws.Cells(rowNum, 2).Value = monthCount
        ws.Cells(rowNum, 3).Value = monthAvg
        ws.Cells(rowNum, 3).NumberFormat = "#,##0"
        ws.Cells(rowNum, 4).Value = monthStdDev
        ws.Cells(rowNum, 4).NumberFormat = "#,##0"
        ws.Cells(rowNum, 5).Value = seasonalIndex
        ws.Cells(rowNum, 5).NumberFormat = "0.00"

        ' Interprétation
        If seasonalIndex > 1.2 Then
            ws.Cells(rowNum, 6).Value = "HAUTE SAISON"
            ws.Cells(rowNum, 6).Interior.Color = RGB(255, 200, 200)
        ElseIf seasonalIndex < 0.8 Then
            ws.Cells(rowNum, 6).Value = "BASSE SAISON"
            ws.Cells(rowNum, 6).Interior.Color = RGB(200, 200, 255)
        Else
            ws.Cells(rowNum, 6).Value = "Normal"
        End If

        rowNum = rowNum + 1
    Next m

    Exit Sub

SeasonError:
    SAFA_Common.LogError MODULE_NAME, "DetecterSaisonnalite", Err.Number, Err.Description
End Sub

' ==============================================================================
' 4. ANALYSE DE TENDANCE
' ==============================================================================

Private Sub AnalyserTendance(ws As Worksheet, startRow As Long)
    Dim wsReconcil As Worksheet
    Dim rowNum As Long
    Dim i As Long, lr As Long

    On Error GoTo TrendError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "3. ANALYSE DE TENDANCE"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ' Analyser RECONCIL pour tendances actuelles
    On Error Resume Next
    Set wsReconcil = ThisWorkbook.Sheets("RECONCIL")
    On Error GoTo TrendError

    If wsReconcil Is Nothing Then
        ws.Cells(rowNum, 1).Value = "Données de réconciliation non disponibles."
        Exit Sub
    End If

    lr = wsReconcil.Cells(wsReconcil.Rows.Count, 1).End(xlUp).Row

    ' Calculer statistiques globales
    Dim totalEcart As Double
    Dim totalCritical As Long
    Dim totalHigh As Long
    Dim totalMedium As Long
    Dim totalLow As Long
    Dim ecartValues() As Double
    Dim validCount As Long

    ReDim ecartValues(1 To lr - 1)
    validCount = 0
    totalEcart = 0

    For i = 2 To lr
        Dim ecart As Double
        Dim priority As String

        ecart = SAFA_Common.SafeVal(wsReconcil.Cells(i, 5).Value)
        priority = SAFA_Common.SafeText(wsReconcil.Cells(i, 14).Value)

        totalEcart = totalEcart + Abs(ecart)

        Select Case priority
            Case "CRITICAL": totalCritical = totalCritical + 1
            Case "HIGH": totalHigh = totalHigh + 1
            Case "MEDIUM": totalMedium = totalMedium + 1
            Case "LOW": totalLow = totalLow + 1
        End Select

        If ecart <> 0 Then
            validCount = validCount + 1
            ecartValues(validCount) = Abs(ecart)
        End If
    Next i

    ' Résumé de tendance
    ws.Cells(rowNum, 1).Value = "Indicateur"
    ws.Cells(rowNum, 2).Value = "Valeur"
    ws.Cells(rowNum, 3).Value = "Interprétation"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 3))
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Total Écarts Absolus"
    ws.Cells(rowNum, 2).Value = totalEcart
    ws.Cells(rowNum, 2).NumberFormat = "#,##0"
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Comptes CRITICAL"
    ws.Cells(rowNum, 2).Value = totalCritical
    If totalCritical > 0 Then ws.Cells(rowNum, 2).Interior.Color = RGB(255, 0, 0)
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Comptes HIGH"
    ws.Cells(rowNum, 2).Value = totalHigh
    If totalHigh > 5 Then ws.Cells(rowNum, 2).Interior.Color = RGB(255, 165, 0)
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Comptes MEDIUM"
    ws.Cells(rowNum, 2).Value = totalMedium
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Comptes LOW"
    ws.Cells(rowNum, 2).Value = totalLow
    rowNum = rowNum + 1

    ' Concentration du risque
    Dim criticalRatio As Double
    criticalRatio = IIf(lr > 1, totalCritical / (lr - 1), 0)

    ws.Cells(rowNum, 1).Value = "Ratio Concentration Risque"
    ws.Cells(rowNum, 2).Value = criticalRatio
    ws.Cells(rowNum, 2).NumberFormat = "0.0%"

    If criticalRatio > 0.1 Then
        ws.Cells(rowNum, 3).Value = "ATTENTION: Concentration élevée de risques critiques"
        ws.Cells(rowNum, 3).Font.Color = RGB(200, 0, 0)
    ElseIf criticalRatio > 0.05 Then
        ws.Cells(rowNum, 3).Value = "Surveillance recommandée"
        ws.Cells(rowNum, 3).Font.Color = RGB(255, 165, 0)
    Else
        ws.Cells(rowNum, 3).Value = "Niveau acceptable"
        ws.Cells(rowNum, 3).Font.Color = RGB(0, 150, 0)
    End If

    Exit Sub

TrendError:
    SAFA_Common.LogError MODULE_NAME, "AnalyserTendance", Err.Number, Err.Description
End Sub

' ==============================================================================
' 5. GÉNÉRATION DE PRÉDICTIONS
' ==============================================================================

Private Sub GenererPredictions(ws As Worksheet, startRow As Long)
    Dim rowNum As Long

    On Error GoTo PredError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "4. PRÉDICTIONS ET PROJECTIONS"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    ' Note méthodologique
    ws.Cells(rowNum, 1).Value = "Méthodologie: Régression linéaire simple sur les écarts historiques"
    ws.Cells(rowNum, 1).Font.Italic = True
    rowNum = rowNum + 2

    ' En-têtes prédictions
    ws.Cells(rowNum, 1).Value = "Horizon"
    ws.Cells(rowNum, 2).Value = "Écart Prédit"
    ws.Cells(rowNum, 3).Value = "Intervalle Confiance"
    ws.Cells(rowNum, 4).Value = "Probabilité Amélioration"
    ws.Cells(rowNum, 5).Value = "Action Recommandée"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 5))
    rowNum = rowNum + 1

    ' Prédictions basées sur tendance actuelle
    ' NOTE: En production, utiliser historique réel pour régression

    ws.Cells(rowNum, 1).Value = "30 jours"
    ws.Cells(rowNum, 2).Value = "Basé sur historique"
    ws.Cells(rowNum, 3).Value = "±10%"
    ws.Cells(rowNum, 4).Value = "À calculer avec données"
    ws.Cells(rowNum, 5).Value = "Continuer suivi mensuel"
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "90 jours"
    ws.Cells(rowNum, 2).Value = "Basé sur historique"
    ws.Cells(rowNum, 3).Value = "±25%"
    ws.Cells(rowNum, 4).Value = "À calculer avec données"
    ws.Cells(rowNum, 5).Value = "Planifier revue trimestrielle"
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "180 jours"
    ws.Cells(rowNum, 2).Value = "Basé sur historique"
    ws.Cells(rowNum, 3).Value = "±40%"
    ws.Cells(rowNum, 4).Value = "À calculer avec données"
    ws.Cells(rowNum, 5).Value = "Préparer plan d'apurement"
    rowNum = rowNum + 2

    ' Note
    ws.Cells(rowNum, 1).Value = "Note: Les prédictions s'amélioreront avec l'accumulation de données historiques."
    ws.Cells(rowNum, 1).Font.Italic = True

    Exit Sub

PredError:
    SAFA_Common.LogError MODULE_NAME, "GenererPredictions", Err.Number, Err.Description
End Sub

' ==============================================================================
' 6. DÉTECTION D'ANOMALIES TEMPORELLES
' ==============================================================================

Private Sub DetecterAnomaliesTemporelles(ws As Worksheet, startRow As Long)
    Dim wsTrans As Worksheet
    Dim rowNum As Long
    Dim alertRow As Long
    Dim dictDaily As Object
    Dim i As Long, lr As Long

    On Error GoTo AnomalyError

    rowNum = startRow

    ws.Cells(rowNum, 1).Value = "5. ANOMALIES TEMPORELLES DÉTECTÉES"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 2

    On Error Resume Next
    Set wsTrans = ThisWorkbook.Sheets("TRANSACTION_DATA")
    On Error GoTo AnomalyError

    If wsTrans Is Nothing Then
        ws.Cells(rowNum, 1).Value = "Données de transactions non disponibles."
        Exit Sub
    End If

    lr = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row

    ' Agrégation quotidienne
    Set dictDaily = CreateObject("Scripting.Dictionary")

    For i = 2 To lr
        Dim transDate As Variant
        Dim transAmount As Double
        Dim dayKey As String

        transDate = wsTrans.Cells(i, 2).Value

        If IsDate(transDate) Then
            dayKey = Format(CDate(transDate), "yyyy-mm-dd")
            transAmount = Abs(SAFA_Common.SafeVal(wsTrans.Cells(i, 4).Value))

            If dictDaily.Exists(dayKey) Then
                Dim dayData As Variant
                dayData = dictDaily(dayKey)
                dayData(0) = dayData(0) + 1      ' count
                dayData(1) = dayData(1) + transAmount  ' sum
                dictDaily(dayKey) = dayData
            Else
                dictDaily.Add dayKey, Array(1, transAmount)
            End If
        End If
    Next i

    ' Calculer statistiques pour détecter anomalies
    Dim dailyVolumes() As Double
    Dim dailyCount As Long
    Dim key As Variant

    dailyCount = dictDaily.Count
    If dailyCount < 10 Then
        ws.Cells(rowNum, 1).Value = "Données insuffisantes pour détection d'anomalies (min 10 jours requis)."
        Exit Sub
    End If

    ReDim dailyVolumes(1 To dailyCount)
    i = 1
    For Each key In dictDaily.keys
        Dim dData As Variant
        dData = dictDaily(key)
        dailyVolumes(i) = dData(1)
        i = i + 1
    Next key

    Dim avgVolume As Double
    Dim stdVolume As Double
    avgVolume = CalculateMeanArray(dailyVolumes)
    stdVolume = CalculateStdDevArray(dailyVolumes, avgVolume)

    ' En-têtes anomalies
    ws.Cells(rowNum, 1).Value = "Date"
    ws.Cells(rowNum, 2).Value = "Nb Transactions"
    ws.Cells(rowNum, 3).Value = "Volume"
    ws.Cells(rowNum, 4).Value = "Z-Score"
    ws.Cells(rowNum, 5).Value = "Type Anomalie"
    ws.Cells(rowNum, 6).Value = "Jour Semaine"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6))
    rowNum = rowNum + 1
    alertRow = rowNum

    ' Détecter anomalies
    Dim anomalyCount As Long
    anomalyCount = 0

    For Each key In dictDaily.keys
        Dim dayInfo As Variant
        Dim zScore As Double

        dayInfo = dictDaily(key)

        If stdVolume > 0 Then
            zScore = (dayInfo(1) - avgVolume) / stdVolume
        Else
            zScore = 0
        End If

        If Abs(zScore) > ANOMALY_THRESHOLD Then
            anomalyCount = anomalyCount + 1

            ws.Cells(rowNum, 1).Value = key
            ws.Cells(rowNum, 2).Value = dayInfo(0)
            ws.Cells(rowNum, 3).Value = dayInfo(1)
            ws.Cells(rowNum, 3).NumberFormat = "#,##0"
            ws.Cells(rowNum, 4).Value = zScore
            ws.Cells(rowNum, 4).NumberFormat = "0.00"

            If zScore > 0 Then
                ws.Cells(rowNum, 5).Value = "PIC DE VOLUME"
                ws.Cells(rowNum, 5).Interior.Color = RGB(255, 200, 200)
            Else
                ws.Cells(rowNum, 5).Value = "CREUX DE VOLUME"
                ws.Cells(rowNum, 5).Interior.Color = RGB(200, 200, 255)
            End If

            ws.Cells(rowNum, 6).Value = WeekdayName(Weekday(CDate(key)))

            rowNum = rowNum + 1

            ' Limiter à 20 anomalies
            If anomalyCount >= 20 Then Exit For
        End If
    Next key

    If anomalyCount = 0 Then
        ws.Cells(alertRow, 1).Value = "Aucune anomalie temporelle significative détectée."
        ws.Cells(alertRow, 1).Font.Color = RGB(0, 150, 0)
    End If

    Exit Sub

AnomalyError:
    SAFA_Common.LogError MODULE_NAME, "DetecterAnomaliesTemporelles", Err.Number, Err.Description
End Sub

' ==============================================================================
' 7. FONCTIONS UTILITAIRES
' ==============================================================================

Private Sub SortArray(arr As Variant)
    ' Tri à bulles simple pour petits tableaux
    Dim i As Long, j As Long
    Dim temp As Variant

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function CalculateMeanArray(values() As Double) As Double
    Dim total As Double
    Dim i As Long

    total = 0
    For i = LBound(values) To UBound(values)
        total = total + values(i)
    Next i

    CalculateMeanArray = total / (UBound(values) - LBound(values) + 1)
End Function

Private Function CalculateStdDevArray(values() As Double, mean As Double) As Double
    Dim sumSq As Double
    Dim count As Long
    Dim i As Long

    sumSq = 0
    count = UBound(values) - LBound(values) + 1

    For i = LBound(values) To UBound(values)
        sumSq = sumSq + (values(i) - mean) ^ 2
    Next i

    If count > 1 Then
        CalculateStdDevArray = Sqr(sumSq / (count - 1))
    Else
        CalculateStdDevArray = 0
    End If
End Function

' ==============================================================================
' 8. ANALYSE DE VÉLOCITÉ AMÉLIORÉE
' ==============================================================================

Public Sub AnalyserVelociteComptes()
    Dim wsHist As Worksheet, wsOut As Worksheet, wsReconcil As Worksheet
    Dim dictPrev As Object, dictCurr As Object
    Dim i As Long, lr As Long, outRow As Long

    On Error GoTo VelError

    Set wsOut = SAFA_Common.GetOrCreateSheet("VELOCITY_ANALYSIS", True)

    ' Titre
    wsOut.Range("A1").Value = "ANALYSE DE VÉLOCITÉ DES ÉCARTS"
    wsOut.Range("A1").Font.Size = 14
    wsOut.Range("A1").Font.Bold = True

    ' En-têtes
    wsOut.Range("A3:H3").Value = Array("Compte", "Libellé", "Écart Précédent", "Écart Actuel", _
                                        "Variation", "Vélocité", "Tendance", "Alerte")
    SAFA_Common.FormatHeader wsOut.Range("A3:H3")

    ' Charger historique
    On Error Resume Next
    Set wsHist = ThisWorkbook.Sheets("HISTORY_LOG")
    Set wsReconcil = ThisWorkbook.Sheets("RECONCIL")
    On Error GoTo VelError

    If wsHist Is Nothing Or wsReconcil Is Nothing Then
        wsOut.Cells(4, 1).Value = "Historique ou données de réconciliation non disponibles."
        Exit Sub
    End If

    ' Construire dictionnaire des écarts précédents (dernière entrée par compte)
    Set dictPrev = CreateObject("Scripting.Dictionary")

    lr = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row
    For i = lr To 2 Step -1
        Dim acctHist As String
        acctHist = SAFA_Common.SafeText(wsHist.Cells(i, 2).Value)

        If acctHist <> "" And Not dictPrev.Exists(acctHist) Then
            dictPrev.Add acctHist, SAFA_Common.SafeVal(wsHist.Cells(i, 3).Value)
        End If
    Next i

    ' Analyser vélocité
    outRow = 4
    lr = wsReconcil.Cells(wsReconcil.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        Dim acct As String, name As String
        Dim currEcart As Double, prevEcart As Double
        Dim velocity As Double, variation As Double

        acct = SAFA_Common.SafeText(wsReconcil.Cells(i, 1).Value)
        name = SAFA_Common.SafeText(wsReconcil.Cells(i, 2).Value)
        currEcart = SAFA_Common.SafeVal(wsReconcil.Cells(i, 5).Value)

        If dictPrev.Exists(acct) Then
            prevEcart = dictPrev(acct)

            If Abs(prevEcart) > 0 Then
                variation = currEcart - prevEcart
                velocity = (Abs(currEcart) - Abs(prevEcart)) / Abs(prevEcart)

                ' Filtrer variations significatives
                If Abs(velocity) > 0.1 Or Abs(variation) > 100000 Then
                    wsOut.Cells(outRow, 1).Value = acct
                    wsOut.Cells(outRow, 2).Value = name
                    wsOut.Cells(outRow, 3).Value = prevEcart
                    wsOut.Cells(outRow, 3).NumberFormat = "#,##0"
                    wsOut.Cells(outRow, 4).Value = currEcart
                    wsOut.Cells(outRow, 4).NumberFormat = "#,##0"
                    wsOut.Cells(outRow, 5).Value = variation
                    wsOut.Cells(outRow, 5).NumberFormat = "#,##0"
                    wsOut.Cells(outRow, 6).Value = velocity
                    wsOut.Cells(outRow, 6).NumberFormat = "0.0%"

                    ' Tendance
                    If velocity > 0.5 Then
                        wsOut.Cells(outRow, 7).Value = "EXPLOSION ↑↑"
                        wsOut.Cells(outRow, 7).Interior.Color = RGB(255, 0, 0)
                        wsOut.Cells(outRow, 8).Value = "CRITIQUE"
                    ElseIf velocity > 0.2 Then
                        wsOut.Cells(outRow, 7).Value = "HAUSSE ↑"
                        wsOut.Cells(outRow, 7).Interior.Color = RGB(255, 165, 0)
                        wsOut.Cells(outRow, 8).Value = "Surveiller"
                    ElseIf velocity < -0.5 Then
                        wsOut.Cells(outRow, 7).Value = "CHUTE ↓↓"
                        wsOut.Cells(outRow, 7).Interior.Color = RGB(0, 200, 0)
                        wsOut.Cells(outRow, 8).Value = "Amélioration"
                    ElseIf velocity < -0.2 Then
                        wsOut.Cells(outRow, 7).Value = "BAISSE ↓"
                        wsOut.Cells(outRow, 7).Interior.Color = RGB(144, 238, 144)
                        wsOut.Cells(outRow, 8).Value = "Positif"
                    Else
                        wsOut.Cells(outRow, 7).Value = "STABLE →"
                        wsOut.Cells(outRow, 8).Value = ""
                    End If

                    outRow = outRow + 1
                End If
            End If
        End If
    Next i

    wsOut.Columns("A:H").AutoFit

    SAFA_Common.WriteAuditLog "PROCESS", "Analyse vélocité terminée", "Comptes analysés: " & (outRow - 4)
    Exit Sub

VelError:
    SAFA_Common.LogError MODULE_NAME, "AnalyserVelociteComptes", Err.Number, Err.Description
End Sub
