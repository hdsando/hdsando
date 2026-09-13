Attribute VB_Name = "Report_Generator"
'===============================================================================
' MODULE: Report_Generator
' Description: Generation de rapports PDF, Excel et envoi d'alertes email
' Version: 10.0
' Auteur: S.A.F.A Team
' Date: Decembre 2024
'===============================================================================
Option Explicit

' ===== CONSTANTES =====
Private Const MODULE_NAME As String = "Report_Generator"
Private Const PDF_MARGIN_TOP As Double = 50
Private Const PDF_MARGIN_BOTTOM As Double = 50
Private Const PDF_MARGIN_LEFT As Double = 30
Private Const PDF_MARGIN_RIGHT As Double = 30

' ===== TYPES PERSONNALISES =====
Private Type ReportConfig
    IncludeCharts As Boolean
    IncludeRawData As Boolean
    IncludeSummary As Boolean
    IncludeCompliance As Boolean
    IncludeForensic As Boolean
    OutputFormat As String
    OutputPath As String
    EmailRecipients As String
    EmailOnCritical As Boolean
End Type

Private Type AlertInfo
    AlertType As String
    Severity As String
    Message As String
    Details As String
    Timestamp As Date
    Account As String
    Amount As Double
End Type

' ===== VARIABLES MODULE =====
Private mReportConfig As ReportConfig
Private mAlerts() As AlertInfo
Private mAlertCount As Long

'===============================================================================
' FONCTION PRINCIPALE: GenerateFullReport
' Description: Genere le rapport complet d'audit
'===============================================================================
Public Sub GenerateFullReport(Optional outputPath As String = "")
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim startTime As Double

    startTime = Timer
    Set wb = ThisWorkbook

    ' Initialiser la configuration
    InitializeReportConfig

    If outputPath <> "" Then mReportConfig.OutputPath = outputPath

    ' Generer Executive Summary
    GenerateExecutiveSummary wb

    ' Generer Dashboard Risque
    GenerateRiskDashboard wb

    ' Generer rapport d'alertes
    GenerateAlertReport wb

    ' Generer analyse forensique
    If mReportConfig.IncludeForensic Then
        GenerateForensicReport wb
    End If

    ' Generer rapport conformite
    If mReportConfig.IncludeCompliance Then
        GenerateComplianceReport wb
    End If

    ' Generer echantillon de test (Top 20)
    GenerateSampleSheet wb

    ' Exporter en PDF si demande
    If mReportConfig.OutputFormat = "PDF" Or mReportConfig.OutputFormat = "BOTH" Then
        ExportToPDF wb
    End If

    ' Log de l'operation
    LogReportGeneration "FULL_REPORT", Timer - startTime

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateFullReport", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateExecutiveSummary
' Description: Cree la feuille de synthese executive
'===============================================================================
Public Sub GenerateExecutiveSummary(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsReconcil As Worksheet
    Dim wsAudit As Worksheet
    Dim rowNum As Long

    Set ws = PrepareReportSheet("EXECUTIVE_SUMMARY")

    ' Recuperer les donnees sources
    On Error Resume Next
    Set wsReconcil = wb.Sheets("RECONCIL")
    Set wsAudit = wb.Sheets("AUDIT_REPORT")
    On Error GoTo ErrorHandler

    ' En-tete du rapport
    rowNum = 1
    With ws
        ' Titre principal
        .Cells(rowNum, 1).Value = "RAPPORT EXECUTIF - AUDIT FINANCIER"
        .Cells(rowNum, 1).Font.Size = 18
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Merge
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Font.Color = vbWhite

        rowNum = rowNum + 2

        ' Information de generation
        .Cells(rowNum, 1).Value = "Date de generation:"
        .Cells(rowNum, 2).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Utilisateur:"
        .Cells(rowNum, 2).Value = Environ("USERNAME")
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Version S.A.F.A:"
        .Cells(rowNum, 2).Value = "10.0"
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 2

        ' Section KPIs
        .Cells(rowNum, 1).Value = "INDICATEURS CLES DE PERFORMANCE (KPIs)"
        .Cells(rowNum, 1).Font.Size = 14
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 2

        ' Calculer les KPIs
        Dim totalRecords As Long
        Dim matchedRecords As Long
        Dim orphanRecords As Long
        Dim criticalAlerts As Long
        Dim highAlerts As Long
        Dim totalRisk As Double

        ' Colonnes RECONCIL / AUDIT_REPORT: constantes canoniques SAFA_Common (source: Core_Engine.ConstruireRapprochement)
        Dim ecartRecords As Long, totalProvision As Double
        If Not wsReconcil Is Nothing Then
            totalRecords = Application.WorksheetFunction.CountA(wsReconcil.Columns(SAFA_Common.RECONCIL_COL_COMPTE)) - 1
            On Error Resume Next
            matchedRecords = Application.WorksheetFunction.CountIf(wsReconcil.Columns(SAFA_Common.RECONCIL_COL_STATUT), SAFA_Common.RECONCIL_STATUT_OK)
            ecartRecords = Application.WorksheetFunction.CountIf(wsReconcil.Columns(SAFA_Common.RECONCIL_COL_STATUT), SAFA_Common.RECONCIL_STATUT_ECART)
            orphanRecords = Application.WorksheetFunction.CountIf(wsReconcil.Columns(SAFA_Common.RECONCIL_COL_SOURCE), "*Only")
            totalProvision = Application.WorksheetFunction.Sum(wsReconcil.Columns(SAFA_Common.RECONCIL_COL_PROVISION))
            On Error GoTo ErrorHandler
        End If

        If Not wsAudit Is Nothing Then
            On Error Resume Next
            criticalAlerts = Application.WorksheetFunction.CountIf(wsAudit.Columns(SAFA_Common.AUDIT_COL_NIVEAU), "CRITICAL") _
                           + Application.WorksheetFunction.CountIf(wsAudit.Columns(SAFA_Common.AUDIT_COL_NIVEAU), "FRAUD")
            highAlerts = Application.WorksheetFunction.CountIf(wsAudit.Columns(SAFA_Common.AUDIT_COL_NIVEAU), "HIGH")
            On Error GoTo ErrorHandler
        End If

        ' Afficher les KPIs
        .Cells(rowNum, 1).Value = "Total lignes analysees:"
        .Cells(rowNum, 2).Value = totalRecords
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Lignes rapprochees:"
        .Cells(rowNum, 2).Value = matchedRecords
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        .Cells(rowNum, 3).Value = FormatPercent(IIf(totalRecords > 0, matchedRecords / totalRecords, 0), 1)
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Ecarts a analyser:"
        .Cells(rowNum, 2).Value = ecartRecords
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        If ecartRecords > 0 Then .Cells(rowNum, 2).Interior.Color = RGB(255, 230, 200)
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Orphelins (GL seul / Balance seule):"
        .Cells(rowNum, 2).Value = orphanRecords
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        If orphanRecords > 0 Then .Cells(rowNum, 2).Interior.Color = RGB(255, 200, 200)
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Provision IFRS 9 estimee (XAF):"
        .Cells(rowNum, 2).Value = totalProvision
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Alertes CRITICAL:"
        .Cells(rowNum, 2).Value = criticalAlerts
        If criticalAlerts > 0 Then
            .Cells(rowNum, 2).Interior.Color = RGB(255, 0, 0)
            .Cells(rowNum, 2).Font.Color = vbWhite
        End If
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Alertes HIGH:"
        .Cells(rowNum, 2).Value = highAlerts
        If highAlerts > 0 Then .Cells(rowNum, 2).Interior.Color = RGB(255, 165, 0)
        rowNum = rowNum + 2

        ' Section Recommandations
        .Cells(rowNum, 1).Value = "RECOMMANDATIONS PRIORITAIRES"
        .Cells(rowNum, 1).Font.Size = 14
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 2

        Dim recNum As Integer
        recNum = 1

        If criticalAlerts > 0 Then
            .Cells(rowNum, 1).Value = recNum & ". URGENT: " & criticalAlerts & " alertes critiques necessitent une investigation immediate"
            .Cells(rowNum, 1).Font.Color = RGB(200, 0, 0)
            rowNum = rowNum + 1
            recNum = recNum + 1
        End If

        If orphanRecords > totalRecords * 0.1 Then
            .Cells(rowNum, 1).Value = recNum & ". Plus de 10% d'orphelins detectes - Verifier l'integrite des donnees sources"
            rowNum = rowNum + 1
            recNum = recNum + 1
        End If

        If highAlerts > 5 Then
            .Cells(rowNum, 1).Value = recNum & ". " & highAlerts & " alertes HIGH - Planifier une revue approfondie"
            rowNum = rowNum + 1
            recNum = recNum + 1
        End If

        .Cells(rowNum, 1).Value = recNum & ". Documenter les investigations et conserver les preuves"
        rowNum = rowNum + 2

        ' Section Signature
        .Cells(rowNum, 1).Value = "VALIDATION"
        .Cells(rowNum, 1).Font.Size = 14
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Prepare par:"
        .Cells(rowNum, 3).Value = "________________________"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Revise par:"
        .Cells(rowNum, 3).Value = "________________________"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Approuve par:"
        .Cells(rowNum, 3).Value = "________________________"
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Date:"
        .Cells(rowNum, 3).Value = "________________________"

        ' Mise en forme
        .Columns("A:F").AutoFit
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 20

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateExecutiveSummary", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateRiskDashboard
' Description: Cree le tableau de bord des risques avec graphiques
'===============================================================================
Public Sub GenerateRiskDashboard(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsReconcil As Worksheet
    Dim rowNum As Long
    Dim chartObj As ChartObject

    Set ws = PrepareReportSheet("DASHBOARD_RISQUE")

    Set wsReconcil = Nothing
    On Error Resume Next
    Set wsReconcil = wb.Sheets("RECONCIL")
    On Error GoTo ErrorHandler

    rowNum = 1

    With ws
        ' Titre
        .Cells(rowNum, 1).Value = "TABLEAU DE BORD DES RISQUES"
        .Cells(rowNum, 1).Font.Size = 16
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Merge
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Font.Color = vbWhite
        rowNum = rowNum + 2

        ' Distribution des risques
        .Cells(rowNum, 1).Value = "Distribution par Niveau de Risque"
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        ' En-tetes
        .Cells(rowNum, 1).Value = "Niveau"
        .Cells(rowNum, 2).Value = "Nombre"
        .Cells(rowNum, 3).Value = "Pourcentage"
        .Cells(rowNum, 4).Value = "Montant Total"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Distribution par priorite: nombre ET montant (|Ecart|) - lu depuis RECONCIL (colonnes canoniques)
        Dim prioNames As Variant, prioColors As Variant
        Dim prioCount(0 To 3) As Long, prioSum(0 To 3) As Double
        Dim totalCount As Long, k As Long
        Dim recData As Variant, lastRec As Long, p As Long

        prioNames = Array("CRITICAL", "HIGH", "MEDIUM", "LOW")
        prioColors = Array(RGB(255, 100, 100), RGB(255, 200, 100), RGB(255, 255, 150), RGB(150, 255, 150))

        If Not wsReconcil Is Nothing Then
            lastRec = wsReconcil.Cells(wsReconcil.Rows.Count, SAFA_Common.RECONCIL_COL_COMPTE).End(xlUp).Row
            If lastRec >= 2 Then
                recData = wsReconcil.Range(wsReconcil.Cells(2, 1), wsReconcil.Cells(lastRec, SAFA_Common.RECONCIL_COL_PRIORITY)).Value
                For k = 1 To UBound(recData, 1)
                    For p = 0 To 3
                        If UCase(SAFA_Common.SafeText(recData(k, SAFA_Common.RECONCIL_COL_PRIORITY))) = prioNames(p) Then
                            prioCount(p) = prioCount(p) + 1
                            prioSum(p) = prioSum(p) + Abs(SAFA_Common.SafeVal(recData(k, SAFA_Common.RECONCIL_COL_ECART)))
                            Exit For
                        End If
                    Next p
                Next k
            End If
        End If
        totalCount = prioCount(0) + prioCount(1) + prioCount(2) + prioCount(3)

        For p = 0 To 3
            .Cells(rowNum, 1).Value = prioNames(p)
            .Cells(rowNum, 2).Value = prioCount(p)
            .Cells(rowNum, 3).Value = IIf(totalCount > 0, prioCount(p) / totalCount, 0)
            .Cells(rowNum, 3).NumberFormat = "0.0%"
            .Cells(rowNum, 4).Value = prioSum(p)
            .Cells(rowNum, 4).NumberFormat = "#,##0"
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = prioColors(p)
            rowNum = rowNum + 1
        Next p
        rowNum = rowNum + 1

        ' Graphique de repartition (barres: lisibles et comparables, contrairement au camembert)
        If totalCount > 0 Then
            Set chartObj = .ChartObjects.Add(Left:=300, Top:=50, Width:=350, Height:=250)
            With chartObj.Chart
                .ChartType = xlBarClustered
                .SetSourceData Source:=ws.Range(ws.Cells(rowNum - 5, 1), ws.Cells(rowNum - 2, 2))
                .HasTitle = True
                .ChartTitle.Text = "Comptes par niveau de risque"
                .HasLegend = False
            End With
        End If

        ' Section Top 10 Risques
        rowNum = rowNum + 2
        .Cells(rowNum, 1).Value = "TOP 10 COMPTES A RISQUE"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Rang"
        .Cells(rowNum, 2).Value = "Compte"
        .Cells(rowNum, 3).Value = "Libelle"
        .Cells(rowNum, 4).Value = "Score"
        .Cells(rowNum, 5).Value = "Niveau"
        .Cells(rowNum, 6).Value = "Ecart"
        .Cells(rowNum, 7).Value = "Facteurs"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 7)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 7)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Top 10 reel: RECONCIL trie par Risk Score decroissant (sans modifier la feuille source)
        Dim topAcc As Variant, i As Long
        If Not wsReconcil Is Nothing Then topAcc = GetTopAccountsByScore(wsReconcil, 10)
        If IsArray(topAcc) Then
            For i = 1 To UBound(topAcc, 1)
                .Cells(rowNum, 1).Value = i
                .Cells(rowNum, 2).Value = topAcc(i, 1)
                .Cells(rowNum, 2).NumberFormat = "@"
                .Cells(rowNum, 3).Value = topAcc(i, 2)
                .Cells(rowNum, 4).Value = topAcc(i, 4)
                .Cells(rowNum, 5).Value = topAcc(i, 5)
                .Cells(rowNum, 6).Value = topAcc(i, 3)
                .Cells(rowNum, 6).NumberFormat = "#,##0"
                .Cells(rowNum, 7).Value = topAcc(i, 6)
                Call ColorPriorityCell(.Cells(rowNum, 5))
                rowNum = rowNum + 1
            Next i
        Else
            .Cells(rowNum, 2).Value = "Aucune donnee RECONCIL - lancez d'abord l'analyse"
            rowNum = rowNum + 1
        End If

        ' Mise en forme finale
        .Columns("A:H").AutoFit

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateRiskDashboard", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateAlertReport
' Description: Genere le rapport detaille des alertes
'===============================================================================
Public Sub GenerateAlertReport(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rowNum As Long

    ' Verifier si existe deja
    On Error Resume Next
    Set ws = wb.Sheets("AUDIT_REPORT")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "AUDIT_REPORT"
    End If

    ' AUDIT_REPORT est ecrit par Forensic_Rules / Advanced_AI (layout canonique SAFA_Common.AUDIT_COL_*).
    ' On ne reecrit JAMAIS l'en-tete existant (l'ancien code l'ecrasait avec un layout different).
    With ws
        If SAFA_Common.SafeText(.Cells(1, 1).Value) = "" Then
            rowNum = 1
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Value = Array("Ref", "Categorie", "Risque", "Niveau", "Compte", _
                                                                     "Description", "Valeur", "Impact Est.", "SLA")
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Font.Bold = True
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Interior.Color = RGB(0, 51, 102)
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Font.Color = vbWhite
        End If

        ' Appliquer mise en forme conditionnelle
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        If lastRow > 1 Then
            ' Couleur selon severite (colonne Niveau)
            Dim rng As Range
            Set rng = .Range(.Cells(2, SAFA_Common.AUDIT_COL_NIVEAU), .Cells(lastRow, SAFA_Common.AUDIT_COL_NIVEAU))

            ' Supprimer formatage conditionnel existant
            rng.FormatConditions.Delete

            ' CRITICAL = Rouge
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""CRITICAL""")
                .Interior.Color = RGB(255, 0, 0)
                .Font.Color = vbWhite
            End With

            ' HIGH = Orange
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""HIGH""")
                .Interior.Color = RGB(255, 165, 0)
            End With

            ' MEDIUM = Jaune
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""MEDIUM""")
                .Interior.Color = RGB(255, 255, 0)
            End With

            ' FRAUD = Rouge fonce
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""FRAUD""")
                .Interior.Color = RGB(139, 0, 0)
                .Font.Color = vbWhite
            End With
        End If

        .Columns("A:I").AutoFit

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateAlertReport", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateForensicReport
' Description: Genere le rapport d'analyse forensique
'===============================================================================
Public Sub GenerateForensicReport(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rowNum As Long

    ' FORENSIC_ANALYSIS est la feuille de resultats Benford ecrite par Forensic_Rules (on la LIT, on ne l'ecrase plus).
    ' Le rapport de synthese forensique est ecrit dans FORENSIC_REPORT.
    Dim wsBenford As Worksheet, wsAudit As Worksheet
    On Error Resume Next
    Set wsBenford = wb.Sheets("FORENSIC_ANALYSIS")
    Set wsAudit = wb.Sheets("AUDIT_REPORT")
    On Error GoTo ErrorHandler

    Set ws = PrepareReportSheet("FORENSIC_REPORT")

    rowNum = 1

    With ws
        ' Titre
        .Cells(rowNum, 1).Value = "ANALYSE FORENSIQUE"
        .Cells(rowNum, 1).Font.Size = 16
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Merge
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Interior.Color = RGB(102, 0, 102)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Font.Color = vbWhite
        rowNum = rowNum + 2

        ' Section Benford
        .Cells(rowNum, 1).Value = "1. ANALYSE DE BENFORD"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        rowNum = rowNum + 2

        ' En-tetes Benford
        .Cells(rowNum, 1).Value = "Chiffre"
        .Cells(rowNum, 2).Value = "Attendu (%)"
        .Cells(rowNum, 3).Value = "Observe (%)"
        .Cells(rowNum, 4).Value = "Ecart"
        .Cells(rowNum, 5).Value = "Statut"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Distribution Benford theorique
        Dim benford(1 To 9) As Double
        benford(1) = 30.1: benford(2) = 17.6: benford(3) = 12.5
        benford(4) = 9.7: benford(5) = 7.9: benford(6) = 6.7
        benford(7) = 5.8: benford(8) = 5.1: benford(9) = 4.6

        ' Valeurs reelles: lues dans FORENSIC_ANALYSIS (Forensic_Rules.Lancer_Benford_Enhanced),
        ' tableau en lignes 9..17: A=Chiffre, B=Observe (nb), C=Attendu (nb), D=Freq reelle, E=Freq theorique, F=Ecart, G=Statut
        Dim d As Integer, obsCount(1 To 9) As Double, totalObs As Double
        Dim chiSq As Double, madVal As Double, benfordOK As Boolean
        Dim cfgF As Config_Manager.ForensicConfig
        cfgF = Config_Manager.GetForensicConfig()

        benfordOK = False
        If Not wsBenford Is Nothing Then
            On Error Resume Next
            For d = 1 To 9
                obsCount(d) = SAFA_Common.SafeVal(wsBenford.Cells(8 + d, 2).Value)
                totalObs = totalObs + obsCount(d)
            Next d
            On Error GoTo ErrorHandler
            benfordOK = (totalObs > 0)
        End If

        For d = 1 To 9
            .Cells(rowNum, 1).Value = d
            .Cells(rowNum, 2).Value = benford(d) / 100
            .Cells(rowNum, 2).NumberFormat = "0.0%"
            If benfordOK Then
                .Cells(rowNum, 3).Value = obsCount(d) / totalObs
                .Cells(rowNum, 4).Value = obsCount(d) / totalObs - benford(d) / 100
                .Cells(rowNum, 4).NumberFormat = "+0.0%;-0.0%"
                If Abs(.Cells(rowNum, 4).Value) > 0.05 Then
                    .Cells(rowNum, 5).Value = "ANOMALIE"
                    .Cells(rowNum, 5).Interior.Color = RGB(255, 200, 200)
                Else
                    .Cells(rowNum, 5).Value = "Conforme"
                End If
                ' Chi-squared et MAD recalcules a partir des comptages (pas de parsing de texte)
                chiSq = chiSq + ((obsCount(d) - totalObs * benford(d) / 100) ^ 2) / (totalObs * benford(d) / 100)
                madVal = madVal + Abs(obsCount(d) / totalObs - benford(d) / 100)
            Else
                .Cells(rowNum, 3).Value = "n/a"
            End If
            .Cells(rowNum, 3).NumberFormat = "0.0%"
            rowNum = rowNum + 1
        Next d
        madVal = madVal / 9

        rowNum = rowNum + 1

        ' Metriques Benford
        .Cells(rowNum, 1).Value = "Echantillon (nb montants):"
        .Cells(rowNum, 2).Value = IIf(benfordOK, totalObs, "n/a")
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        If benfordOK And totalObs < cfgF.BenfordMinSampleSize Then .Cells(rowNum, 3).Value = "Echantillon < " & cfgF.BenfordMinSampleSize & " : test peu fiable"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Chi-squared:"
        .Cells(rowNum, 2).Value = IIf(benfordOK, Round(chiSq, 2), "n/a")
        .Cells(rowNum, 3).Value = "(Seuil critique 5%, 8 ddl: " & cfgF.BenfordChiSquaredCritical & ")"
        If benfordOK And chiSq > cfgF.BenfordChiSquaredCritical Then .Cells(rowNum, 2).Interior.Color = RGB(255, 200, 200)
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "MAD:"
        .Cells(rowNum, 2).Value = IIf(benfordOK, Round(madVal, 4), "n/a")
        .Cells(rowNum, 3).Value = "(< " & cfgF.BenfordMADExcellent & " Excellent, < " & cfgF.BenfordMADAcceptable & " Acceptable, < " & cfgF.BenfordMADMarginal & " Marginal, sinon Non conforme)"
        If benfordOK Then
            If madVal >= cfgF.BenfordMADMarginal Then
                .Cells(rowNum, 4).Value = "NON CONFORME"
                .Cells(rowNum, 4).Interior.Color = RGB(255, 100, 100)
            ElseIf madVal >= cfgF.BenfordMADAcceptable Then
                .Cells(rowNum, 4).Value = "MARGINAL"
                .Cells(rowNum, 4).Interior.Color = RGB(255, 230, 150)
            Else
                .Cells(rowNum, 4).Value = "CONFORME"
                .Cells(rowNum, 4).Interior.Color = RGB(200, 255, 200)
            End If
        End If
        rowNum = rowNum + 2

        ' Section Patterns
        .Cells(rowNum, 1).Value = "2. PATTERNS DETECTES"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Pattern"
        .Cells(rowNum, 2).Value = "Occurrences"
        .Cells(rowNum, 3).Value = "Risque"
        .Cells(rowNum, 4).Value = "Description"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Occurrences reelles: comptage par categorie dans AUDIT_REPORT (colonne Categorie)
        Call WritePatternRow(ws, rowNum, wsAudit, "KEYWORD", "Mots-cles suspects (FRD-001)", "CADEAU, URGENT, MANUEL, OVERRIDE...")
        Call WritePatternRow(ws, rowNum, wsAudit, "WEEKEND", "Transactions week-end (FRD-002)", "Operations > " & Format(cfgF.WeekendThreshold, "#,##0") & " XAF samedi/dimanche")
        Call WritePatternRow(ws, rowNum, wsAudit, "THRESHOLD", "Juste sous seuil (FRD-003)", "Montants proches des seuils LAB/FT")
        Call WritePatternRow(ws, rowNum, wsAudit, "LAYERING", "Circularite / layering (FRD-004)", "Round-tripping, empilement de flux")
        Call WritePatternRow(ws, rowNum, wsAudit, "STRUCTURING", "Saucissonnage (FRD-005)", "Fractionnement de transactions")
        Call WritePatternRow(ws, rowNum, wsAudit, "DUPLICATE", "Doublons (FRD-006)", "Transactions identiques")
        Call WritePatternRow(ws, rowNum, wsAudit, "PATTERN", "Montants ronds / repetitifs (FRD-007)", "Patterns de montants")
        Call WritePatternRow(ws, rowNum, wsAudit, "TIMING", "Timing suspect (FRD-008)", "Fin de mois, heures atypiques")
        rowNum = rowNum + 1

        ' Section Z-Score
        .Cells(rowNum, 1).Value = "3. ANOMALIES STATISTIQUES (Z-SCORE)"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Ref"
        .Cells(rowNum, 2).Value = "Compte"
        .Cells(rowNum, 3).Value = "Niveau"
        .Cells(rowNum, 4).Value = "Montant"
        .Cells(rowNum, 5).Value = "Risque"
        .Cells(rowNum, 6).Value = "Details (moyenne, ecart-type, n)"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Alertes IA-xxx (Z-Score) reelles depuis AUDIT_REPORT (max 15)
        If Not wsAudit Is Nothing Then
            Dim lastA As Long, ia As Long, nZ As Long
            lastA = wsAudit.Cells(wsAudit.Rows.Count, SAFA_Common.AUDIT_COL_REF).End(xlUp).Row
            For ia = 2 To lastA
                If Left(SAFA_Common.SafeText(wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_REF).Value), 3) = "IA-" Then
                    .Cells(rowNum, 1).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_REF).Value
                    .Cells(rowNum, 2).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_COMPTE).Value
                    .Cells(rowNum, 2).NumberFormat = "@"
                    .Cells(rowNum, 3).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_NIVEAU).Value
                    .Cells(rowNum, 4).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_VALEUR).Value
                    .Cells(rowNum, 4).NumberFormat = "#,##0"
                    .Cells(rowNum, 5).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_RISQUE).Value
                    .Cells(rowNum, 6).Value = wsAudit.Cells(ia, SAFA_Common.AUDIT_COL_DESCRIPTION).Value
                    Call ColorPriorityCell(.Cells(rowNum, 3))
                    rowNum = rowNum + 1
                    nZ = nZ + 1
                    If nZ >= 15 Then Exit For
                End If
            Next ia
            If nZ = 0 Then .Cells(rowNum, 1).Value = "Aucune anomalie Z-Score detectee"
        End If

        .Columns("A:H").AutoFit

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateForensicReport", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateComplianceReport
' Description: Genere le rapport de conformite reglementaire
'===============================================================================
Public Sub GenerateComplianceReport(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rowNum As Long

    ' COMPLIANCE_CHECK est la feuille de resultats ecrite par Regulatory_Compliance (on la LIT, on ne l'ecrase plus).
    ' Le rapport de synthese est ecrit dans COMPLIANCE_REPORT.
    Dim wsCheck As Worksheet
    On Error Resume Next
    Set wsCheck = wb.Sheets("COMPLIANCE_CHECK")
    On Error GoTo ErrorHandler

    Set ws = PrepareReportSheet("COMPLIANCE_REPORT")

    rowNum = 1

    With ws
        ' Titre
        .Cells(rowNum, 1).Value = "RAPPORT DE CONFORMITE REGLEMENTAIRE"
        .Cells(rowNum, 1).Font.Size = 16
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Merge
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Interior.Color = RGB(0, 100, 0)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 8)).Font.Color = vbWhite
        rowNum = rowNum + 2

        ' Section COBAC
        .Cells(rowNum, 1).Value = "1. CONFORMITE COBAC"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(230, 230, 230)
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Regle"
        .Cells(rowNum, 2).Value = "Description"
        .Cells(rowNum, 3).Value = "Seuil"
        .Cells(rowNum, 4).Value = "Resultat"
        .Cells(rowNum, 5).Value = "Statut"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        Dim cfgR As Config_Manager.RegulatoryConfig
        cfgR = Config_Manager.GetRegulatoryConfig()

        Call WriteComplianceRow(ws, rowNum, wsCheck, "COBAC-001", "Suspens > " & cfgR.CobacSuspensLimitDays & " jours", cfgR.CobacSuspensLimitDays & " jours")
        Call WriteComplianceRow(ws, rowNum, wsCheck, "COBAC-002", "Transit non apure J+" & cfgR.CobacTransitLimitDays, cfgR.CobacTransitLimitDays & " jours")
        Call WriteComplianceRow(ws, rowNum, wsCheck, "COBAC-003", "Couverture provisions", "> 50%")
        rowNum = rowNum + 1

        ' Section OHADA
        .Cells(rowNum, 1).Value = "2. CONFORMITE OHADA/SYSCOHADA"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(230, 230, 230)
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Regle"
        .Cells(rowNum, 2).Value = "Description"
        .Cells(rowNum, 3).Value = "Seuil"
        .Cells(rowNum, 4).Value = "Resultat"
        .Cells(rowNum, 5).Value = "Statut"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        Call WriteComplianceRow(ws, rowNum, wsCheck, "OHADA-001", "Equilibre bilan (Actif=Passif)", "Ecart < 1000")
        Call WriteComplianceRow(ws, rowNum, wsCheck, "OHADA-002", "Coherence sens comptes", "Par classe")
        rowNum = rowNum + 1

        ' Section LAB/FT
        .Cells(rowNum, 1).Value = "3. CONFORMITE LAB/FT"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(230, 230, 230)
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Regle"
        .Cells(rowNum, 2).Value = "Description"
        .Cells(rowNum, 3).Value = "Seuil"
        .Cells(rowNum, 4).Value = "Resultat"
        .Cells(rowNum, 5).Value = "Statut"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        Call WriteComplianceRow(ws, rowNum, wsCheck, "LAB-001", "Transactions > seuil declaration", Format(cfgR.LabftDeclarationThreshold, "#,##0") & " XAF")
        Call WriteComplianceRow(ws, rowNum, wsCheck, "LAB-002", "Detection structuration", "Cumul > " & Format(cfgR.LabftStructuringThreshold, "#,##0") & " XAF")
        rowNum = rowNum + 1

        If wsCheck Is Nothing Then
            .Cells(rowNum, 1).Value = "Feuille COMPLIANCE_CHECK absente: executez Regulatory_Compliance.Lancer_Verification_Conformite"
            .Cells(rowNum, 1).Font.Italic = True
            .Cells(rowNum, 1).Font.Color = RGB(200, 0, 0)
            rowNum = rowNum + 1
        End If
        rowNum = rowNum + 1

        ' Signature
        .Cells(rowNum, 1).Value = "Rapport genere le: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(rowNum, 1).Font.Italic = True

        .Columns("A:H").AutoFit

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateComplianceReport", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: ExportToPDF
' Description: Exporte les rapports en PDF
'===============================================================================
Public Sub ExportToPDF(wb As Workbook, Optional outputPath As String = "")
    On Error GoTo ErrorHandler

    Dim pdfPath As String
    Dim sheetsToExport As Variant
    Dim ws As Worksheet
    Dim i As Integer

    ' Determiner le chemin de sortie
    If outputPath = "" Then
        pdfPath = wb.Path & "\" & "SAFA_Report_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    Else
        pdfPath = outputPath
    End If

    ' Liste des feuilles a exporter
    sheetsToExport = Array("EXECUTIVE_SUMMARY", "DASHBOARD_RISQUE", "AUDIT_REPORT", _
                          "RECONCIL", "FORENSIC_REPORT", "FORENSIC_ANALYSIS", _
                          "COMPLIANCE_REPORT", "COMPLIANCE_CHECK", "ECHANTILLON_TEST")

    ' Selectionner les feuilles existantes
    Dim firstSheet As Boolean
    firstSheet = True

    For i = LBound(sheetsToExport) To UBound(sheetsToExport)
        On Error Resume Next
        Set ws = wb.Sheets(sheetsToExport(i))
        If Not ws Is Nothing Then
            If firstSheet Then
                ws.Select
                firstSheet = False
            Else
                ws.Select False ' Ajouter a la selection
            End If
        End If
        Set ws = Nothing
        On Error GoTo ErrorHandler
    Next i

    ' Exporter en PDF
    If Not firstSheet Then
        ActiveSheet.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

        MsgBox "Rapport PDF genere: " & vbNewLine & pdfPath, vbInformation, "Export PDF"
    End If

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "ExportToPDF", Err.Number, Err.Description
    MsgBox "Erreur lors de l'export PDF: " & Err.Description, vbExclamation, "Erreur"
End Sub

'===============================================================================
' FONCTION: SendAlertEmail
' Description: Envoie une alerte par email via Outlook
'===============================================================================
Public Sub SendAlertEmail(recipient As String, subject As String, body As String, _
                         Optional attachmentPath As String = "", _
                         Optional highPriority As Boolean = False)
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim outlookMail As Object

    ' Creer instance Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        MsgBox "Microsoft Outlook n'est pas installe ou accessible.", vbExclamation, "Erreur Email"
        Exit Sub
    End If

    ' Creer email
    Set outlookMail = outlookApp.CreateItem(0) ' olMailItem

    With outlookMail
        .To = recipient
        .subject = "[S.A.F.A] " & subject
        .body = "=== ALERTE AUTOMATIQUE S.A.F.A ===" & vbNewLine & vbNewLine & _
                body & vbNewLine & vbNewLine & _
                "---" & vbNewLine & _
                "Ce message a ete genere automatiquement par S.A.F.A v10.0" & vbNewLine & _
                "Date: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbNewLine & _
                "Utilisateur: " & Environ("USERNAME")

        If attachmentPath <> "" Then
            If Dir(attachmentPath) <> "" Then
                .Attachments.Add attachmentPath
            End If
        End If

        If highPriority Then
            .Importance = 2 ' olImportanceHigh
        End If

        .Display ' Afficher pour verification (ou .Send pour envoi direct)
    End With

    Set outlookMail = Nothing
    Set outlookApp = Nothing

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "SendAlertEmail", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: SendCriticalAlerts
' Description: Envoie automatiquement les alertes critiques par email
'===============================================================================
Public Sub SendCriticalAlerts(wb As Workbook, recipient As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim criticalCount As Long
    Dim alertBody As String

    On Error Resume Next
    Set ws = wb.Sheets("AUDIT_REPORT")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    criticalCount = 0
    alertBody = "ALERTES CRITIQUES DETECTEES:" & vbNewLine & vbNewLine

    ' Parcourir les alertes (layout canonique AUDIT_COL_*)
    For i = 2 To lastRow
        Dim niveau As String
        niveau = UCase(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.AUDIT_COL_NIVEAU).Value))
        If niveau = "CRITICAL" Or niveau = "FRAUD" Then
            criticalCount = criticalCount + 1
            alertBody = alertBody & criticalCount & ". [" & ws.Cells(i, SAFA_Common.AUDIT_COL_REF).Value & "] " & _
                        ws.Cells(i, SAFA_Common.AUDIT_COL_CATEGORIE).Value & " - " & ws.Cells(i, SAFA_Common.AUDIT_COL_RISQUE).Value & vbNewLine
            alertBody = alertBody & "   Compte: " & ws.Cells(i, SAFA_Common.AUDIT_COL_COMPTE).Value & vbNewLine
            alertBody = alertBody & "   Details: " & ws.Cells(i, SAFA_Common.AUDIT_COL_DESCRIPTION).Value & vbNewLine
            alertBody = alertBody & "   Montant: " & Format(ws.Cells(i, SAFA_Common.AUDIT_COL_VALEUR).Value, "#,##0") & " XAF" & vbNewLine & vbNewLine
        End If
    Next i

    ' Envoyer si alertes critiques trouvees
    If criticalCount > 0 Then
        alertBody = alertBody & "---" & vbNewLine & _
                   "Total alertes critiques: " & criticalCount & vbNewLine & _
                   "Action requise: Investigation immediate"

        SendAlertEmail recipient, _
                      "URGENT - " & criticalCount & " Alertes Critiques", _
                      alertBody, _
                      "", _
                      True
    End If

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "SendCriticalAlerts", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: GenerateSampleSheet
' Description: Cree l'echantillon de test (Top 20)
'===============================================================================
Public Sub GenerateSampleSheet(wb As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsReconcil As Worksheet
    Dim rowNum As Long
    Dim i As Long

    Set ws = PrepareReportSheet("ECHANTILLON_TEST")

    On Error Resume Next
    Set wsReconcil = wb.Sheets("RECONCIL")
    On Error GoTo ErrorHandler

    rowNum = 1

    With ws
        ' Titre
        .Cells(rowNum, 1).Value = "ECHANTILLON DE TEST - TOP 20 RISQUES"
        .Cells(rowNum, 1).Font.Size = 14
        .Cells(rowNum, 1).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Merge
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Font.Color = vbWhite
        rowNum = rowNum + 2

        ' En-tetes
        .Cells(rowNum, 1).Value = "RANG"
        .Cells(rowNum, 2).Value = "COMPTE"
        .Cells(rowNum, 3).Value = "LIBELLE"
        .Cells(rowNum, 4).Value = "ECART"
        .Cells(rowNum, 5).Value = "SCORE"
        .Cells(rowNum, 6).Value = "NIVEAU"
        .Cells(rowNum, 7).Value = "FACTEURS"
        .Cells(rowNum, 8).Value = "STATUT_INV"
        .Cells(rowNum, 9).Value = "COMMENTAIRE"
        .Cells(rowNum, 10).Value = "CONCLUSION"

        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Top 20 reel: RECONCIL trie par Risk Score decroissant (colonnes canoniques, feuille source intacte)
        Dim topAcc As Variant
        If Not wsReconcil Is Nothing Then topAcc = GetTopAccountsByScore(wsReconcil, 20)
        If IsArray(topAcc) Then
            For i = 1 To UBound(topAcc, 1)
                .Cells(rowNum, 1).Value = i
                .Cells(rowNum, 2).Value = topAcc(i, 1)
                .Cells(rowNum, 2).NumberFormat = "@"
                .Cells(rowNum, 3).Value = topAcc(i, 2)
                .Cells(rowNum, 4).Value = topAcc(i, 3)
                .Cells(rowNum, 4).NumberFormat = "#,##0"
                .Cells(rowNum, 5).Value = topAcc(i, 4)
                .Cells(rowNum, 6).Value = topAcc(i, 5)
                .Cells(rowNum, 7).Value = topAcc(i, 6)
                .Cells(rowNum, 8).Value = "A INVESTIGUER"

                Select Case UCase(SAFA_Common.SafeText(topAcc(i, 5)))
                    Case "CRITICAL"
                        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 200, 200)
                    Case "HIGH"
                        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 230, 200)
                    Case "MEDIUM"
                        .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 255, 200)
                End Select
                rowNum = rowNum + 1
            Next i
        End If

        ' Ajouter lignes vides si moins de 20
        Do While rowNum < 24
            .Cells(rowNum, 1).Value = rowNum - 3
            rowNum = rowNum + 1
        Loop

        ' Zone de signature
        rowNum = rowNum + 2
        .Cells(rowNum, 1).Value = "Investigations realisees par:"
        .Cells(rowNum, 4).Value = "________________________"
        rowNum = rowNum + 1
        .Cells(rowNum, 1).Value = "Date:"
        .Cells(rowNum, 4).Value = "________________________"

        .Columns("A:J").AutoFit

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "GenerateSampleSheet", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTIONS UTILITAIRES
'===============================================================================

Private Sub InitializeReportConfig()
    With mReportConfig
        .IncludeCharts = True
        .IncludeRawData = False
        .IncludeSummary = True
        .IncludeCompliance = True
        .IncludeForensic = True
        .OutputFormat = "XLSX"
        .OutputPath = ThisWorkbook.Path
        .EmailOnCritical = False
    End With
End Sub

Private Sub LogReportGeneration(reportType As String, duration As Double)
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("HISTORY_LOG")
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = reportType
    ws.Cells(lastRow, 3).Value = Environ("USERNAME")
    ws.Cells(lastRow, 4).Value = Round(duration, 2) & " sec"
    ws.Cells(lastRow, 5).Value = "SUCCESS"
End Sub

' Gestion d'erreurs centralisee (SAFA_Common)
Private Sub LogError(moduleName As String, procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    Call SAFA_Common.LogError(moduleName, procName, errNum, errDesc)
End Sub

'===============================================================================
' FONCTION: PrepareReportSheet
' Description: Recupere ou cree une feuille de rapport et la vide proprement
'              (supprime TCD et graphiques avant Clear, sinon erreur 1004)
'===============================================================================
Private Function PrepareReportSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim pt As Object, co As Object

    Set ws = Core_Engine.GetOrCreateSheet(sheetName, False)

    On Error Resume Next
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0

    Set PrepareReportSheet = ws
End Function

'===============================================================================
' FONCTION: GetTopAccountsByScore
' Description: Retourne les N comptes au Risk Score le plus eleve, SANS trier
'              la feuille RECONCIL (lecture en memoire). Selection partielle O(n*N).
' Retour: tableau (1..N, 1..6) = Compte, Libelle, Ecart, Score, Priority, Facteurs
'         ou Empty si aucune donnee
'===============================================================================
Private Function GetTopAccountsByScore(wsReconcil As Worksheet, topN As Long) As Variant
    Dim lastRow As Long, n As Long, i As Long, k As Long
    Dim data As Variant, used() As Boolean
    Dim bestIdx As Long, bestScore As Double, sc As Double
    Dim result() As Variant

    lastRow = wsReconcil.Cells(wsReconcil.Rows.Count, SAFA_Common.RECONCIL_COL_COMPTE).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    data = wsReconcil.Range(wsReconcil.Cells(2, 1), wsReconcil.Cells(lastRow, SAFA_Common.RECONCIL_COL_PRIORITY)).Value
    n = UBound(data, 1)
    If topN > n Then topN = n
    If topN < 1 Then Exit Function

    ReDim used(1 To n)
    ReDim result(1 To topN, 1 To 6)

    For k = 1 To topN
        bestIdx = 0: bestScore = -1
        For i = 1 To n
            If Not used(i) Then
                sc = SAFA_Common.SafeVal(data(i, SAFA_Common.RECONCIL_COL_SCORE))
                If sc > bestScore Then
                    bestScore = sc
                    bestIdx = i
                End If
            End If
        Next i
        If bestIdx = 0 Then Exit For
        used(bestIdx) = True
        result(k, 1) = SAFA_Common.SafeText(data(bestIdx, SAFA_Common.RECONCIL_COL_COMPTE))
        result(k, 2) = SAFA_Common.SafeText(data(bestIdx, SAFA_Common.RECONCIL_COL_LIBELLE))
        result(k, 3) = SAFA_Common.SafeVal(data(bestIdx, SAFA_Common.RECONCIL_COL_ECART))
        result(k, 4) = bestScore
        result(k, 5) = SAFA_Common.SafeText(data(bestIdx, SAFA_Common.RECONCIL_COL_PRIORITY))
        result(k, 6) = SAFA_Common.SafeText(data(bestIdx, SAFA_Common.RECONCIL_COL_FACTEURS))
    Next k

    GetTopAccountsByScore = result
End Function

'===============================================================================
' PROCEDURE: ColorPriorityCell - couleur standard selon niveau
'===============================================================================
Private Sub ColorPriorityCell(cell As Range)
    Select Case UCase(SAFA_Common.SafeText(cell.Value))
        Case "CRITICAL", "FRAUD"
            cell.Interior.Color = RGB(255, 0, 0): cell.Font.Color = vbWhite
        Case "HIGH", "MAJOR"
            cell.Interior.Color = RGB(255, 165, 0)
        Case "MEDIUM"
            cell.Interior.Color = RGB(255, 255, 0)
        Case "LOW"
            cell.Interior.Color = RGB(200, 255, 200)
    End Select
End Sub

'===============================================================================
' PROCEDURE: WritePatternRow - ligne du tableau "patterns detectes" avec
'            comptage reel dans AUDIT_REPORT (colonne Categorie)
'===============================================================================
Private Sub WritePatternRow(ws As Worksheet, ByRef rowNum As Long, wsAudit As Worksheet, _
                            category As String, label As String, description As String)
    Dim cnt As Long, risk As String
    cnt = 0
    If Not wsAudit Is Nothing Then
        On Error Resume Next
        cnt = Application.WorksheetFunction.CountIf(wsAudit.Columns(SAFA_Common.AUDIT_COL_CATEGORIE), category)
        On Error GoTo 0
    End If
    If cnt = 0 Then
        risk = "-"
    ElseIf cnt >= 10 Then
        risk = "ELEVE"
    ElseIf cnt >= 3 Then
        risk = "MOYEN"
    Else
        risk = "FAIBLE"
    End If
    ws.Cells(rowNum, 1).Value = label
    ws.Cells(rowNum, 2).Value = cnt
    ws.Cells(rowNum, 3).Value = risk
    ws.Cells(rowNum, 4).Value = description
    If risk = "ELEVE" Then ws.Cells(rowNum, 3).Interior.Color = RGB(255, 150, 150)
    If risk = "MOYEN" Then ws.Cells(rowNum, 3).Interior.Color = RGB(255, 230, 150)
    rowNum = rowNum + 1
End Sub

'===============================================================================
' PROCEDURE: WriteComplianceRow - ligne de regle avec resultat reel lu dans
'            COMPLIANCE_CHECK (colonnes canoniques COMPLIANCE_COL_*)
'===============================================================================
Private Sub WriteComplianceRow(ws As Worksheet, ByRef rowNum As Long, wsCheck As Worksheet, _
                               ruleId As String, label As String, seuil As String)
    Dim lastRow As Long, i As Long
    Dim statut As String, details As String

    statut = "Non execute": details = ""
    If Not wsCheck Is Nothing Then
        lastRow = wsCheck.Cells(wsCheck.Rows.Count, SAFA_Common.COMPLIANCE_COL_ID).End(xlUp).Row
        For i = 2 To lastRow
            If UCase(SAFA_Common.SafeText(wsCheck.Cells(i, SAFA_Common.COMPLIANCE_COL_ID).Value)) = UCase(ruleId) Then
                statut = SAFA_Common.SafeText(wsCheck.Cells(i, SAFA_Common.COMPLIANCE_COL_STATUT).Value)
                details = SAFA_Common.SafeText(wsCheck.Cells(i, SAFA_Common.COMPLIANCE_COL_DETAILS).Value, 200)
                Exit For
            End If
        Next i
    End If

    ws.Cells(rowNum, 1).Value = ruleId
    ws.Cells(rowNum, 2).Value = label
    ws.Cells(rowNum, 3).Value = seuil
    ws.Cells(rowNum, 4).Value = details
    ws.Cells(rowNum, 5).Value = statut

    Select Case UCase(statut)
        Case "CONFORME", "OK", "PASS"
            ws.Cells(rowNum, 5).Interior.Color = RGB(200, 255, 200)
        Case "NON CONFORME", "FAIL", "KO", "CRITICAL"
            ws.Cells(rowNum, 5).Interior.Color = RGB(255, 150, 150)
        Case "ATTENTION", "WARNING", "A VERIFIER"
            ws.Cells(rowNum, 5).Interior.Color = RGB(255, 230, 150)
        Case "NON EXECUTE"
            ws.Cells(rowNum, 5).Font.Italic = True
    End Select
    rowNum = rowNum + 1
End Sub

'===============================================================================
' FONCTION: CreateBenfordChart
' Description: Cree un graphique de comparaison Benford
'===============================================================================
Public Sub CreateBenfordChart(ws As Worksheet, dataRange As Range, chartTitle As String)
    On Error GoTo ErrorHandler

    Dim chartObj As ChartObject

    Set chartObj = ws.ChartObjects.Add(Left:=400, Top:=50, Width:=450, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=dataRange
        .HasTitle = True
        .ChartTitle.Text = chartTitle

        ' Personnaliser les couleurs
        .SeriesCollection(1).Interior.Color = RGB(0, 100, 200)
        If .SeriesCollection.Count > 1 Then
            .SeriesCollection(2).Interior.Color = RGB(200, 50, 50)
        End If

        ' Legendes
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Axes
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Premier chiffre"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Frequence (%)"
    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "CreateBenfordChart", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: FormatReportSheet
' Description: Applique le formatage standard aux feuilles de rapport
'===============================================================================
Public Sub FormatReportSheet(ws As Worksheet)
    On Error GoTo ErrorHandler

    With ws
        ' Police par defaut
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10

        ' Zoom
        ActiveWindow.Zoom = 85

        ' Masquer quadrillage
        ActiveWindow.DisplayGridlines = False

        ' Figer les volets (ligne 1)
        .Rows(2).Select
        ActiveWindow.FreezePanes = True

        ' Retour cellule A1
        .Range("A1").Select
    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "FormatReportSheet", Err.Number, Err.Description
End Sub
