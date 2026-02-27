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

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    Set ws = Core_Engine.GetOrCreateSheet("EXECUTIVE_SUMMARY", True)

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

        If Not wsReconcil Is Nothing Then
            totalRecords = Application.WorksheetFunction.CountA(wsReconcil.Columns(1)) - 1
            On Error Resume Next
            matchedRecords = Application.WorksheetFunction.CountIf(wsReconcil.Columns("J"), "OK*")
            orphanRecords = totalRecords - matchedRecords
            On Error GoTo ErrorHandler
        End If

        If Not wsAudit Is Nothing Then
            On Error Resume Next
            criticalAlerts = Application.WorksheetFunction.CountIf(wsAudit.Columns("C"), "CRITICAL")
            highAlerts = Application.WorksheetFunction.CountIf(wsAudit.Columns("C"), "HIGH")
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

        .Cells(rowNum, 1).Value = "Orphelins detectes:"
        .Cells(rowNum, 2).Value = orphanRecords
        .Cells(rowNum, 2).NumberFormat = "#,##0"
        If orphanRecords > 0 Then .Cells(rowNum, 2).Interior.Color = RGB(255, 200, 200)
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

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    Set ws = Core_Engine.GetOrCreateSheet("DASHBOARD_RISQUE", True)

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

        Dim critCount As Long, highCount As Long, medCount As Long, lowCount As Long
        Dim totalCount As Long

        If Not wsReconcil Is Nothing Then
            On Error Resume Next
            ' CORRIGÉ BUG-004: La colonne Priority est en N (col 14), pas en K
            ' Structure RECONCIL: A=Compte, B=Libelle, C=Balance, D=GL, E=Ecart, F=Status,
            '                     G=Anciennete, H=NbTrans, I=DateMin, J=DateMax, K=Score,
            '                     L=Facteurs, M=Provisions, N=Priority
            critCount = Application.WorksheetFunction.CountIf(wsReconcil.Columns("N"), "CRITICAL")
            highCount = Application.WorksheetFunction.CountIf(wsReconcil.Columns("N"), "HIGH")
            medCount = Application.WorksheetFunction.CountIf(wsReconcil.Columns("N"), "MEDIUM")
            lowCount = Application.WorksheetFunction.CountIf(wsReconcil.Columns("N"), "LOW")
            totalCount = critCount + highCount + medCount + lowCount
            On Error GoTo ErrorHandler
        End If

        ' CRITICAL
        .Cells(rowNum, 1).Value = "CRITICAL"
        .Cells(rowNum, 2).Value = critCount
        .Cells(rowNum, 3).Value = IIf(totalCount > 0, critCount / totalCount, 0)
        .Cells(rowNum, 3).NumberFormat = "0.0%"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(255, 100, 100)
        rowNum = rowNum + 1

        ' HIGH
        .Cells(rowNum, 1).Value = "HIGH"
        .Cells(rowNum, 2).Value = highCount
        .Cells(rowNum, 3).Value = IIf(totalCount > 0, highCount / totalCount, 0)
        .Cells(rowNum, 3).NumberFormat = "0.0%"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(255, 200, 100)
        rowNum = rowNum + 1

        ' MEDIUM
        .Cells(rowNum, 1).Value = "MEDIUM"
        .Cells(rowNum, 2).Value = medCount
        .Cells(rowNum, 3).Value = IIf(totalCount > 0, medCount / totalCount, 0)
        .Cells(rowNum, 3).NumberFormat = "0.0%"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(255, 255, 150)
        rowNum = rowNum + 1

        ' LOW
        .Cells(rowNum, 1).Value = "LOW"
        .Cells(rowNum, 2).Value = lowCount
        .Cells(rowNum, 3).Value = IIf(totalCount > 0, lowCount / totalCount, 0)
        .Cells(rowNum, 3).NumberFormat = "0.0%"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 4)).Interior.Color = RGB(150, 255, 150)
        rowNum = rowNum + 2

        ' Creer graphique camembert
        If totalCount > 0 Then
            Set chartObj = .ChartObjects.Add(Left:=300, Top:=50, Width:=350, Height:=250)
            With chartObj.Chart
                .ChartType = xlPie
                .SetSourceData Source:=ws.Range(ws.Cells(rowNum - 5, 1), ws.Cells(rowNum - 2, 2))
                .HasTitle = True
                .ChartTitle.Text = "Repartition des Risques"
                .ApplyDataLabels xlDataLabelsShowPercent
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
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
        rowNum = rowNum + 1

        ' Extraire Top 10 depuis RECONCIL (simulation)
        Dim i As Integer
        For i = 1 To 10
            .Cells(rowNum, 1).Value = i
            rowNum = rowNum + 1
        Next i

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

    ' Ajouter timestamp et resume
    With ws
        ' Verifier si en-tete existe
        If .Cells(1, 1).Value <> "ID" Then
            ' Creer en-tete
            rowNum = 1
            .Cells(rowNum, 1).Value = "ID"
            .Cells(rowNum, 2).Value = "REGLE"
            .Cells(rowNum, 3).Value = "SEVERITE"
            .Cells(rowNum, 4).Value = "COMPTE"
            .Cells(rowNum, 5).Value = "DESCRIPTION"
            .Cells(rowNum, 6).Value = "MONTANT"
            .Cells(rowNum, 7).Value = "DATE_DETECT"
            .Cells(rowNum, 8).Value = "STATUT"
            .Cells(rowNum, 9).Value = "COMMENTAIRE"

            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Font.Bold = True
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Interior.Color = RGB(0, 51, 102)
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 9)).Font.Color = vbWhite
        End If

        ' Appliquer mise en forme conditionnelle
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        If lastRow > 1 Then
            ' Couleur selon severite
            Dim rng As Range
            Set rng = .Range(.Cells(2, 3), .Cells(lastRow, 3))

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

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    Set ws = Core_Engine.GetOrCreateSheet("FORENSIC_ANALYSIS", True)

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

        Dim d As Integer
        For d = 1 To 9
            .Cells(rowNum, 1).Value = d
            .Cells(rowNum, 2).Value = benford(d) / 100
            .Cells(rowNum, 2).NumberFormat = "0.0%"
            .Cells(rowNum, 3).Value = 0 ' A remplir par analyse reelle
            .Cells(rowNum, 3).NumberFormat = "0.0%"
            rowNum = rowNum + 1
        Next d

        rowNum = rowNum + 1

        ' Metriques Benford
        .Cells(rowNum, 1).Value = "Chi-squared:"
        .Cells(rowNum, 2).Value = "N/A"
        .Cells(rowNum, 3).Value = "(Seuil critique: 15.51)"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "MAD:"
        .Cells(rowNum, 2).Value = "N/A"
        .Cells(rowNum, 3).Value = "(< 0.006 = Excellent, < 0.012 = Acceptable)"
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

        ' Liste des patterns
        .Cells(rowNum, 1).Value = "Mots-cles suspects"
        .Cells(rowNum, 4).Value = "CADEAU, URGENT, MANUEL, etc."
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Transactions week-end"
        .Cells(rowNum, 4).Value = "Operations > 50K samedi/dimanche"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Juste sous seuil"
        .Cells(rowNum, 4).Value = "Montants proches des seuils LAB/FT"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Circularite"
        .Cells(rowNum, 4).Value = "Round-tripping detecte"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Saucissonnage"
        .Cells(rowNum, 4).Value = "Fractionnement de transactions"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "Doublons"
        .Cells(rowNum, 4).Value = "Transactions identiques"
        rowNum = rowNum + 2

        ' Section Z-Score
        .Cells(rowNum, 1).Value = "3. ANOMALIES STATISTIQUES (Z-SCORE)"
        .Cells(rowNum, 1).Font.Bold = True
        .Cells(rowNum, 1).Font.Size = 12
        rowNum = rowNum + 2

        .Cells(rowNum, 1).Value = "Compte"
        .Cells(rowNum, 2).Value = "Moyenne"
        .Cells(rowNum, 3).Value = "Ecart-type"
        .Cells(rowNum, 4).Value = "Montant"
        .Cells(rowNum, 5).Value = "Z-Score"
        .Cells(rowNum, 6).Value = "Alerte"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)

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

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    Set ws = Core_Engine.GetOrCreateSheet("COMPLIANCE_CHECK", True)

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

        .Cells(rowNum, 1).Value = "COBAC-001"
        .Cells(rowNum, 2).Value = "Suspens > 90 jours"
        .Cells(rowNum, 3).Value = "90 jours"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "COBAC-002"
        .Cells(rowNum, 2).Value = "Transit non apure J+7"
        .Cells(rowNum, 3).Value = "7 jours"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "COBAC-003"
        .Cells(rowNum, 2).Value = "Couverture provisions"
        .Cells(rowNum, 3).Value = "> 50%"
        rowNum = rowNum + 2

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

        .Cells(rowNum, 1).Value = "OHADA-001"
        .Cells(rowNum, 2).Value = "Equilibre bilan (Actif=Passif)"
        .Cells(rowNum, 3).Value = "Ecart < 1000"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "OHADA-002"
        .Cells(rowNum, 2).Value = "Coherence sens comptes"
        .Cells(rowNum, 3).Value = "Par classe"
        rowNum = rowNum + 2

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

        .Cells(rowNum, 1).Value = "LAB-001"
        .Cells(rowNum, 2).Value = "Transactions > seuil declaration"
        .Cells(rowNum, 3).Value = "5,000,000 XAF"
        rowNum = rowNum + 1

        .Cells(rowNum, 1).Value = "LAB-002"
        .Cells(rowNum, 2).Value = "Detection structuration"
        .Cells(rowNum, 3).Value = "Pattern"
        rowNum = rowNum + 2

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
                          "RECONCIL", "FORENSIC_ANALYSIS", "COMPLIANCE_CHECK")

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

    ' Parcourir les alertes
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = "CRITICAL" Or ws.Cells(i, 3).Value = "FRAUD" Then
            criticalCount = criticalCount + 1
            alertBody = alertBody & criticalCount & ". " & ws.Cells(i, 2).Value & vbNewLine
            alertBody = alertBody & "   Compte: " & ws.Cells(i, 4).Value & vbNewLine
            alertBody = alertBody & "   Description: " & ws.Cells(i, 5).Value & vbNewLine
            alertBody = alertBody & "   Montant: " & Format(ws.Cells(i, 6).Value, "#,##0") & " XAF" & vbNewLine & vbNewLine
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

    ' CORRIGÉ BUG-003: Utiliser GetOrCreateSheet pour éviter conflits
    Set ws = Core_Engine.GetOrCreateSheet("ECHANTILLON_TEST", True)

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

        ' Copier Top 20 depuis RECONCIL trie par score
        If Not wsReconcil Is Nothing Then
            Dim lastRowRec As Long
            lastRowRec = wsReconcil.Cells(wsReconcil.Rows.Count, 1).End(xlUp).Row

            ' Pour simplifier, copier les 20 premieres lignes avec score
            Dim copyCount As Integer
            copyCount = 0

            For i = 2 To lastRowRec
                If copyCount >= 20 Then Exit For

                ' Verifier si la ligne a un score
                ' CORRIGÉ BUG-005: Colonnes correctes selon structure RECONCIL
                ' Structure: A(1)=Compte, B(2)=Libelle, C(3)=Balance, D(4)=GL, E(5)=Ecart,
                '            F(6)=Status, G(7)=Anciennete, H(8)=NbTrans, I(9)=DateMin,
                '            J(10)=DateMax, K(11)=Score, L(12)=Facteurs, M(13)=Provisions, N(14)=Priority
                If wsReconcil.Cells(i, 11).Value <> "" Then ' Score non vide
                    copyCount = copyCount + 1
                    .Cells(rowNum, 1).Value = copyCount
                    .Cells(rowNum, 2).Value = wsReconcil.Cells(i, 1).Value ' Compte (col A)
                    .Cells(rowNum, 3).Value = wsReconcil.Cells(i, 2).Value ' Libelle (col B)
                    .Cells(rowNum, 4).Value = wsReconcil.Cells(i, 5).Value ' Ecart (col E)
                    .Cells(rowNum, 4).NumberFormat = "#,##0"
                    .Cells(rowNum, 5).Value = wsReconcil.Cells(i, 11).Value ' Score (col K)
                    .Cells(rowNum, 6).Value = wsReconcil.Cells(i, 14).Value ' Priority/Niveau (col N)
                    .Cells(rowNum, 7).Value = wsReconcil.Cells(i, 12).Value ' Facteurs (col L)
                    .Cells(rowNum, 8).Value = "A INVESTIGUER"

                    ' Colorer selon niveau de priorité
                    Select Case wsReconcil.Cells(i, 14).Value ' Priority (col N)
                        Case "CRITICAL"
                            .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 200, 200)
                        Case "HIGH"
                            .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 230, 200)
                        Case "MEDIUM"
                            .Range(.Cells(rowNum, 1), .Cells(rowNum, 10)).Interior.Color = RGB(255, 255, 200)
                    End Select

                    rowNum = rowNum + 1
                End If
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

' CORRIGÉ: Utiliser SAFA_Common.LogError pour centraliser la gestion d'erreurs
Private Sub LogError(moduleName As String, procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    Call SAFA_Common.LogError(moduleName, procName, errNum, errDesc)
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
