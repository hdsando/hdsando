Attribute VB_Name = "Config_Manager"
'===============================================================================
' MODULE: Config_Manager
' Description: Gestion de la configuration externalisee (JSON)
' Version: 10.0
' Auteur: S.A.F.A Team
' Date: Decembre 2024
'===============================================================================
Option Explicit

' ===== CONSTANTES =====
Private Const MODULE_NAME As String = "Config_Manager"
Private Const CONFIG_SHEET_NAME As String = "CONFIG_DATA"
Private Const DEFAULT_CONFIG_FILE As String = "settings.json"
Private Const DEFAULT_KEYWORDS_FILE As String = "fraud_keywords.json"

' ===== TYPES PERSONNALISES =====
Public Type GeneralConfig
    ApplicationName As String
    Version As String
    DefaultCurrency As String
    DefaultTolerance As Double
    MaxRowsMemory As Long
    DefaultSolId As String
End Type

Public Type ForensicConfig
    BenfordChiSquaredCritical As Double
    BenfordMADExcellent As Double
    BenfordMADAcceptable As Double
    BenfordMADMarginal As Double
    BenfordMinSampleSize As Long
    ZScoreWarning As Double
    ZScoreCritical As Double
    ZScoreMinTransactions As Long
    StructuringThreshold As Double
    StructuringMinCount As Long
    LayeringThreshold As Double
    LayeringMinCount As Long
    WeekendThreshold As Double
End Type

Public Type RegulatoryConfig
    CobacSuspensLimitDays As Long
    CobacTransitLimitDays As Long
    BeacReserveRatio As Double
    LabftDeclarationThreshold As Double
    LabftStructuringThreshold As Double
    PrudentialLargeExposureRatio As Double
End Type

Public Type RiskScoringConfig
    WeightAmount As Integer
    WeightAge As Integer
    WeightOrphan As Integer
    WeightVolume As Integer
    WeightSensitivity As Integer
    ThresholdCritical As Integer
    ThresholdHigh As Integer
    ThresholdMedium As Integer
    AmountTier1 As Double
    AmountTier2 As Double
    AmountTier3 As Double
    AgeTier1 As Long
    AgeTier2 As Long
    AgeTier3 As Long
End Type

Public Type IFRS9Config
    Stage3Days As Long
    Stage3Rate As Double
    Stage2HighDays As Long
    Stage2HighRate As Double
    Stage2LowDays As Long
    Stage2LowRate As Double
    Stage1HighDays As Long
    Stage1HighRate As Double
    Stage1LowRate As Double
End Type

Public Type UIConfig
    DefaultZoom As Integer
    ColorSuccess As String
    ColorWarning As String
    ColorError As String
    ColorPrimary As String
    ColorNeutral As String
End Type

Public Type FullConfig
    General As GeneralConfig
    Forensic As ForensicConfig
    Regulatory As RegulatoryConfig
    RiskScoring As RiskScoringConfig
    IFRS9 As IFRS9Config
    UI As UIConfig
End Type

' ===== VARIABLES MODULE =====
Private mConfig As FullConfig
Private mConfigLoaded As Boolean
Private mFraudKeywords As Collection
Private mExclusionPatterns As Collection

'===============================================================================
' FONCTION PRINCIPALE: LoadConfiguration
' Description: Charge la configuration depuis le fichier JSON ou les defauts
'===============================================================================
Public Function LoadConfiguration(Optional configPath As String = "") As Boolean
    On Error GoTo ErrorHandler

    ' Initialiser avec les valeurs par defaut
    SetDefaultConfiguration

    ' Tenter de charger depuis le fichier JSON
    If configPath = "" Then
        configPath = ThisWorkbook.Path & "\config\" & DEFAULT_CONFIG_FILE
    End If

    If Dir(configPath) <> "" Then
        LoadFromJSONFile configPath
    Else
        ' Tenter de charger depuis la feuille CONFIG_DATA
        LoadFromConfigSheet
    End If

    ' Charger les mots-cles de fraude
    LoadFraudKeywords

    mConfigLoaded = True
    LoadConfiguration = True

    Exit Function

ErrorHandler:
    LogError MODULE_NAME, "LoadConfiguration", Err.Number, Err.Description
    SetDefaultConfiguration
    mConfigLoaded = True
    LoadConfiguration = False
End Function

'===============================================================================
' FONCTION: SetDefaultConfiguration
' Description: Definit les valeurs par defaut de configuration
'===============================================================================
Private Sub SetDefaultConfiguration()
    ' General
    With mConfig.General
        .ApplicationName = "S.A.F.A"
        .Version = "10.0"
        .DefaultCurrency = "XAF"
        .DefaultTolerance = 100
        .MaxRowsMemory = 500000
        .DefaultSolId = "799"
    End With

    ' Forensic
    With mConfig.Forensic
        .BenfordChiSquaredCritical = 15.51
        .BenfordMADExcellent = 0.006
        .BenfordMADAcceptable = 0.012
        .BenfordMADMarginal = 0.015
        .BenfordMinSampleSize = 100
        .ZScoreWarning = 2.5
        .ZScoreCritical = 3.5
        .ZScoreMinTransactions = 5
        .StructuringThreshold = 50000
        .StructuringMinCount = 3
        .LayeringThreshold = 1000000
        .LayeringMinCount = 4
        .WeekendThreshold = 50000
    End With

    ' Regulatory
    With mConfig.Regulatory
        .CobacSuspensLimitDays = 90
        .CobacTransitLimitDays = 7
        .BeacReserveRatio = 0.07
        .LabftDeclarationThreshold = 5000000
        .LabftStructuringThreshold = 4500000
        .PrudentialLargeExposureRatio = 0.25
    End With

    ' Risk Scoring
    With mConfig.RiskScoring
        .WeightAmount = 30
        .WeightAge = 30
        .WeightOrphan = 20
        .WeightVolume = 10
        .WeightSensitivity = 10
        .ThresholdCritical = 70
        .ThresholdHigh = 50
        .ThresholdMedium = 30
        .AmountTier1 = 10000000
        .AmountTier2 = 1000000
        .AmountTier3 = 100000
        .AgeTier1 = 365
        .AgeTier2 = 180
        .AgeTier3 = 90
    End With

    ' IFRS 9
    With mConfig.IFRS9
        .Stage3Days = 360
        .Stage3Rate = 1#
        .Stage2HighDays = 180
        .Stage2HighRate = 0.5
        .Stage2LowDays = 90
        .Stage2LowRate = 0.25
        .Stage1HighDays = 30
        .Stage1HighRate = 0.1
        .Stage1LowRate = 0.01
    End With

    ' UI
    With mConfig.UI
        .DefaultZoom = 85
        .ColorSuccess = "#00C000"
        .ColorWarning = "#FF8000"
        .ColorError = "#FF0000"
        .ColorPrimary = "#003366"
        .ColorNeutral = "#808080"
    End With
End Sub

'===============================================================================
' FONCTION: LoadFromJSONFile
' Description: Charge la configuration depuis un fichier JSON
'===============================================================================
Private Sub LoadFromJSONFile(filePath As String)
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim jsonContent As String
    Dim lines() As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)

    jsonContent = ts.ReadAll
    ts.Close

    ' Parser simple JSON (sans bibliotheque externe)
    ' Extraction des valeurs principales
    mConfig.General.DefaultTolerance = ExtractJSONNumber(jsonContent, "default_tolerance", 100)
    mConfig.General.DefaultCurrency = ExtractJSONString(jsonContent, "default_currency", "XAF")
    mConfig.General.MaxRowsMemory = ExtractJSONNumber(jsonContent, "max_rows_memory", 500000)
    mConfig.General.DefaultSolId = ExtractJSONString(jsonContent, "default_sol_id", "799")

    ' Forensic
    mConfig.Forensic.BenfordChiSquaredCritical = ExtractJSONNumber(jsonContent, "chi_squared_critical_value", 15.51)
    mConfig.Forensic.BenfordMADMarginal = ExtractJSONNumber(jsonContent, "mad_threshold_marginal", 0.015)
    mConfig.Forensic.ZScoreWarning = ExtractJSONNumber(jsonContent, "threshold_warning", 2.5)
    mConfig.Forensic.ZScoreCritical = ExtractJSONNumber(jsonContent, "threshold_critical", 3.5)

    ' Regulatory
    mConfig.Regulatory.CobacSuspensLimitDays = ExtractJSONNumber(jsonContent, "suspens_limit_days", 90)
    mConfig.Regulatory.CobacTransitLimitDays = ExtractJSONNumber(jsonContent, "transit_limit_days", 7)
    mConfig.Regulatory.LabftDeclarationThreshold = ExtractJSONNumber(jsonContent, "declaration_threshold_xaf", 5000000)

    Set fso = Nothing

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "LoadFromJSONFile", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: LoadFromConfigSheet
' Description: Charge la configuration depuis une feuille Excel cachee
'===============================================================================
Private Sub LoadFromConfigSheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim paramName As String
    Dim paramValue As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        paramName = UCase(Trim(ws.Cells(i, 1).Value))
        paramValue = ws.Cells(i, 2).Value

        Select Case paramName
            ' General
            Case "DEFAULT_TOLERANCE"
                mConfig.General.DefaultTolerance = CDbl(paramValue)
            Case "DEFAULT_CURRENCY"
                mConfig.General.DefaultCurrency = CStr(paramValue)
            Case "MAX_ROWS_MEMORY"
                mConfig.General.MaxRowsMemory = CLng(paramValue)
            Case "DEFAULT_SOL_ID"
                mConfig.General.DefaultSolId = CStr(paramValue)

            ' Forensic
            Case "BENFORD_CHI_CRITICAL"
                mConfig.Forensic.BenfordChiSquaredCritical = CDbl(paramValue)
            Case "BENFORD_MAD_MARGINAL"
                mConfig.Forensic.BenfordMADMarginal = CDbl(paramValue)
            Case "ZSCORE_WARNING"
                mConfig.Forensic.ZScoreWarning = CDbl(paramValue)
            Case "ZSCORE_CRITICAL"
                mConfig.Forensic.ZScoreCritical = CDbl(paramValue)
            Case "STRUCTURING_THRESHOLD"
                mConfig.Forensic.StructuringThreshold = CDbl(paramValue)
            Case "LAYERING_THRESHOLD"
                mConfig.Forensic.LayeringThreshold = CDbl(paramValue)

            ' Regulatory
            Case "COBAC_SUSPENS_DAYS"
                mConfig.Regulatory.CobacSuspensLimitDays = CLng(paramValue)
            Case "COBAC_TRANSIT_DAYS"
                mConfig.Regulatory.CobacTransitLimitDays = CLng(paramValue)
            Case "LABFT_THRESHOLD"
                mConfig.Regulatory.LabftDeclarationThreshold = CDbl(paramValue)

            ' Risk Scoring
            Case "RISK_THRESHOLD_CRITICAL"
                mConfig.RiskScoring.ThresholdCritical = CInt(paramValue)
            Case "RISK_THRESHOLD_HIGH"
                mConfig.RiskScoring.ThresholdHigh = CInt(paramValue)
            Case "RISK_THRESHOLD_MEDIUM"
                mConfig.RiskScoring.ThresholdMedium = CInt(paramValue)
        End Select
    Next i

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "LoadFromConfigSheet", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTION: SaveConfigurationToSheet
' Description: Sauvegarde la configuration dans une feuille Excel
'===============================================================================
Public Sub SaveConfigurationToSheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rowNum As Long

    ' Supprimer si existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Creer la feuille
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = CONFIG_SHEET_NAME

    rowNum = 1

    With ws
        ' En-tetes
        .Cells(rowNum, 1).Value = "PARAMETRE"
        .Cells(rowNum, 2).Value = "VALEUR"
        .Cells(rowNum, 3).Value = "DESCRIPTION"
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 3)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 3)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 3)).Font.Color = vbWhite
        rowNum = rowNum + 1

        ' Section General
        .Cells(rowNum, 1).Value = "=== GENERAL ==="
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        AddConfigRow ws, rowNum, "DEFAULT_TOLERANCE", mConfig.General.DefaultTolerance, "Tolerance d'ecart (XAF)"
        AddConfigRow ws, rowNum, "DEFAULT_CURRENCY", mConfig.General.DefaultCurrency, "Devise par defaut"
        AddConfigRow ws, rowNum, "MAX_ROWS_MEMORY", mConfig.General.MaxRowsMemory, "Lignes max en memoire"
        AddConfigRow ws, rowNum, "DEFAULT_SOL_ID", mConfig.General.DefaultSolId, "Code agence par defaut"

        ' Section Forensic
        .Cells(rowNum, 1).Value = "=== FORENSIC ==="
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        AddConfigRow ws, rowNum, "BENFORD_CHI_CRITICAL", mConfig.Forensic.BenfordChiSquaredCritical, "Chi-squared seuil critique"
        AddConfigRow ws, rowNum, "BENFORD_MAD_MARGINAL", mConfig.Forensic.BenfordMADMarginal, "MAD seuil marginal"
        AddConfigRow ws, rowNum, "ZSCORE_WARNING", mConfig.Forensic.ZScoreWarning, "Z-Score seuil warning"
        AddConfigRow ws, rowNum, "ZSCORE_CRITICAL", mConfig.Forensic.ZScoreCritical, "Z-Score seuil critique"
        AddConfigRow ws, rowNum, "STRUCTURING_THRESHOLD", mConfig.Forensic.StructuringThreshold, "Seuil saucissonnage"
        AddConfigRow ws, rowNum, "LAYERING_THRESHOLD", mConfig.Forensic.LayeringThreshold, "Seuil circularite"
        AddConfigRow ws, rowNum, "WEEKEND_THRESHOLD", mConfig.Forensic.WeekendThreshold, "Seuil weekend"

        ' Section Regulatory
        .Cells(rowNum, 1).Value = "=== REGULATORY ==="
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        AddConfigRow ws, rowNum, "COBAC_SUSPENS_DAYS", mConfig.Regulatory.CobacSuspensLimitDays, "COBAC limite suspens (jours)"
        AddConfigRow ws, rowNum, "COBAC_TRANSIT_DAYS", mConfig.Regulatory.CobacTransitLimitDays, "COBAC limite transit (jours)"
        AddConfigRow ws, rowNum, "LABFT_THRESHOLD", mConfig.Regulatory.LabftDeclarationThreshold, "Seuil declaration LAB/FT"
        AddConfigRow ws, rowNum, "PRUDENTIAL_LARGE_EXPOSURE", mConfig.Regulatory.PrudentialLargeExposureRatio, "Ratio grand risque"

        ' Section Risk Scoring
        .Cells(rowNum, 1).Value = "=== RISK SCORING ==="
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1

        AddConfigRow ws, rowNum, "RISK_WEIGHT_AMOUNT", mConfig.RiskScoring.WeightAmount, "Poids facteur montant"
        AddConfigRow ws, rowNum, "RISK_WEIGHT_AGE", mConfig.RiskScoring.WeightAge, "Poids facteur age"
        AddConfigRow ws, rowNum, "RISK_WEIGHT_ORPHAN", mConfig.RiskScoring.WeightOrphan, "Poids facteur orphelin"
        AddConfigRow ws, rowNum, "RISK_THRESHOLD_CRITICAL", mConfig.RiskScoring.ThresholdCritical, "Seuil CRITICAL"
        AddConfigRow ws, rowNum, "RISK_THRESHOLD_HIGH", mConfig.RiskScoring.ThresholdHigh, "Seuil HIGH"
        AddConfigRow ws, rowNum, "RISK_THRESHOLD_MEDIUM", mConfig.RiskScoring.ThresholdMedium, "Seuil MEDIUM"

        ' Mise en forme
        .Columns("A:C").AutoFit
        .Visible = xlSheetVeryHidden

    End With

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "SaveConfigurationToSheet", Err.Number, Err.Description
End Sub

Private Sub AddConfigRow(ws As Worksheet, ByRef rowNum As Long, paramName As String, paramValue As Variant, description As String)
    ws.Cells(rowNum, 1).Value = paramName
    ws.Cells(rowNum, 2).Value = paramValue
    ws.Cells(rowNum, 3).Value = description
    rowNum = rowNum + 1
End Sub

'===============================================================================
' FONCTION: LoadFraudKeywords
' Description: Charge les mots-cles de fraude depuis le fichier JSON
'===============================================================================
Private Sub LoadFraudKeywords()
    On Error GoTo ErrorHandler

    Dim keywordsPath As String
    Dim fso As Object
    Dim ts As Object
    Dim jsonContent As String

    Set mFraudKeywords = New Collection
    Set mExclusionPatterns = New Collection

    keywordsPath = ThisWorkbook.Path & "\config\" & DEFAULT_KEYWORDS_FILE

    ' Mots-cles par defaut si fichier non trouve
    If Dir(keywordsPath) = "" Then
        SetDefaultFraudKeywords
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(keywordsPath, 1, False)
    jsonContent = ts.ReadAll
    ts.Close

    ' Extraction des mots-cles (parser simple)
    ExtractKeywordsFromJSON jsonContent

    Set fso = Nothing

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "LoadFraudKeywords", Err.Number, Err.Description
    SetDefaultFraudKeywords
End Sub

Private Sub SetDefaultFraudKeywords()
    ' High Risk
    mFraudKeywords.Add "CADEAU", "HR1"
    mFraudKeywords.Add "GIFT", "HR2"
    mFraudKeywords.Add "GRATIFICATION", "HR3"
    mFraudKeywords.Add "BAKCHICH", "HR4"
    mFraudKeywords.Add "POT DE VIN", "HR5"

    ' Medium Risk
    mFraudKeywords.Add "URGENT", "MR1"
    mFraudKeywords.Add "EMERGENCY", "MR2"
    mFraudKeywords.Add "MANUEL", "MR3"
    mFraudKeywords.Add "EXCEPTION", "MR4"
    mFraudKeywords.Add "OVERRIDE", "MR5"
    mFraudKeywords.Add "BYPASS", "MR6"
    mFraudKeywords.Add "VIP", "MR7"
    mFraudKeywords.Add "DIRECTEUR", "MR8"

    ' Correction patterns
    mFraudKeywords.Add "CORRECTION", "CP1"
    mFraudKeywords.Add "AJUSTEMENT", "CP2"
    mFraudKeywords.Add "ANNULATION", "CP3"
    mFraudKeywords.Add "EXTOURNE", "CP4"

    ' Exclusions
    mExclusionPatterns.Add "AUTO", "EX1"
    mExclusionPatterns.Add "AUTOMATIQUE", "EX2"
    mExclusionPatterns.Add "SYSTEM", "EX3"
    mExclusionPatterns.Add "BATCH", "EX4"
End Sub

Private Sub ExtractKeywordsFromJSON(jsonContent As String)
    On Error Resume Next

    ' Parser simplifie - extraction des termes entre guillemets apres "terms"
    Dim pos As Long
    Dim endPos As Long
    Dim term As String
    Dim counter As Long

    counter = 0

    ' Chercher les termes high_risk
    pos = InStr(1, jsonContent, """high_risk""", vbTextCompare)
    If pos > 0 Then
        ExtractTermsFromSection jsonContent, pos, "HR", counter
    End If

    ' Chercher les termes medium_risk
    pos = InStr(1, jsonContent, """medium_risk""", vbTextCompare)
    If pos > 0 Then
        ExtractTermsFromSection jsonContent, pos, "MR", counter
    End If

    ' Chercher les termes correction_patterns
    pos = InStr(1, jsonContent, """correction_patterns""", vbTextCompare)
    If pos > 0 Then
        ExtractTermsFromSection jsonContent, pos, "CP", counter
    End If

    ' Chercher exclusion_patterns
    pos = InStr(1, jsonContent, """exclusion_patterns""", vbTextCompare)
    If pos > 0 Then
        ExtractTermsFromSection jsonContent, pos, "EX", counter, True
    End If
End Sub

Private Sub ExtractTermsFromSection(jsonContent As String, startPos As Long, prefix As String, ByRef counter As Long, Optional isExclusion As Boolean = False)
    On Error Resume Next

    Dim termsPos As Long
    Dim bracketEnd As Long
    Dim term As String
    Dim termStart As Long
    Dim termEnd As Long
    Dim searchPos As Long

    ' Trouver "terms": [
    termsPos = InStr(startPos, jsonContent, """terms""", vbTextCompare)
    If termsPos = 0 Then Exit Sub

    bracketEnd = InStr(termsPos, jsonContent, "]")
    If bracketEnd = 0 Then Exit Sub

    ' Extraire les termes entre guillemets
    searchPos = termsPos
    Do
        termStart = InStr(searchPos, jsonContent, """")
        If termStart = 0 Or termStart > bracketEnd Then Exit Do

        termEnd = InStr(termStart + 1, jsonContent, """")
        If termEnd = 0 Or termEnd > bracketEnd Then Exit Do

        term = Mid(jsonContent, termStart + 1, termEnd - termStart - 1)

        ' Ignorer les mots-cles JSON
        If term <> "terms" And Len(term) > 2 Then
            counter = counter + 1
            If isExclusion Then
                mExclusionPatterns.Add term, prefix & counter
            Else
                mFraudKeywords.Add term, prefix & counter
            End If
        End If

        searchPos = termEnd + 1
    Loop
End Sub

'===============================================================================
' FONCTIONS D'ACCES AUX CONFIGURATIONS
'===============================================================================

Public Function GetConfig() As FullConfig
    If Not mConfigLoaded Then LoadConfiguration
    GetConfig = mConfig
End Function

Public Function GetGeneralConfig() As GeneralConfig
    If Not mConfigLoaded Then LoadConfiguration
    GetGeneralConfig = mConfig.General
End Function

Public Function GetForensicConfig() As ForensicConfig
    If Not mConfigLoaded Then LoadConfiguration
    GetForensicConfig = mConfig.Forensic
End Function

Public Function GetRegulatoryConfig() As RegulatoryConfig
    If Not mConfigLoaded Then LoadConfiguration
    GetRegulatoryConfig = mConfig.Regulatory
End Function

Public Function GetRiskScoringConfig() As RiskScoringConfig
    If Not mConfigLoaded Then LoadConfiguration
    GetRiskScoringConfig = mConfig.RiskScoring
End Function

Public Function GetIFRS9Config() As IFRS9Config
    If Not mConfigLoaded Then LoadConfiguration
    GetIFRS9Config = mConfig.IFRS9
End Function

Public Function GetFraudKeywords() As Collection
    If mFraudKeywords Is Nothing Then LoadFraudKeywords
    Set GetFraudKeywords = mFraudKeywords
End Function

Public Function GetExclusionPatterns() As Collection
    If mExclusionPatterns Is Nothing Then LoadFraudKeywords
    Set GetExclusionPatterns = mExclusionPatterns
End Function

'===============================================================================
' FONCTIONS DE MODIFICATION
'===============================================================================

Public Sub SetConfigValue(section As String, paramName As String, paramValue As Variant)
    On Error GoTo ErrorHandler

    Select Case UCase(section)
        Case "GENERAL"
            Select Case UCase(paramName)
                Case "DEFAULT_TOLERANCE"
                    mConfig.General.DefaultTolerance = CDbl(paramValue)
                Case "DEFAULT_CURRENCY"
                    mConfig.General.DefaultCurrency = CStr(paramValue)
                Case "MAX_ROWS_MEMORY"
                    mConfig.General.MaxRowsMemory = CLng(paramValue)
                Case "DEFAULT_SOL_ID"
                    mConfig.General.DefaultSolId = CStr(paramValue)
            End Select

        Case "FORENSIC"
            Select Case UCase(paramName)
                Case "ZSCORE_WARNING"
                    mConfig.Forensic.ZScoreWarning = CDbl(paramValue)
                Case "ZSCORE_CRITICAL"
                    mConfig.Forensic.ZScoreCritical = CDbl(paramValue)
                Case "BENFORD_MAD_MARGINAL"
                    mConfig.Forensic.BenfordMADMarginal = CDbl(paramValue)
            End Select

        Case "REGULATORY"
            Select Case UCase(paramName)
                Case "COBAC_SUSPENS_DAYS"
                    mConfig.Regulatory.CobacSuspensLimitDays = CLng(paramValue)
                Case "LABFT_THRESHOLD"
                    mConfig.Regulatory.LabftDeclarationThreshold = CDbl(paramValue)
            End Select

        Case "RISK_SCORING"
            Select Case UCase(paramName)
                Case "THRESHOLD_CRITICAL"
                    mConfig.RiskScoring.ThresholdCritical = CInt(paramValue)
                Case "THRESHOLD_HIGH"
                    mConfig.RiskScoring.ThresholdHigh = CInt(paramValue)
                Case "THRESHOLD_MEDIUM"
                    mConfig.RiskScoring.ThresholdMedium = CInt(paramValue)
            End Select
    End Select

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "SetConfigValue", Err.Number, Err.Description
End Sub

'===============================================================================
' FONCTIONS UTILITAIRES JSON
'===============================================================================

Private Function ExtractJSONString(jsonContent As String, key As String, defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    pos = InStr(1, jsonContent, """" & key & """", vbTextCompare)
    If pos = 0 Then
        ExtractJSONString = defaultValue
        Exit Function
    End If

    ' Chercher le : puis la valeur
    valueStart = InStr(pos, jsonContent, ":")
    If valueStart = 0 Then
        ExtractJSONString = defaultValue
        Exit Function
    End If

    ' Chercher les guillemets
    valueStart = InStr(valueStart, jsonContent, """")
    If valueStart = 0 Then
        ExtractJSONString = defaultValue
        Exit Function
    End If

    valueEnd = InStr(valueStart + 1, jsonContent, """")
    If valueEnd = 0 Then
        ExtractJSONString = defaultValue
        Exit Function
    End If

    ExtractJSONString = Mid(jsonContent, valueStart + 1, valueEnd - valueStart - 1)
    Exit Function

ErrorHandler:
    ExtractJSONString = defaultValue
End Function

Private Function ExtractJSONNumber(jsonContent As String, key As String, defaultValue As Double) As Double
    On Error GoTo ErrorHandler

    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim valueStr As String
    Dim i As Long

    pos = InStr(1, jsonContent, """" & key & """", vbTextCompare)
    If pos = 0 Then
        ExtractJSONNumber = defaultValue
        Exit Function
    End If

    ' Chercher le :
    valueStart = InStr(pos, jsonContent, ":")
    If valueStart = 0 Then
        ExtractJSONNumber = defaultValue
        Exit Function
    End If

    valueStart = valueStart + 1

    ' Sauter les espaces
    Do While Mid(jsonContent, valueStart, 1) = " " Or Mid(jsonContent, valueStart, 1) = vbTab
        valueStart = valueStart + 1
    Loop

    ' Extraire la valeur numerique
    valueStr = ""
    For i = valueStart To Len(jsonContent)
        Dim c As String
        c = Mid(jsonContent, i, 1)
        If c Like "[0-9.]" Or c = "-" Then
            valueStr = valueStr & c
        Else
            Exit For
        End If
    Next i

    If valueStr = "" Then
        ExtractJSONNumber = defaultValue
    Else
        ExtractJSONNumber = CDbl(valueStr)
    End If

    Exit Function

ErrorHandler:
    ExtractJSONNumber = defaultValue
End Function

'===============================================================================
' FONCTION: ValidateConfiguration
' Description: Valide les parametres de configuration
'===============================================================================
Public Function ValidateConfiguration() As Boolean
    On Error GoTo ErrorHandler

    Dim errors As Collection
    Set errors = New Collection

    ' Validation General
    If mConfig.General.DefaultTolerance < 0 Then
        errors.Add "DEFAULT_TOLERANCE doit etre >= 0"
    End If

    If mConfig.General.MaxRowsMemory < 1000 Then
        errors.Add "MAX_ROWS_MEMORY doit etre >= 1000"
    End If

    ' Validation Forensic
    If mConfig.Forensic.ZScoreWarning <= 0 Then
        errors.Add "ZSCORE_WARNING doit etre > 0"
    End If

    If mConfig.Forensic.ZScoreCritical <= mConfig.Forensic.ZScoreWarning Then
        errors.Add "ZSCORE_CRITICAL doit etre > ZSCORE_WARNING"
    End If

    If mConfig.Forensic.BenfordChiSquaredCritical <= 0 Then
        errors.Add "BENFORD_CHI_CRITICAL doit etre > 0"
    End If

    ' Validation Risk Scoring
    Dim totalWeight As Integer
    totalWeight = mConfig.RiskScoring.WeightAmount + mConfig.RiskScoring.WeightAge + _
                 mConfig.RiskScoring.WeightOrphan + mConfig.RiskScoring.WeightVolume + _
                 mConfig.RiskScoring.WeightSensitivity

    If totalWeight <> 100 Then
        errors.Add "La somme des poids Risk Scoring doit etre 100 (actuel: " & totalWeight & ")"
    End If

    If mConfig.RiskScoring.ThresholdCritical <= mConfig.RiskScoring.ThresholdHigh Then
        errors.Add "THRESHOLD_CRITICAL doit etre > THRESHOLD_HIGH"
    End If

    ' Afficher les erreurs
    If errors.Count > 0 Then
        Dim errMsg As String
        Dim err As Variant
        errMsg = "Erreurs de configuration detectees:" & vbNewLine & vbNewLine

        For Each err In errors
            errMsg = errMsg & "- " & err & vbNewLine
        Next err

        MsgBox errMsg, vbExclamation, "Validation Configuration"
        ValidateConfiguration = False
    Else
        ValidateConfiguration = True
    End If

    Exit Function

ErrorHandler:
    LogError MODULE_NAME, "ValidateConfiguration", Err.Number, Err.Description
    ValidateConfiguration = False
End Function

'===============================================================================
' FONCTION: ExportConfigurationToJSON
' Description: Exporte la configuration actuelle vers un fichier JSON
'===============================================================================
Public Sub ExportConfigurationToJSON(Optional filePath As String = "")
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim jsonStr As String

    If filePath = "" Then
        filePath = ThisWorkbook.Path & "\config\settings_export.json"
    End If

    ' Construire le JSON
    jsonStr = "{" & vbNewLine
    jsonStr = jsonStr & "  ""general"": {" & vbNewLine
    jsonStr = jsonStr & "    ""default_tolerance"": " & mConfig.General.DefaultTolerance & "," & vbNewLine
    jsonStr = jsonStr & "    ""default_currency"": """ & mConfig.General.DefaultCurrency & """," & vbNewLine
    jsonStr = jsonStr & "    ""max_rows_memory"": " & mConfig.General.MaxRowsMemory & "," & vbNewLine
    jsonStr = jsonStr & "    ""default_sol_id"": """ & mConfig.General.DefaultSolId & """" & vbNewLine
    jsonStr = jsonStr & "  }," & vbNewLine

    jsonStr = jsonStr & "  ""forensic"": {" & vbNewLine
    jsonStr = jsonStr & "    ""zscore_warning"": " & mConfig.Forensic.ZScoreWarning & "," & vbNewLine
    jsonStr = jsonStr & "    ""zscore_critical"": " & mConfig.Forensic.ZScoreCritical & "," & vbNewLine
    jsonStr = jsonStr & "    ""benford_chi_critical"": " & mConfig.Forensic.BenfordChiSquaredCritical & "," & vbNewLine
    jsonStr = jsonStr & "    ""benford_mad_marginal"": " & mConfig.Forensic.BenfordMADMarginal & vbNewLine
    jsonStr = jsonStr & "  }," & vbNewLine

    jsonStr = jsonStr & "  ""regulatory"": {" & vbNewLine
    jsonStr = jsonStr & "    ""cobac_suspens_days"": " & mConfig.Regulatory.CobacSuspensLimitDays & "," & vbNewLine
    jsonStr = jsonStr & "    ""cobac_transit_days"": " & mConfig.Regulatory.CobacTransitLimitDays & "," & vbNewLine
    jsonStr = jsonStr & "    ""labft_threshold"": " & mConfig.Regulatory.LabftDeclarationThreshold & vbNewLine
    jsonStr = jsonStr & "  }," & vbNewLine

    jsonStr = jsonStr & "  ""risk_scoring"": {" & vbNewLine
    jsonStr = jsonStr & "    ""threshold_critical"": " & mConfig.RiskScoring.ThresholdCritical & "," & vbNewLine
    jsonStr = jsonStr & "    ""threshold_high"": " & mConfig.RiskScoring.ThresholdHigh & "," & vbNewLine
    jsonStr = jsonStr & "    ""threshold_medium"": " & mConfig.RiskScoring.ThresholdMedium & vbNewLine
    jsonStr = jsonStr & "  }" & vbNewLine

    jsonStr = jsonStr & "}"

    ' Ecrire le fichier
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write jsonStr
    ts.Close

    MsgBox "Configuration exportee: " & filePath, vbInformation, "Export"

    Set fso = Nothing

    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "ExportConfigurationToJSON", Err.Number, Err.Description
End Sub

' CORRIGÉ: Utiliser SAFA_Common.LogError pour centraliser la gestion d'erreurs
Private Sub LogError(moduleName As String, procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    ' Déléguer à SAFA_Common si disponible, sinon Debug.Print
    Call SAFA_Common.LogError(moduleName, procName, errNum, errDesc)
End Sub
