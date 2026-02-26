VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings
   Caption         =   "S.A.F.A v10.0 - Parametres"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' FORMULAIRE: frmSettings
' Description: Interface de configuration des parametres S.A.F.A
' Version: 10.0
' Auteur: S.A.F.A Team
' Date: Decembre 2024
'===============================================================================
Option Explicit

' ===== CONSTANTES =====
Private Const FORM_NAME As String = "frmSettings"

' ===== VARIABLES =====
Private mSettingsChanged As Boolean

'===============================================================================
' EVENEMENT: UserForm_Initialize
' Description: Initialisation du formulaire
'===============================================================================
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Configuration du formulaire
    Me.Caption = "S.A.F.A v10.0 - Parametres"
    Me.Width = 500
    Me.Height = 520

    ' Charger la configuration actuelle
    LoadCurrentSettings

    mSettingsChanged = False
End Sub

'===============================================================================
' PROCEDURE: LoadCurrentSettings
' Description: Charge les parametres actuels dans les controles
'===============================================================================
Private Sub LoadCurrentSettings()
    On Error Resume Next

    Dim cfg As FullConfig
    cfg = Config_Manager.GetConfig()

    ' Onglet General
    Me.txtTolerance.Value = cfg.General.DefaultTolerance
    Me.txtSolId.Value = cfg.General.DefaultSolId
    Me.cboDevise.Value = cfg.General.DefaultCurrency
    Me.txtMaxRows.Value = cfg.General.MaxRowsMemory

    ' Onglet Forensic
    Me.txtZScoreWarning.Value = cfg.Forensic.ZScoreWarning
    Me.txtZScoreCritical.Value = cfg.Forensic.ZScoreCritical
    Me.txtBenfordMAD.Value = cfg.Forensic.BenfordMADMarginal
    Me.txtStructuringThreshold.Value = cfg.Forensic.StructuringThreshold
    Me.txtWeekendThreshold.Value = cfg.Forensic.WeekendThreshold

    ' Onglet Regulatory
    Me.txtCobacSuspens.Value = cfg.Regulatory.CobacSuspensLimitDays
    Me.txtCobacTransit.Value = cfg.Regulatory.CobacTransitLimitDays
    Me.txtLabftThreshold.Value = cfg.Regulatory.LabftDeclarationThreshold

    ' Onglet Risk Scoring
    Me.txtRiskCritical.Value = cfg.RiskScoring.ThresholdCritical
    Me.txtRiskHigh.Value = cfg.RiskScoring.ThresholdHigh
    Me.txtRiskMedium.Value = cfg.RiskScoring.ThresholdMedium

    ' Populer liste des devises
    Me.cboDevise.Clear
    Me.cboDevise.AddItem "XAF"
    Me.cboDevise.AddItem "EUR"
    Me.cboDevise.AddItem "USD"
    Me.cboDevise.AddItem "XOF"
    Me.cboDevise.AddItem "GBP"
    Me.cboDevise.AddItem "NGN"
    Me.cboDevise.AddItem "GHS"
    Me.cboDevise.AddItem "KES"
    Me.cboDevise.AddItem "ZAR"

    Me.cboDevise.Value = cfg.General.DefaultCurrency
End Sub

'===============================================================================
' PROCEDURE: SaveSettings
' Description: Sauvegarde les parametres modifies
'===============================================================================
Private Sub SaveSettings()
    On Error GoTo ErrorHandler

    ' Valider les entrees
    If Not ValidateInputs() Then Exit Sub

    ' Sauvegarder General
    Config_Manager.SetConfigValue "GENERAL", "DEFAULT_TOLERANCE", CDbl(Me.txtTolerance.Value)
    Config_Manager.SetConfigValue "GENERAL", "DEFAULT_SOL_ID", Me.txtSolId.Value
    Config_Manager.SetConfigValue "GENERAL", "DEFAULT_CURRENCY", Me.cboDevise.Value
    Config_Manager.SetConfigValue "GENERAL", "MAX_ROWS_MEMORY", CLng(Me.txtMaxRows.Value)

    ' Sauvegarder Forensic
    Config_Manager.SetConfigValue "FORENSIC", "ZSCORE_WARNING", CDbl(Me.txtZScoreWarning.Value)
    Config_Manager.SetConfigValue "FORENSIC", "ZSCORE_CRITICAL", CDbl(Me.txtZScoreCritical.Value)
    Config_Manager.SetConfigValue "FORENSIC", "BENFORD_MAD_MARGINAL", CDbl(Me.txtBenfordMAD.Value)

    ' Sauvegarder Regulatory
    Config_Manager.SetConfigValue "REGULATORY", "COBAC_SUSPENS_DAYS", CLng(Me.txtCobacSuspens.Value)
    Config_Manager.SetConfigValue "REGULATORY", "LABFT_THRESHOLD", CDbl(Me.txtLabftThreshold.Value)

    ' Sauvegarder Risk Scoring
    Config_Manager.SetConfigValue "RISK_SCORING", "THRESHOLD_CRITICAL", CInt(Me.txtRiskCritical.Value)
    Config_Manager.SetConfigValue "RISK_SCORING", "THRESHOLD_HIGH", CInt(Me.txtRiskHigh.Value)
    Config_Manager.SetConfigValue "RISK_SCORING", "THRESHOLD_MEDIUM", CInt(Me.txtRiskMedium.Value)

    ' Sauvegarder dans la feuille cachee
    Config_Manager.SaveConfigurationToSheet

    MsgBox "Parametres sauvegardes avec succes!", vbInformation, "Sauvegarde"

    mSettingsChanged = False
    Me.Hide

    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de la sauvegarde: " & Err.Description, vbExclamation, "Erreur"
End Sub

'===============================================================================
' FUNCTION: ValidateInputs
' Description: Valide les entrees utilisateur
'===============================================================================
Private Function ValidateInputs() As Boolean
    On Error GoTo ErrorHandler

    Dim errors As String
    errors = ""

    ' Validation General
    If Not IsNumeric(Me.txtTolerance.Value) Or CDbl(Me.txtTolerance.Value) < 0 Then
        errors = errors & "- Tolerance doit etre un nombre >= 0" & vbNewLine
    End If

    If Len(Me.txtSolId.Value) <> 3 Then
        errors = errors & "- SOL ID doit contenir 3 caracteres" & vbNewLine
    End If

    ' Validation Forensic
    If Not IsNumeric(Me.txtZScoreWarning.Value) Or CDbl(Me.txtZScoreWarning.Value) <= 0 Then
        errors = errors & "- Z-Score Warning doit etre > 0" & vbNewLine
    End If

    If Not IsNumeric(Me.txtZScoreCritical.Value) Or CDbl(Me.txtZScoreCritical.Value) <= CDbl(Me.txtZScoreWarning.Value) Then
        errors = errors & "- Z-Score Critical doit etre > Z-Score Warning" & vbNewLine
    End If

    ' Validation Risk Scoring
    If Not IsNumeric(Me.txtRiskCritical.Value) Or CInt(Me.txtRiskCritical.Value) <= CInt(Me.txtRiskHigh.Value) Then
        errors = errors & "- Seuil Critical doit etre > Seuil High" & vbNewLine
    End If

    If Not IsNumeric(Me.txtRiskHigh.Value) Or CInt(Me.txtRiskHigh.Value) <= CInt(Me.txtRiskMedium.Value) Then
        errors = errors & "- Seuil High doit etre > Seuil Medium" & vbNewLine
    End If

    If errors <> "" Then
        MsgBox "Erreurs de validation:" & vbNewLine & vbNewLine & errors, vbExclamation, "Validation"
        ValidateInputs = False
    Else
        ValidateInputs = True
    End If

    Exit Function

ErrorHandler:
    MsgBox "Erreur de validation: " & Err.Description, vbExclamation, "Erreur"
    ValidateInputs = False
End Function

'===============================================================================
' EVENEMENT: btnSave_Click
' Description: Sauvegarde les parametres
'===============================================================================
Private Sub btnSave_Click()
    SaveSettings
End Sub

'===============================================================================
' EVENEMENT: btnCancel_Click
' Description: Annule les modifications
'===============================================================================
Private Sub btnCancel_Click()
    If mSettingsChanged Then
        If MsgBox("Des modifications non sauvegardees seront perdues." & vbNewLine & _
                  "Voulez-vous vraiment quitter?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
            Exit Sub
        End If
    End If
    Me.Hide
End Sub

'===============================================================================
' EVENEMENT: btnReset_Click
' Description: Reinitialise les parametres par defaut
'===============================================================================
Private Sub btnReset_Click()
    If MsgBox("Reinitialiser tous les parametres aux valeurs par defaut?", _
              vbQuestion + vbYesNo, "Confirmation") = vbYes Then

        ' Reinitialiser la configuration
        Config_Manager.LoadConfiguration ""

        ' Recharger dans le formulaire
        LoadCurrentSettings

        MsgBox "Parametres reinitialises.", vbInformation, "Reset"
        mSettingsChanged = True
    End If
End Sub

'===============================================================================
' EVENEMENT: btnExport_Click
' Description: Exporte la configuration en JSON
'===============================================================================
Private Sub btnExport_Click()
    On Error GoTo ErrorHandler

    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="safa_config_" & Format(Now, "yyyymmdd") & ".json", _
        FileFilter:="JSON Files (*.json),*.json", _
        Title:="Exporter la configuration")

    If filePath <> "False" Then
        Config_Manager.ExportConfigurationToJSON filePath
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de l'export: " & Err.Description, vbExclamation, "Erreur"
End Sub

'===============================================================================
' EVENEMENT: Changement de valeur
' Description: Marque les parametres comme modifies
'===============================================================================
Private Sub txtTolerance_Change()
    mSettingsChanged = True
End Sub

Private Sub txtSolId_Change()
    mSettingsChanged = True
End Sub

Private Sub cboDevise_Change()
    mSettingsChanged = True
End Sub

Private Sub txtZScoreWarning_Change()
    mSettingsChanged = True
End Sub

Private Sub txtZScoreCritical_Change()
    mSettingsChanged = True
End Sub

Private Sub txtRiskCritical_Change()
    mSettingsChanged = True
End Sub

Private Sub txtRiskHigh_Change()
    mSettingsChanged = True
End Sub

Private Sub txtRiskMedium_Change()
    mSettingsChanged = True
End Sub

'===============================================================================
' EVENEMENT: UserForm_QueryClose
' Description: Gestion de la fermeture du formulaire
'===============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And mSettingsChanged Then
        If MsgBox("Des modifications non sauvegardees seront perdues." & vbNewLine & _
                  "Voulez-vous vraiment quitter?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

'===============================================================================
' PROCEDURE: CreateControls
' Description: Creation des controles du formulaire (pour reference)
'===============================================================================
' Les controles suivants doivent etre crees dans l'editeur VBA:
'
' === ONGLET GENERAL ===
' - lblTolerance: Label "Tolerance (XAF):"
' - txtTolerance: TextBox pour la tolerance
' - lblSolId: Label "SOL ID par defaut:"
' - txtSolId: TextBox pour le SOL ID
' - lblDevise: Label "Devise par defaut:"
' - cboDevise: ComboBox pour la devise
' - lblMaxRows: Label "Max lignes memoire:"
' - txtMaxRows: TextBox pour max rows
'
' === ONGLET FORENSIC ===
' - lblZScoreWarning: Label "Z-Score Warning:"
' - txtZScoreWarning: TextBox
' - lblZScoreCritical: Label "Z-Score Critical:"
' - txtZScoreCritical: TextBox
' - lblBenfordMAD: Label "Benford MAD Marginal:"
' - txtBenfordMAD: TextBox
' - lblStructuringThreshold: Label "Seuil Structuration:"
' - txtStructuringThreshold: TextBox
' - lblWeekendThreshold: Label "Seuil Weekend:"
' - txtWeekendThreshold: TextBox
'
' === ONGLET REGULATORY ===
' - lblCobacSuspens: Label "COBAC Suspens (jours):"
' - txtCobacSuspens: TextBox
' - lblCobacTransit: Label "COBAC Transit (jours):"
' - txtCobacTransit: TextBox
' - lblLabftThreshold: Label "Seuil LAB/FT (XAF):"
' - txtLabftThreshold: TextBox
'
' === ONGLET RISK SCORING ===
' - lblRiskCritical: Label "Seuil CRITICAL:"
' - txtRiskCritical: TextBox
' - lblRiskHigh: Label "Seuil HIGH:"
' - txtRiskHigh: TextBox
' - lblRiskMedium: Label "Seuil MEDIUM:"
' - txtRiskMedium: TextBox
'
' === BOUTONS ===
' - btnSave: CommandButton "Sauvegarder"
' - btnCancel: CommandButton "Annuler"
' - btnReset: CommandButton "Reinitialiser"
' - btnExport: CommandButton "Exporter JSON"
