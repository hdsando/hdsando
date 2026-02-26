Attribute VB_Name = "FormBuilder"
'===============================================================================
' MODULE: FormBuilder
' Description: Generateur automatique des formulaires S.A.F.A
' Version: 10.0
' Usage: Executer CreateAllForms() une seule fois pour generer les interfaces
'===============================================================================
Option Explicit

Private Const FORM_WIDTH As Single = 560
Private Const FORM_HEIGHT As Single = 500

'===============================================================================
' PROCEDURE PRINCIPALE: CreateAllForms
' Description: Cree tous les formulaires necessaires
'===============================================================================
Public Sub CreateAllForms()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Creer le formulaire Cockpit
    Call CreateCockpitForm

    ' Creer le formulaire Login
    Call CreateLoginForm

    ' Creer le formulaire Settings
    Call CreateSettingsForm

    Application.ScreenUpdating = True

    MsgBox "Formulaires crees avec succes!" & vbCrLf & vbCrLf & _
           "Pour lancer S.A.F.A:" & vbCrLf & _
           "1. Appuyez sur ALT+F11 pour ouvrir l'editeur VBA" & vbCrLf & _
           "2. Double-cliquez sur USF_Cockpit" & vbCrLf & _
           "3. Appuyez sur F5 pour executer", _
           vbInformation, "S.A.F.A - Installation"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la creation des formulaires:" & vbCrLf & _
           Err.Description, vbCritical, "Erreur"
End Sub

'===============================================================================
' FONCTION: CreateCockpitForm
' Description: Cree le formulaire principal USF_Cockpit
'===============================================================================
Public Sub CreateCockpitForm()
    On Error GoTo ErrHandler

    Dim VBProj As Object
    Dim VBComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim topPos As Single

    Set VBProj = ThisWorkbook.VBProject

    ' Supprimer si existe
    On Error Resume Next
    VBProj.VBComponents.Remove VBProj.VBComponents("USF_Cockpit")
    On Error GoTo ErrHandler

    ' Creer le UserForm
    Set VBComp = VBProj.VBComponents.Add(3) ' vbext_ct_MSForm
    VBComp.Name = "USF_Cockpit"

    Set frm = VBComp.Designer

    ' Proprietes du formulaire
    With frm
        .Caption = "S.A.F.A v10.0 - Cockpit de Pilotage"
        .Width = FORM_WIDTH
        .Height = FORM_HEIGHT
        .BackColor = &HFFFFFF ' Blanc
        .StartUpPosition = 1 ' Center
    End With

    topPos = 10

    ' ===== TITRE =====
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTitle", True)
    With ctrl
        .Caption = "S.A.F.A v10.0"
        .Left = 10: .Top = topPos: .Width = 540: .Height = 30
        .Font.Size = 18: .Font.Bold = True
        .TextAlign = 2 ' Center
        .ForeColor = &H663300 ' Bleu fonce
        .BackColor = &HFFFFFF
    End With

    topPos = topPos + 35

    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSubtitle", True)
    With ctrl
        .Caption = "System for Automated Financial Audit"
        .Left = 10: .Top = topPos: .Width = 540: .Height = 18
        .Font.Size = 10: .Font.Italic = True
        .TextAlign = 2
        .ForeColor = &H808080
    End With

    topPos = topPos + 30

    ' ===== SECTION IMPORT =====
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraImport", True)
    With ctrl
        .Caption = "1. Import des Donnees"
        .Left = 10: .Top = topPos: .Width = 265: .Height = 120
        .Font.Bold = True
    End With

    ' Bouton Import Balance
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnImportBal", True)
    With ctrl
        .Caption = "Importer Balance"
        .Left = 20: .Top = topPos + 25: .Width = 120: .Height = 28
        .BackColor = &H80FF80 ' Vert clair
    End With

    ' Status Balance
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatusBal", True)
    With ctrl
        .Caption = "En attente..."
        .Left = 145: .Top = topPos + 30: .Width = 120: .Height = 18
        .ForeColor = &HFF ' Rouge
    End With

    ' Shape indicateur Balance
    Set ctrl = frm.Controls.Add("Forms.Label.1", "shpBalanceStatus", True)
    With ctrl
        .Caption = ""
        .Left = 250: .Top = topPos + 30: .Width = 15: .Height = 15
        .BackColor = &HFF ' Rouge
        .BorderStyle = 1
    End With

    ' Bouton Import GL
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnImportGL", True)
    With ctrl
        .Caption = "Importer GL Proof"
        .Left = 20: .Top = topPos + 60: .Width = 120: .Height = 28
        .BackColor = &H80FF80
    End With

    ' Status GL
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatusGL", True)
    With ctrl
        .Caption = "En attente..."
        .Left = 145: .Top = topPos + 65: .Width = 120: .Height = 18
        .ForeColor = &HFF
    End With

    ' Shape indicateur GL
    Set ctrl = frm.Controls.Add("Forms.Label.1", "shpGLStatus", True)
    With ctrl
        .Caption = ""
        .Left = 250: .Top = topPos + 65: .Width = 15: .Height = 15
        .BackColor = &HFF
        .BorderStyle = 1
    End With

    ' ===== SECTION PARAMETRES =====
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraParams", True)
    With ctrl
        .Caption = "2. Parametres"
        .Left = 285: .Top = topPos: .Width = 265: .Height = 120
        .Font.Bold = True
    End With

    ' Label SOL ID
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblSolID", True)
    With ctrl
        .Caption = "Code Agence (SOL ID):"
        .Left = 295: .Top = topPos + 25: .Width = 120: .Height = 18
    End With

    ' TextBox SOL ID
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtSolID", True)
    With ctrl
        .Left = 420: .Top = topPos + 22: .Width = 60: .Height = 22
        .Text = "799"
        .MaxLength = 3
    End With

    ' Label Tolerance
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTolerance", True)
    With ctrl
        .Caption = "Tolerance (XAF):"
        .Left = 295: .Top = topPos + 55: .Width = 120: .Height = 18
    End With

    ' TextBox Tolerance
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtTolerance", True)
    With ctrl
        .Left = 420: .Top = topPos + 52: .Width = 80: .Height = 22
        .Text = "100"
    End With

    ' Label Devise
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblDevise", True)
    With ctrl
        .Caption = "Devise:"
        .Left = 295: .Top = topPos + 85: .Width = 60: .Height = 18
    End With

    ' ComboBox Devise
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cboDevise", True)
    With ctrl
        .Left = 360: .Top = topPos + 82: .Width = 70: .Height = 22
        .Style = 2 ' fmStyleDropDownList
        .AddItem "XAF"
        .AddItem "EUR"
        .AddItem "USD"
        .AddItem "XOF"
        .ListIndex = 0
    End With

    topPos = topPos + 130

    ' ===== SECTION ANALYSE =====
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraAnalyse", True)
    With ctrl
        .Caption = "3. Analyse"
        .Left = 10: .Top = topPos: .Width = 540: .Height = 80
        .Font.Bold = True
    End With

    ' Bouton Lancer
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLancer", True)
    With ctrl
        .Caption = "LANCER L'ANALYSE COMPLETE"
        .Left = 20: .Top = topPos + 25: .Width = 250: .Height = 40
        .Font.Size = 12: .Font.Bold = True
        .BackColor = &HFF8000 ' Orange
        .ForeColor = &HFFFFFF ' Blanc
    End With

    ' Labels derniere analyse
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblLastAnalysis", True)
    With ctrl
        .Caption = "Aucune analyse effectuee"
        .Left = 285: .Top = topPos + 25: .Width = 250: .Height = 18
        .Font.Bold = True
    End With

    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblAnalysisDetails", True)
    With ctrl
        .Caption = ""
        .Left = 285: .Top = topPos + 45: .Width = 250: .Height = 18
    End With

    topPos = topPos + 90

    ' ===== SECTION RAPPORTS =====
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraRapports", True)
    With ctrl
        .Caption = "4. Rapports et Analyses"
        .Left = 10: .Top = topPos: .Width = 540: .Height = 80
        .Font.Bold = True
    End With

    ' Bouton Dashboard
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnVoirDashboard", True)
    With ctrl
        .Caption = "Dashboard"
        .Left = 20: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    ' Bouton Alertes
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnVoirAlertes", True)
    With ctrl
        .Caption = "Alertes"
        .Left = 105: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    ' Bouton Compliance
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnVoirCompliance", True)
    With ctrl
        .Caption = "Conformite"
        .Left = 190: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    ' Bouton Analyse Temporelle
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnAnalyseTemporelle", True)
    With ctrl
        .Caption = "Temporel"
        .Left = 275: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    ' Bouton Analyse Reseau
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnAnalyseReseau", True)
    With ctrl
        .Caption = "Reseau"
        .Left = 360: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    ' Bouton Diagnostic
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnDiagnostic", True)
    With ctrl
        .Caption = "Diagnostic"
        .Left = 445: .Top = topPos + 25: .Width = 80: .Height = 28
    End With

    topPos = topPos + 90

    ' ===== SECTION ACTIONS =====
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraActions", True)
    With ctrl
        .Caption = "Actions"
        .Left = 10: .Top = topPos: .Width = 540: .Height = 55
        .Font.Bold = True
    End With

    ' Bouton Reset
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnReset", True)
    With ctrl
        .Caption = "Reinitialiser"
        .Left = 20: .Top = topPos + 20: .Width = 100: .Height = 25
        .BackColor = &HC0C0FF ' Rouge clair
    End With

    ' Bouton Configuration
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnConfiguration", True)
    With ctrl
        .Caption = "Configuration"
        .Left = 130: .Top = topPos + 20: .Width = 100: .Height = 25
    End With

    ' Bouton Aide
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblHelp", True)
    With ctrl
        .Caption = "[?] Aide"
        .Left = 350: .Top = topPos + 22: .Width = 60: .Height = 20
        .ForeColor = &HFF0000 ' Bleu
        .Font.Underline = True
    End With

    ' Bouton Quitter
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnQuitter", True)
    With ctrl
        .Caption = "Quitter"
        .Left = 445: .Top = topPos + 20: .Width = 80: .Height = 25
    End With

    ' ===== AJOUTER LE CODE =====
    Call AddCockpitCode(VBComp)

    Exit Sub

ErrHandler:
    MsgBox "Erreur CreateCockpitForm: " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: AddCockpitCode
' Description: Ajoute le code VBA au formulaire Cockpit
'===============================================================================
Private Sub AddCockpitCode(VBComp As Object)
    On Error GoTo ErrHandler

    Dim codeModule As Object
    Dim code As String

    Set codeModule = VBComp.codeModule

    ' Le code est deja dans le fichier USF_Cockpit.frm
    ' On ajoute juste les declarations de base si besoin

    code = "Option Explicit" & vbCrLf & vbCrLf
    code = code & "' Constantes couleurs" & vbCrLf
    code = code & "Private Const COLOR_SUCCESS As Long = &HC000&" & vbCrLf
    code = code & "Private Const COLOR_WARNING As Long = &H80FF&" & vbCrLf
    code = code & "Private Const COLOR_ERROR As Long = &HFF&" & vbCrLf
    code = code & "Private Const COLOR_NEUTRAL As Long = &H808080" & vbCrLf
    code = code & "Private Const COLOR_PRIMARY As Long = &H804000" & vbCrLf

    codeModule.AddFromString code

    Exit Sub

ErrHandler:
    Debug.Print "Erreur AddCockpitCode: " & Err.Description
End Sub

'===============================================================================
' FONCTION: CreateLoginForm
' Description: Cree le formulaire de connexion frmLogin
'===============================================================================
Public Sub CreateLoginForm()
    On Error GoTo ErrHandler

    Dim VBProj As Object
    Dim VBComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim topPos As Single

    Set VBProj = ThisWorkbook.VBProject

    ' Supprimer si existe
    On Error Resume Next
    VBProj.VBComponents.Remove VBProj.VBComponents("frmLogin")
    On Error GoTo ErrHandler

    ' Creer le UserForm
    Set VBComp = VBProj.VBComponents.Add(3)
    VBComp.Name = "frmLogin"

    Set frm = VBComp.Designer

    With frm
        .Caption = "S.A.F.A - Connexion"
        .Width = 320
        .Height = 220
        .BackColor = &HFFFFFF
        .StartUpPosition = 1
    End With

    topPos = 20

    ' Logo/Titre
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblLogo", True)
    With ctrl
        .Caption = "S.A.F.A v10.0"
        .Left = 10: .Top = topPos: .Width = 300: .Height = 25
        .Font.Size = 14: .Font.Bold = True
        .TextAlign = 2
        .ForeColor = &H663300
    End With

    topPos = topPos + 40

    ' Label Username
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblUsername", True)
    With ctrl
        .Caption = "Utilisateur:"
        .Left = 20: .Top = topPos: .Width = 80: .Height = 18
    End With

    ' TextBox Username
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtUsername", True)
    With ctrl
        .Left = 100: .Top = topPos - 3: .Width = 180: .Height = 22
    End With

    topPos = topPos + 35

    ' Label Password
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblPassword", True)
    With ctrl
        .Caption = "Mot de passe:"
        .Left = 20: .Top = topPos: .Width = 80: .Height = 18
    End With

    ' TextBox Password
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtPassword", True)
    With ctrl
        .Left = 100: .Top = topPos - 3: .Width = 180: .Height = 22
        .PasswordChar = "*"
    End With

    topPos = topPos + 45

    ' Bouton Connexion
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnLogin", True)
    With ctrl
        .Caption = "Connexion"
        .Left = 60: .Top = topPos: .Width = 100: .Height = 30
        .Font.Bold = True
        .BackColor = &H80FF80
    End With

    ' Bouton Annuler
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With ctrl
        .Caption = "Annuler"
        .Left = 170: .Top = topPos: .Width = 80: .Height = 30
    End With

    topPos = topPos + 40

    ' Label erreur
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblError", True)
    With ctrl
        .Caption = ""
        .Left = 20: .Top = topPos: .Width = 280: .Height = 18
        .ForeColor = &HFF
        .TextAlign = 2
    End With

    Exit Sub

ErrHandler:
    MsgBox "Erreur CreateLoginForm: " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: CreateSettingsForm
' Description: Cree le formulaire de parametres frmSettings
'===============================================================================
Public Sub CreateSettingsForm()
    On Error GoTo ErrHandler

    Dim VBProj As Object
    Dim VBComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim topPos As Single

    Set VBProj = ThisWorkbook.VBProject

    ' Supprimer si existe
    On Error Resume Next
    VBProj.VBComponents.Remove VBProj.VBComponents("frmSettings")
    On Error GoTo ErrHandler

    ' Creer le UserForm
    Set VBComp = VBProj.VBComponents.Add(3)
    VBComp.Name = "frmSettings"

    Set frm = VBComp.Designer

    With frm
        .Caption = "S.A.F.A - Configuration"
        .Width = 400
        .Height = 350
        .BackColor = &HFFFFFF
        .StartUpPosition = 1
    End With

    topPos = 15

    ' Frame General
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraGeneral", True)
    With ctrl
        .Caption = "Parametres Generaux"
        .Left = 10: .Top = topPos: .Width = 380: .Height = 100
        .Font.Bold = True
    End With

    ' Tolerance
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblTol", True)
    With ctrl
        .Caption = "Tolerance par defaut:"
        .Left = 25: .Top = topPos + 25: .Width = 120: .Height = 18
    End With

    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtDefaultTolerance", True)
    With ctrl
        .Left = 150: .Top = topPos + 22: .Width = 100: .Height = 22
        .Text = "100"
    End With

    ' Z-Score Warning
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblZWarn", True)
    With ctrl
        .Caption = "Z-Score Warning:"
        .Left = 25: .Top = topPos + 55: .Width = 120: .Height = 18
    End With

    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtZScoreWarning", True)
    With ctrl
        .Left = 150: .Top = topPos + 52: .Width = 60: .Height = 22
        .Text = "2.5"
    End With

    ' Z-Score Critical
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblZCrit", True)
    With ctrl
        .Caption = "Z-Score Critical:"
        .Left = 220: .Top = topPos + 55: .Width = 100: .Height = 18
    End With

    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtZScoreCritical", True)
    With ctrl
        .Left = 325: .Top = topPos + 52: .Width = 50: .Height = 22
        .Text = "3.5"
    End With

    topPos = topPos + 110

    ' Frame Regulatory
    Set ctrl = frm.Controls.Add("Forms.Frame.1", "fraRegulatory", True)
    With ctrl
        .Caption = "Parametres Reglementaires"
        .Left = 10: .Top = topPos: .Width = 380: .Height = 100
        .Font.Bold = True
    End With

    ' COBAC Suspens
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblCobac", True)
    With ctrl
        .Caption = "COBAC Suspens (jours):"
        .Left = 25: .Top = topPos + 25: .Width = 130: .Height = 18
    End With

    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtCobacSuspens", True)
    With ctrl
        .Left = 160: .Top = topPos + 22: .Width = 50: .Height = 22
        .Text = "90"
    End With

    ' LAB/FT Seuil
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblLabft", True)
    With ctrl
        .Caption = "LAB/FT Seuil (XAF):"
        .Left = 25: .Top = topPos + 55: .Width = 130: .Height = 18
    End With

    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtLabftThreshold", True)
    With ctrl
        .Left = 160: .Top = topPos + 52: .Width = 100: .Height = 22
        .Text = "5000000"
    End With

    topPos = topPos + 115

    ' Boutons
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    With ctrl
        .Caption = "Sauvegarder"
        .Left = 100: .Top = topPos: .Width = 100: .Height = 30
        .Font.Bold = True
        .BackColor = &H80FF80
    End With

    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With ctrl
        .Caption = "Fermer"
        .Left = 210: .Top = topPos: .Width = 80: .Height = 30
    End With

    Exit Sub

ErrHandler:
    MsgBox "Erreur CreateSettingsForm: " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: ShowCockpit
' Description: Affiche le formulaire Cockpit (point d'entree principal)
'===============================================================================
Public Sub ShowCockpit()
    On Error GoTo ErrHandler

    ' Verifier si le formulaire existe
    Dim VBComp As Object
    On Error Resume Next
    Set VBComp = ThisWorkbook.VBProject.VBComponents("USF_Cockpit")
    On Error GoTo ErrHandler

    If VBComp Is Nothing Then
        ' Creer le formulaire si n'existe pas
        If MsgBox("Le formulaire Cockpit n'existe pas." & vbCrLf & vbCrLf & _
                  "Voulez-vous le creer maintenant?", _
                  vbYesNo + vbQuestion, "S.A.F.A") = vbYes Then
            Call CreateCockpitForm
        Else
            Exit Sub
        End If
    End If

    ' Afficher le formulaire
    VBA.UserForms.Add("USF_Cockpit").Show

    Exit Sub

ErrHandler:
    MsgBox "Erreur: " & Err.Description & vbCrLf & vbCrLf & _
           "Astuce: Executez d'abord FormBuilder.CreateAllForms", _
           vbExclamation, "Erreur"
End Sub

'===============================================================================
' FONCTION: CheckFormsExist
' Description: Verifie si tous les formulaires necessaires existent
'===============================================================================
Public Function CheckFormsExist() As Boolean
    On Error Resume Next

    Dim VBProj As Object
    Set VBProj = ThisWorkbook.VBProject

    If VBProj.VBComponents("USF_Cockpit") Is Nothing Then
        CheckFormsExist = False
        Exit Function
    End If

    If VBProj.VBComponents("frmLogin") Is Nothing Then
        CheckFormsExist = False
        Exit Function
    End If

    CheckFormsExist = True
End Function
