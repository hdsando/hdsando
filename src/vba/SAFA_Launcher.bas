Attribute VB_Name = "SAFA_Launcher"
'===============================================================================
' MODULE: SAFA_Launcher
' Description: Point d'entree principal de S.A.F.A
' Version: 10.0
' Usage: Executer Demarrer_SAFA() pour lancer l'application
'===============================================================================
Option Explicit

'===============================================================================
' PROCEDURE: Demarrer_SAFA
' Description: Lance l'application S.A.F.A (point d'entree principal)
'===============================================================================
Public Sub Demarrer_SAFA()
    On Error GoTo ErrHandler

    ' Verifier si les formulaires existent
    If Not FormBuilder.CheckFormsExist() Then
        Dim rep As Integer
        rep = MsgBox("Les formulaires S.A.F.A ne sont pas installes." & vbCrLf & vbCrLf & _
                     "Voulez-vous les creer maintenant?" & vbCrLf & vbCrLf & _
                     "(Cette operation ne prend que quelques secondes)", _
                     vbYesNo + vbQuestion, "S.A.F.A - Installation")

        If rep = vbYes Then
            Call FormBuilder.CreateAllForms
        Else
            Exit Sub
        End If
    End If

    ' Charger la configuration
    Call Config_Manager.LoadConfiguration

    ' Afficher le Cockpit
    Call FormBuilder.ShowCockpit

    Exit Sub

ErrHandler:
    MsgBox "Erreur au demarrage:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Code erreur: " & Err.Number, vbCritical, "Erreur S.A.F.A"
End Sub

'===============================================================================
' PROCEDURE: Installation_Complete
' Description: Effectue une installation complete de S.A.F.A
'===============================================================================
Public Sub Installation_Complete()
    On Error GoTo ErrHandler

    Dim startTime As Double
    startTime = Timer

    MsgBox "Installation de S.A.F.A v10.0" & vbCrLf & vbCrLf & _
           "Cette procedure va:" & vbCrLf & _
           "1. Creer les formulaires (interface)" & vbCrLf & _
           "2. Initialiser la configuration" & vbCrLf & _
           "3. Creer les feuilles necessaires" & vbCrLf & vbCrLf & _
           "Cliquez OK pour continuer.", vbInformation, "Installation S.A.F.A"

    Application.ScreenUpdating = False

    ' 1. Creer les formulaires
    Call FormBuilder.CreateAllForms

    ' 2. Charger/Initialiser la configuration
    Call Config_Manager.LoadConfiguration

    ' 3. Creer les feuilles de base
    Call CreerFeuillesBase

    Application.ScreenUpdating = True

    MsgBox "Installation terminee avec succes!" & vbCrLf & vbCrLf & _
           "Duree: " & Format(Timer - startTime, "0.0") & " secondes" & vbCrLf & vbCrLf & _
           "Pour lancer S.A.F.A, executez:" & vbCrLf & _
           "  SAFA_Launcher.Demarrer_SAFA", _
           vbInformation, "Installation Complete"

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'installation:" & vbCrLf & _
           Err.Description, vbCritical, "Erreur"
End Sub

'===============================================================================
' PROCEDURE: CreerFeuillesBase
' Description: Cree les feuilles Excel de base necessaires
'===============================================================================
Private Sub CreerFeuillesBase()
    On Error Resume Next

    Dim ws As Worksheet

    ' Feuille PARAM
    Set ws = ThisWorkbook.Sheets("PARAM")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "PARAM"
        ws.Range("A1").Value = "Parametre"
        ws.Range("B1").Value = "Valeur"
        ws.Range("A1:B1").Font.Bold = True

        ws.Range("A2").Value = "Tolerance"
        ws.Range("B2").Value = 100

        ws.Range("A3").Value = "Date Analyse"
        ws.Range("B3").Value = ""

        ws.Range("D1").Value = "Patterns Exclusion"
        ws.Range("D2").Value = "SUSPENS"
        ws.Range("D3").Value = "TRANSIT"
        ws.Range("D4").Value = "ATTENTE"

        ws.Range("F1").Value = "SOL ID"
        ws.Range("F2").Value = "799"

        ws.Columns("A:F").AutoFit
    End If

    ' Feuille USERS (pour securite)
    Set ws = ThisWorkbook.Sheets("USERS")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "USERS"

        ' En-tetes etendus pour Security_Enhanced
        ws.Range("A1:L1").Value = Array("Username", "PasswordHash", "Salt", "Role", _
                                        "CreatedDate", "LastLogin", "IsActive", _
                                        "FailedAttempts", "LastFailedAttempt", _
                                        "MustChangePassword", "PasswordExpiry", "Email")
        ws.Range("A1:L1").Font.Bold = True

        ' Utilisateur admin par defaut (mot de passe: admin123)
        ws.Range("A2").Value = "admin"
        ws.Range("B2").Value = "DEFAULT_HASH_TO_CHANGE"
        ws.Range("C2").Value = "UNIQUE_SALT"
        ws.Range("D2").Value = "ADMIN"
        ws.Range("E2").Value = Date
        ws.Range("G2").Value = True
        ws.Range("H2").Value = 0
        ws.Range("J2").Value = True ' Doit changer mot de passe

        ws.Visible = xlSheetVeryHidden
        ws.Columns("A:L").AutoFit
    End If

    ' Feuille AUDIT_TRAIL
    Set ws = ThisWorkbook.Sheets("AUDIT_TRAIL")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "AUDIT_TRAIL"
        ws.Range("A1:F1").Value = Array("Timestamp", "Type", "Details", "User", "Module", "Hash")
        ws.Range("A1:F1").Font.Bold = True
        ws.Visible = xlSheetVeryHidden
    End If

    On Error GoTo 0
End Sub

'===============================================================================
' PROCEDURE: Afficher_Version
' Description: Affiche les informations de version
'===============================================================================
Public Sub Afficher_Version()
    Dim cfg As Config_Manager.GeneralConfig
    cfg = Config_Manager.GetGeneralConfig()

    MsgBox cfg.ApplicationName & " - System for Automated Financial Audit" & vbCrLf & _
           "Version " & cfg.Version & vbCrLf & vbCrLf & _
           "Modules installes:" & vbCrLf & _
           "  - Core_Engine (Rapprochement)" & vbCrLf & _
           "  - Forensic_Rules (Detection Fraude)" & vbCrLf & _
           "  - Advanced_AI (Z-Score, Clustering)" & vbCrLf & _
           "  - Regulatory_Compliance (COBAC, OHADA)" & vbCrLf & _
           "  - Security_Module (Authentification)" & vbCrLf & _
           "  - Security_Enhanced (Crypto avancee)" & vbCrLf & _
           "  - Temporal_Analysis (Tendances)" & vbCrLf & _
           "  - Network_Analysis (Graphes)" & vbCrLf & _
           "  - Report_Generator (Rapports)" & vbCrLf & _
           "  - Config_Manager (Configuration)" & vbCrLf & _
           "  - Data_Ingestion (Import)" & vbCrLf & _
           "  - Auto_Diagnostic (Sante systeme)" & vbCrLf & _
           "  - Batch_Automation (Traitement auto)" & vbCrLf & vbCrLf & _
           "(c) 2024 S.A.F.A Team", _
           vbInformation, "A propos de " & cfg.ApplicationName
End Sub

'===============================================================================
' PROCEDURE: Aide_Rapide
' Description: Affiche l'aide rapide
'===============================================================================
Public Sub Aide_Rapide()
    MsgBox "GUIDE DE DEMARRAGE RAPIDE S.A.F.A" & vbCrLf & vbCrLf & _
           "1. INSTALLATION (premiere fois seulement):" & vbCrLf & _
           "   Executez: SAFA_Launcher.Installation_Complete" & vbCrLf & vbCrLf & _
           "2. UTILISATION:" & vbCrLf & _
           "   Executez: SAFA_Launcher.Demarrer_SAFA" & vbCrLf & vbCrLf & _
           "3. DANS LE COCKPIT:" & vbCrLf & _
           "   a) Importez votre fichier Balance Finacle" & vbCrLf & _
           "   b) Importez votre fichier GL Proof" & vbCrLf & _
           "   c) Entrez le code agence (SOL ID)" & vbCrLf & _
           "   d) Cliquez 'Lancer l'Analyse'" & vbCrLf & vbCrLf & _
           "4. RACCOURCIS:" & vbCrLf & _
           "   F5 = Rafraichir les indicateurs" & vbCrLf & _
           "   Entree = Lancer l'analyse" & vbCrLf & _
           "   Echap = Fermer", _
           vbInformation, "Aide S.A.F.A"
End Sub
