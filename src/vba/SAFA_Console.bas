Attribute VB_Name = "SAFA_Console"
'===============================================================================
' MODULE: SAFA_Console
' Description: Version console de S.A.F.A (sans UserForm)
' Version: 10.0
' Usage: Fonctionne avec InputBox/MsgBox - pas besoin de formulaires
'===============================================================================
Option Explicit

'===============================================================================
' PROCEDURE PRINCIPALE: Demarrer
' Description: Lance S.A.F.A en mode console (sans formulaire)
'===============================================================================
Public Sub Demarrer()
    Dim choix As String
    Dim continuer As Boolean

    ' Initialiser
    Call Config_Manager.LoadConfiguration
    Call InitialiserFeuillesBase

    continuer = True

    Do While continuer
        choix = AfficherMenuPrincipal()

        Select Case choix
            Case "1"
                Call ImporterBalance
            Case "2"
                Call ImporterGLProof
            Case "3"
                Call DefinirParametres
            Case "4"
                Call LancerAnalyse
            Case "5"
                Call AfficherRapports
            Case "6"
                Call AnalysesAvancees
            Case "7"
                Call AfficherStatut
            Case "8"
                Call Reinitialiser
            Case "0", ""
                continuer = False
            Case Else
                MsgBox "Choix invalide. Veuillez reessayer.", vbExclamation
        End Select

        ' Rafraichir le panneau de statut de la feuille MENU (si presente)
        On Error Resume Next
        Call SAFA_Menu.RefreshStatus
        On Error GoTo 0
    Loop

    MsgBox "Merci d'avoir utilise S.A.F.A!", vbInformation, "Au revoir"
End Sub

'===============================================================================
' FONCTION: AfficherMenuPrincipal
'===============================================================================
Private Function AfficherMenuPrincipal() As String
    Dim menu As String
    Dim cfg As Config_Manager.GeneralConfig
    cfg = Config_Manager.GetGeneralConfig()

    menu = "===== " & cfg.ApplicationName & " v" & cfg.Version & " =====" & vbCrLf & vbCrLf
    menu = menu & "1 - Importer Balance Finacle" & vbCrLf
    menu = menu & "2 - Importer GL Proof" & vbCrLf
    menu = menu & "3 - Definir Parametres" & vbCrLf
    menu = menu & "4 - LANCER L'ANALYSE" & vbCrLf
    menu = menu & "5 - Voir les Rapports" & vbCrLf
    menu = menu & "6 - Analyses Avancees" & vbCrLf
    menu = menu & "7 - Afficher Statut" & vbCrLf
    menu = menu & "8 - Reinitialiser" & vbCrLf
    menu = menu & "0 - Quitter" & vbCrLf & vbCrLf
    menu = menu & "Entrez votre choix (0-8):"

    AfficherMenuPrincipal = InputBox(menu, cfg.ApplicationName)
End Function

'===============================================================================
' PROCEDURE: ImporterBalance
'===============================================================================
Public Sub ImporterBalance()
    Dim filePath As String
    Dim result As Data_Ingestion.ImportResult

    filePath = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm,Fichiers CSV (*.csv),*.csv,Tous (*.*),*.*", _
        Title:="Selectionner le fichier Balance Finacle")

    If filePath = "False" Or filePath = "" Then
        MsgBox "Import annule.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    result = Data_Ingestion.ImportFile(filePath, "BALANCE_RAW", True)
    Application.ScreenUpdating = True

    If result.Success Then
        MsgBox "Balance importee avec succes!" & vbCrLf & vbCrLf & _
               "Lignes importees: " & result.RowsImported & vbCrLf & _
               "Duree: " & Format(result.Duration, "0.00") & " sec", _
               vbInformation, "Import Balance"
    Else
        MsgBox "Erreur lors de l'import:" & vbCrLf & result.Errors, vbCritical, "Erreur"
    End If
End Sub

'===============================================================================
' PROCEDURE: ImporterGLProof
'===============================================================================
Public Sub ImporterGLProof()
    Dim filePath As String
    Dim result As Data_Ingestion.ImportResult

    filePath = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm,Fichiers CSV (*.csv),*.csv,Fichiers TXT (*.txt),*.txt,Tous (*.*),*.*", _
        Title:="Selectionner le fichier GL Proof")

    If filePath = "False" Or filePath = "" Then
        MsgBox "Import annule.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    result = Data_Ingestion.ImportFile(filePath, "GLPROOF_RAW", True)
    Application.ScreenUpdating = True

    If result.Success Then
        MsgBox "GL Proof importe avec succes!" & vbCrLf & vbCrLf & _
               "Lignes importees: " & result.RowsImported & vbCrLf & _
               "Duree: " & Format(result.Duration, "0.00") & " sec", _
               vbInformation, "Import GL Proof"
    Else
        MsgBox "Erreur lors de l'import:" & vbCrLf & result.Errors, vbCritical, "Erreur"
    End If
End Sub

'===============================================================================
' PROCEDURE: DefinirParametres
'===============================================================================
Public Sub DefinirParametres()
    Dim solId As String
    Dim tolerance As String
    Dim wsParam As Worksheet

    On Error Resume Next
    Set wsParam = ThisWorkbook.Sheets("PARAM")
    On Error GoTo 0

    If wsParam Is Nothing Then
        Call InitialiserFeuillesBase
        Set wsParam = ThisWorkbook.Sheets("PARAM")
    End If

    ' SOL ID
    solId = InputBox("Code Agence (SOL ID):" & vbCrLf & vbCrLf & _
                     "Valeur actuelle: " & wsParam.Range("F2").Value, _
                     "Parametre SOL ID", wsParam.Range("F2").Value)

    If solId <> "" Then
        wsParam.Range("F2").Value = solId
    End If

    ' Tolerance
    tolerance = InputBox("Tolerance d'ecart (XAF):" & vbCrLf & vbCrLf & _
                         "Valeur actuelle: " & wsParam.Range("B2").Value, _
                         "Parametre Tolerance", wsParam.Range("B2").Value)

    If tolerance <> "" And IsNumeric(tolerance) Then
        wsParam.Range("B2").Value = CDbl(tolerance)
    End If

    MsgBox "Parametres sauvegardes!" & vbCrLf & vbCrLf & _
           "SOL ID: " & wsParam.Range("F2").Value & vbCrLf & _
           "Tolerance: " & wsParam.Range("B2").Value & " XAF", _
           vbInformation, "Parametres"
End Sub

'===============================================================================
' PROCEDURE: LancerAnalyse
'===============================================================================
Public Sub LancerAnalyse()
    Dim wsBal As Worksheet, wsGL As Worksheet
    Dim nbBal As Long, nbGL As Long
    Dim rep As Integer

    ' Verifier les donnees
    On Error Resume Next
    Set wsBal = ThisWorkbook.Sheets("BALANCE_RAW")
    Set wsGL = ThisWorkbook.Sheets("GLPROOF_RAW")
    On Error GoTo 0

    If wsBal Is Nothing Then
        nbBal = 0
    Else
        nbBal = Application.CountA(wsBal.Range("A:A"))
    End If

    If wsGL Is Nothing Then
        nbGL = 0
    Else
        nbGL = Application.CountA(wsGL.Range("A:A"))
    End If

    ' Verification
    If nbBal < 2 Then
        MsgBox "ERREUR: Balance non importee ou vide." & vbCrLf & vbCrLf & _
               "Utilisez l'option 1 pour importer la Balance.", vbExclamation, "Donnees manquantes"
        Exit Sub
    End If

    If nbGL < 2 Then
        MsgBox "ERREUR: GL Proof non importe ou vide." & vbCrLf & vbCrLf & _
               "Utilisez l'option 2 pour importer le GL Proof.", vbExclamation, "Donnees manquantes"
        Exit Sub
    End If

    ' Confirmation
    rep = MsgBox("Pret a lancer l'analyse:" & vbCrLf & vbCrLf & _
                 "- Balance: " & nbBal & " lignes" & vbCrLf & _
                 "- GL Proof: " & nbGL & " lignes" & vbCrLf & vbCrLf & _
                 "Continuer?", vbYesNo + vbQuestion, "Confirmation")

    If rep = vbNo Then Exit Sub

    ' Lancer l'analyse
    Application.ScreenUpdating = False
    Call Core_Engine.Lancer_Traitement_Complet
    Application.ScreenUpdating = True

    MsgBox "Analyse terminee!" & vbCrLf & vbCrLf & _
           "Consultez les feuilles:" & vbCrLf & _
           "- RECONCIL (rapprochement)" & vbCrLf & _
           "- AUDIT_REPORT (alertes)" & vbCrLf & _
           "- DASHBOARD_RISQUE (tableau de bord)", _
           vbInformation, "Analyse Complete"
End Sub

'===============================================================================
' PROCEDURE: AfficherRapports
'===============================================================================
Public Sub AfficherRapports()
    On Error GoTo ErrHandler

    Dim choix As String
    Dim menu As String

    ' Verifier si l'analyse de base a ete faite
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Aucun rapport disponible." & vbCrLf & vbCrLf & _
               "Lancez d'abord une analyse (option 4 du menu principal).", _
               vbExclamation, "Analyse requise"
        Exit Sub
    End If

    menu = "===== RAPPORTS =====" & vbCrLf & vbCrLf
    menu = menu & "--- Voir les feuilles ---" & vbCrLf
    menu = menu & "1 - RECONCIL (Rapprochement)" & vbCrLf
    menu = menu & "2 - AUDIT_REPORT (Alertes)" & vbCrLf
    menu = menu & "3 - DASHBOARD_RISQUE (Tableau de bord)" & vbCrLf
    menu = menu & "4 - COMPLIANCE_CHECK (Conformite)" & vbCrLf
    menu = menu & "5 - EXECUTIVE_SUMMARY (Synthese)" & vbCrLf
    menu = menu & vbCrLf & "--- Actions ---" & vbCrLf
    menu = menu & "6 - GENERER tous les rapports" & vbCrLf
    menu = menu & "7 - Exporter en PDF" & vbCrLf
    menu = menu & "0 - Retour" & vbCrLf & vbCrLf
    menu = menu & "Choix:"

    choix = InputBox(menu, "Rapports")

    Select Case choix
        Case "1"
            Call ActiverFeuille("RECONCIL")
        Case "2"
            Call ActiverFeuille("AUDIT_REPORT")
        Case "3"
            If Not Core_Engine.FeuilleExiste("DASHBOARD_RISQUE") Then
                Call Report_Generator.GenerateRiskDashboard(ThisWorkbook)
            End If
            Call ActiverFeuille("DASHBOARD_RISQUE")
        Case "4"
            If Not Core_Engine.FeuilleExiste("COMPLIANCE_CHECK") Then
                Call Regulatory_Compliance.Lancer_Verification_Conformite
            End If
            Call ActiverFeuille("COMPLIANCE_CHECK")
        Case "5"
            If Not Core_Engine.FeuilleExiste("EXECUTIVE_SUMMARY") Then
                Call Report_Generator.GenerateExecutiveSummary(ThisWorkbook)
            End If
            Call ActiverFeuille("EXECUTIVE_SUMMARY")
        Case "6"
            Call GenererTousRapports
        Case "7"
            Call Report_Generator.ExportToPDF(ThisWorkbook)
            MsgBox "Export PDF termine!", vbInformation
    End Select

    Exit Sub

ErrHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, "Erreur"
End Sub

'===============================================================================
' PROCEDURE: AnalysesAvancees
'===============================================================================
Public Sub AnalysesAvancees()
    On Error GoTo ErrHandler

    Dim choix As String
    Dim menu As String
    Dim analyseOK As Boolean

    ' Verifier si l'analyse de base a ete faite
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "ATTENTION: Aucune analyse de base effectuee." & vbCrLf & vbCrLf & _
               "Retournez au menu principal et lancez l'option 4 d'abord.", _
               vbExclamation, "Analyse requise"
        Exit Sub
    End If

    menu = "===== ANALYSES AVANCEES =====" & vbCrLf & vbCrLf
    menu = menu & "1 - Analyse Temporelle (tendances)" & vbCrLf
    menu = menu & "2 - Analyse Reseau (circuits)" & vbCrLf
    menu = menu & "3 - Diagnostic Systeme" & vbCrLf
    menu = menu & "4 - IA par Compte (Z-Score)" & vbCrLf
    menu = menu & "5 - Analyse Benford" & vbCrLf
    menu = menu & "6 - Conformite Reglementaire" & vbCrLf
    menu = menu & "7 - Generer Tous les Rapports" & vbCrLf
    menu = menu & "0 - Retour" & vbCrLf & vbCrLf
    menu = menu & "Choix:"

    choix = InputBox(menu, "Analyses Avancees")

    If choix = "" Or choix = "0" Then Exit Sub

    Application.ScreenUpdating = False
    analyseOK = False

    Select Case choix
        Case "1"
            Call Temporal_Analysis.Lancer_Analyse_Temporelle
            analyseOK = Core_Engine.FeuilleExiste("TEMPORAL_ANALYSIS")
            If analyseOK Then
                MsgBox "Analyse temporelle terminee!" & vbCrLf & "Voir feuille TEMPORAL_ANALYSIS", vbInformation
            End If

        Case "2"
            If Core_Engine.FeuilleExiste("TRANSACTION_DATA") Then
                Call Network_Analysis.Lancer_Analyse_Reseau
                analyseOK = Core_Engine.FeuilleExiste("NETWORK_ANALYSIS")
                If analyseOK Then
                    MsgBox "Analyse reseau terminee!" & vbCrLf & "Voir feuille NETWORK_ANALYSIS", vbInformation
                End If
            Else
                MsgBox "Feuille TRANSACTION_DATA requise." & vbCrLf & _
                       "Relancez l'analyse complete (option 4 menu principal)", vbExclamation
            End If

        Case "3"
            Dim health As Auto_Diagnostic.SystemHealth
            health = Auto_Diagnostic.LancerDiagnosticComplet()
            analyseOK = Core_Engine.FeuilleExiste("SYSTEM_DIAGNOSTIC")
            If analyseOK Then
                MsgBox "Diagnostic termine!" & vbCrLf & vbCrLf & _
                       "Score global: " & health.OverallScore & "/100" & vbCrLf & _
                       "Voir feuille SYSTEM_DIAGNOSTIC", vbInformation
            End If

        Case "4"
            If Core_Engine.FeuilleExiste("TRANSACTION_DATA") Then
                Call Advanced_AI.Lancer_IA_Par_Compte
                MsgBox "Analyse IA terminee!" & vbCrLf & "Voir feuille AUDIT_REPORT", vbInformation
                analyseOK = True
            Else
                MsgBox "Feuille TRANSACTION_DATA requise." & vbCrLf & _
                       "Relancez l'analyse complete (option 4 menu principal)", vbExclamation
            End If

        Case "5"
            If Core_Engine.FeuilleExiste("TRANSACTION_DATA") Then
                Call Forensic_Rules.Lancer_Benford
                MsgBox "Analyse Benford terminee!", vbInformation
                analyseOK = True
            Else
                MsgBox "Feuille TRANSACTION_DATA requise." & vbCrLf & _
                       "Relancez l'analyse complete (option 4 menu principal)", vbExclamation
            End If

        Case "6"
            Call Regulatory_Compliance.Lancer_Verification_Conformite
            analyseOK = Core_Engine.FeuilleExiste("COMPLIANCE_CHECK")
            If analyseOK Then
                MsgBox "Controles conformite termines!" & vbCrLf & "Voir feuille COMPLIANCE_CHECK", vbInformation
            End If

        Case "7"
            ' Generer tous les rapports
            Call GenererTousRapports
            analyseOK = True
    End Select

    Application.ScreenUpdating = True

    If Not analyseOK And choix <> "0" And choix <> "" Then
        MsgBox "L'analyse n'a pas produit de resultats." & vbCrLf & _
               "Verifiez les donnees sources.", vbExclamation
    End If

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur dans AnalysesAvancees:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Erreur"
End Sub

'===============================================================================
' PROCEDURE: GenererTousRapports
'===============================================================================
Public Sub GenererTousRapports()
    On Error Resume Next

    MsgBox "Generation de tous les rapports..." & vbCrLf & _
           "Cela peut prendre quelques instants.", vbInformation

    Application.ScreenUpdating = False

    ' Generer les rapports
    Call Report_Generator.GenerateExecutiveSummary(ThisWorkbook)
    Call Report_Generator.GenerateRiskDashboard(ThisWorkbook)
    Call Report_Generator.GenerateForensicReport(ThisWorkbook)
    Call Report_Generator.GenerateComplianceReport(ThisWorkbook)
    Call Report_Generator.GenerateSampleSheet(ThisWorkbook)

    ' Finaliser rapport AI
    Call Advanced_AI.Finaliser_Rapport

    Application.ScreenUpdating = True

    MsgBox "Tous les rapports ont ete generes!" & vbCrLf & vbCrLf & _
           "Feuilles creees:" & vbCrLf & _
           "- EXECUTIVE_SUMMARY" & vbCrLf & _
           "- DASHBOARD_RISQUE" & vbCrLf & _
           "- FORENSIC_ANALYSIS" & vbCrLf & _
           "- COMPLIANCE_CHECK" & vbCrLf & _
           "- ECHANTILLON_TEST", vbInformation, "Rapports Generes"
End Sub

'===============================================================================
' PROCEDURE: AfficherStatut
'===============================================================================
Public Sub AfficherStatut()
    Dim status As String
    Dim nbBal As Long, nbGL As Long, nbRec As Long, nbAlerts As Long

    status = "===== STATUT S.A.F.A =====" & vbCrLf & vbCrLf

    ' Balance
    If Core_Engine.FeuilleExiste("BALANCE_RAW") Then
        nbBal = Application.CountA(ThisWorkbook.Sheets("BALANCE_RAW").Range("A:A")) - 1
        status = status & "Balance: OK (" & nbBal & " lignes)" & vbCrLf
    Else
        status = status & "Balance: NON IMPORTEE" & vbCrLf
    End If

    ' GL Proof
    If Core_Engine.FeuilleExiste("GLPROOF_RAW") Then
        nbGL = Application.CountA(ThisWorkbook.Sheets("GLPROOF_RAW").Range("A:A")) - 1
        status = status & "GL Proof: OK (" & nbGL & " lignes)" & vbCrLf
    Else
        status = status & "GL Proof: NON IMPORTE" & vbCrLf
    End If

    ' Reconciliation
    If Core_Engine.FeuilleExiste("RECONCIL") Then
        nbRec = Application.CountA(ThisWorkbook.Sheets("RECONCIL").Range("A:A")) - 1
        status = status & "Rapprochement: OK (" & nbRec & " comptes)" & vbCrLf
    Else
        status = status & "Rapprochement: NON EFFECTUE" & vbCrLf
    End If

    ' Alertes
    If Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        nbAlerts = Application.CountA(ThisWorkbook.Sheets("AUDIT_REPORT").Range("A:A")) - 1
        status = status & "Alertes: " & nbAlerts & " detectees" & vbCrLf
    End If

    ' Parametres
    status = status & vbCrLf & "--- Parametres ---" & vbCrLf
    On Error Resume Next
    status = status & "SOL ID: " & ThisWorkbook.Sheets("PARAM").Range("F2").Value & vbCrLf
    status = status & "Tolerance: " & ThisWorkbook.Sheets("PARAM").Range("B2").Value & " XAF" & vbCrLf
    On Error GoTo 0

    MsgBox status, vbInformation, "Statut"
End Sub

'===============================================================================
' PROCEDURE: Reinitialiser
'===============================================================================
Public Sub Reinitialiser()
    Dim rep As Integer

    rep = MsgBox("Voulez-vous vraiment reinitialiser?" & vbCrLf & vbCrLf & _
                 "Toutes les donnees seront effacees.", _
                 vbYesNo + vbExclamation, "Confirmation")

    If rep = vbYes Then
        Call Core_Engine.Reinitialiser_Outil_Parametrable(False)
        MsgBox "Reinitialisation effectuee.", vbInformation
    End If
End Sub

'===============================================================================
' PROCEDURE: ActiverFeuille
'===============================================================================
Private Sub ActiverFeuille(nomFeuille As String)
    On Error Resume Next
    If Core_Engine.FeuilleExiste(nomFeuille) Then
        ThisWorkbook.Sheets(nomFeuille).Visible = xlSheetVisible
        ThisWorkbook.Sheets(nomFeuille).Activate
    Else
        MsgBox "La feuille " & nomFeuille & " n'existe pas encore." & vbCrLf & _
               "Lancez d'abord une analyse.", vbExclamation
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' PROCEDURE: InitialiserFeuillesBase
'===============================================================================
Private Sub InitialiserFeuillesBase()
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
        ' Patterns d'exclusion (Non-Proofable): comptes techniques, PAS les suspens/transit
        ' (les suspens et transit doivent etre analyses pour les controles COBAC)
        ws.Range("D1").Value = "Patterns Exclusion"
        ws.Range("D2").Value = "ISO"
        ws.Range("D3").Value = "NOSTRO"
        ws.Range("D4").Value = "VOSTRO"
        ws.Range("F1").Value = "SOL ID"
        ws.Range("F2").Value = "799"
    End If

    On Error GoTo 0
End Sub

'===============================================================================
' RACCOURCI: Lancement rapide
'===============================================================================
Public Sub LancerSAFA()
    Call Demarrer
End Sub

'===============================================================================
' POINTS D'ENTREE DIRECTS (boutons de la feuille MENU - SAFA_Menu.BuildMenu)
' Chaque bouton appelle une action unique, sans passer par les menus InputBox.
'===============================================================================

Public Sub OuvrirDashboard()
    On Error Resume Next
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    If Not Core_Engine.FeuilleExiste("DASHBOARD_RISQUE") Then Call Report_Generator.GenerateRiskDashboard(ThisWorkbook)
    Call ActiverFeuille("DASHBOARD_RISQUE")
End Sub

Public Sub OuvrirAlertes()
    On Error Resume Next
    If Not Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        MsgBox "Aucune alerte: lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    Call Report_Generator.GenerateAlertReport(ThisWorkbook)
    Call ActiverFeuille("AUDIT_REPORT")
End Sub

Public Sub OuvrirConformite()
    On Error Resume Next
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    If Not Core_Engine.FeuilleExiste("COMPLIANCE_CHECK") Then Call Regulatory_Compliance.Lancer_Verification_Conformite
    Call ActiverFeuille("COMPLIANCE_CHECK")
End Sub

Public Sub OuvrirSynthese()
    On Error Resume Next
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    If Not Core_Engine.FeuilleExiste("EXECUTIVE_SUMMARY") Then Call Report_Generator.GenerateExecutiveSummary(ThisWorkbook)
    Call ActiverFeuille("EXECUTIVE_SUMMARY")
End Sub

Public Sub OuvrirEchantillon()
    On Error Resume Next
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    Call Report_Generator.GenerateSampleSheet(ThisWorkbook)
    Call ActiverFeuille("ECHANTILLON_TEST")
End Sub

Public Sub OuvrirJournalAudit()
    On Error Resume Next
    If Core_Engine.FeuilleExiste("AUDIT_TRAIL") Then
        Call ActiverFeuille("AUDIT_TRAIL")
    Else
        MsgBox "Le journal d'audit n'existe pas encore.", vbInformation, "S.A.F.A"
    End If
End Sub

Public Sub ExporterPDF()
    On Error GoTo ErrHandler
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    Call Report_Generator.ExportToPDF(ThisWorkbook)
    Exit Sub
ErrHandler:
    MsgBox "Export PDF impossible: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub LancerTemporel()
    On Error GoTo ErrHandler
    If Not Core_Engine.FeuilleExiste("RECONCIL") Then
        MsgBox "Lancez d'abord l'analyse complete.", vbExclamation, "S.A.F.A": Exit Sub
    End If
    Application.ScreenUpdating = False
    Call Temporal_Analysis.Lancer_Analyse_Temporelle
    Application.ScreenUpdating = True
    Call ActiverFeuille("TEMPORAL_ANALYSIS")
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Analyse temporelle: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub LancerReseau()
    On Error GoTo ErrHandler
    If Not Core_Engine.FeuilleExiste("TRANSACTION_DATA") Then
        MsgBox "Lancez d'abord l'analyse complete (TRANSACTION_DATA requis).", vbExclamation, "S.A.F.A": Exit Sub
    End If
    Application.ScreenUpdating = False
    Call Network_Analysis.Lancer_Analyse_Reseau
    Application.ScreenUpdating = True
    Call ActiverFeuille("NETWORK_ANALYSIS")
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Analyse reseau: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub LancerDiagnostic()
    On Error GoTo ErrHandler
    Dim health As Auto_Diagnostic.SystemHealth
    Application.ScreenUpdating = False
    health = Auto_Diagnostic.LancerDiagnosticComplet()
    Application.ScreenUpdating = True
    Call ActiverFeuille("SYSTEM_DIAGNOSTIC")
    MsgBox "Diagnostic termine - score global: " & health.OverallScore & "/100", vbInformation, "S.A.F.A"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Diagnostic: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub LancerTests()
    On Error GoTo ErrHandler
    Call SAFA_Tests.RunSmokeTestUI
    Exit Sub
ErrHandler:
    MsgBox "Tests: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub ChargerDonneesDemo()
    On Error GoTo ErrHandler
    If MsgBox("Generer un jeu de donnees de demonstration ?" & vbCrLf & vbCrLf & _
              "Les feuilles BALANCE_RAW et GLPROOF_RAW seront REMPLACEES par des donnees fictives" & vbCrLf & _
              "(300 comptes, ~4 000 ecritures, anomalies injectees pour tester les detecteurs).", _
              vbYesNo + vbQuestion, "Donnees de demonstration") = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    Call Demo_Data.GenerateDemoData
    Application.ScreenUpdating = True
    MsgBox "Donnees de demonstration chargees." & vbCrLf & "Cliquez maintenant sur LANCER L'ANALYSE COMPLETE.", vbInformation, "S.A.F.A"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Donnees demo: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub AfficherConfiguration()
    On Error Resume Next
    Dim cfg As Config_Manager.FullConfig
    cfg = Config_Manager.GetConfig()
    MsgBox "CONFIGURATION ACTIVE" & vbCrLf & vbCrLf & _
           "Tolerance: " & cfg.General.DefaultTolerance & " XAF" & vbCrLf & _
           "Devise: " & cfg.General.DefaultCurrency & "   SOL ID: " & cfg.General.DefaultSolId & vbCrLf & vbCrLf & _
           "Z-Score warning / critical: " & cfg.Forensic.ZScoreWarning & " / " & cfg.Forensic.ZScoreCritical & vbCrLf & _
           "Benford Chi2 critique: " & cfg.Forensic.BenfordChiSquaredCritical & "   MAD marginal: " & cfg.Forensic.BenfordMADMarginal & vbCrLf & _
           "Weekend seuil: " & Format(cfg.Forensic.WeekendThreshold, "#,##0") & " XAF" & vbCrLf & vbCrLf & _
           "COBAC suspens / transit: " & cfg.Regulatory.CobacSuspensLimitDays & " j / " & cfg.Regulatory.CobacTransitLimitDays & " j" & vbCrLf & _
           "LAB/FT declaration: " & Format(cfg.Regulatory.LabftDeclarationThreshold, "#,##0") & " XAF" & vbCrLf & _
           "LAB/FT structuration: " & Format(cfg.Regulatory.LabftStructuringThreshold, "#,##0") & " XAF" & vbCrLf & vbCrLf & _
           "Seuils risque C/H/M: " & cfg.RiskScoring.ThresholdCritical & " / " & cfg.RiskScoring.ThresholdHigh & " / " & cfg.RiskScoring.ThresholdMedium & vbCrLf & vbCrLf & _
           "Modifiable dans la feuille CONFIG_DATA (Config_Manager.SaveConfigurationToSheet) ou config\settings.json", _
           vbInformation, "S.A.F.A - Configuration"
End Sub

Public Sub AfficherAide()
    MsgBox "S.A.F.A v" & SAFA_Common.SAFA_VERSION & " - System for Automated Financial Audit" & vbCrLf & vbCrLf & _
           "1. Importer Balance Finacle  (ou Donnees demo pour tester)" & vbCrLf & _
           "2. Importer GL Proof" & vbCrLf & _
           "3. Verifier Parametres (SOL ID, tolerance)" & vbCrLf & _
           "4. LANCER L'ANALYSE COMPLETE" & vbCrLf & _
           "5. Consulter Dashboard / Alertes / Conformite / Synthese" & vbCrLf & _
           "6. Analyses avancees: Temporel, Reseau, Diagnostic" & vbCrLf & vbCrLf & _
           "Tests: le bouton Tests genere des donnees demo, lance tout le pipeline" & vbCrLf & _
           "et verifie que chaque anomalie injectee est detectee (feuille TEST_RESULTS).", _
           vbInformation, "Aide S.A.F.A"
End Sub

'===============================================================================
' CALIBRATION AUTOMATIQUE ET PARTAGE
'===============================================================================

Public Sub AfficherCalibration()
    ' Relance la calibration sur les donnees importees et affiche la feuille CALIBRATION
    On Error GoTo ErrHandler
    If Not Core_Engine.FeuilleExiste("BALANCE_RAW") Or Not Core_Engine.FeuilleExiste("GLPROOF_RAW") Then
        MsgBox "Importez d'abord la Balance et le GL Proof.", vbExclamation, "Calibration": Exit Sub
    End If
    Application.ScreenUpdating = False
    Call Auto_Calibration.CalibrerNumerotation(False)
    Application.ScreenUpdating = True
    Call ActiverFeuille("CALIBRATION")
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Calibration: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub PreparerPartage()
    ' Produit une copie distribuable du classeur: donnees effacees, configuration embarquee,
    ' compte admin garanti, journal conserve. Le destinataire n'a besoin que d'Excel (Windows)
    ' avec les macros activees - aucune installation ni acces VBA requis.
    On Error GoTo ErrHandler
    Dim dest As String, rep As Integer

    rep = MsgBox("Preparer une copie de S.A.F.A a partager ?" & vbCrLf & vbCrLf & _
                 "La copie contiendra les modules et la configuration, mais AUCUNE donnee importee" & vbCrLf & _
                 "(Balance, GL Proof, resultats). Le classeur courant n'est pas modifie.", _
                 vbYesNo + vbQuestion, "Preparer partage")
    If rep = vbNo Then Exit Sub

    dest = Application.GetSaveAsFilename( _
        InitialFileName:=ThisWorkbook.Path & "\SAFA_partage_" & Format(Date, "yyyymmdd") & ".xlsm", _
        FileFilter:="Classeur Excel avec macros (*.xlsm),*.xlsm", Title:="Enregistrer la copie a partager")
    If dest = "False" Or dest = "" Then Exit Sub

    Application.ScreenUpdating = False
    ' 1. Sauvegarder le classeur courant tel quel
    ThisWorkbook.Save
    ' 2. Embarquer la configuration (feuille CONFIG_DATA) et garantir le compte admin
    Call Config_Manager.SaveConfigurationToSheet
    On Error Resume Next
    Call Security_Enhanced.EnsureDefaultAdmin
    On Error GoTo ErrHandler
    ' 3. Copie
    ThisWorkbook.SaveCopyAs dest
    Application.ScreenUpdating = True

    ' 4. Nettoyer la copie: l'ouvrir, effacer les donnees, refermer
    Dim wbCopy As Workbook, ws As Worksheet, keep As Variant, k As Variant, isKeep As Boolean
    keep = Array("MENU", "PARAM", "CONFIG_DATA", "USERS", "AUDIT_TRAIL")
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wbCopy = Workbooks.Open(dest)
    For Each ws In wbCopy.Worksheets
        isKeep = False
        For Each k In keep
            If UCase(ws.Name) = CStr(k) Then isKeep = True
        Next k
        If Not isKeep And wbCopy.Worksheets.Count > 1 Then ws.Delete
    Next ws
    On Error Resume Next
    wbCopy.Sheets("AUDIT_TRAIL").Rows("2:1048576").ClearContents
    wbCopy.Sheets("MENU").Activate
    On Error GoTo ErrHandler
    wbCopy.Save
    wbCopy.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Call Core_Engine.WriteToAuditLog("SHARE", "Copie distribuable creee: " & dest)
    MsgBox "Copie prete a partager:" & vbCrLf & dest & vbCrLf & vbCrLf & _
           "Le destinataire ouvre le fichier, clique 'Activer le contenu', et utilise le MENU." & vbCrLf & _
           "Compte admin initial: admin / Admin@2024! (changement obligatoire).", vbInformation, "Preparer partage"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True: Application.DisplayAlerts = True: Application.EnableEvents = True
    MsgBox "Preparer partage: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

'===============================================================================
' GL MONITORING (regles de la Knowledge Sharing Session DAI du 09/09/2026)
'===============================================================================

Public Sub LancerGLMonitoring()
    On Error GoTo ErrHandler
    If Not Core_Engine.FeuilleExiste("BALANCE_DATA") Then
        MsgBox "Lancez d'abord l'analyse complete (BALANCE_DATA requis).", vbExclamation, "GL Monitoring": Exit Sub
    End If
    Application.ScreenUpdating = False
    Call GL_Monitoring.Lancer_GL_Monitoring(True)
    Application.ScreenUpdating = True
    Call ActiverFeuille("GL_MONITORING")
    On Error Resume Next
    Call SAFA_Menu.RefreshStatus
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "GL Monitoring: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub LancerQualiteProofs()
    On Error GoTo ErrHandler
    Call Proof_Quality.AnalyserDossierProofs
    Call ActiverFeuille("PROOF_QUALITY")
    Exit Sub
ErrHandler:
    MsgBox "Qualite des proofs: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

Public Sub ImporterBalancePrecedente()
    ' Balance de la periode precedente (meme format que la Balance) -> BALANCE_PREV_RAW (variations GLM-007)
    Dim filePath As String, result As Data_Ingestion.ImportResult
    filePath = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm,Fichiers CSV (*.csv),*.csv,Tous (*.*),*.*", _
        Title:="Selectionner la Balance de la periode PRECEDENTE (N-1)")
    If filePath = "False" Or filePath = "" Then Exit Sub
    Application.ScreenUpdating = False
    result = Data_Ingestion.ImportFile(filePath, "BALANCE_PREV_RAW", True)
    Application.ScreenUpdating = True
    If result.Success Then
        MsgBox "Balance N-1 importee: " & result.RowsImported & " lignes." & vbCrLf & "Relancez GL Monitoring pour evaluer les variations de charges.", vbInformation, "S.A.F.A"
        Call Core_Engine.WriteToAuditLog("IMPORT", "Balance N-1 importee: " & result.RowsImported & " lignes")
    Else
        MsgBox "Erreur lors de l'import:" & vbCrLf & result.Errors, vbCritical, "Erreur Import"
    End If
End Sub

Public Sub ImporterListeProofables()
    ' Liste des comptes proofables tenue par le Controle Interne (colonne A = numero de compte) -> PROOFABLE_LIST
    Dim filePath As String, result As Data_Ingestion.ImportResult
    filePath = Application.GetOpenFilename( _
        FileFilter:="Fichiers Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm,Fichiers CSV (*.csv),*.csv,Tous (*.*),*.*", _
        Title:="Selectionner la liste des comptes proofables du Controle Interne")
    If filePath = "False" Or filePath = "" Then Exit Sub
    Application.ScreenUpdating = False
    result = Data_Ingestion.ImportFile(filePath, "PROOFABLE_LIST", True)
    Application.ScreenUpdating = True
    If result.Success Then
        MsgBox "Liste CI importee: " & result.RowsImported & " lignes (colonne A = numero de compte)." & vbCrLf & _
               "Relancez GL Monitoring pour verifier l'exhaustivite (GLM-001).", vbInformation, "S.A.F.A"
        Call Core_Engine.WriteToAuditLog("IMPORT", "Liste comptes proofables CI importee: " & result.RowsImported & " lignes")
    Else
        MsgBox "Erreur lors de l'import:" & vbCrLf & result.Errors, vbCritical, "Erreur Import"
    End If
End Sub
