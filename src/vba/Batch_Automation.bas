Attribute VB_Name = "Batch_Automation"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: BATCH_AUTOMATION v10.0
' ==============================================================================
' Description: Automatisation et traitement par lots
' Fonctionnalités:
'   - Traitement sans intervention humaine
'   - Import automatique depuis dossier surveillé
'   - Planification et scheduling
'   - Export automatique des rapports
'   - Alertes automatiques par email
'   - Journalisation complète
' ==============================================================================

' --- CONSTANTES ---
Private Const MODULE_NAME As String = "Batch_Automation"
Private Const LOG_FILE_NAME As String = "SAFA_Batch_Log.txt"

' --- TYPES ---
Public Type BatchConfig
    WatchFolder As String
    OutputFolder As String
    BalancePattern As String
    GLProofPattern As String
    AutoEmail As Boolean
    EmailRecipient As String
    AutoArchive As Boolean
    ArchiveFolder As String
    Silent As Boolean
End Type

Public Type BatchResult
    StartTime As Date
    EndTime As Date
    Duration As Double
    FilesProcessed As Integer
    Errors As String
    RecordsImported As Long
    AlertsGenerated As Long
    ReportPath As String
    Success As Boolean
End Type

' --- VARIABLES MODULE ---
Private mConfig As BatchConfig
Private mResult As BatchResult

' ==============================================================================
' 1. POINT D'ENTRÉE PRINCIPAL - TRAITEMENT COMPLET AUTOMATISÉ
' ==============================================================================

Public Function ExecuterTraitementBatch(Optional configPath As String = "") As BatchResult
    On Error GoTo BatchError

    mResult.StartTime = Now
    mResult.FilesProcessed = 0
    mResult.RecordsImported = 0
    mResult.AlertsGenerated = 0
    mResult.Errors = ""
    mResult.Success = False

    ' Initialiser configuration
    If configPath <> "" Then
        Call ChargerConfigurationBatch(configPath)
    Else
        Call SetDefaultBatchConfig
    End If

    ' Journaliser le démarrage
    Call LogBatch("INFO", "=== DÉMARRAGE TRAITEMENT BATCH ===")
    Call LogBatch("INFO", "Version S.A.F.A: " & SAFA_Common.SAFA_VERSION)
    Call LogBatch("INFO", "Utilisateur: " & Environ("USERNAME"))

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 1: DIAGNOSTIC RAPIDE
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 1: Diagnostic système...")

    If Not Auto_Diagnostic.DiagnosticRapide() Then
        Call LogBatch("ERROR", "Diagnostic échoué - Traitement annulé")
        mResult.Errors = "Échec diagnostic système"
        GoTo BatchEnd
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 2: IMPORT AUTOMATIQUE DES FICHIERS
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 2: Recherche et import des fichiers...")

    If mConfig.WatchFolder <> "" Then
        Call ImporterFichiersAutomatiquement
    Else
        Call LogBatch("WARN", "Pas de dossier surveillé configuré - Import manuel requis")
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 3: TRAITEMENT PRINCIPAL
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 3: Lancement du traitement principal...")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Appeler le traitement complet du Core_Engine
    On Error Resume Next
    Call Core_Engine.Lancer_Traitement_Complet
    If Err.Number <> 0 Then
        Call LogBatch("ERROR", "Erreur traitement principal: " & Err.Description)
        mResult.Errors = mResult.Errors & "Traitement principal: " & Err.Description & "; "
        Err.Clear
    End If
    On Error GoTo BatchError

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 4: ANALYSES AVANCÉES
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 4: Analyses avancées...")

    ' Analyse forensique
    On Error Resume Next
    Call Forensic_Rules.Lancer_Analyse_Forensique
    If Err.Number <> 0 Then
        Call LogBatch("WARN", "Erreur analyse forensique: " & Err.Description)
        Err.Clear
    End If

    ' Analyse Benford
    Call Forensic_Rules.Lancer_Benford_Enhanced
    If Err.Number <> 0 Then
        Call LogBatch("WARN", "Erreur analyse Benford: " & Err.Description)
        Err.Clear
    End If

    ' Conformité réglementaire
    Call Regulatory_Compliance.Lancer_Verification_Conformite
    If Err.Number <> 0 Then
        Call LogBatch("WARN", "Erreur vérification conformité: " & Err.Description)
        Err.Clear
    End If

    ' Analyse temporelle
    Call Temporal_Analysis.Lancer_Analyse_Temporelle
    If Err.Number <> 0 Then
        Call LogBatch("WARN", "Erreur analyse temporelle: " & Err.Description)
        Err.Clear
    End If

    ' Analyse réseau
    Call Network_Analysis.Lancer_Analyse_Reseau
    If Err.Number <> 0 Then
        Call LogBatch("WARN", "Erreur analyse réseau: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo BatchError

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 5: GÉNÉRATION DES RAPPORTS
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 5: Génération des rapports...")

    Dim reportPath As String
    If mConfig.OutputFolder <> "" Then
        reportPath = mConfig.OutputFolder & "\SAFA_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    Else
        reportPath = ThisWorkbook.Path & "\SAFA_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    End If

    On Error Resume Next
    Call Report_Generator.GenerateFullReport(reportPath)
    If Err.Number = 0 Then
        Call Report_Generator.ExportToPDF(ThisWorkbook, reportPath)
        mResult.ReportPath = reportPath
        Call LogBatch("INFO", "Rapport généré: " & reportPath)
    Else
        Call LogBatch("WARN", "Erreur génération rapport: " & Err.Description)
        Err.Clear
    End If
    On Error GoTo BatchError

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 6: COMPTAGE DES ALERTES
    ' ═══════════════════════════════════════════════════════════════
    Call LogBatch("INFO", "Étape 6: Compilation des résultats...")

    If SAFA_Common.FeuilleExiste("AUDIT_REPORT") Then
        Dim wsAudit As Worksheet
        Set wsAudit = ThisWorkbook.Sheets("AUDIT_REPORT")
        mResult.AlertsGenerated = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row - 1
    End If

    If SAFA_Common.FeuilleExiste("RECONCIL") Then
        Dim wsReconcil As Worksheet
        Set wsReconcil = ThisWorkbook.Sheets("RECONCIL")
        mResult.RecordsImported = wsReconcil.Cells(wsReconcil.Rows.Count, 1).End(xlUp).Row - 1
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 7: ALERTES EMAIL AUTOMATIQUES
    ' ═══════════════════════════════════════════════════════════════
    If mConfig.AutoEmail And mConfig.EmailRecipient <> "" Then
        Call LogBatch("INFO", "Étape 7: Envoi alertes email...")

        Dim criticalCount As Long
        criticalCount = 0

        If SAFA_Common.FeuilleExiste("AUDIT_REPORT") Then
            On Error Resume Next
            criticalCount = Application.WorksheetFunction.CountIf(wsAudit.Columns("C"), "CRITICAL")
            On Error GoTo BatchError
        End If

        If criticalCount > 0 Then
            On Error Resume Next
            Call Report_Generator.SendCriticalAlerts(ThisWorkbook, mConfig.EmailRecipient)
            If Err.Number = 0 Then
                Call LogBatch("INFO", "Alertes email envoyées: " & criticalCount & " critiques")
            Else
                Call LogBatch("WARN", "Échec envoi email: " & Err.Description)
                Err.Clear
            End If
            On Error GoTo BatchError
        Else
            Call LogBatch("INFO", "Pas d'alerte critique - Email non envoyé")
        End If
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' ÉTAPE 8: ARCHIVAGE
    ' ═══════════════════════════════════════════════════════════════
    If mConfig.AutoArchive And mConfig.ArchiveFolder <> "" Then
        Call LogBatch("INFO", "Étape 8: Archivage...")
        Call ArchiverFichiersTraites
    End If

    ' Succès
    mResult.Success = True
    Call LogBatch("INFO", "=== TRAITEMENT BATCH TERMINÉ AVEC SUCCÈS ===")

BatchEnd:
    mResult.EndTime = Now
    mResult.Duration = (mResult.EndTime - mResult.StartTime) * 86400 ' Secondes

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' Log final
    Call LogBatch("INFO", "Durée totale: " & Format(mResult.Duration, "0.0") & " secondes")
    Call LogBatch("INFO", "Fichiers traités: " & mResult.FilesProcessed)
    Call LogBatch("INFO", "Enregistrements importés: " & mResult.RecordsImported)
    Call LogBatch("INFO", "Alertes générées: " & mResult.AlertsGenerated)

    If mResult.Errors <> "" Then
        Call LogBatch("WARN", "Erreurs rencontrées: " & mResult.Errors)
    End If

    ' Journaliser dans audit trail
    SAFA_Common.WriteAuditLog "BATCH", "Traitement batch terminé", _
                               "Success: " & mResult.Success & _
                               ", Durée: " & Format(mResult.Duration, "0.0") & "s" & _
                               ", Alertes: " & mResult.AlertsGenerated

    ExecuterTraitementBatch = mResult
    Exit Function

BatchError:
    mResult.Errors = mResult.Errors & "Erreur fatale: " & Err.Description & "; "
    Call LogBatch("ERROR", "ERREUR FATALE: " & Err.Number & " - " & Err.Description)
    mResult.Success = False
    Resume BatchEnd
End Function

' ==============================================================================
' 2. CONFIGURATION
' ==============================================================================

Private Sub SetDefaultBatchConfig()
    mConfig.WatchFolder = ThisWorkbook.Path & "\Input"
    mConfig.OutputFolder = ThisWorkbook.Path & "\Output"
    mConfig.BalancePattern = "*Balance*.xlsx"
    mConfig.GLProofPattern = "*GLProof*.xlsx"
    mConfig.AutoEmail = False
    mConfig.EmailRecipient = ""
    mConfig.AutoArchive = True
    mConfig.ArchiveFolder = ThisWorkbook.Path & "\Archive"
    mConfig.Silent = True
End Sub

Public Sub ChargerConfigurationBatch(configPath As String)
    ' Charge la configuration depuis un fichier texte
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String

    On Error GoTo ConfigError

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(configPath) Then
        Call SetDefaultBatchConfig
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(configPath, 1, False)

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        If line <> "" And Left(line, 1) <> "#" Then
            parts = Split(line, "=")

            If UBound(parts) >= 1 Then
                Select Case UCase(Trim(parts(0)))
                    Case "WATCHFOLDER"
                        mConfig.WatchFolder = Trim(parts(1))
                    Case "OUTPUTFOLDER"
                        mConfig.OutputFolder = Trim(parts(1))
                    Case "BALANCEPATTERN"
                        mConfig.BalancePattern = Trim(parts(1))
                    Case "GLPROOFPATTERN"
                        mConfig.GLProofPattern = Trim(parts(1))
                    Case "AUTOEMAIL"
                        mConfig.AutoEmail = (UCase(Trim(parts(1))) = "TRUE")
                    Case "EMAILRECIPIENT"
                        mConfig.EmailRecipient = Trim(parts(1))
                    Case "AUTOARCHIVE"
                        mConfig.AutoArchive = (UCase(Trim(parts(1))) = "TRUE")
                    Case "ARCHIVEFOLDER"
                        mConfig.ArchiveFolder = Trim(parts(1))
                    Case "SILENT"
                        mConfig.Silent = (UCase(Trim(parts(1))) = "TRUE")
                End Select
            End If
        End If
    Loop

    ts.Close
    Exit Sub

ConfigError:
    Call SetDefaultBatchConfig
End Sub

Public Sub SauvegarderConfigurationBatch(configPath As String)
    ' Sauvegarde la configuration dans un fichier texte
    Dim fso As Object
    Dim ts As Object

    On Error GoTo SaveError

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(configPath, True)

    ts.WriteLine "# Configuration Batch S.A.F.A"
    ts.WriteLine "# Généré le " & Format(Now, "dd/mm/yyyy hh:nn")
    ts.WriteLine ""
    ts.WriteLine "WatchFolder=" & mConfig.WatchFolder
    ts.WriteLine "OutputFolder=" & mConfig.OutputFolder
    ts.WriteLine "BalancePattern=" & mConfig.BalancePattern
    ts.WriteLine "GLProofPattern=" & mConfig.GLProofPattern
    ts.WriteLine "AutoEmail=" & IIf(mConfig.AutoEmail, "TRUE", "FALSE")
    ts.WriteLine "EmailRecipient=" & mConfig.EmailRecipient
    ts.WriteLine "AutoArchive=" & IIf(mConfig.AutoArchive, "TRUE", "FALSE")
    ts.WriteLine "ArchiveFolder=" & mConfig.ArchiveFolder
    ts.WriteLine "Silent=" & IIf(mConfig.Silent, "TRUE", "FALSE")

    ts.Close
    Exit Sub

SaveError:
    ' Silently fail
End Sub

' ==============================================================================
' 3. IMPORT AUTOMATIQUE DES FICHIERS
' ==============================================================================

Private Sub ImporterFichiersAutomatiquement()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim balanceFile As String
    Dim glProofFile As String

    On Error GoTo ImportError

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(mConfig.WatchFolder) Then
        Call LogBatch("WARN", "Dossier surveillé non trouvé: " & mConfig.WatchFolder)
        Exit Sub
    End If

    Set folder = fso.GetFolder(mConfig.WatchFolder)

    ' Chercher fichiers Balance et GL Proof
    For Each file In folder.Files
        If file.Name Like mConfig.BalancePattern Then
            If balanceFile = "" Or file.DateLastModified > fso.GetFile(balanceFile).DateLastModified Then
                balanceFile = file.Path
            End If
        End If

        If file.Name Like mConfig.GLProofPattern Then
            If glProofFile = "" Or file.DateLastModified > fso.GetFile(glProofFile).DateLastModified Then
                glProofFile = file.Path
            End If
        End If
    Next file

    ' Importer Balance
    If balanceFile <> "" Then
        Call LogBatch("INFO", "Import Balance: " & balanceFile)

        Dim balanceResult As Data_Ingestion.ImportResult
        balanceResult = Data_Ingestion.ImportFile(balanceFile, "BALANCE_DATA", True)

        If balanceResult.Success Then
            Call LogBatch("INFO", "Balance importé: " & balanceResult.RowsImported & " lignes")
            mResult.FilesProcessed = mResult.FilesProcessed + 1
            mResult.RecordsImported = mResult.RecordsImported + balanceResult.RowsImported
        Else
            Call LogBatch("ERROR", "Échec import Balance: " & balanceResult.Errors)
            mResult.Errors = mResult.Errors & "Balance: " & balanceResult.Errors & "; "
        End If
    Else
        Call LogBatch("WARN", "Aucun fichier Balance trouvé avec pattern: " & mConfig.BalancePattern)
    End If

    ' Importer GL Proof
    If glProofFile <> "" Then
        Call LogBatch("INFO", "Import GL Proof: " & glProofFile)

        Dim glResult As Data_Ingestion.ImportResult
        glResult = Data_Ingestion.ImportFile(glProofFile, "GLPROOF_DATA", True)

        If glResult.Success Then
            Call LogBatch("INFO", "GL Proof importé: " & glResult.RowsImported & " lignes")
            mResult.FilesProcessed = mResult.FilesProcessed + 1
            mResult.RecordsImported = mResult.RecordsImported + glResult.RowsImported
        Else
            Call LogBatch("ERROR", "Échec import GL Proof: " & glResult.Errors)
            mResult.Errors = mResult.Errors & "GL Proof: " & glResult.Errors & "; "
        End If
    Else
        Call LogBatch("WARN", "Aucun fichier GL Proof trouvé avec pattern: " & mConfig.GLProofPattern)
    End If

    Exit Sub

ImportError:
    Call LogBatch("ERROR", "Erreur import automatique: " & Err.Description)
    mResult.Errors = mResult.Errors & "Import: " & Err.Description & "; "
End Sub

' ==============================================================================
' 4. ARCHIVAGE
' ==============================================================================

Private Sub ArchiverFichiersTraites()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim archiveSubFolder As String

    On Error GoTo ArchiveError

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Créer sous-dossier d'archive avec date
    archiveSubFolder = mConfig.ArchiveFolder & "\" & Format(Now, "yyyymmdd_hhnnss")

    If Not fso.FolderExists(mConfig.ArchiveFolder) Then
        fso.CreateFolder mConfig.ArchiveFolder
    End If

    fso.CreateFolder archiveSubFolder

    ' Déplacer fichiers traités
    If fso.FolderExists(mConfig.WatchFolder) Then
        Set folder = fso.GetFolder(mConfig.WatchFolder)

        For Each file In folder.Files
            If file.Name Like mConfig.BalancePattern Or file.Name Like mConfig.GLProofPattern Then
                fso.MoveFile file.Path, archiveSubFolder & "\" & file.Name
                Call LogBatch("INFO", "Archivé: " & file.Name)
            End If
        Next file
    End If

    Call LogBatch("INFO", "Archivage terminé: " & archiveSubFolder)
    Exit Sub

ArchiveError:
    Call LogBatch("WARN", "Erreur archivage: " & Err.Description)
End Sub

' ==============================================================================
' 5. JOURNALISATION
' ==============================================================================

Private Sub LogBatch(level As String, message As String)
    ' Écrit dans le fichier log et dans Debug
    Dim fso As Object
    Dim ts As Object
    Dim logPath As String
    Dim logLine As String

    logLine = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & message

    ' Debug output
    Debug.Print logLine

    ' Fichier log
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = ThisWorkbook.Path & "\" & LOG_FILE_NAME

    Set ts = fso.OpenTextFile(logPath, 8, True) ' 8 = ForAppending
    If Not ts Is Nothing Then
        ts.WriteLine logLine
        ts.Close
    End If
    On Error GoTo 0

    ' Afficher dans barre de statut si non silencieux
    If Not mConfig.Silent Then
        Application.StatusBar = "[S.A.F.A] " & message
    End If
End Sub

' ==============================================================================
' 6. MODE LIGNE DE COMMANDE
' ==============================================================================

Public Sub ExecuterEnModeCommande()
    ' Fonction appelable depuis un script VBScript externe
    ' Usage: cscript.exe RunSAFA.vbs
    '
    ' Exemple de script VBScript (RunSAFA.vbs):
    ' ---------------------------------------------
    ' Set objExcel = CreateObject("Excel.Application")
    ' objExcel.Visible = False
    ' Set objWorkbook = objExcel.Workbooks.Open("C:\SAFA\SAFA_v10.xlsm")
    ' objExcel.Run "Batch_Automation.ExecuterEnModeCommande"
    ' objWorkbook.Save
    ' objWorkbook.Close
    ' objExcel.Quit
    ' ---------------------------------------------

    Dim result As BatchResult

    ' Mode silencieux
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Charger config depuis chemin standard
    Dim configPath As String
    configPath = ThisWorkbook.Path & "\batch_config.txt"

    ' Exécuter traitement
    result = ExecuterTraitementBatch(configPath)

    ' Sauvegarder
    ThisWorkbook.Save

    ' Log résultat
    If result.Success Then
        Call LogBatch("INFO", "Mode commande terminé avec succès")
    Else
        Call LogBatch("ERROR", "Mode commande terminé avec erreurs: " & result.Errors)
    End If
End Sub

' ==============================================================================
' 7. PLANIFICATION WINDOWS TASK SCHEDULER
' ==============================================================================

Public Sub GenererScriptPlanification()
    ' Génère les scripts pour planifier l'exécution automatique
    Dim fso As Object
    Dim ts As Object
    Dim basePath As String
    Dim vbsPath As String
    Dim batPath As String

    On Error GoTo ScriptError

    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.Path

    ' ═══════════════════════════════════════════════════════════════
    ' Script VBScript
    ' ═══════════════════════════════════════════════════════════════
    vbsPath = basePath & "\RunSAFA.vbs"
    Set ts = fso.CreateTextFile(vbsPath, True)

    ts.WriteLine "' Script automatisé S.A.F.A v" & SAFA_Common.SAFA_VERSION
    ts.WriteLine "' Généré le " & Format(Now, "dd/mm/yyyy hh:nn")
    ts.WriteLine "'"
    ts.WriteLine "' Usage: cscript.exe RunSAFA.vbs"
    ts.WriteLine "' Ou via Windows Task Scheduler"
    ts.WriteLine ""
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine ""
    ts.WriteLine "Dim objExcel"
    ts.WriteLine "Dim objWorkbook"
    ts.WriteLine ""
    ts.WriteLine "Set objExcel = CreateObject(""Excel.Application"")"
    ts.WriteLine "objExcel.Visible = False"
    ts.WriteLine "objExcel.DisplayAlerts = False"
    ts.WriteLine ""
    ts.WriteLine "Set objWorkbook = objExcel.Workbooks.Open(""" & ThisWorkbook.FullName & """)"
    ts.WriteLine ""
    ts.WriteLine "If Err.Number = 0 Then"
    ts.WriteLine "    objExcel.Run ""Batch_Automation.ExecuterEnModeCommande"""
    ts.WriteLine "    objWorkbook.Save"
    ts.WriteLine "    objWorkbook.Close False"
    ts.WriteLine "End If"
    ts.WriteLine ""
    ts.WriteLine "objExcel.Quit"
    ts.WriteLine ""
    ts.WriteLine "Set objWorkbook = Nothing"
    ts.WriteLine "Set objExcel = Nothing"
    ts.WriteLine ""
    ts.WriteLine "WScript.Quit IIf(Err.Number = 0, 0, 1)"

    ts.Close

    ' ═══════════════════════════════════════════════════════════════
    ' Script BAT
    ' ═══════════════════════════════════════════════════════════════
    batPath = basePath & "\RunSAFA.bat"
    Set ts = fso.CreateTextFile(batPath, True)

    ts.WriteLine "@echo off"
    ts.WriteLine "REM Script automatisé S.A.F.A v" & SAFA_Common.SAFA_VERSION
    ts.WriteLine "REM Généré le " & Format(Now, "dd/mm/yyyy hh:nn")
    ts.WriteLine ""
    ts.WriteLine "cd /d """ & basePath & """"
    ts.WriteLine "cscript.exe //NoLogo RunSAFA.vbs"
    ts.WriteLine ""
    ts.WriteLine "if %ERRORLEVEL% EQU 0 ("
    ts.WriteLine "    echo Traitement S.A.F.A terminé avec succès"
    ts.WriteLine ") else ("
    ts.WriteLine "    echo ERREUR: Traitement S.A.F.A échoué"
    ts.WriteLine ")"
    ts.WriteLine ""
    ts.WriteLine "pause"

    ts.Close

    ' ═══════════════════════════════════════════════════════════════
    ' Configuration batch par défaut
    ' ═══════════════════════════════════════════════════════════════
    Dim configPath As String
    configPath = basePath & "\batch_config.txt"

    If Not fso.FileExists(configPath) Then
        Call SetDefaultBatchConfig
        Call SauvegarderConfigurationBatch(configPath)
    End If

    MsgBox "Scripts de planification générés:" & vbCrLf & vbCrLf & _
           "1. " & vbsPath & vbCrLf & _
           "2. " & batPath & vbCrLf & _
           "3. " & configPath & vbCrLf & vbCrLf & _
           "Pour planifier l'exécution automatique:" & vbCrLf & _
           "1. Ouvrir Windows Task Scheduler" & vbCrLf & _
           "2. Créer une tâche planifiée" & vbCrLf & _
           "3. Action: Démarrer un programme" & vbCrLf & _
           "4. Programme: cscript.exe" & vbCrLf & _
           "5. Arguments: """ & vbsPath & """", _
           vbInformation, "Scripts Générés"

    Exit Sub

ScriptError:
    MsgBox "Erreur génération scripts: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' 8. UTILITAIRES
' ==============================================================================

Public Function GetBatchConfig() As BatchConfig
    GetBatchConfig = mConfig
End Function

Public Function GetLastBatchResult() As BatchResult
    GetLastBatchResult = mResult
End Function

Public Sub AfficherDernierResultat()
    MsgBox "Dernier traitement batch:" & vbCrLf & vbCrLf & _
           "Début: " & Format(mResult.StartTime, "dd/mm/yyyy hh:nn:ss") & vbCrLf & _
           "Fin: " & Format(mResult.EndTime, "dd/mm/yyyy hh:nn:ss") & vbCrLf & _
           "Durée: " & Format(mResult.Duration, "0.0") & " secondes" & vbCrLf & _
           "Fichiers: " & mResult.FilesProcessed & vbCrLf & _
           "Enregistrements: " & mResult.RecordsImported & vbCrLf & _
           "Alertes: " & mResult.AlertsGenerated & vbCrLf & _
           "Succès: " & IIf(mResult.Success, "OUI", "NON") & vbCrLf & _
           IIf(mResult.Errors <> "", "Erreurs: " & mResult.Errors, ""), _
           IIf(mResult.Success, vbInformation, vbExclamation), _
           "Résultat Batch"
End Sub
