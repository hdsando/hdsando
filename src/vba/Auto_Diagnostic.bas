Attribute VB_Name = "Auto_Diagnostic"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: AUTO_DIAGNOSTIC v10.0
' ==============================================================================
' Description: Auto-diagnostic et vérification de l'intégrité du système
' Fonctionnalités:
'   - Vérification des feuilles requises
'   - Validation de la configuration
'   - Test de l'intégrité de l'audit trail
'   - Diagnostic de la connectivité (Outlook, FileSystem)
'   - Vérification des dépendances
'   - Rapport de santé système
' ==============================================================================

' --- CONSTANTES ---
Private Const MODULE_NAME As String = "Auto_Diagnostic"

' --- TYPES ---
Public Type DiagnosticResult
    TestName As String
    Category As String
    Status As String      ' PASS, WARN, FAIL
    Message As String
    Recommendation As String
    Timestamp As Date
End Type

Public Type SystemHealth
    OverallStatus As String
    PassCount As Integer
    WarnCount As Integer
    FailCount As Integer
    CriticalIssues As String
    LastCheck As Date
End Type

' --- VARIABLES MODULE ---
Private mResults As Collection
Private mHealth As SystemHealth

' ==============================================================================
' 1. POINT D'ENTRÉE PRINCIPAL
' ==============================================================================

Public Function LancerDiagnosticComplet() As SystemHealth
    Dim wsOut As Worksheet
    Dim startTime As Double

    On Error GoTo DiagError

    startTime = Timer

    ' Initialiser
    Set mResults = New Collection
    mHealth.PassCount = 0
    mHealth.WarnCount = 0
    mHealth.FailCount = 0
    mHealth.CriticalIssues = ""
    mHealth.LastCheck = Now

    ' Créer feuille de rapport
    Set wsOut = SAFA_Common.GetOrCreateSheet("SYSTEM_DIAGNOSTIC", True)

    ' Titre
    wsOut.Range("A1").Value = "DIAGNOSTIC SYSTÈME S.A.F.A v" & SAFA_Common.SAFA_VERSION
    wsOut.Range("A1").Font.Size = 16
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1:F1").Merge
    wsOut.Range("A1:F1").Interior.Color = RGB(0, 51, 102)
    wsOut.Range("A1:F1").Font.Color = vbWhite

    wsOut.Range("A2").Value = "Date d'exécution: " & Format(Now, "dd/mm/yyyy hh:nn:ss")

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 1: VÉRIFICATION DES FEUILLES REQUISES
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerFeuilles

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 2: VÉRIFICATION DE LA CONFIGURATION
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerConfiguration

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 3: VÉRIFICATION DE L'AUDIT TRAIL
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerAuditTrail

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 4: VÉRIFICATION DES DÉPENDANCES
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerDependances

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 5: VÉRIFICATION DE LA SÉCURITÉ
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerSecurite

    ' ═══════════════════════════════════════════════════════════════
    ' SECTION 6: VÉRIFICATION DES PERFORMANCES
    ' ═══════════════════════════════════════════════════════════════
    Call DiagnostiquerPerformances

    ' Générer rapport
    Call GenererRapportDiagnostic(wsOut)

    ' Déterminer statut global
    If mHealth.FailCount > 0 Then
        mHealth.OverallStatus = "CRITICAL"
    ElseIf mHealth.WarnCount > 3 Then
        mHealth.OverallStatus = "WARNING"
    ElseIf mHealth.WarnCount > 0 Then
        mHealth.OverallStatus = "ACCEPTABLE"
    Else
        mHealth.OverallStatus = "HEALTHY"
    End If

    SAFA_Common.WriteAuditLog "DIAGNOSTIC", "Diagnostic complet terminé", _
                               "Status: " & mHealth.OverallStatus & _
                               ", Pass: " & mHealth.PassCount & _
                               ", Warn: " & mHealth.WarnCount & _
                               ", Fail: " & mHealth.FailCount

    LancerDiagnosticComplet = mHealth
    Exit Function

DiagError:
    SAFA_Common.LogError MODULE_NAME, "LancerDiagnosticComplet", Err.Number, Err.Description
    mHealth.OverallStatus = "ERROR"
    mHealth.CriticalIssues = "Erreur diagnostic: " & Err.Description
    LancerDiagnosticComplet = mHealth
End Function

' ==============================================================================
' 2. DIAGNOSTIC DES FEUILLES
' ==============================================================================

Private Sub DiagnostiquerFeuilles()
    ' Feuilles obligatoires
    Dim requiredSheets As Variant
    requiredSheets = Array("BALANCE_DATA", "GLPROOF_DATA", "RECONCIL", "AUDIT_REPORT", "AUDIT_TRAIL")

    Dim optionalSheets As Variant
    optionalSheets = Array("TRANSACTION_DATA", "HISTORY_LOG", "PARAM", "CONFIG", "USERS")

    Dim sheetName As Variant
    Dim result As DiagnosticResult

    ' Vérifier feuilles obligatoires
    For Each sheetName In requiredSheets
        result.TestName = "Feuille " & CStr(sheetName)
        result.Category = "Structure"
        result.Timestamp = Now

        If SAFA_Common.FeuilleExiste(CStr(sheetName)) Then
            result.Status = "PASS"
            result.Message = "Présente"
            result.Recommendation = ""
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "Absente - sera créée automatiquement"
            result.Recommendation = "Exécuter le traitement pour créer"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If

        mResults.Add result
    Next sheetName

    ' Vérifier feuilles optionnelles
    For Each sheetName In optionalSheets
        result.TestName = "Feuille " & CStr(sheetName) & " (optionnel)"
        result.Category = "Structure"
        result.Timestamp = Now

        If SAFA_Common.FeuilleExiste(CStr(sheetName)) Then
            result.Status = "PASS"
            result.Message = "Présente"
            result.Recommendation = ""
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "INFO"
            result.Message = "Absente - optionnelle"
            result.Recommendation = ""
        End If

        mResults.Add result
    Next sheetName

    ' Vérifier l'intégrité des données si feuilles présentes
    If SAFA_Common.FeuilleExiste("BALANCE_DATA") Then
        Dim wsBalance As Worksheet
        Set wsBalance = ThisWorkbook.Sheets("BALANCE_DATA")
        Dim balanceRows As Long
        balanceRows = wsBalance.Cells(wsBalance.Rows.Count, 1).End(xlUp).Row - 1

        result.TestName = "Données Balance"
        result.Category = "Données"
        result.Timestamp = Now

        If balanceRows > 0 Then
            result.Status = "PASS"
            result.Message = balanceRows & " lignes"
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "Aucune donnée"
            result.Recommendation = "Importer les données Balance"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If

        mResults.Add result
    End If

    If SAFA_Common.FeuilleExiste("GLPROOF_DATA") Then
        Dim wsGL As Worksheet
        Set wsGL = ThisWorkbook.Sheets("GLPROOF_DATA")
        Dim glRows As Long
        glRows = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row - 1

        result.TestName = "Données GL Proof"
        result.Category = "Données"
        result.Timestamp = Now

        If glRows > 0 Then
            result.Status = "PASS"
            result.Message = glRows & " lignes"
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "Aucune donnée"
            result.Recommendation = "Importer les données GL Proof"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If

        mResults.Add result
    End If
End Sub

' ==============================================================================
' 3. DIAGNOSTIC DE LA CONFIGURATION
' ==============================================================================

Private Sub DiagnostiquerConfiguration()
    Dim result As DiagnosticResult

    ' Vérifier la feuille PARAM
    If SAFA_Common.FeuilleExiste("PARAM") Then
        Dim wsParam As Worksheet
        Set wsParam = ThisWorkbook.Sheets("PARAM")

        result.TestName = "Configuration Tolérance"
        result.Category = "Configuration"
        result.Timestamp = Now

        Dim tolerance As Double
        tolerance = SAFA_Common.SafeVal(wsParam.Range("B2").Value)

        If tolerance > 0 And tolerance < 10000000 Then
            result.Status = "PASS"
            result.Message = "Tolérance = " & Format(tolerance, "#,##0")
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "Tolérance non définie ou invalide"
            result.Recommendation = "Vérifier la valeur de tolérance dans PARAM"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If

        mResults.Add result

        ' Vérifier SOL_ID
        result.TestName = "Configuration SOL_ID"
        result.Category = "Configuration"
        result.Timestamp = Now

        Dim solId As String
        solId = SAFA_Common.SafeText(wsParam.Range("B3").Value)

        If Len(solId) >= 3 Then
            result.Status = "PASS"
            result.Message = "SOL_ID = " & solId
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "SOL_ID non défini"
            result.Recommendation = "Définir SOL_ID dans PARAM"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If

        mResults.Add result
    Else
        result.TestName = "Feuille PARAM"
        result.Category = "Configuration"
        result.Timestamp = Now
        result.Status = "WARN"
        result.Message = "Feuille PARAM absente - Utilisation des valeurs par défaut"
        result.Recommendation = "Créer et configurer la feuille PARAM"
        mHealth.WarnCount = mHealth.WarnCount + 1
        mResults.Add result
    End If

    ' Vérifier les macros activées
    result.TestName = "Macros VBA"
    result.Category = "Configuration"
    result.Timestamp = Now
    result.Status = "PASS"
    result.Message = "Macros activées (sinon ce code ne s'exécuterait pas)"
    mHealth.PassCount = mHealth.PassCount + 1
    mResults.Add result
End Sub

' ==============================================================================
' 4. DIAGNOSTIC DE L'AUDIT TRAIL
' ==============================================================================

Private Sub DiagnostiquerAuditTrail()
    Dim result As DiagnosticResult

    If Not SAFA_Common.FeuilleExiste("AUDIT_TRAIL") Then
        result.TestName = "Audit Trail - Existence"
        result.Category = "Sécurité"
        result.Timestamp = Now
        result.Status = "INFO"
        result.Message = "Audit trail non initialisé"
        result.Recommendation = "Sera créé au premier log"
        mResults.Add result
        Exit Sub
    End If

    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("AUDIT_TRAIL")

    Dim lastRow As Long
    lastRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row

    result.TestName = "Audit Trail - Taille"
    result.Category = "Sécurité"
    result.Timestamp = Now

    If lastRow > 1 Then
        result.Status = "PASS"
        result.Message = lastRow - 1 & " entrées"
        mHealth.PassCount = mHealth.PassCount + 1

        ' Vérifier si trop gros
        If lastRow > 100000 Then
            result.Status = "WARN"
            result.Message = lastRow - 1 & " entrées - Archive recommandée"
            result.Recommendation = "Archiver et purger l'audit trail"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If
    Else
        result.Status = "PASS"
        result.Message = "Vide"
    End If
    mResults.Add result

    ' Vérifier intégrité des hash (échantillon)
    If lastRow > 10 Then
        result.TestName = "Audit Trail - Intégrité Hash"
        result.Category = "Sécurité"
        result.Timestamp = Now

        Dim integrityOK As Boolean
        integrityOK = True

        Dim i As Long
        For i = 3 To WorksheetFunction.Min(lastRow, 20)
            Dim prevHash As String
            prevHash = SAFA_Common.SafeText(wsAudit.Cells(i - 1, 6).Value)

            Dim storedPrevHash As String
            storedPrevHash = SAFA_Common.SafeText(wsAudit.Cells(i, 7).Value)

            If prevHash <> storedPrevHash And storedPrevHash <> "" Then
                integrityOK = False
                Exit For
            End If
        Next i

        If integrityOK Then
            result.Status = "PASS"
            result.Message = "Chaîne de hash valide"
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "FAIL"
            result.Message = "ALERTE: Rupture de chaîne détectée!"
            result.Recommendation = "Investigation sécurité requise immédiatement"
            mHealth.FailCount = mHealth.FailCount + 1
            mHealth.CriticalIssues = mHealth.CriticalIssues & "Intégrité audit trail compromise; "
        End If
        mResults.Add result
    End If
End Sub

' ==============================================================================
' 5. DIAGNOSTIC DES DÉPENDANCES
' ==============================================================================

Private Sub DiagnostiquerDependances()
    Dim result As DiagnosticResult

    ' Test FileSystemObject
    result.TestName = "FileSystemObject"
    result.Category = "Dépendances"
    result.Timestamp = Now

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Err.Number = 0 And Not fso Is Nothing Then
        result.Status = "PASS"
        result.Message = "Disponible"
        mHealth.PassCount = mHealth.PassCount + 1
    Else
        result.Status = "FAIL"
        result.Message = "Non disponible - Import/Export impossible"
        result.Recommendation = "Vérifier les références VBA"
        mHealth.FailCount = mHealth.FailCount + 1
    End If
    Err.Clear
    On Error GoTo 0
    mResults.Add result

    ' Test Outlook
    result.TestName = "Microsoft Outlook"
    result.Category = "Dépendances"
    result.Timestamp = Now

    On Error Resume Next
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")

    If Err.Number = 0 And Not olApp Is Nothing Then
        result.Status = "PASS"
        result.Message = "Disponible pour alertes email"
        mHealth.PassCount = mHealth.PassCount + 1
        Set olApp = Nothing
    Else
        result.Status = "WARN"
        result.Message = "Non disponible - Alertes email désactivées"
        result.Recommendation = "Installer Outlook ou désactiver alertes email"
        mHealth.WarnCount = mHealth.WarnCount + 1
    End If
    Err.Clear
    On Error GoTo 0
    mResults.Add result

    ' Test Dictionary
    result.TestName = "Scripting.Dictionary"
    result.Category = "Dépendances"
    result.Timestamp = Now

    On Error Resume Next
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Err.Number = 0 And Not dict Is Nothing Then
        result.Status = "PASS"
        result.Message = "Disponible"
        mHealth.PassCount = mHealth.PassCount + 1
    Else
        result.Status = "FAIL"
        result.Message = "Non disponible - Traitement impossible"
        result.Recommendation = "Vérifier Microsoft Scripting Runtime"
        mHealth.FailCount = mHealth.FailCount + 1
        mHealth.CriticalIssues = mHealth.CriticalIssues & "Dictionary non disponible; "
    End If
    Err.Clear
    On Error GoTo 0
    mResults.Add result

    ' Test PDF Export
    result.TestName = "Export PDF"
    result.Category = "Dépendances"
    result.Timestamp = Now

    On Error Resume Next
    ' Test minimal
    Dim canExportPDF As Boolean
    canExportPDF = (Application.Version >= 12)

    If canExportPDF Then
        result.Status = "PASS"
        result.Message = "Disponible (Excel " & Application.Version & ")"
        mHealth.PassCount = mHealth.PassCount + 1
    Else
        result.Status = "WARN"
        result.Message = "Export PDF peut ne pas fonctionner"
        result.Recommendation = "Utiliser Excel 2007 ou supérieur"
        mHealth.WarnCount = mHealth.WarnCount + 1
    End If
    On Error GoTo 0
    mResults.Add result
End Sub

' ==============================================================================
' 6. DIAGNOSTIC DE LA SÉCURITÉ
' ==============================================================================

Private Sub DiagnostiquerSecurite()
    Dim result As DiagnosticResult

    ' Vérifier si feuille USERS existe et a des utilisateurs
    If SAFA_Common.FeuilleExiste("USERS") Then
        Dim wsUsers As Worksheet
        Set wsUsers = ThisWorkbook.Sheets("USERS")

        result.TestName = "Utilisateurs configurés"
        result.Category = "Sécurité"
        result.Timestamp = Now

        Dim userCount As Long
        userCount = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row - 1

        If userCount > 0 Then
            result.Status = "PASS"
            result.Message = userCount & " utilisateur(s)"
            mHealth.PassCount = mHealth.PassCount + 1
        Else
            result.Status = "WARN"
            result.Message = "Aucun utilisateur configuré"
            result.Recommendation = "Configurer au moins un utilisateur admin"
            mHealth.WarnCount = mHealth.WarnCount + 1
        End If
        mResults.Add result

        ' Vérifier si mot de passe par défaut toujours actif
        result.TestName = "Mot de passe par défaut"
        result.Category = "Sécurité"
        result.Timestamp = Now

        Dim defaultHash As String
        defaultHash = SAFA_Common.ComputeHash("admin123" & "SAFA_SALT_2024")

        Dim i As Long
        Dim defaultFound As Boolean
        defaultFound = False

        For i = 2 To userCount + 1
            If wsUsers.Cells(i, 2).Value = defaultHash Then
                defaultFound = True
                Exit For
            End If
        Next i

        If defaultFound Then
            result.Status = "WARN"
            result.Message = "Mot de passe par défaut toujours actif"
            result.Recommendation = "Changer le mot de passe admin immédiatement"
            mHealth.WarnCount = mHealth.WarnCount + 1
        Else
            result.Status = "PASS"
            result.Message = "Pas de mot de passe par défaut"
            mHealth.PassCount = mHealth.PassCount + 1
        End If
        mResults.Add result
    Else
        result.TestName = "Module Authentification"
        result.Category = "Sécurité"
        result.Timestamp = Now
        result.Status = "INFO"
        result.Message = "Non initialisé"
        result.Recommendation = "Le module d'authentification sera créé à la première connexion"
        mResults.Add result
    End If

    ' Vérifier protection du classeur
    result.TestName = "Protection Classeur"
    result.Category = "Sécurité"
    result.Timestamp = Now

    If ThisWorkbook.ProtectStructure Then
        result.Status = "PASS"
        result.Message = "Structure protégée"
        mHealth.PassCount = mHealth.PassCount + 1
    Else
        result.Status = "INFO"
        result.Message = "Structure non protégée"
        result.Recommendation = "Protéger le classeur en production"
    End If
    mResults.Add result
End Sub

' ==============================================================================
' 7. DIAGNOSTIC DES PERFORMANCES
' ==============================================================================

Private Sub DiagnostiquerPerformances()
    Dim result As DiagnosticResult

    ' Taille du fichier
    result.TestName = "Taille du fichier"
    result.Category = "Performances"
    result.Timestamp = Now

    Dim fileSize As Double
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ThisWorkbook.FullName) Then
        fileSize = fso.GetFile(ThisWorkbook.FullName).Size / 1024 / 1024
    End If
    On Error GoTo 0

    If fileSize < 50 Then
        result.Status = "PASS"
        result.Message = Format(fileSize, "0.0") & " MB"
        mHealth.PassCount = mHealth.PassCount + 1
    ElseIf fileSize < 100 Then
        result.Status = "WARN"
        result.Message = Format(fileSize, "0.0") & " MB - Performance dégradée possible"
        result.Recommendation = "Archiver les données historiques"
        mHealth.WarnCount = mHealth.WarnCount + 1
    Else
        result.Status = "WARN"
        result.Message = Format(fileSize, "0.0") & " MB - ATTENTION: Fichier volumineux"
        result.Recommendation = "Purger et archiver immédiatement"
        mHealth.WarnCount = mHealth.WarnCount + 1
    End If
    mResults.Add result

    ' Nombre de feuilles
    result.TestName = "Nombre de feuilles"
    result.Category = "Performances"
    result.Timestamp = Now

    Dim sheetCount As Integer
    sheetCount = ThisWorkbook.Sheets.Count

    If sheetCount < 20 Then
        result.Status = "PASS"
        result.Message = sheetCount & " feuilles"
        mHealth.PassCount = mHealth.PassCount + 1
    Else
        result.Status = "WARN"
        result.Message = sheetCount & " feuilles - Envisager nettoyage"
        result.Recommendation = "Supprimer les feuilles inutilisées"
        mHealth.WarnCount = mHealth.WarnCount + 1
    End If
    mResults.Add result

    ' Mémoire disponible
    result.TestName = "Version Excel"
    result.Category = "Performances"
    result.Timestamp = Now
    result.Status = "PASS"
    result.Message = "Excel " & Application.Version & " (" & IIf(Application.Version >= 16, "64-bit capable", "32-bit") & ")"
    mHealth.PassCount = mHealth.PassCount + 1
    mResults.Add result
End Sub

' ==============================================================================
' 8. GÉNÉRATION DU RAPPORT
' ==============================================================================

Private Sub GenererRapportDiagnostic(ws As Worksheet)
    Dim rowNum As Long
    Dim result As Variant

    rowNum = 4

    ' Résumé global
    ws.Cells(rowNum, 1).Value = "RÉSUMÉ"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Tests réussis:"
    ws.Cells(rowNum, 2).Value = mHealth.PassCount
    ws.Cells(rowNum, 2).Interior.Color = RGB(0, 200, 0)
    ws.Cells(rowNum, 2).Font.Color = vbWhite
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Avertissements:"
    ws.Cells(rowNum, 2).Value = mHealth.WarnCount
    If mHealth.WarnCount > 0 Then
        ws.Cells(rowNum, 2).Interior.Color = RGB(255, 255, 0)
    End If
    rowNum = rowNum + 1

    ws.Cells(rowNum, 1).Value = "Échecs:"
    ws.Cells(rowNum, 2).Value = mHealth.FailCount
    If mHealth.FailCount > 0 Then
        ws.Cells(rowNum, 2).Interior.Color = RGB(255, 0, 0)
        ws.Cells(rowNum, 2).Font.Color = vbWhite
    End If
    rowNum = rowNum + 2

    ' Détails par catégorie
    ws.Cells(rowNum, 1).Value = "DÉTAILS DES TESTS"
    ws.Cells(rowNum, 1).Font.Bold = True
    ws.Cells(rowNum, 1).Font.Size = 12
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Interior.Color = RGB(200, 200, 200)
    rowNum = rowNum + 1

    ' En-têtes
    ws.Cells(rowNum, 1).Value = "Test"
    ws.Cells(rowNum, 2).Value = "Catégorie"
    ws.Cells(rowNum, 3).Value = "Statut"
    ws.Cells(rowNum, 4).Value = "Message"
    ws.Cells(rowNum, 5).Value = "Recommandation"
    SAFA_Common.FormatHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 5))
    rowNum = rowNum + 1

    ' Écrire résultats
    For Each result In mResults
        Dim r As DiagnosticResult
        r = result

        ws.Cells(rowNum, 1).Value = r.TestName
        ws.Cells(rowNum, 2).Value = r.Category
        ws.Cells(rowNum, 3).Value = r.Status
        ws.Cells(rowNum, 4).Value = r.Message
        ws.Cells(rowNum, 5).Value = r.Recommendation

        ' Colorer selon statut
        Select Case r.Status
            Case "PASS"
                ws.Cells(rowNum, 3).Interior.Color = RGB(0, 200, 0)
                ws.Cells(rowNum, 3).Font.Color = vbWhite
            Case "WARN"
                ws.Cells(rowNum, 3).Interior.Color = RGB(255, 255, 0)
            Case "FAIL"
                ws.Cells(rowNum, 3).Interior.Color = RGB(255, 0, 0)
                ws.Cells(rowNum, 3).Font.Color = vbWhite
            Case "INFO"
                ws.Cells(rowNum, 3).Interior.Color = RGB(173, 216, 230)
        End Select

        rowNum = rowNum + 1
    Next result

    ' Problèmes critiques
    If mHealth.CriticalIssues <> "" Then
        rowNum = rowNum + 1
        ws.Cells(rowNum, 1).Value = "⚠️ PROBLÈMES CRITIQUES:"
        ws.Cells(rowNum, 1).Font.Bold = True
        ws.Cells(rowNum, 1).Font.Color = RGB(200, 0, 0)
        rowNum = rowNum + 1
        ws.Cells(rowNum, 1).Value = mHealth.CriticalIssues
        ws.Cells(rowNum, 1).Font.Color = RGB(200, 0, 0)
    End If

    ws.Columns("A:F").AutoFit
End Sub

' ==============================================================================
' 9. DIAGNOSTIC RAPIDE (SANS RAPPORT)
' ==============================================================================

Public Function DiagnosticRapide() As Boolean
    ' Vérifie rapidement les éléments essentiels
    ' Retourne True si tout est OK, False sinon
    Dim allOK As Boolean
    allOK = True

    ' Test Dictionary
    On Error Resume Next
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then allOK = False
    Err.Clear
    On Error GoTo 0

    ' Test FileSystem
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then allOK = False
    Err.Clear
    On Error GoTo 0

    DiagnosticRapide = allOK
End Function

' ==============================================================================
' 10. VÉRIFICATION AU DÉMARRAGE
' ==============================================================================

Public Sub VerifierAuDemarrage()
    ' À appeler depuis Workbook_Open
    ' Exécute un diagnostic rapide et affiche une alerte si problème

    If Not DiagnosticRapide() Then
        MsgBox "ATTENTION: Des problèmes de configuration ont été détectés." & vbCrLf & _
               "Veuillez exécuter le diagnostic complet via le menu S.A.F.A.", _
               vbExclamation, "S.A.F.A - Alerte Système"
    End If

    ' Initialiser le générateur aléatoire
    SAFA_Common.InitializeRandomizer
End Sub
