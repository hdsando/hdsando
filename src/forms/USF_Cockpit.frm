VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Cockpit
   Caption         =   "S.A.F.A v10.0 - Cockpit de Pilotage"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   OleObjectBlob   =   "USF_Cockpit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Cockpit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' USERFORM: USF_COCKPIT v10.0 (Enhanced)
' ==============================================================================
' Description: Interface principale de pilotage de l'audit
' Fonctionnalités:
'   - Import de données avec validation
'   - Configuration des paramètres
'   - Lancement de l'analyse
'   - Indicateurs visuels d'état
'   - Accès aux rapports
' ==============================================================================

Option Explicit

' --- CONSTANTES ---
Private Const COLOR_SUCCESS As Long = &HC000&     ' Vert
Private Const COLOR_WARNING As Long = &H80FF&     ' Orange
Private Const COLOR_ERROR As Long = &HFF&         ' Rouge
Private Const COLOR_NEUTRAL As Long = &H808080    ' Gris
Private Const COLOR_PRIMARY As Long = &H804000    ' Bleu foncé

' ==============================================================================
' INITIALISATION
' ==============================================================================

Private Sub UserForm_Initialize()
    ' Configuration initiale du formulaire
    On Error Resume Next

    ' Titre avec version
    Me.Caption = "S.A.F.A v" & Core_Engine.SAFA_VERSION & " - Cockpit de Pilotage"

    ' Charger les paramètres sauvegardés
    LoadSavedParameters

    ' Mettre à jour les indicateurs visuels
    UpdateAllStatuses

    ' Log de l'ouverture
    Call Core_Engine.WriteToAuditLog("UI", "Cockpit ouvert")

    On Error GoTo 0
End Sub

Private Sub LoadSavedParameters()
    ' Charger les paramètres depuis la feuille PARAM
    On Error Resume Next

    Dim wsParam As Worksheet
    Set wsParam = ThisWorkbook.Sheets("PARAM")

    If Not wsParam Is Nothing Then
        ' SOL ID
        Me.txtSolID.Value = wsParam.Range("F2").Value
        If Me.txtSolID.Value = "" Then Me.txtSolID.Value = "799"

        ' Tolérance
        Me.txtTolerance.Value = wsParam.Range("B2").Value
        If Me.txtTolerance.Value = "" Then Me.txtTolerance.Value = "100"

        ' Devise
        If Not Me.cboDevise Is Nothing Then
            Me.cboDevise.Clear
            Me.cboDevise.AddItem "XAF"
            Me.cboDevise.AddItem "EUR"
            Me.cboDevise.AddItem "USD"
            Me.cboDevise.AddItem "XOF"
            Me.cboDevise.Value = "XAF"
        End If
    End If

    On Error GoTo 0
End Sub

' ==============================================================================
' BOUTONS D'ACTION
' ==============================================================================

Private Sub btnReset_Click()
    ' Réinitialisation de l'outil
    Dim rep As Integer, repHist As Integer

    rep = MsgBox("Voulez-vous vraiment effacer toutes les données et rapports ?" & vbCrLf & vbCrLf & _
                 "Cette action est irréversible.", _
                 vbYesNo + vbExclamation + vbDefaultButton2, _
                 "Confirmation de nettoyage")

    If rep = vbNo Then Exit Sub

    repHist = MsgBox("Voulez-vous conserver l'historique pour l'analyse de vélocité ?" & vbCrLf & vbCrLf & _
                     "[OUI] = Audit mensuel récurrent (Recommandé)" & vbCrLf & _
                     "[NON] = Remise à zéro totale" & vbCrLf & _
                     "[ANNULER] = Retour", _
                     vbYesNoCancel + vbQuestion, _
                     "Mode de nettoyage")

    If repHist = vbCancel Then Exit Sub

    Me.Hide
    DoEvents

    ' Appel au moteur Core
    Call Core_Engine.Reinitialiser_Outil_Parametrable(keepHistory:=(repHist = vbYes))

    Me.Show
    UpdateAllStatuses

    MsgBox "Nettoyage effectué avec succès.", vbInformation, "S.A.F.A"
End Sub

Private Sub btnImportBal_Click()
    ' Import du fichier Balance
    Me.Hide
    DoEvents

    Call Core_Engine.Importer_Source_Balance

    Me.Show
    UpdateAllStatuses
End Sub

Private Sub btnImportGL_Click()
    ' Import du fichier GL Proof
    Me.Hide
    DoEvents

    Call Core_Engine.Importer_Source_GLProof

    Me.Show
    UpdateAllStatuses
End Sub

Private Sub btnLancer_Click()
    ' Lancement de l'analyse complète
    Dim wsBal As Worksheet, wsGL As Worksheet
    Dim nbLignesBal As Long, nbLignesGL As Long

    ' ═══════════════════════════════════════════════════════════════
    ' 1. VALIDATION DES DONNÉES
    ' ═══════════════════════════════════════════════════════════════

    On Error Resume Next
    Set wsBal = ThisWorkbook.Sheets("BALANCE_RAW")
    Set wsGL = ThisWorkbook.Sheets("GLPROOF_RAW")
    On Error GoTo 0

    ' Vérification Balance
    If wsBal Is Nothing Then
        nbLignesBal = 0
    Else
        nbLignesBal = Application.CountA(wsBal.Range("A:A"))
    End If

    ' Vérification GL
    If wsGL Is Nothing Then
        nbLignesGL = 0
    Else
        nbLignesGL = Application.CountA(wsGL.Range("A:A"))
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' 2. MESSAGES D'ERREUR
    ' ═══════════════════════════════════════════════════════════════

    If nbLignesBal < 2 Then
        MsgBox "ERREUR: La Balance n'est pas importée ou est vide." & vbCrLf & vbCrLf & _
               "Veuillez importer le fichier Balance (onglet BALANCE_RAW).", _
               vbExclamation, "Données manquantes"
        Exit Sub
    End If

    If nbLignesGL < 2 Then
        MsgBox "ERREUR: Le GL Proof n'est pas importé ou est vide." & vbCrLf & vbCrLf & _
               "Veuillez importer le fichier GL Proof (onglet GLPROOF_RAW).", _
               vbExclamation, "Données manquantes"
        Exit Sub
    End If

    If Trim(Me.txtSolID.Value) = "" Then
        MsgBox "ERREUR: Le Code Agence (SOL ID) est requis.", _
               vbExclamation, "Paramètre manquant"
        Me.txtSolID.SetFocus
        Exit Sub
    End If

    ' Validation numérique
    If Not IsNumeric(Me.txtTolerance.Value) Then
        Me.txtTolerance.Value = "100"
    End If

    ' ═══════════════════════════════════════════════════════════════
    ' 3. SAUVEGARDE DES PARAMÈTRES
    ' ═══════════════════════════════════════════════════════════════

    On Error Resume Next
    With ThisWorkbook.Sheets("PARAM")
        .Range("F1").Value = "SOL ID"
        .Range("F2").Value = Trim(Me.txtSolID.Value)
        .Range("B1").Value = "Tolérance"
        .Range("B2").Value = CDbl(Me.txtTolerance.Value)
    End With
    On Error GoTo 0

    ' ═══════════════════════════════════════════════════════════════
    ' 4. CONFIRMATION ET LANCEMENT
    ' ═══════════════════════════════════════════════════════════════

    Dim confirmMsg As String
    confirmMsg = "Prêt à lancer l'analyse avec les paramètres suivants:" & vbCrLf & vbCrLf & _
                 "• Balance: " & nbLignesBal & " lignes" & vbCrLf & _
                 "• GL Proof: " & nbLignesGL & " lignes" & vbCrLf & _
                 "• Code Agence: " & Me.txtSolID.Value & vbCrLf & _
                 "• Tolérance: " & Me.txtTolerance.Value & " XAF" & vbCrLf & vbCrLf & _
                 "Continuer ?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    ' Fermer le formulaire et lancer
    Me.Hide
    DoEvents

    Call Core_Engine.Lancer_Traitement_Complet

    Unload Me
End Sub

Private Sub btnQuitter_Click()
    ' Fermeture du formulaire
    Call Core_Engine.WriteToAuditLog("UI", "Cockpit fermé par l'utilisateur")
    Unload Me
End Sub

' ==============================================================================
' BOUTONS RAPPORTS (OPTIONNELS)
' ==============================================================================

Private Sub btnVoirDashboard_Click()
    ' Naviguer vers le Dashboard
    On Error Resume Next
    If Core_Engine.FeuilleExiste("DASHBOARD_RISQUE") Then
        Me.Hide
        ThisWorkbook.Sheets("DASHBOARD_RISQUE").Activate
        Unload Me
    Else
        MsgBox "Le Dashboard n'est pas encore généré." & vbCrLf & _
               "Lancez d'abord une analyse.", vbInformation, "S.A.F.A"
    End If
    On Error GoTo 0
End Sub

Private Sub btnVoirAlertes_Click()
    ' Naviguer vers les alertes
    On Error Resume Next
    If Core_Engine.FeuilleExiste("AUDIT_REPORT") Then
        Me.Hide
        ThisWorkbook.Sheets("AUDIT_REPORT").Activate
        Unload Me
    Else
        MsgBox "Aucune alerte générée." & vbCrLf & _
               "Lancez d'abord une analyse.", vbInformation, "S.A.F.A"
    End If
    On Error GoTo 0
End Sub

Private Sub btnVoirCompliance_Click()
    ' Naviguer vers la conformité
    On Error Resume Next
    If Core_Engine.FeuilleExiste("COMPLIANCE_CHECK") Then
        Me.Hide
        ThisWorkbook.Sheets("COMPLIANCE_CHECK").Activate
        Unload Me
    Else
        MsgBox "Aucun rapport de conformité généré." & vbCrLf & _
               "Lancez d'abord une analyse.", vbInformation, "S.A.F.A"
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' MISE À JOUR DES INDICATEURS
' ==============================================================================

Private Sub UpdateAllStatuses()
    ' Met à jour tous les indicateurs visuels
    Call UpdateBalanceStatus
    Call UpdateGLStatus
    Call UpdateAnalysisStatus
End Sub

Private Sub UpdateBalanceStatus()
    ' Indicateur Balance
    Dim lineCount As Long

    On Error Resume Next
    If Core_Engine.FeuilleExiste("BALANCE_RAW") Then
        lineCount = Application.CountA(ThisWorkbook.Sheets("BALANCE_RAW").Range("A:A"))
    Else
        lineCount = 0
    End If
    On Error GoTo 0

    If lineCount > 1 Then
        Me.lblStatusBal.Caption = "OK (" & Format(lineCount, "#,##0") & " lignes)"
        Me.lblStatusBal.ForeColor = COLOR_SUCCESS
        Me.lblStatusBal.Font.Bold = True
        Me.shpBalanceStatus.BackColor = COLOR_SUCCESS
    Else
        Me.lblStatusBal.Caption = "En attente d'import..."
        Me.lblStatusBal.ForeColor = COLOR_ERROR
        Me.lblStatusBal.Font.Bold = False
        Me.shpBalanceStatus.BackColor = COLOR_ERROR
    End If
End Sub

Private Sub UpdateGLStatus()
    ' Indicateur GL Proof
    Dim lineCount As Long

    On Error Resume Next
    If Core_Engine.FeuilleExiste("GLPROOF_RAW") Then
        lineCount = Application.CountA(ThisWorkbook.Sheets("GLPROOF_RAW").Range("A:A"))
    Else
        lineCount = 0
    End If
    On Error GoTo 0

    If lineCount > 1 Then
        Me.lblStatusGL.Caption = "OK (" & Format(lineCount, "#,##0") & " lignes)"
        Me.lblStatusGL.ForeColor = COLOR_SUCCESS
        Me.lblStatusGL.Font.Bold = True
        Me.shpGLStatus.BackColor = COLOR_SUCCESS
    Else
        Me.lblStatusGL.Caption = "En attente d'import..."
        Me.lblStatusGL.ForeColor = COLOR_ERROR
        Me.lblStatusGL.Font.Bold = False
        Me.shpGLStatus.BackColor = COLOR_ERROR
    End If
End Sub

Private Sub UpdateAnalysisStatus()
    ' Indicateur dernière analyse
    On Error Resume Next

    If Core_Engine.FeuilleExiste("RECONCIL") Then
        Dim wsR As Worksheet
        Set wsR = ThisWorkbook.Sheets("RECONCIL")

        If wsR.Cells(2, 1).Value <> "" Then
            Dim nbEcarts As Long, nbCritical As Long
            nbEcarts = Application.CountIf(wsR.Range("F:F"), "Ecart à analyser")
            nbCritical = Application.CountIf(wsR.Range("N:N"), "CRITICAL")

            Me.lblLastAnalysis.Caption = "Dernière analyse disponible"
            Me.lblAnalysisDetails.Caption = nbEcarts & " écarts | " & nbCritical & " critiques"

            If nbCritical > 0 Then
                Me.lblAnalysisDetails.ForeColor = COLOR_ERROR
            ElseIf nbEcarts > 0 Then
                Me.lblAnalysisDetails.ForeColor = COLOR_WARNING
            Else
                Me.lblAnalysisDetails.ForeColor = COLOR_SUCCESS
            End If
        Else
            Me.lblLastAnalysis.Caption = "Aucune analyse effectuée"
            Me.lblAnalysisDetails.Caption = ""
        End If
    Else
        Me.lblLastAnalysis.Caption = "Aucune analyse effectuée"
        Me.lblAnalysisDetails.Caption = ""
    End If

    On Error GoTo 0
End Sub

' ==============================================================================
' VALIDATION DES ENTRÉES
' ==============================================================================

Private Sub txtSolID_Change()
    ' Validation du SOL ID (3 chiffres)
    Dim val As String
    val = Me.txtSolID.Value

    If Len(val) > 3 Then
        Me.txtSolID.Value = Left(val, 3)
    End If
End Sub

Private Sub txtTolerance_Change()
    ' Validation de la tolérance (numérique)
    Dim val As String
    val = Me.txtTolerance.Value

    ' Supprimer les caractères non numériques
    Dim i As Long, result As String
    For i = 1 To Len(val)
        If Mid(val, i, 1) Like "[0-9]" Then
            result = result & Mid(val, i, 1)
        End If
    Next i

    If result <> val Then
        Me.txtTolerance.Value = result
    End If
End Sub

' ==============================================================================
' RACCOURCIS CLAVIER
' ==============================================================================

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Gestion des raccourcis clavier
    Select Case KeyCode
        Case vbKeyEscape
            ' Échap = Fermer
            btnQuitter_Click
        Case vbKeyReturn
            ' Entrée = Lancer (si données OK)
            If Me.btnLancer.Enabled Then
                btnLancer_Click
            End If
        Case vbKeyF5
            ' F5 = Rafraîchir les indicateurs
            UpdateAllStatuses
    End Select
End Sub

' ==============================================================================
' AIDE CONTEXTUELLE
' ==============================================================================

Private Sub lblHelp_Click()
    ' Afficher l'aide
    Dim helpMsg As String

    helpMsg = "S.A.F.A - System for Automated Financial Audit" & vbCrLf & _
              "Version " & Core_Engine.SAFA_VERSION & vbCrLf & vbCrLf & _
              "ÉTAPES D'UTILISATION:" & vbCrLf & _
              "1. Importer le fichier Balance Finacle" & vbCrLf & _
              "2. Importer le fichier GL Proof" & vbCrLf & _
              "3. Renseigner le Code Agence (SOL ID)" & vbCrLf & _
              "4. Cliquer sur 'Lancer l'Analyse'" & vbCrLf & vbCrLf & _
              "RACCOURCIS:" & vbCrLf & _
              "• Entrée: Lancer l'analyse" & vbCrLf & _
              "• Échap: Fermer" & vbCrLf & _
              "• F5: Rafraîchir les indicateurs" & vbCrLf & vbCrLf & _
              "FONCTIONNALITÉS:" & vbCrLf & _
              "• Rapprochement Balance/GL automatique" & vbCrLf & _
              "• Détection de fraudes (Benford, patterns)" & vbCrLf & _
              "• Analyse IA (Z-Score, Vélocité)" & vbCrLf & _
              "• Conformité réglementaire (COBAC, OHADA)" & vbCrLf & _
              "• Scoring de risque automatique"

    MsgBox helpMsg, vbInformation, "Aide S.A.F.A"
End Sub
