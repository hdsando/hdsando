VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: THISWORKBOOK v10.0
' ==============================================================================
' Description: Gestion des événements du classeur
' Fonctionnalités:
'   - Mode Application (feuilles cachées au démarrage)
'   - Lancement différé du Cockpit
'   - Protection de l'intégrité du classeur
'   - Initialisation de l'environnement
' ==============================================================================

Option Explicit

' --- CONSTANTES ---
Private Const STARTUP_DELAY As String = "00:00:01"  ' Délai avant affichage Cockpit

' ==============================================================================
' ÉVÉNEMENT: OUVERTURE DU CLASSEUR
' ==============================================================================

Private Sub Workbook_Open()
    Dim ws As Worksheet

    On Error Resume Next

    Application.ScreenUpdating = False

    ' ═══════════════════════════════════════════════════════════════
    ' 1. INITIALISATION DE L'ENVIRONNEMENT
    ' ═══════════════════════════════════════════════════════════════

    ' S'assurer que les feuilles essentielles existent
    Call EnsureRequiredSheets

    ' Charger la configuration (settings.json ou CONFIG_DATA, sinon defauts)
    Call Config_Manager.LoadConfiguration

    ' ═══════════════════════════════════════════════════════════════
    ' 2. FEUILLE MENU (cockpit a boutons, sans UserForm)
    ' ═══════════════════════════════════════════════════════════════

    If Not SheetExists("MENU") Then Call SAFA_Menu.BuildMenu
    ThisWorkbook.Sheets("MENU").Visible = xlSheetVisible
    ThisWorkbook.Sheets("MENU").Activate
    Call SAFA_Menu.RefreshStatus

    ' Masquer uniquement les feuilles techniques (les resultats restent visibles)
    Call Masquer_Feuilles_Techniques_Silencieux

    Application.ScreenUpdating = True

    Call Core_Engine.WriteToAuditLog("SESSION", "S.A.F.A ouvert")

    On Error GoTo 0
End Sub

' ==============================================================================
' ÉVÉNEMENT: FERMETURE DU CLASSEUR
' ==============================================================================

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyer avant fermeture

    On Error Resume Next

    ' Log de fermeture
    Call Core_Engine.WriteToAuditLog("SESSION", "S.A.F.A fermé")

    ' Réinitialiser la StatusBar
    Application.StatusBar = False

    On Error GoTo 0
End Sub

' ==============================================================================
' ÉVÉNEMENT: SAUVEGARDE DU CLASSEUR
' ==============================================================================

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Actions avant sauvegarde

    On Error Resume Next

    ' Log de sauvegarde
    Call Core_Engine.WriteToAuditLog("SESSION", "Classeur sauvegardé")

    ' S'assurer que les feuilles de résultats sont visibles pour la sauvegarde
    ' (optionnel - dépend des besoins)

    On Error GoTo 0
End Sub

' ==============================================================================
' ÉVÉNEMENT: ACTIVATION D'UNE FEUILLE
' ==============================================================================

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Gérer l'activation des feuilles

    On Error Resume Next

    ' Zoom par défaut
    ActiveWindow.Zoom = 85

    ' Position par défaut
    Sh.Range("A1").Select

    On Error GoTo 0
End Sub

' ==============================================================================
' FONCTIONS UTILITAIRES
' ==============================================================================

Private Sub EnsureRequiredSheets()
    ' S'assurer que les feuilles essentielles existent
    Dim requiredSheets As Variant
    Dim sheetName As Variant

    requiredSheets = Array("PARAM", "BALANCE_RAW", "GLPROOF_RAW")

    For Each sheetName In requiredSheets
        If Not SheetExists(CStr(sheetName)) Then
            Dim newSheet As Worksheet
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newSheet.Name = CStr(sheetName)

            ' Configuration initiale selon le type de feuille
            Select Case CStr(sheetName)
                Case "PARAM"
                    Call SetupParamSheet(newSheet)
            End Select
        End If
    Next sheetName
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub SetupParamSheet(ws As Worksheet)
    ' Configuration de la feuille PARAM
    With ws
        .Cells.Clear

        ' En-têtes
        .Range("A1").Value = "PARAMÈTRES S.A.F.A"
        .Range("A1").Font.Bold = True

        .Range("A2").Value = "Tolérance (XAF)"
        .Range("B2").Value = 100

        .Range("C2").Value = "Devise"
        .Range("D2").Value = "XAF"

        .Range("E2").Value = "SOL ID"
        .Range("F2").Value = "799"

        ' Exclusions
        .Range("A4").Value = "Patterns d'exclusion (Non-Proofable)"
        .Range("D4").Value = "ISO"
        .Range("D5").Value = "NOSTRO"
        .Range("D6").Value = "VOSTRO"
    End With
End Sub

' ==============================================================================
' PROCÉDURES PUBLIQUES (ACCESSIBLES VIA MACROS)
' ==============================================================================

Public Sub Lancer_Demarrage_Differe()
    ' Point d'entree historique: sequence de demarrage puis menu
    Call Core_Engine.Sequence_Demarrage
    Call SAFA_Menu.OuvrirMenu
End Sub

Public Sub Afficher_Cockpit()
    ' Afficher le cockpit (feuille MENU a boutons; le mode console reste disponible via SAFA_Console.Demarrer)
    Call SAFA_Menu.OuvrirMenu
End Sub

Private Sub Masquer_Feuilles_Techniques_Silencieux()
    Dim ws As Worksheet
    Dim technicalSheets As Variant
    technicalSheets = Array("AUDIT_TRAIL", "CONFIG_DATA", "USERS", "HISTORY_LOG")
    For Each ws In ThisWorkbook.Worksheets
        If IsInArray(ws.Name, technicalSheets) Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub

Public Sub Afficher_Toutes_Feuilles()
    ' Utilitaire pour debug - Afficher toutes les feuilles
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

    MsgBox "Toutes les feuilles sont maintenant visibles.", vbInformation
End Sub

Public Sub Masquer_Feuilles_Techniques()
    ' Remettre en mode Application
    Dim ws As Worksheet
    Dim technicalSheets As Variant

    technicalSheets = Array("AUDIT_TRAIL", "CONFIG", "HISTORY_LOG")

    For Each ws In ThisWorkbook.Worksheets
        If IsInArray(ws.Name, technicalSheets) Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    MsgBox "Feuilles techniques masquées.", vbInformation
End Sub

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i)) = UCase(val) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
