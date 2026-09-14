Attribute VB_Name = "SAFA_Menu"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SAFA_MENU v10.0
' ==============================================================================
' Description: Cockpit a boutons sur une feuille Excel ("MENU").
'   - Aucun UserForm (.frm/.frx) requis: fonctionne sur toutes les versions
'   - Les boutons sont des formes (Shapes) reliees aux procedures publiques
'     de SAFA_Console via OnAction
'   - Un panneau de statut (cellules) est rafraichi par RefreshStatus
' ==============================================================================

Private Const MENU_SHEET As String = "MENU"

' Cellules du panneau de statut (colonne I = valeurs)
Private Const CELL_STATUS_BAL As String = "I5"
Private Const CELL_STATUS_GL As String = "I6"
Private Const CELL_STATUS_SOL As String = "I9"
Private Const CELL_STATUS_TOL As String = "I10"
Private Const CELL_STATUS_LAST As String = "I13"
Private Const CELL_STATUS_ECARTS As String = "I14"
Private Const CELL_STATUS_CRIT As String = "I15"

' Couleurs (Long BGR)
Private Const CLR_NAVY As Long = 6697728      ' RGB(0, 51, 102)
Private Const CLR_BLUE As Long = 10040115     ' RGB(51, 102, 153)
Private Const CLR_GREEN As Long = 3329330     ' RGB(50, 205, 50) -> ajuste ci-dessous
Private Const CLR_ORANGE As Long = 33023      ' RGB(255, 128, 0)
Private Const CLR_GREY As Long = 8421504      ' RGB(128, 128, 128)
Private Const CLR_TEAL As Long = 8421376      ' RGB(0, 128, 128)
Private Const CLR_RED As Long = 3355647       ' RGB(255, 51, 51)

' ==============================================================================
' CONSTRUCTION DU MENU
' ==============================================================================

Public Sub BuildMenu()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim shp As Shape

    Application.ScreenUpdating = False

    ' Recuperer ou creer la feuille, puis la vider
    Set ws = SAFA_Common.GetOrCreateSheet(MENU_SHEET, False)
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.Clear
    ws.Move Before:=ThisWorkbook.Sheets(1)

    ' Mise en page
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:K").ColumnWidth = 13
    ws.Rows("1:32").RowHeight = 18
    ws.Cells.Interior.Color = RGB(247, 248, 250)
    ws.Cells.Font.Name = "Calibri"

    ' Bandeau titre
    With ws.Range("B1:K2")
        .Merge
        .Value = "S.A.F.A v" & SAFA_Common.SAFA_VERSION & "   -   System for Automated Financial Audit"
        .Font.Size = 18: .Font.Bold = True: .Font.Color = vbWhite
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows("1:2").RowHeight = 22
    With ws.Range("B3")
        .Value = "Rapprochement Balance / GL Proof - Forensique - IA - Conformite COBAC / OHADA / LAB-FT"
        .Font.Italic = True: .Font.Color = RGB(90, 90, 90)
    End With

    ' Sections (libelles)
    Call SectionLabel(ws, "B4", "1. DONNEES")
    Call SectionLabel(ws, "H4", "STATUT")
    Call SectionLabel(ws, "B8", "2. PARAMETRES")
    Call SectionLabel(ws, "B12", "3. ANALYSE")
    Call SectionLabel(ws, "B17", "4. RAPPORTS")
    Call SectionLabel(ws, "B22", "5. ANALYSES AVANCEES")
    Call SectionLabel(ws, "B28", "6. SYSTEME")

    ' Panneau de statut (libelles)
    ws.Range("H5").Value = "Balance:"
    ws.Range("H6").Value = "GL Proof:"
    ws.Range("H9").Value = "SOL ID:"
    ws.Range("H10").Value = "Tolerance:"
    ws.Range("H13").Value = "Derniere analyse:"
    ws.Range("H14").Value = "Ecarts a analyser:"
    ws.Range("H15").Value = "Comptes CRITICAL:"
    ws.Range("H5:H15").Font.Color = RGB(90, 90, 90)
    ws.Range("I5:K15").Font.Bold = True
    ws.Range("H4:K15").Borders(xlEdgeLeft).LineStyle = xlContinuous
    ws.Range("H4:K15").Borders(xlEdgeLeft).Color = RGB(200, 200, 200)

    ' --- 1. DONNEES ---
    Call AddButton(ws, "btnImportBal", "Importer Balance", "B5", 2, CLR_BLUE, "SAFA_Console.ImporterBalance")
    Call AddButton(ws, "btnImportGL", "Importer GL Proof", "D5", 2, CLR_BLUE, "SAFA_Console.ImporterGLProof")
    Call AddButton(ws, "btnDemo", "Donnees demo", "F5", 2, CLR_TEAL, "SAFA_Console.ChargerDonneesDemo")

    ' --- 2. PARAMETRES ---
    Call AddButton(ws, "btnParams", "Parametres", "B9", 2, CLR_GREY, "SAFA_Console.DefinirParametres")
    Call AddButton(ws, "btnConfig", "Configuration", "D9", 2, CLR_GREY, "SAFA_Console.AfficherConfiguration")

    ' --- 3. ANALYSE ---
    Call AddButton(ws, "btnRun", "LANCER L'ANALYSE COMPLETE", "B13", 4, CLR_ORANGE, "SAFA_Console.LancerAnalyse", 34, 13)

    ' --- 4. RAPPORTS ---
    Call AddButton(ws, "btnDash", "Dashboard", "B18", 2, CLR_NAVY, "SAFA_Console.OuvrirDashboard")
    Call AddButton(ws, "btnAlerts", "Alertes", "D18", 2, CLR_NAVY, "SAFA_Console.OuvrirAlertes")
    Call AddButton(ws, "btnComp", "Conformite", "F18", 2, CLR_NAVY, "SAFA_Console.OuvrirConformite")
    Call AddButton(ws, "btnExec", "Synthese", "H18", 2, CLR_NAVY, "SAFA_Console.OuvrirSynthese")
    Call AddButton(ws, "btnAll", "Generer tous", "B20", 2, CLR_BLUE, "SAFA_Console.GenererTousRapports")
    Call AddButton(ws, "btnPDF", "Export PDF", "D20", 2, CLR_BLUE, "SAFA_Console.ExporterPDF")
    Call AddButton(ws, "btnSample", "Echantillon Top 20", "F20", 2, CLR_BLUE, "SAFA_Console.OuvrirEchantillon")

    ' --- 5. AVANCE ---
    Call AddButton(ws, "btnTemp", "Temporel", "B23", 2, CLR_TEAL, "SAFA_Console.LancerTemporel")
    Call AddButton(ws, "btnNet", "Reseau", "D23", 2, CLR_TEAL, "SAFA_Console.LancerReseau")
    Call AddButton(ws, "btnDiag", "Diagnostic", "F23", 2, CLR_TEAL, "SAFA_Console.LancerDiagnostic")
    Call AddButton(ws, "btnTests", "Tests auto", "H23", 2, CLR_TEAL, "SAFA_Console.LancerTests")
    Call AddButton(ws, "btnGLM", "GL Monitoring", "B25", 2, CLR_ORANGE, "SAFA_Console.LancerGLMonitoring")
    Call AddButton(ws, "btnProofs", "Qualite proofs", "D25", 2, CLR_ORANGE, "SAFA_Console.LancerQualiteProofs")
    Call AddButton(ws, "btnPrev", "Balance N-1", "F25", 2, CLR_BLUE, "SAFA_Console.ImporterBalancePrecedente")
    Call AddButton(ws, "btnCI", "Liste proofables", "H25", 2, CLR_BLUE, "SAFA_Console.ImporterListeProofables")

    ' --- 6. SYSTEME ---
    Call AddButton(ws, "btnReset", "Reinitialiser", "B29", 2, CLR_RED, "SAFA_Console.Reinitialiser")
    Call AddButton(ws, "btnTrail", "Journal d'audit", "D29", 2, CLR_GREY, "SAFA_Console.OuvrirJournalAudit")
    Call AddButton(ws, "btnHelp", "Aide", "F29", 2, CLR_GREY, "SAFA_Console.AfficherAide")
    Call AddButton(ws, "btnConsole", "Mode console", "H29", 2, CLR_GREY, "SAFA_Console.Demarrer")
    Call AddButton(ws, "btnShare", "Preparer partage", "J29", 2, CLR_NAVY, "SAFA_Console.PreparerPartage")
    Call AddButton(ws, "btnCalib", "Calibration", "J25", 2, CLR_BLUE, "SAFA_Console.AfficherCalibration")

    ws.Range("B32").Value = "Astuce: le bouton Tests auto charge des donnees demo, execute tout le pipeline et verifie chaque detecteur. GL Monitoring: regles de la seance DAI du 09/09/2026."
    ws.Range("B32").Font.Size = 9: ws.Range("B32").Font.Color = RGB(120, 120, 120)

    ws.Range("A1").Select
    Call RefreshStatus

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur BuildMenu: " & Err.Description, vbCritical, "S.A.F.A"
End Sub

' ==============================================================================
' PANNEAU DE STATUT
' ==============================================================================

Public Sub RefreshStatus()
    On Error Resume Next

    Dim ws As Worksheet
    Dim n As Long, nEcarts As Long, nCrit As Long

    Set ws = ThisWorkbook.Sheets(MENU_SHEET)
    If ws Is Nothing Then Exit Sub

    ' Balance
    n = 0
    If SAFA_Common.FeuilleExiste("BALANCE_RAW") Then n = Application.CountA(ThisWorkbook.Sheets("BALANCE_RAW").Range("A:A"))
    Call SetStatus(ws.Range(CELL_STATUS_BAL), IIf(n > 1, "OK - " & Format(n, "#,##0") & " lignes", "En attente d'import"), n > 1)

    ' GL Proof
    n = 0
    If SAFA_Common.FeuilleExiste("GLPROOF_RAW") Then n = Application.CountA(ThisWorkbook.Sheets("GLPROOF_RAW").Range("A:A"))
    Call SetStatus(ws.Range(CELL_STATUS_GL), IIf(n > 1, "OK - " & Format(n, "#,##0") & " lignes", "En attente d'import"), n > 1)

    ' Parametres
    If SAFA_Common.FeuilleExiste("PARAM") Then
        ws.Range(CELL_STATUS_SOL).Value = "'" & SAFA_Common.SafeText(ThisWorkbook.Sheets("PARAM").Range("F2").Value)
        ws.Range(CELL_STATUS_TOL).Value = Format(SAFA_Common.SafeVal(ThisWorkbook.Sheets("PARAM").Range("B2").Value), "#,##0") & " XAF"
    End If

    ' Analyse
    If SAFA_Common.FeuilleExiste("RECONCIL") Then
        With ThisWorkbook.Sheets("RECONCIL")
            n = Application.CountA(.Columns(SAFA_Common.RECONCIL_COL_COMPTE)) - 1
            nEcarts = Application.CountIf(.Columns(SAFA_Common.RECONCIL_COL_STATUT), SAFA_Common.RECONCIL_STATUT_ECART)
            nCrit = Application.CountIf(.Columns(SAFA_Common.RECONCIL_COL_PRIORITY), "CRITICAL")
        End With
        If n > 0 Then
            Call SetStatus(ws.Range(CELL_STATUS_LAST), Format(Now, "dd/mm/yyyy hh:nn") & " - " & Format(n, "#,##0") & " comptes", True)
            Call SetStatus(ws.Range(CELL_STATUS_ECARTS), Format(nEcarts, "#,##0"), nEcarts = 0)
            Call SetStatus(ws.Range(CELL_STATUS_CRIT), Format(nCrit, "#,##0"), nCrit = 0)
        Else
            Call SetStatus(ws.Range(CELL_STATUS_LAST), "Aucune analyse", False)
            ws.Range(CELL_STATUS_ECARTS).Value = "-": ws.Range(CELL_STATUS_CRIT).Value = "-"
        End If
    Else
        Call SetStatus(ws.Range(CELL_STATUS_LAST), "Aucune analyse", False)
        ws.Range(CELL_STATUS_ECARTS).Value = "-": ws.Range(CELL_STATUS_CRIT).Value = "-"
        ws.Range(CELL_STATUS_ECARTS).Interior.ColorIndex = xlNone
        ws.Range(CELL_STATUS_CRIT).Interior.ColorIndex = xlNone
    End If
End Sub

Public Sub OuvrirMenu()
    On Error Resume Next
    If Not SAFA_Common.FeuilleExiste(MENU_SHEET) Then Call BuildMenu
    ThisWorkbook.Sheets(MENU_SHEET).Visible = xlSheetVisible
    ThisWorkbook.Sheets(MENU_SHEET).Activate
    Call RefreshStatus
End Sub

' ==============================================================================
' HELPERS
' ==============================================================================

Private Sub SectionLabel(ws As Worksheet, addr As String, caption As String)
    With ws.Range(addr)
        .Value = caption
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(0, 51, 102)
    End With
End Sub

Private Sub SetStatus(cell As Range, txt As String, ok As Boolean)
    cell.Value = txt
    If ok Then
        cell.Interior.Color = RGB(214, 240, 220)
        cell.Font.Color = RGB(20, 100, 50)
    Else
        cell.Interior.Color = RGB(252, 226, 226)
        cell.Font.Color = RGB(160, 30, 30)
    End If
End Sub

' Cree un bouton (forme arrondie) ancre sur une cellule, large de nbCols colonnes
Private Sub AddButton(ws As Worksheet, shapeName As String, caption As String, anchor As String, _
                      nbCols As Long, fillColor As Long, macro As String, _
                      Optional heightPt As Single = 26, Optional fontSize As Single = 10)
    Dim r As Range, shp As Shape
    Dim w As Single, i As Long

    Set r = ws.Range(anchor)
    w = 0
    For i = 0 To nbCols - 1
        w = w + r.Offset(0, i).Width
    Next i
    w = w - 6

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, r.Left + 2, r.Top + 2, w, heightPt)
    With shp
        .Name = shapeName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.18
        .Shadow.Visible = msoFalse
        .OnAction = macro
        .Placement = xlFreeFloating
        With .TextFrame2
            .TextRange.Text = caption
            .TextRange.Font.Size = fontSize
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2
            .WordWrap = msoTrue
        End With
    End With
End Sub
