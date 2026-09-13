Attribute VB_Name = "Proof_Quality"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: PROOF_QUALITY v10.1
' ==============================================================================
' Description: Controles de FORME des justificatifs de comptes (proofs) Excel,
'   d'apres la Knowledge Sharing Session DAI du 09/09/2026:
'     - totaux en formules, jamais saisis a la main
'     - aucune ligne / colonne / feuille masquee
'     - pas de transactions qui s'annulent (elles ne doivent plus figurer)
'     - un copier-coller du detail du compte n'est pas un proof
'     - le total du proof doit cadrer avec le solde GL a la date d'arrete
'   Usage: bouton "Qualite proofs" -> choisir le dossier contenant les fichiers
'   Resultat: feuille PROOF_QUALITY (une ligne par feuille analysee)
' ==============================================================================

Private Const MAX_ROWS_SCAN As Long = 5000
Private Const MAX_COLS_SCAN As Long = 60

Private Type SheetCheck
    HiddenRows As Long
    HiddenCols As Long
    Formulas As Long
    NumericCells As Long
    HardcodedTotals As Long
    ProofTotal As Double
    CancelPairs As Long
    LooksLikeStatement As Boolean
    Account As String
    Observations As String
End Type

Public Sub AnalyserDossierProofs(Optional dossier As String = "")
    On Error GoTo ErrHandler

    Dim fso As Object, fld As Object, f As Object
    Dim wsOut As Worksheet, r As Long, nFiles As Long, nNonConf As Long
    Dim wb As Workbook, ws As Worksheet, chk As SheetCheck
    Dim dictGL As Object, verdict As String, glBal As Variant, ecart As Variant
    Dim prevUpd As Boolean, prevEv As Boolean, prevAl As Boolean

    If dossier = "" Then dossier = ChoisirDossier("Selectionnez le dossier contenant les proofs (fichiers Excel)")
    If dossier = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then
        MsgBox "Dossier introuvable: " & dossier, vbExclamation, "Qualite des proofs"
        Exit Sub
    End If

    Set dictGL = ChargerSoldesGL()

    Set wsOut = SAFA_Common.GetOrCreateSheet("PROOF_QUALITY", True)
    wsOut.Range("A1:N1").Value = Array("Fichier", "Feuille", "Compte detecte", "Total proof", "Solde GL", "Ecart proof/GL", _
                                       "Lignes masquees", "Colonnes masquees", "Totaux saisis (non formule)", _
                                       "Paires annulantes", "Formules", "Copie du releve ?", "Verdict", "Observations")
    SAFA_Common.FormatHeader wsOut.Range("A1:N1")
    r = 2

    prevUpd = Application.ScreenUpdating: prevEv = Application.EnableEvents: prevAl = Application.DisplayAlerts
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.DisplayAlerts = False

    Set fld = fso.GetFolder(dossier)
    For Each f In fld.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb") And Left(f.Name, 2) <> "~$" Then
            nFiles = nFiles + 1
            Application.StatusBar = "Proof " & nFiles & ": " & f.Name
            Set wb = Nothing
            On Error Resume Next
            Set wb = Workbooks.Open(f.Path, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
            On Error GoTo ErrHandler
            If wb Is Nothing Then
                wsOut.Cells(r, 1).Value = f.Name
                wsOut.Cells(r, 13).Value = "NON LISIBLE"
                wsOut.Cells(r, 13).Interior.Color = RGB(255, 170, 170)
                r = r + 1
            Else
                For Each ws In wb.Worksheets
                    chk = AnalyserFeuille(ws)
                    glBal = "": ecart = ""
                    If chk.Account <> "" Then
                        If dictGL.Exists(chk.Account) Then
                            glBal = dictGL(chk.Account)
                            If chk.ProofTotal <> 0 Then ecart = Round(CDbl(glBal) - chk.ProofTotal, 0)
                        End If
                    End If

                    verdict = "CONFORME"
                    If ws.Visible <> xlSheetVisible Then verdict = "NON CONFORME": chk.Observations = chk.Observations & "Feuille masquee; "
                    If chk.HiddenRows > 0 Or chk.HiddenCols > 0 Then verdict = "NON CONFORME"
                    If chk.HardcodedTotals > 0 Then verdict = "NON CONFORME"
                    If chk.CancelPairs > 0 Then verdict = "NON CONFORME"
                    If chk.LooksLikeStatement Then verdict = "NON CONFORME"
                    If IsNumeric(ecart) Then
                        If ecart <> "" Then
                            If Abs(CDbl(ecart)) > 100 Then verdict = "NON CONFORME": chk.Observations = chk.Observations & "Total proof different du solde GL; "
                        End If
                    End If
                    If chk.NumericCells = 0 And chk.Formulas = 0 Then verdict = "VIDE / NON EVALUABLE"

                    With wsOut
                        .Cells(r, 1).Value = f.Name
                        .Cells(r, 2).Value = ws.Name
                        .Cells(r, 3).Value = IIf(chk.Account = "", "", "'" & chk.Account)
                        .Cells(r, 4).Value = IIf(chk.ProofTotal = 0, "", chk.ProofTotal)
                        .Cells(r, 5).Value = glBal
                        .Cells(r, 6).Value = ecart
                        .Cells(r, 7).Value = chk.HiddenRows
                        .Cells(r, 8).Value = chk.HiddenCols
                        .Cells(r, 9).Value = chk.HardcodedTotals
                        .Cells(r, 10).Value = chk.CancelPairs
                        .Cells(r, 11).Value = chk.Formulas
                        .Cells(r, 12).Value = IIf(chk.LooksLikeStatement, "OUI", "NON")
                        .Cells(r, 13).Value = verdict
                        .Cells(r, 14).Value = chk.Observations
                        Select Case verdict
                            Case "CONFORME": .Cells(r, 13).Interior.Color = RGB(200, 255, 200)
                            Case "NON CONFORME": .Cells(r, 13).Interior.Color = RGB(255, 170, 170): nNonConf = nNonConf + 1
                            Case Else: .Cells(r, 13).Interior.Color = RGB(230, 230, 230)
                        End Select
                    End With
                    r = r + 1
                Next ws
                wb.Close SaveChanges:=False
            End If
        End If
    Next f

    Application.StatusBar = False
    Application.ScreenUpdating = prevUpd: Application.EnableEvents = prevEv: Application.DisplayAlerts = prevAl

    wsOut.Columns("A:N").AutoFit
    wsOut.Columns("N").ColumnWidth = 60
    wsOut.Columns("N").WrapText = True
    wsOut.Range("D2:F" & r).NumberFormat = "#,##0"
    If r > 2 Then wsOut.Range("A1:N" & (r - 1)).AutoFilter

    Call SAFA_Common.WriteAuditLog("PROCESS", "Qualite des proofs", nFiles & " fichiers, " & nNonConf & " feuilles non conformes, dossier " & dossier)

    MsgBox "Analyse terminee: " & nFiles & " fichier(s), " & (r - 2) & " feuille(s), " & nNonConf & " non conforme(s)." & vbCrLf & vbCrLf & _
           "Voir la feuille PROOF_QUALITY.", vbInformation, "Qualite des proofs"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Call SAFA_Common.LogError("Proof_Quality", "AnalyserDossierProofs", Err.Number, Err.Description)
    MsgBox "Erreur: " & Err.Description, vbCritical, "Qualite des proofs"
End Sub

' ==============================================================================
' ANALYSE D'UNE FEUILLE
' ==============================================================================

Private Function AnalyserFeuille(ws As Worksheet) As SheetCheck
    Dim chk As SheetCheck
    Dim ur As Range, nRows As Long, nCols As Long, i As Long, j As Long
    Dim c As Range, v As Variant, txt As String
    Dim dictVals As Object, colCounts() As Long, bestCol As Long, bestCnt As Long
    Dim headerHits As Long, re As Object, m As Object, seenPair As Object

    On Error GoTo ErrHandler
    Set ur = ws.UsedRange
    If ur Is Nothing Then AnalyserFeuille = chk: Exit Function
    nRows = Application.WorksheetFunction.Min(ur.Rows.Count, MAX_ROWS_SCAN)
    nCols = Application.WorksheetFunction.Min(ur.Columns.Count, MAX_COLS_SCAN)
    If nRows < 1 Or nCols < 1 Then AnalyserFeuille = chk: Exit Function

    ' Lignes et colonnes masquees (filtre automatique inclus)
    For i = 1 To nRows
        If ur.Rows(i).EntireRow.Hidden Then chk.HiddenRows = chk.HiddenRows + 1
    Next i
    For j = 1 To nCols
        If ur.Columns(j).EntireColumn.Hidden Then chk.HiddenCols = chk.HiddenCols + 1
    Next j
    If chk.HiddenRows > 0 Then chk.Observations = chk.Observations & chk.HiddenRows & " ligne(s) masquee(s)" & IIf(ws.AutoFilterMode, " (filtre actif)", "") & "; "
    If chk.HiddenCols > 0 Then chk.Observations = chk.Observations & chk.HiddenCols & " colonne(s) masquee(s); "

    ' Formules
    On Error Resume Next
    chk.Formulas = ur.SpecialCells(xlCellTypeFormulas).Count
    If Err.Number <> 0 Then chk.Formulas = 0: Err.Clear
    On Error GoTo ErrHandler

    ' Colonne de montants (celle qui contient le plus de nombres), numero de compte, en-tetes de releve
    ReDim colCounts(1 To nCols)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\b[0-9]{10,16}\b"
    re.Global = False
    For i = 1 To nRows
        For j = 1 To nCols
            Set c = ur.Cells(i, j)
            v = c.Value
            If IsNumeric(v) And Not IsEmpty(v) And Not IsDate(v) Then
                If VarType(v) <> vbString Then
                    chk.NumericCells = chk.NumericCells + 1
                    colCounts(j) = colCounts(j) + 1
                End If
            ElseIf VarType(v) = vbString Then
                If i <= 40 Then
                    txt = UCase(CStr(v))
                    If chk.Account = "" Then
                        If re.Test(txt) Then chk.Account = SAFA_Common.NormalizeAccountKey(re.Execute(txt)(0).Value)
                    End If
                    If InStr(txt, "TRAN DATE") > 0 Or InStr(txt, "VALUE DATE") > 0 Or InStr(txt, "NARRATION") > 0 Or _
                       InStr(txt, "PARTICULARS") > 0 Or InStr(txt, "TRAN ID") > 0 Or InStr(txt, "POSTED BY") > 0 Then headerHits = headerHits + 1
                End If
            End If
        Next j
    Next i
    For j = 1 To nCols
        If colCounts(j) > bestCnt Then bestCnt = colCounts(j): bestCol = j
    Next j

    ' Totaux saisis a la main: cellule "TOTAL" puis premier nombre a droite sans formule
    Dim found As Range, firstAddr As String, k As Long
    Set found = ur.Find(What:="TOTAL", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not found Is Nothing Then
        firstAddr = found.Address
        Do
            For k = 1 To 8
                If found.Column + k <= ws.Columns.Count Then
                    Set c = ws.Cells(found.Row, found.Column + k)
                    If IsNumeric(c.Value) And Not IsEmpty(c.Value) And VarType(c.Value) <> vbString Then
                        If Not c.HasFormula Then
                            chk.HardcodedTotals = chk.HardcodedTotals + 1
                            chk.Observations = chk.Observations & "Total saisi en " & c.Address(False, False) & "; "
                        End If
                        chk.ProofTotal = CDbl(c.Value)
                        Exit For
                    End If
                End If
            Next k
            Set found = ur.FindNext(found)
        Loop While Not found Is Nothing And found.Address <> firstAddr
    End If

    ' Paires qui s'annulent dans la colonne de montants (v et -v)
    If bestCol > 0 And bestCnt >= 2 Then
        Set dictVals = CreateObject("Scripting.Dictionary")
        Set seenPair = CreateObject("Scripting.Dictionary")
        For i = 1 To nRows
            v = ur.Cells(i, bestCol).Value
            If IsNumeric(v) And Not IsEmpty(v) And VarType(v) <> vbString Then
                If CDbl(v) <> 0 And Not ur.Cells(i, bestCol).HasFormula Then
                    txt = Format(Abs(CDbl(v)), "0.00")
                    If dictVals.Exists(txt) Then
                        If Sgn(dictVals(txt)) <> Sgn(CDbl(v)) And Not seenPair.Exists(txt) Then
                            chk.CancelPairs = chk.CancelPairs + 1
                            seenPair(txt) = 1
                        End If
                    Else
                        dictVals(txt) = CDbl(v)
                    End If
                End If
            End If
        Next i
        If chk.CancelPairs > 0 Then chk.Observations = chk.Observations & chk.CancelPairs & " paire(s) de montants qui s'annulent (a retirer du proof); "
    End If

    ' Copier-coller du releve de compte: en-tetes de journal, beaucoup de lignes, aucune formule
    If headerHits >= 2 And chk.Formulas = 0 And bestCnt > 15 Then
        chk.LooksLikeStatement = True
        chk.Observations = chk.Observations & "Ressemble a une copie du detail du compte (en-tetes de releve, " & bestCnt & " lignes, aucune formule): ce n'est pas un justificatif; "
    End If
    If chk.Formulas = 0 And chk.NumericCells > 0 And Not chk.LooksLikeStatement Then
        chk.Observations = chk.Observations & "Aucune formule dans la feuille; "
    End If

    AnalyserFeuille = chk
    Exit Function
ErrHandler:
    chk.Observations = chk.Observations & "Erreur d'analyse: " & Err.Description & "; "
    AnalyserFeuille = chk
End Function

' ==============================================================================
' HELPERS
' ==============================================================================

Private Function ChargerSoldesGL() As Object
    ' Soldes GL a la date d'arrete (GLPROOF_DATA col. D, sinon RECONCIL col. C)
    Dim d As Object, ws As Worksheet, lr As Long, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If SAFA_Common.FeuilleExiste("GLPROOF_DATA") Then
        Set ws = ThisWorkbook.Sheets("GLPROOF_DATA")
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lr
            k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(ws.Cells(i, 1).Value))
            If k <> "" And Not d.Exists(k) Then d.Add k, SAFA_Common.SafeVal(ws.Cells(i, 4).Value)
        Next i
    ElseIf SAFA_Common.FeuilleExiste("RECONCIL") Then
        Set ws = ThisWorkbook.Sheets("RECONCIL")
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lr
            k = SAFA_Common.NormalizeAccountKey(SAFA_Common.SafeText(ws.Cells(i, SAFA_Common.RECONCIL_COL_COMPTE).Value))
            If k <> "" And Not d.Exists(k) Then d.Add k, SAFA_Common.SafeVal(ws.Cells(i, SAFA_Common.RECONCIL_COL_SOLDE_GL).Value)
        Next i
    End If
    Set ChargerSoldesGL = d
End Function

Private Function ChoisirDossier(titre As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = titre
    If fd.Show = -1 Then ChoisirDossier = fd.SelectedItems(1) Else ChoisirDossier = ""
End Function
