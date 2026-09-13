' ==============================================================================
' S.A.F.A v10.0 - INSTALLATEUR AUTOMATIQUE (Windows Script Host)
' ==============================================================================
' Double-cliquez sur ce fichier: il cree SAFA.xlsm a la racine du projet et
' y importe tous les modules VBA de src\vba (transcodes en Windows-1252 pour
' eviter les accents corrompus), injecte le code ThisWorkbook, construit la
' feuille MENU et (optionnel) charge les donnees de demonstration.
'
' PREREQUIS (une seule fois):
'   Excel > Fichier > Options > Centre de gestion de la confidentialite
'   > Parametres... > Parametres des macros
'   > cocher "Acces approuve au modele d'objet du projet VBA"
' ==============================================================================
Option Explicit

Dim fso, shell, scriptDir, rootDir, srcDir, tmpDir, logPath, xlsmPath
Dim xl, wb, log
Dim imported, skipped, errors

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
rootDir = fso.GetParentFolderName(scriptDir)
srcDir = fso.BuildPath(rootDir, "src\vba")
tmpDir = fso.BuildPath(scriptDir, "_tmp_ansi")
logPath = fso.BuildPath(scriptDir, "install_log.txt")
xlsmPath = fso.BuildPath(rootDir, "SAFA.xlsm")

imported = "" : skipped = "" : errors = 0

Set log = fso.CreateTextFile(logPath, True)
LogLine "S.A.F.A - Installation demarree " & Now
LogLine "Racine: " & rootDir

If Not fso.FolderExists(srcDir) Then
    Fail "Dossier introuvable: " & srcDir & vbCrLf & "Placez ce script dans le dossier install\ du projet."
End If

' ---- Sauvegarde d'un SAFA.xlsm existant ----
If fso.FileExists(xlsmPath) Then
    Dim bak
    bak = fso.BuildPath(rootDir, "SAFA_backup_" & Stamp() & ".xlsm")
    fso.CopyFile xlsmPath, bak, True
    LogLine "Sauvegarde: " & bak
End If

' ---- Excel ----
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then Fail "Impossible de demarrer Excel: " & Err.Description
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Set wb = xl.Workbooks.Add
' 52 = xlOpenXMLWorkbookMacroEnabled
On Error Resume Next
If fso.FileExists(xlsmPath) Then fso.DeleteFile xlsmPath, True
wb.SaveAs xlsmPath, 52
If Err.Number <> 0 Then
    CleanupExcel
    Fail "Echec de l'enregistrement de " & xlsmPath & vbCrLf & Err.Description
End If
On Error GoTo 0
LogLine "Classeur cree: " & xlsmPath

' ---- Verifier l'acces au projet VBA ----
On Error Resume Next
Dim nComp
nComp = wb.VBProject.VBComponents.Count
If Err.Number <> 0 Then
    Err.Clear
    CleanupExcel
    MsgBox "L'acces au projet VBA n'est pas autorise." & vbCrLf & vbCrLf & _
           "Dans Excel:" & vbCrLf & _
           "  Fichier > Options > Centre de gestion de la confidentialite" & vbCrLf & _
           "  > Parametres du Centre de gestion de la confidentialite..." & vbCrLf & _
           "  > Parametres des macros" & vbCrLf & _
           "  > cochez 'Acces approuve au modele d'objet du projet VBA'" & vbCrLf & vbCrLf & _
           "Puis relancez Install_SAFA.vbs.", 48, "S.A.F.A - Installation"
    WScript.Quit 1
End If
On Error GoTo 0

' ---- Transcodage UTF-8 -> Windows-1252 et import ----
If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
fso.CreateFolder tmpDir

Dim f, name, ext, ansiPath
' SAFA_Common en premier (constantes partagees), puis le reste
ImportModule "SAFA_Common.bas"
For Each f In fso.GetFolder(srcDir).Files
    name = f.Name
    ext = LCase(fso.GetExtensionName(name))
    If ext = "bas" And LCase(name) <> "safa_common.bas" Then
        If LCase(name) = "formbuilder.bas" Or LCase(name) = "safa_launcher.bas" Then
            skipped = skipped & name & " "
            LogLine "Ignore (obsolete): " & name
        Else
            ImportModule name
        End If
    End If
Next

' ---- ThisWorkbook.cls : injection du code (sans l'en-tete VERSION/Attribute) ----
InjectThisWorkbook

' ---- Construire le MENU ----
On Error Resume Next
xl.Run "'" & wb.Name & "'!SAFA_Menu.BuildMenu"
If Err.Number <> 0 Then
    LogLine "AVERTISSEMENT BuildMenu: " & Err.Description
    errors = errors + 1
    Err.Clear
Else
    LogLine "Feuille MENU construite"
End If
On Error GoTo 0

' ---- Donnees demo (optionnel) ----
Dim rep
rep = MsgBox("Modules importes." & vbCrLf & vbCrLf & _
             "Charger un jeu de donnees de demonstration (300 comptes, ~4 000 ecritures) pour tester immediatement ?", _
             36, "S.A.F.A - Installation")
If rep = 6 Then
    On Error Resume Next
    xl.Run "'" & wb.Name & "'!Demo_Data.GenerateDemoData"
    If Err.Number <> 0 Then
        LogLine "AVERTISSEMENT Demo_Data: " & Err.Description
        errors = errors + 1
        Err.Clear
    Else
        LogLine "Donnees demo chargees"
    End If
    xl.Run "'" & wb.Name & "'!SAFA_Menu.RefreshStatus"
    On Error GoTo 0
End If

' ---- Sauvegarde et ouverture ----
wb.Save
LogLine "Classeur sauvegarde"

On Error Resume Next
fso.DeleteFolder tmpDir, True
On Error GoTo 0

xl.EnableEvents = True
xl.DisplayAlerts = True
xl.Visible = True
On Error Resume Next
wb.Sheets("MENU").Activate
On Error GoTo 0

LogLine "Installation terminee. Modules: " & imported
log.Close

MsgBox "Installation terminee !" & vbCrLf & vbCrLf & _
       "Fichier: " & xlsmPath & vbCrLf & _
       "Modules importes: " & imported & vbCrLf & _
       IIf(errors > 0, vbCrLf & errors & " avertissement(s) - voir install_log.txt" & vbCrLf, "") & vbCrLf & _
       "Utilisez les boutons de la feuille MENU." & vbCrLf & _
       "Pour tester: bouton 'Tests auto' (donnees demo + pipeline + verifications).", _
       64, "S.A.F.A - Installation"

' ==============================================================================
' PROCEDURES
' ==============================================================================

Sub ImportModule(fileName)
    Dim src, dst
    src = fso.BuildPath(srcDir, fileName)
    If Not fso.FileExists(src) Then
        LogLine "Absent: " & fileName
        Exit Sub
    End If
    dst = fso.BuildPath(tmpDir, fileName)
    On Error Resume Next
    TranscodeToAnsi src, dst
    If Err.Number <> 0 Then
        LogLine "ERREUR transcodage " & fileName & ": " & Err.Description
        errors = errors + 1
        Err.Clear
        Exit Sub
    End If
    wb.VBProject.VBComponents.Import dst
    If Err.Number <> 0 Then
        LogLine "ERREUR import " & fileName & ": " & Err.Description
        errors = errors + 1
        Err.Clear
    Else
        imported = imported & Replace(fileName, ".bas", "") & " "
        LogLine "Importe: " & fileName
    End If
    On Error GoTo 0
End Sub

Sub InjectThisWorkbook()
    Dim src, dst, txt, lines, i, startIdx, code, cm
    src = fso.BuildPath(srcDir, "ThisWorkbook.cls")
    If Not fso.FileExists(src) Then
        LogLine "ThisWorkbook.cls absent - ignore"
        Exit Sub
    End If
    dst = fso.BuildPath(tmpDir, "ThisWorkbook.txt")
    On Error Resume Next
    TranscodeToAnsi src, dst
    txt = fso.OpenTextFile(dst, 1).ReadAll
    lines = Split(txt, vbLf)
    ' Sauter l'en-tete: VERSION / BEGIN..END / Attribute VB_*
    startIdx = 0
    For i = 0 To UBound(lines)
        If Left(Trim(Replace(lines(i), vbCr, "")), 10) = "Attribute " Then startIdx = i + 1
        If i > 20 Then Exit For
    Next
    code = ""
    For i = startIdx To UBound(lines)
        code = code & Replace(lines(i), vbCr, "") & vbCrLf
    Next
    Set cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString code
    If Err.Number <> 0 Then
        LogLine "ERREUR ThisWorkbook: " & Err.Description
        errors = errors + 1
        Err.Clear
    Else
        LogLine "Code ThisWorkbook injecte"
        imported = imported & "ThisWorkbook "
    End If
    On Error GoTo 0
End Sub

Sub TranscodeToAnsi(srcPath, dstPath)
    ' Les .bas du depot sont en UTF-8; l'editeur VBA importe en ANSI (Windows-1252).
    Dim sIn, sOut, txt
    Set sIn = CreateObject("ADODB.Stream")
    sIn.Type = 2 : sIn.Charset = "utf-8" : sIn.Open
    sIn.LoadFromFile srcPath
    txt = sIn.ReadText(-1)
    sIn.Close
    Set sOut = CreateObject("ADODB.Stream")
    sOut.Type = 2 : sOut.Charset = "windows-1252" : sOut.Open
    sOut.WriteText txt
    sOut.SaveToFile dstPath, 2
    sOut.Close
End Sub

Sub LogLine(msg)
    On Error Resume Next
    log.WriteLine Now & " | " & msg
End Sub

Sub CleanupExcel()
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set wb = Nothing : Set xl = Nothing
End Sub

Sub Fail(msg)
    On Error Resume Next
    LogLine "ECHEC: " & msg
    log.Close
    MsgBox msg, 16, "S.A.F.A - Installation"
    WScript.Quit 1
End Sub

Function Stamp()
    Dim d : d = Now
    Stamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & _
            Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
End Function
