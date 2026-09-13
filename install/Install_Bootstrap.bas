Attribute VB_Name = "Install_Bootstrap"
Option Explicit

' ==============================================================================
' S.A.F.A v10 - INSTALLATION SANS SCRIPT (module d'amorcage)
' ==============================================================================
' Usage (une seule fois):
'   1. Creez un classeur vide, enregistrez-le en .xlsm (ex: C:\SAFA\SAFA.xlsm),
'      DANS LE DOSSIER qui contient src\vba (racine du projet).
'   2. ALT+F11 > Fichier > Importer un fichier... > ce fichier (Install_Bootstrap.bas)
'   3. ALT+F8 > Installer_SAFA > Executer
' Le module importe tous les .bas de src\vba (avec conversion UTF-8 -> ANSI pour
' les accents), injecte ThisWorkbook.cls, construit la feuille MENU et propose
' de charger les donnees de demonstration.
' Prerequis: Excel > Options > Centre de gestion de la confidentialite
'            > Parametres des macros > "Acces approuve au modele d'objet du projet VBA"
' ==============================================================================

Public Sub Installer_SAFA()
    Dim srcDir As String, tmpDir As String, f As String
    Dim imported As String, errs As String, n As Long

    On Error GoTo ErrHandler

    ' --- 1. Verifier que le classeur est enregistre en .xlsm ---
    If LCase(Right(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        MsgBox "Enregistrez d'abord ce classeur au format .xlsm (Fichier > Enregistrer sous)," & vbCrLf & _
               "dans le dossier racine du projet (celui qui contient src\vba).", vbExclamation, "S.A.F.A - Installation"
        Exit Sub
    End If

    ' --- 2. Verifier l'acces au projet VBA ---
    Dim nComp As Long
    On Error Resume Next
    nComp = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "L'acces au projet VBA n'est pas autorise." & vbCrLf & vbCrLf & _
               "Fichier > Options > Centre de gestion de la confidentialite" & vbCrLf & _
               "> Parametres du Centre de gestion de la confidentialite... > Parametres des macros" & vbCrLf & _
               "> cochez 'Acces approuve au modele d'objet du projet VBA'" & vbCrLf & vbCrLf & _
               "Puis relancez Installer_SAFA.", vbCritical, "S.A.F.A - Installation"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' --- 3. Localiser src\vba ---
    srcDir = ThisWorkbook.Path & "\src\vba"
    If Dir(srcDir & "\SAFA_Common.bas") = "" Then
        srcDir = ChoisirDossier("Selectionnez le dossier src\vba du projet S.A.F.A")
        If srcDir = "" Then Exit Sub
        If Dir(srcDir & "\SAFA_Common.bas") = "" Then
            MsgBox "SAFA_Common.bas introuvable dans " & srcDir, vbCritical, "S.A.F.A - Installation"
            Exit Sub
        End If
    End If

    tmpDir = Environ("TEMP") & "\safa_install_" & Format(Now, "yyyymmddhhnnss")
    MkDir tmpDir

    Application.ScreenUpdating = False
    Application.StatusBar = "S.A.F.A: import des modules..."

    ' --- 4. Importer SAFA_Common en premier, puis les autres ---
    Call ImporterModule(srcDir, tmpDir, "SAFA_Common.bas", imported, errs)
    f = Dir(srcDir & "\*.bas")
    Do While f <> ""
        Select Case LCase(f)
            Case "safa_common.bas", "install_bootstrap.bas", "formbuilder.bas", "safa_launcher.bas"
                ' deja importe ou obsolete
            Case Else
                Call ImporterModule(srcDir, tmpDir, f, imported, errs)
        End Select
        f = Dir()
    Loop

    ' --- 5. ThisWorkbook.cls ---
    Call InjecterThisWorkbook(srcDir, tmpDir, imported, errs)

    ' --- 6. Nettoyage temp ---
    On Error Resume Next
    Kill tmpDir & "\*.*"
    RmDir tmpDir
    On Error GoTo ErrHandler

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' --- 7. Construire le MENU (les modules importes sont compiles a l'appel) ---
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!SAFA_Menu.BuildMenu"
    If Err.Number <> 0 Then errs = errs & "BuildMenu: " & Err.Description & vbCrLf: Err.Clear
    On Error GoTo ErrHandler

    ' --- 8. Donnees demo (optionnel) ---
    If MsgBox("Modules importes: " & imported & vbCrLf & vbCrLf & _
              "Charger un jeu de donnees de demonstration (300 comptes, ~4 000 ecritures) pour tester ?", _
              vbYesNo + vbQuestion, "S.A.F.A - Installation") = vbYes Then
        On Error Resume Next
        Application.Run "'" & ThisWorkbook.Name & "'!Demo_Data.GenerateDemoData"
        If Err.Number <> 0 Then errs = errs & "Demo_Data: " & Err.Description & vbCrLf: Err.Clear
        Application.Run "'" & ThisWorkbook.Name & "'!SAFA_Menu.RefreshStatus"
        On Error GoTo ErrHandler
    End If

    ThisWorkbook.Save

    If errs = "" Then
        MsgBox "Installation terminee !" & vbCrLf & vbCrLf & _
               "Utilisez les boutons de la feuille MENU." & vbCrLf & _
               "Pour tester: bouton 'Tests auto'." & vbCrLf & vbCrLf & _
               "Vous pouvez supprimer le module Install_Bootstrap (ALT+F11, clic droit > Supprimer).", _
               vbInformation, "S.A.F.A - Installation"
    Else
        MsgBox "Installation terminee avec avertissements:" & vbCrLf & vbCrLf & errs, vbExclamation, "S.A.F.A - Installation"
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur d'installation: " & Err.Description, vbCritical, "S.A.F.A - Installation"
End Sub

' ------------------------------------------------------------------------------

Private Sub ImporterModule(srcDir As String, tmpDir As String, fileName As String, ByRef imported As String, ByRef errs As String)
    Dim modName As String, dst As String, comp As Object
    modName = Left(fileName, Len(fileName) - 4)
    dst = tmpDir & "\" & fileName

    On Error Resume Next
    Call TranscoderVersAnsi(srcDir & "\" & fileName, dst)
    If Err.Number <> 0 Then
        errs = errs & fileName & " (transcodage): " & Err.Description & vbCrLf: Err.Clear
        Exit Sub
    End If

    ' Remplacer un module existant du meme nom
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents(modName)
    Err.Clear
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp

    ThisWorkbook.VBProject.VBComponents.Import dst
    If Err.Number <> 0 Then
        errs = errs & fileName & ": " & Err.Description & vbCrLf: Err.Clear
    Else
        imported = imported & modName & " "
    End If
    On Error GoTo 0
End Sub

Private Sub InjecterThisWorkbook(srcDir As String, tmpDir As String, ByRef imported As String, ByRef errs As String)
    Dim src As String, dst As String, txt As String, lines() As String
    Dim i As Long, startIdx As Long, code As String, cm As Object

    src = srcDir & "\ThisWorkbook.cls"
    If Dir(src) = "" Then Exit Sub
    dst = tmpDir & "\ThisWorkbook.txt"

    On Error Resume Next
    Call TranscoderVersAnsi(src, dst)
    txt = LireFichier(dst)
    lines = Split(Replace(txt, vbCrLf, vbLf), vbLf)

    ' Sauter l'en-tete VERSION / BEGIN..END / Attribute VB_*
    startIdx = 0
    For i = 0 To UBound(lines)
        If Left(Trim(lines(i)), 10) = "Attribute " Then startIdx = i + 1
        If i > 20 Then Exit For
    Next i
    For i = startIdx To UBound(lines)
        code = code & lines(i) & vbCrLf
    Next i

    Set cm = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString code
    If Err.Number <> 0 Then
        errs = errs & "ThisWorkbook: " & Err.Description & vbCrLf: Err.Clear
    Else
        imported = imported & "ThisWorkbook "
    End If
    On Error GoTo 0
End Sub

Private Sub TranscoderVersAnsi(srcPath As String, dstPath As String)
    ' Les .bas du depot sont en UTF-8; l'editeur VBA importe en ANSI (Windows-1252)
    Dim sIn As Object, sOut As Object, txt As String
    Set sIn = CreateObject("ADODB.Stream")
    sIn.Type = 2: sIn.Charset = "utf-8": sIn.Open
    sIn.LoadFromFile srcPath
    txt = sIn.ReadText(-1)
    sIn.Close
    Set sOut = CreateObject("ADODB.Stream")
    sOut.Type = 2: sOut.Charset = "windows-1252": sOut.Open
    sOut.WriteText txt
    sOut.SaveToFile dstPath, 2
    sOut.Close
End Sub

Private Function LireFichier(p As String) As String
    Dim ff As Integer
    ff = FreeFile
    Open p For Input As #ff
    LireFichier = Input$(LOF(ff), ff)
    Close #ff
End Function

Private Function ChoisirDossier(titre As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    fd.Title = titre
    If fd.Show = -1 Then ChoisirDossier = fd.SelectedItems(1) Else ChoisirDossier = ""
End Function
