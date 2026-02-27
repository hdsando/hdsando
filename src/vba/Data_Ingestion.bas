Attribute VB_Name = "Data_Ingestion"
'===============================================================================
' MODULE: Data_Ingestion
' Description: Import et traitement des donnees multi-formats (Excel, CSV, TXT)
' Version: 10.0
' Auteur: S.A.F.A Team
' Date: Decembre 2024
'===============================================================================
Option Explicit

' ===== CONSTANTES =====
Private Const MODULE_NAME As String = "Data_Ingestion"
Private Const MAX_SCAN_ROWS As Long = 50
Private Const MAX_SCAN_COLS As Long = 50

' ===== TYPES PERSONNALISES =====
Public Type ColumnMapping
    ColumnIndex As Long
    ColumnName As String
    DataType As String
    IsRequired As Boolean
    DefaultValue As Variant
End Type

Public Type FileInfo
    FilePath As String
    FileName As String
    FileType As String
    FileSize As Long
    SheetCount As Long
    RowCount As Long
    ColCount As Long
    HasHeaders As Boolean
    Encoding As String
    Delimiter As String
End Type

Public Type ImportResult
    Success As Boolean
    RowsImported As Long
    RowsSkipped As Long
    Errors As String
    Duration As Double
    TargetSheet As String
End Type

' ===== VARIABLES MODULE =====
Private mColumnMappings As Collection
Private mLastImportResult As ImportResult

'===============================================================================
' FONCTION PRINCIPALE: ImportFile
' Description: Import intelligent de fichiers (detection automatique du format)
'===============================================================================
Public Function ImportFile(filePath As String, targetSheetName As String, _
                          Optional hasHeaders As Boolean = True, _
                          Optional delimiter As String = "") As ImportResult
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim fileInfo As FileInfo
    Dim result As ImportResult

    startTime = Timer

    ' Analyser le fichier
    fileInfo = AnalyzeFile(filePath)

    If fileInfo.FilePath = "" Then
        result.Success = False
        result.Errors = "Fichier non trouve ou inaccessible"
        ImportFile = result
        Exit Function
    End If

    ' Importer selon le type
    Select Case UCase(fileInfo.FileType)
        Case "XLSX", "XLS", "XLSM"
            result = ImportExcelFile(filePath, targetSheetName, hasHeaders)

        Case "CSV"
            If delimiter = "" Then delimiter = DetectDelimiter(filePath)
            result = ImportCSVFile(filePath, targetSheetName, hasHeaders, delimiter)

        Case "TXT"
            If delimiter = "" Then delimiter = vbTab
            result = ImportTextFile(filePath, targetSheetName, hasHeaders, delimiter)

        Case Else
            result.Success = False
            result.Errors = "Format de fichier non supporte: " & fileInfo.FileType
    End Select

    result.Duration = Timer - startTime
    result.TargetSheet = targetSheetName
    mLastImportResult = result

    ImportFile = result
    Exit Function

ErrorHandler:
    result.Success = False
    result.Errors = "Erreur " & Err.Number & ": " & Err.Description
    ImportFile = result
    LogError MODULE_NAME, "ImportFile", Err.Number, Err.Description
End Function

'===============================================================================
' FONCTION: AnalyzeFile
' Description: Analyse un fichier et retourne ses caracteristiques
'===============================================================================
Public Function AnalyzeFile(filePath As String) As FileInfo
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim file As Object
    Dim info As FileInfo

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        AnalyzeFile = info
        Exit Function
    End If

    Set file = fso.GetFile(filePath)

    With info
        .FilePath = filePath
        .FileName = fso.GetFileName(filePath)
        .FileType = UCase(fso.GetExtensionName(filePath))
        .FileSize = file.Size
        .Encoding = "UTF-8" ' Par defaut
        .HasHeaders = True
    End With

    ' Pour les fichiers Excel, compter les feuilles
    If info.FileType = "XLSX" Or info.FileType = "XLS" Or info.FileType = "XLSM" Then
        Dim tempWb As Workbook
        On Error Resume Next
        Set tempWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        If Not tempWb Is Nothing Then
            info.SheetCount = tempWb.Sheets.Count
            info.RowCount = tempWb.Sheets(1).UsedRange.Rows.Count
            info.ColCount = tempWb.Sheets(1).UsedRange.Columns.Count
            tempWb.Close SaveChanges:=False
        End If
        On Error GoTo ErrorHandler
    End If

    ' Pour CSV/TXT, detecter le delimiteur
    If info.FileType = "CSV" Or info.FileType = "TXT" Then
        info.Delimiter = DetectDelimiter(filePath)
        info.RowCount = CountFileLines(filePath)
    End If

    Set fso = Nothing
    AnalyzeFile = info
    Exit Function

ErrorHandler:
    LogError MODULE_NAME, "AnalyzeFile", Err.Number, Err.Description
    AnalyzeFile = info
End Function

'===============================================================================
' FONCTION: ImportExcelFile
' Description: Importe un fichier Excel
'===============================================================================
Public Function ImportExcelFile(filePath As String, targetSheetName As String, _
                               Optional hasHeaders As Boolean = True) As ImportResult
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim result As ImportResult
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    ' Ouvrir le fichier source
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

    ' Prendre la premiere feuille
    Set sourceWs = sourceWb.Sheets(1)

    ' Determiner la plage de donnees
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then
        result.Success = False
        result.Errors = "Fichier vide ou format invalide"
        sourceWb.Close SaveChanges:=False
        ImportExcelFile = result
        Exit Function
    End If

    Set dataRange = sourceWs.Range(sourceWs.Cells(1, 1), sourceWs.Cells(lastRow, lastCol))

    ' Preparer la feuille cible
    Set targetWs = PrepareTargetSheet(targetSheetName)

    ' Copier les donnees
    dataRange.Copy targetWs.Range("A1")

    ' Nettoyer
    Application.CutCopyMode = False
    sourceWb.Close SaveChanges:=False

    result.Success = True
    result.RowsImported = lastRow - IIf(hasHeaders, 1, 0)
    result.TargetSheet = targetSheetName

    Application.ScreenUpdating = True
    ImportExcelFile = result
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    result.Success = False
    result.Errors = "Erreur " & Err.Number & ": " & Err.Description
    ImportExcelFile = result
    LogError MODULE_NAME, "ImportExcelFile", Err.Number, Err.Description

    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
End Function

'===============================================================================
' FONCTION: ImportCSVFile
' Description: Importe un fichier CSV avec gestion des encodages
'===============================================================================
Public Function ImportCSVFile(filePath As String, targetSheetName As String, _
                             Optional hasHeaders As Boolean = True, _
                             Optional delimiter As String = ",") As ImportResult
    On Error GoTo ErrorHandler

    Dim targetWs As Worksheet
    Dim result As ImportResult
    Dim qt As QueryTable

    ' Preparer la feuille cible
    Set targetWs = PrepareTargetSheet(targetSheetName)

    ' Utiliser QueryTable pour import robuste
    Set qt = targetWs.QueryTables.Add( _
        Connection:="TEXT;" & filePath, _
        Destination:=targetWs.Range("A1"))

    With qt
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = (delimiter = vbTab)
        .TextFileSemicolonDelimiter = (delimiter = ";")
        .TextFileCommaDelimiter = (delimiter = ",")
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1) ' General
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ' Compter les lignes importees
    result.RowsImported = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    If hasHeaders Then result.RowsImported = result.RowsImported - 1

    result.Success = True
    result.TargetSheet = targetSheetName

    ImportCSVFile = result
    Exit Function

ErrorHandler:
    result.Success = False
    result.Errors = "Erreur " & Err.Number & ": " & Err.Description
    ImportCSVFile = result
    LogError MODULE_NAME, "ImportCSVFile", Err.Number, Err.Description
End Function

'===============================================================================
' FONCTION: ImportTextFile
' Description: Importe un fichier texte delimite
'===============================================================================
Public Function ImportTextFile(filePath As String, targetSheetName As String, _
                              Optional hasHeaders As Boolean = True, _
                              Optional delimiter As String = vbTab) As ImportResult
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim targetWs As Worksheet
    Dim result As ImportResult
    Dim lineContent As String
    Dim fields() As String
    Dim rowNum As Long
    Dim colNum As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)

    Set targetWs = PrepareTargetSheet(targetSheetName)

    rowNum = 1

    ' Desactiver les mises a jour ecran
    Application.ScreenUpdating = False

    Do Until ts.AtEndOfStream
        lineContent = ts.ReadLine

        ' Parser la ligne
        fields = Split(lineContent, delimiter)

        ' Ecrire les champs
        For colNum = 0 To UBound(fields)
            targetWs.Cells(rowNum, colNum + 1).Value = Trim(fields(colNum))
        Next colNum

        rowNum = rowNum + 1

        ' Limite de securite
        If rowNum > 1000000 Then Exit Do
    Loop

    ts.Close

    Application.ScreenUpdating = True

    result.Success = True
    result.RowsImported = rowNum - 1 - IIf(hasHeaders, 1, 0)
    result.TargetSheet = targetSheetName

    Set fso = Nothing
    ImportTextFile = result
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    result.Success = False
    result.Errors = "Erreur " & Err.Number & ": " & Err.Description
    ImportTextFile = result
    LogError MODULE_NAME, "ImportTextFile", Err.Number, Err.Description
End Function

'===============================================================================
' FONCTION: ImportAllSheets
' Description: Importe toutes les feuilles d'un classeur Excel
'===============================================================================
Public Function ImportAllSheets(filePath As String, prefix As String) As ImportResult
    On Error GoTo ErrorHandler

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim result As ImportResult
    Dim totalRows As Long
    Dim sheetCount As Long

    Application.ScreenUpdating = False

    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

    For Each sourceWs In sourceWb.Sheets
        ' Ignorer les feuilles vides
        If sourceWs.UsedRange.Rows.Count > 1 Then
            ' Creer la feuille cible
            Set targetWs = PrepareTargetSheet(prefix & "_" & Left(sourceWs.Name, 20))

            ' Copier les donnees
            sourceWs.UsedRange.Copy targetWs.Range("A1")

            totalRows = totalRows + sourceWs.UsedRange.Rows.Count
            sheetCount = sheetCount + 1
        End If
    Next sourceWs

    Application.CutCopyMode = False
    sourceWb.Close SaveChanges:=False

    Application.ScreenUpdating = True

    result.Success = True
    result.RowsImported = totalRows
    result.TargetSheet = sheetCount & " feuilles importees"

    ImportAllSheets = result
    Exit Function

ErrorHandler:
    Application.ScreenUpdating = True
    result.Success = False
    result.Errors = "Erreur " & Err.Number & ": " & Err.Description
    ImportAllSheets = result
    LogError MODULE_NAME, "ImportAllSheets", Err.Number, Err.Description

    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
End Function

'===============================================================================
' FONCTION: DetectColumns
' Description: Detection automatique des colonnes (Balance, GL Proof, etc.)
'===============================================================================
Public Function DetectColumns(ws As Worksheet, fileType As String) As Collection
    On Error GoTo ErrorHandler

    Dim mappings As New Collection
    Dim colMap As ColumnMapping
    Dim col As Long
    Dim headerValue As String
    Dim lastCol As Long
    Dim scanRow As Long

    ' Scanner les premieres lignes pour trouver l'en-tete
    For scanRow = 1 To MAX_SCAN_ROWS
        If ws.Cells(scanRow, 1).Value <> "" Then
            lastCol = ws.Cells(scanRow, ws.Columns.Count).End(xlToLeft).Column
            If lastCol > 1 Then Exit For
        End If
    Next scanRow

    ' Mapper les colonnes selon le type de fichier
    For col = 1 To lastCol
        headerValue = UCase(Trim(ws.Cells(scanRow, col).Value))

        Select Case fileType
            Case "BALANCE"
                colMap = MapBalanceColumn(headerValue, col)
            Case "GLPROOF"
                colMap = MapGLProofColumn(headerValue, col)
            Case Else
                colMap = MapGenericColumn(headerValue, col)
        End Select

        If colMap.ColumnName <> "" Then
            mappings.Add colMap, colMap.ColumnName
        End If
    Next col

    Set DetectColumns = mappings
    Exit Function

ErrorHandler:
    LogError MODULE_NAME, "DetectColumns", Err.Number, Err.Description
    Set DetectColumns = New Collection
End Function

'===============================================================================
' FONCTION: MapBalanceColumn
' Description: Mapping des colonnes Balance Finacle
'===============================================================================
Private Function MapBalanceColumn(headerValue As String, colIndex As Long) As ColumnMapping
    Dim colMap As ColumnMapping

    colMap.ColumnIndex = colIndex
    colMap.DataType = "TEXT"

    Select Case True
        Case InStr(headerValue, "ACCOUNT") > 0 Or InStr(headerValue, "COMPTE") > 0 Or InStr(headerValue, "ACID") > 0
            colMap.ColumnName = "ACCOUNT_NO"
            colMap.IsRequired = True

        Case InStr(headerValue, "BALANCE") > 0 Or InStr(headerValue, "SOLDE") > 0
            colMap.ColumnName = "BALANCE"
            colMap.DataType = "NUMBER"
            colMap.IsRequired = True

        Case InStr(headerValue, "CURRENCY") > 0 Or InStr(headerValue, "DEVISE") > 0 Or InStr(headerValue, "CCY") > 0
            colMap.ColumnName = "CURRENCY"

        Case InStr(headerValue, "NAME") > 0 Or InStr(headerValue, "LIBELLE") > 0 Or InStr(headerValue, "INTITULE") > 0
            colMap.ColumnName = "ACCOUNT_NAME"

        Case InStr(headerValue, "BRANCH") > 0 Or InStr(headerValue, "SOL") > 0 Or InStr(headerValue, "AGENCE") > 0
            colMap.ColumnName = "SOL_ID"

        Case InStr(headerValue, "DATE") > 0
            colMap.ColumnName = "VALUE_DATE"
            colMap.DataType = "DATE"

        Case InStr(headerValue, "DEBIT") > 0 Or InStr(headerValue, "DR") > 0
            colMap.ColumnName = "DEBIT"
            colMap.DataType = "NUMBER"

        Case InStr(headerValue, "CREDIT") > 0 Or InStr(headerValue, "CR") > 0
            colMap.ColumnName = "CREDIT"
            colMap.DataType = "NUMBER"

        Case Else
            colMap.ColumnName = ""
    End Select

    MapBalanceColumn = colMap
End Function

'===============================================================================
' FONCTION: MapGLProofColumn
' Description: Mapping des colonnes GL Proof
'===============================================================================
Private Function MapGLProofColumn(headerValue As String, colIndex As Long) As ColumnMapping
    Dim colMap As ColumnMapping

    colMap.ColumnIndex = colIndex
    colMap.DataType = "TEXT"

    Select Case True
        Case InStr(headerValue, "GL") > 0 And InStr(headerValue, "CODE") > 0
            colMap.ColumnName = "GL_CODE"
            colMap.IsRequired = True

        Case InStr(headerValue, "ACCOUNT") > 0 Or InStr(headerValue, "FOLIO") > 0
            colMap.ColumnName = "ACCOUNT_NO"
            colMap.IsRequired = True

        Case InStr(headerValue, "AMOUNT") > 0 Or InStr(headerValue, "MONTANT") > 0
            colMap.ColumnName = "AMOUNT"
            colMap.DataType = "NUMBER"
            colMap.IsRequired = True

        Case InStr(headerValue, "TRAN") > 0 And (InStr(headerValue, "DATE") > 0 Or InStr(headerValue, "DT") > 0)
            colMap.ColumnName = "TRAN_DATE"
            colMap.DataType = "DATE"

        Case InStr(headerValue, "VALUE") > 0 And InStr(headerValue, "DATE") > 0
            colMap.ColumnName = "VALUE_DATE"
            colMap.DataType = "DATE"

        Case InStr(headerValue, "NARR") > 0 Or InStr(headerValue, "DESC") > 0 Or InStr(headerValue, "LIBELLE") > 0
            colMap.ColumnName = "NARRATION"

        Case InStr(headerValue, "REF") > 0 Or InStr(headerValue, "TRAN") > 0 And InStr(headerValue, "ID") > 0
            colMap.ColumnName = "TRAN_ID"

        Case InStr(headerValue, "PART") > 0 Or InStr(headerValue, "TYPE") > 0
            colMap.ColumnName = "PART_TRAN_TYPE"

        Case InStr(headerValue, "SOL") > 0 Or InStr(headerValue, "BRANCH") > 0
            colMap.ColumnName = "SOL_ID"

        Case Else
            colMap.ColumnName = ""
    End Select

    MapGLProofColumn = colMap
End Function

'===============================================================================
' FONCTION: MapGenericColumn
' Description: Mapping generique des colonnes
'===============================================================================
Private Function MapGenericColumn(headerValue As String, colIndex As Long) As ColumnMapping
    Dim colMap As ColumnMapping

    colMap.ColumnIndex = colIndex
    colMap.ColumnName = "COL_" & colIndex
    colMap.DataType = "TEXT"
    colMap.IsRequired = False

    MapGenericColumn = colMap
End Function

'===============================================================================
' FONCTION: ValidateImportedData
' Description: Valide les donnees importees
'===============================================================================
Public Function ValidateImportedData(ws As Worksheet, fileType As String) As String
    On Error GoTo ErrorHandler

    Dim errors As String
    Dim lastRow As Long
    Dim col As Long
    Dim mappings As Collection
    Dim colMap As ColumnMapping
    Dim emptyCount As Long
    Dim i As Long

    errors = ""
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        ValidateImportedData = "ERREUR: Aucune donnee importee"
        Exit Function
    End If

    ' Detecter les colonnes
    Set mappings = DetectColumns(ws, fileType)

    ' Verifier les colonnes requises
    Select Case fileType
        Case "BALANCE"
            If Not CollectionContains(mappings, "ACCOUNT_NO") Then
                errors = errors & "- Colonne ACCOUNT_NO manquante" & vbNewLine
            End If
            If Not CollectionContains(mappings, "BALANCE") Then
                errors = errors & "- Colonne BALANCE manquante" & vbNewLine
            End If

        Case "GLPROOF"
            If Not CollectionContains(mappings, "ACCOUNT_NO") Then
                errors = errors & "- Colonne ACCOUNT_NO manquante" & vbNewLine
            End If
            If Not CollectionContains(mappings, "AMOUNT") Then
                errors = errors & "- Colonne AMOUNT manquante" & vbNewLine
            End If
    End Select

    ' Verifier les lignes vides
    emptyCount = 0
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = "" Then
            emptyCount = emptyCount + 1
        End If
    Next i

    If emptyCount > lastRow * 0.1 Then
        errors = errors & "- " & emptyCount & " lignes vides detectees (>" & 10 & "%)" & vbNewLine
    End If

    If errors = "" Then
        ValidateImportedData = "OK"
    Else
        ValidateImportedData = errors
    End If

    Exit Function

ErrorHandler:
    ValidateImportedData = "ERREUR: " & Err.Description
    LogError MODULE_NAME, "ValidateImportedData", Err.Number, Err.Description
End Function

'===============================================================================
' FONCTIONS UTILITAIRES
'===============================================================================

Private Function PrepareTargetSheet(sheetName As String) As Worksheet
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    ' Supprimer si existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Creer nouvelle feuille
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = Left(sheetName, 31) ' Excel limite a 31 caracteres

    Set PrepareTargetSheet = ws
    Exit Function

ErrorHandler:
    LogError MODULE_NAME, "PrepareTargetSheet", Err.Number, Err.Description
    Set PrepareTargetSheet = Nothing
End Function

Private Function DetectDelimiter(filePath As String) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim firstLine As String
    Dim commaCount As Long
    Dim semicolonCount As Long
    Dim tabCount As Long
    Dim pipeCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)

    If Not ts.AtEndOfStream Then
        firstLine = ts.ReadLine
    End If

    ts.Close

    ' Compter les delimiteurs potentiels
    commaCount = Len(firstLine) - Len(Replace(firstLine, ",", ""))
    semicolonCount = Len(firstLine) - Len(Replace(firstLine, ";", ""))
    tabCount = Len(firstLine) - Len(Replace(firstLine, vbTab, ""))
    pipeCount = Len(firstLine) - Len(Replace(firstLine, "|", ""))

    ' Determiner le delimiteur le plus frequent
    If tabCount >= commaCount And tabCount >= semicolonCount And tabCount >= pipeCount Then
        DetectDelimiter = vbTab
    ElseIf semicolonCount >= commaCount And semicolonCount >= pipeCount Then
        DetectDelimiter = ";"
    ElseIf pipeCount >= commaCount Then
        DetectDelimiter = "|"
    Else
        DetectDelimiter = ","
    End If

    Set fso = Nothing
    Exit Function

ErrorHandler:
    DetectDelimiter = ","
    LogError MODULE_NAME, "DetectDelimiter", Err.Number, Err.Description
End Function

Private Function CountFileLines(filePath As String) As Long
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim ts As Object
    Dim lineCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)

    lineCount = 0
    Do Until ts.AtEndOfStream
        ts.ReadLine
        lineCount = lineCount + 1
    Loop

    ts.Close
    Set fso = Nothing

    CountFileLines = lineCount
    Exit Function

ErrorHandler:
    CountFileLines = 0
    LogError MODULE_NAME, "CountFileLines", Err.Number, Err.Description
End Function

Private Function CollectionContains(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim item As Variant
    item = col(key)
    CollectionContains = (Err.Number = 0)
    Err.Clear
End Function

Public Function GetLastImportResult() As ImportResult
    GetLastImportResult = mLastImportResult
End Function

'===============================================================================
' FONCTION: CleanImportedData
' Description: Nettoie les donnees importees (espaces, caracteres speciaux)
'===============================================================================
Public Sub CleanImportedData(ws As Worksheet, Optional trimSpaces As Boolean = True, _
                            Optional removeSpecialChars As Boolean = False)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim cellValue As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Application.ScreenUpdating = False

    For Each cell In dataRange
        If Not IsEmpty(cell.Value) Then
            cellValue = CStr(cell.Value)

            ' Supprimer espaces en debut/fin
            If trimSpaces Then
                cellValue = Trim(cellValue)
            End If

            ' Supprimer caracteres speciaux
            If removeSpecialChars Then
                cellValue = CleanSpecialChars(cellValue)
            End If

            If CStr(cell.Value) <> cellValue Then
                cell.Value = cellValue
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    LogError MODULE_NAME, "CleanImportedData", Err.Number, Err.Description
End Sub

Private Function CleanSpecialChars(text As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String

    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        ' Garder alphanumerique, espaces et ponctuations courantes
        If c Like "[A-Za-z0-9 .,;:!?()-_/@]" Then
            result = result & c
        End If
    Next i

    CleanSpecialChars = result
End Function

'===============================================================================
' FONCTION: ConvertColumnToNumber
' Description: Convertit une colonne texte en numerique
'===============================================================================
Public Sub ConvertColumnToNumber(ws As Worksheet, colIndex As Long)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim numValue As Double

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lastRow ' Ignorer l'en-tete
        cellValue = CStr(ws.Cells(i, colIndex).Value)

        ' Nettoyer la valeur
        cellValue = Replace(cellValue, " ", "")
        cellValue = Replace(cellValue, Chr(160), "") ' Espace insecable
        cellValue = Replace(cellValue, ",", ".")

        ' Supprimer les symboles de devise
        cellValue = Replace(cellValue, "XAF", "")
        cellValue = Replace(cellValue, "EUR", "")
        cellValue = Replace(cellValue, "USD", "")
        cellValue = Replace(cellValue, "$", "")
        cellValue = Replace(cellValue, "€", "")
        cellValue = Replace(cellValue, "F", "")

        On Error Resume Next
        numValue = CDbl(Trim(cellValue))
        If Err.Number = 0 Then
            ws.Cells(i, colIndex).Value = numValue
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next i

    ' Formater la colonne
    ws.Columns(colIndex).NumberFormat = "#,##0.00"

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    LogError MODULE_NAME, "ConvertColumnToNumber", Err.Number, Err.Description
End Sub

' CORRIGÉ: Utiliser SAFA_Common.LogError pour centraliser la gestion d'erreurs
Private Sub LogError(moduleName As String, procName As String, errNum As Long, errDesc As String)
    On Error Resume Next
    Call SAFA_Common.LogError(moduleName, procName, errNum, errDesc)
End Sub
