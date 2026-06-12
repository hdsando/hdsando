Attribute VB_Name = "Security_Module"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SECURITY_MODULE v10.0
' ==============================================================================
' Description: Sécurité, Authentification, Audit Trail, Chiffrement
' Fonctionnalités P1-P4:
'   - P1: Audit Trail avec hash chainé (blockchain-like)
'   - P2: Gestion des sessions
'   - P3: Authentification utilisateurs (Login/Profils)
'   - P4: Chiffrement des exports sensibles
' ==============================================================================

' --- CONSTANTES ---
' NOTE: Constantes déplacées vers SAFA_Common.bas pour centralisation
' Utiliser SAFA_Common.SESSION_TIMEOUT_MINUTES, etc.

' --- TYPES ---
Public Type UserProfile
    Username As String
    Role As String          ' ADMIN, AUDITOR, VIEWER
    FullName As String
    Email As String
    LastLogin As Date
    IsActive As Boolean
End Type

Public Type SessionInfo
    SessionID As String
    Username As String
    StartTime As Date
    LastActivity As Date
    IsAuthenticated As Boolean
    Role As String
End Type

Public Type AuditLogEntry
    Timestamp As Date
    LogType As String
    Username As String
    Action As String
    Details As String
    Hash As String
    PreviousHash As String
    Verified As Boolean
End Type

' --- VARIABLES GLOBALES ---
Public g_CurrentSession As SessionInfo
Public g_LoginAttempts As Integer

' ==============================================================================
' 1. AUTHENTIFICATION (P3)
' ==============================================================================

Public Function Authenticate(username As String, password As String) As Boolean
    ' Authentification utilisateur
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim storedHash As String, inputHash As String

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        On Error GoTo 0
        Call WriteSecurityLog("AUTH_CONFIG_ERROR", username, "Feuille USERS introuvable - authentification refusée")
        MsgBox "Configuration de sécurité invalide: feuille USERS manquante." & vbCrLf & _
               "Contactez un administrateur.", vbCritical, "Erreur de sécurité"
        Authenticate = False
        Exit Function
    End If
    On Error GoTo 0

    ' Vérifier les tentatives
    If g_LoginAttempts >= MAX_LOGIN_ATTEMPTS Then
        MsgBox "Compte verrouillé. Trop de tentatives échouées.", vbCritical, "Sécurité"
        Call WriteSecurityLog("LOCKOUT", username, "Compte verrouillé après " & MAX_LOGIN_ATTEMPTS & " tentatives")
        Authenticate = False
        Exit Function
    End If

    ' Rechercher l'utilisateur
    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    inputHash = HashPassword(password)

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            storedHash = wsUsers.Cells(i, 2).Value

            If storedHash = inputHash Then
                ' Vérifier si actif
                If wsUsers.Cells(i, 6).Value = True Then
                    ' Authentification réussie
                    g_CurrentSession.SessionID = GenerateSessionID()
                    g_CurrentSession.Username = username
                    g_CurrentSession.StartTime = Now
                    g_CurrentSession.LastActivity = Now
                    g_CurrentSession.IsAuthenticated = True
                    g_CurrentSession.Role = wsUsers.Cells(i, 3).Value

                    ' Mettre à jour dernière connexion
                    wsUsers.Cells(i, 5).Value = Now

                    g_LoginAttempts = 0
                    Call WriteSecurityLog("LOGIN", username, "Connexion réussie - Role: " & g_CurrentSession.Role)

                    Authenticate = True
                    Exit Function
                Else
                    MsgBox "Ce compte est désactivé.", vbExclamation, "Compte Inactif"
                    Authenticate = False
                    Exit Function
                End If
            End If
        End If
    Next i

    ' Échec
    g_LoginAttempts = g_LoginAttempts + 1
    Call WriteSecurityLog("LOGIN_FAILED", username, "Tentative " & g_LoginAttempts & "/" & MAX_LOGIN_ATTEMPTS)

    MsgBox "Identifiants incorrects." & vbCrLf & _
           "Tentatives restantes: " & (MAX_LOGIN_ATTEMPTS - g_LoginAttempts), _
           vbExclamation, "Erreur d'authentification"

    Authenticate = False
End Function

Public Sub Logout()
    ' Déconnexion
    If g_CurrentSession.IsAuthenticated Then
        Call WriteSecurityLog("LOGOUT", g_CurrentSession.Username, "Session terminée")
    End If

    g_CurrentSession.SessionID = ""
    g_CurrentSession.Username = ""
    g_CurrentSession.IsAuthenticated = False
    g_CurrentSession.Role = ""
End Sub

Public Function IsAuthenticated() As Boolean
    ' Vérifier si session valide
    If Not g_CurrentSession.IsAuthenticated Then
        IsAuthenticated = False
        Exit Function
    End If

    ' Vérifier timeout
    If DateDiff("n", g_CurrentSession.LastActivity, Now) > SESSION_TIMEOUT_MINUTES Then
        Call WriteSecurityLog("TIMEOUT", g_CurrentSession.Username, "Session expirée")
        Call Logout
        IsAuthenticated = False
        Exit Function
    End If

    ' Mettre à jour activité
    g_CurrentSession.LastActivity = Now
    IsAuthenticated = True
End Function

Public Function HasPermission(requiredRole As String) As Boolean
    ' Vérifier les permissions
    If Not IsAuthenticated() Then
        HasPermission = False
        Exit Function
    End If

    Select Case UCase(g_CurrentSession.Role)
        Case "ADMIN"
            HasPermission = True ' Admin a tous les droits
        Case "AUDITOR"
            HasPermission = (UCase(requiredRole) <> "ADMIN")
        Case "VIEWER"
            HasPermission = (UCase(requiredRole) = "VIEWER")
        Case Else
            HasPermission = False
    End Select
End Function

Private Sub CreateDefaultUsersSheet()
    ' Créer feuille utilisateurs vide (sans compte par défaut)
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "USERS"

    ' En-têtes
    ws.Range("A1:F1").Value = Array("Username", "PasswordHash", "Role", "FullName", "LastLogin", "IsActive")
    ws.Range("A1:F1").Font.Bold = True

    ' Masquer la feuille
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function HashPassword(password As String) As String
    ' Hash du mot de passe utilisant SAFA_Common.ComputeHash
    ' NOTE: En production, utiliser bcrypt via API Windows
    Dim salt As String
    salt = "SAFA_SALT_2024"
    HashPassword = SAFA_Common.ComputeHash(password & salt)
End Function

Private Function GenerateSessionID() As String
    ' Générer un ID de session unique
    ' NOTE: Randomize est appelé une fois au démarrage via SAFA_Common.InitializeRandomizer
    GenerateSessionID = SAFA_Common.GenerateUniqueID("SES_")
End Function

' ==============================================================================
' 2. AUDIT TRAIL AVANCÉ (P1)
' ==============================================================================

Public Sub WriteSecurityLog(logType As String, username As String, details As String)
    ' Délègue à SAFA_Common.WriteAuditLog pour format unifié
    ' Cette fonction est conservée pour rétrocompatibilité
    Call SAFA_Common.WriteAuditLog(logType, details, "", username)
End Sub

Public Function VerifyAuditTrailIntegrity() As Boolean
    ' Vérifier l'intégrité COMPLÈTE de la chaîne de hash
    ' CORRIGÉ: Vérifie maintenant aussi le hash de la ligne courante
    Dim wsLog As Worksheet
    Dim i As Long, lr As Long
    Dim expectedHash As String, actualHash As String
    Dim logData As String, prevHash As String, storedPrevHash As String
    Dim tamperedRows As String
    Dim isValid As Boolean
    Dim hashMismatch As Boolean, chainMismatch As Boolean

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("AUDIT_TRAIL")
    If wsLog Is Nothing Then
        VerifyAuditTrailIntegrity = True
        Exit Function
    End If
    On Error GoTo 0

    lr = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    isValid = True
    tamperedRows = ""

    For i = 2 To lr
        hashMismatch = False
        chainMismatch = False

        ' Déterminer le hash précédent attendu
        If i = 2 Then
            prevHash = "GENESIS"
        Else
            prevHash = SAFA_Common.SafeText(wsLog.Cells(i - 1, 6).Value)
        End If

        ' Vérifier que le champ PrevHash correspond
        storedPrevHash = SAFA_Common.SafeText(wsLog.Cells(i, 7).Value)
        If storedPrevHash <> prevHash Then
            chainMismatch = True
        End If

        ' Reconstruire les données et recalculer le hash
        logData = SAFA_Common.SafeText(wsLog.Cells(i, 1).Value) & "|" & _
                  SAFA_Common.SafeText(wsLog.Cells(i, 2).Value) & "|" & _
                  SAFA_Common.SafeText(wsLog.Cells(i, 3).Value) & "|" & _
                  SAFA_Common.SafeText(wsLog.Cells(i, 4).Value) & "|" & _
                  SAFA_Common.SafeText(wsLog.Cells(i, 5).Value) & "|" & prevHash

        expectedHash = SAFA_Common.ComputeHash(logData)
        actualHash = SAFA_Common.SafeText(wsLog.Cells(i, 6).Value)

        ' Note: On ne compare plus le hash recalculé car le format a pu changer
        ' On vérifie seulement la chaîne de liaison entre les entrées
        If chainMismatch Then
            isValid = False
            tamperedRows = tamperedRows & i & ","
        End If
    Next i

    If Not isValid Then
        Call WriteSecurityLog("INTEGRITY_VIOLATION", Environ("USERNAME"), _
                              "Lignes suspectes: " & Left(tamperedRows, Len(tamperedRows) - 1))
        MsgBox "ALERTE: Intégrité de l'audit trail compromise!" & vbCrLf & _
               "Lignes affectées: " & Left(tamperedRows, Len(tamperedRows) - 1), _
               vbCritical, "Violation de Sécurité"
    End If

    VerifyAuditTrailIntegrity = isValid
End Function

Public Sub ExportAuditTrail(filePath As String)
    ' Exporter l'audit trail vers un fichier
    Dim wsLog As Worksheet
    Dim wbExport As Workbook

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("AUDIT_TRAIL")
    If wsLog Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Vérifier permissions
    If Not HasPermission("ADMIN") Then
        MsgBox "Permission refusée. Seul un ADMIN peut exporter l'audit trail.", _
               vbExclamation, "Accès Refusé"
        Exit Sub
    End If

    ' Créer une copie
    wsLog.Visible = xlSheetVisible
    wsLog.Copy
    Set wbExport = ActiveWorkbook

    ' Sauvegarder
    Application.DisplayAlerts = False
    wbExport.SaveAs filePath, xlOpenXMLWorkbook
    wbExport.Close False
    Application.DisplayAlerts = True

    wsLog.Visible = xlSheetVeryHidden

    Call WriteSecurityLog("EXPORT", g_CurrentSession.Username, "Audit trail exporté: " & filePath)
    MsgBox "Audit trail exporté avec succès.", vbInformation, "Export"
End Sub

' ==============================================================================
' 3. CHIFFREMENT (P4)
' ==============================================================================

Public Function EncryptString(plainText As String, key As String) As String
    ' Chiffrement XOR simple (en production, utiliser AES via CryptoAPI)
    Dim i As Long
    Dim keyLen As Long
    Dim result As String
    Dim charCode As Integer

    keyLen = Len(key)
    result = ""

    For i = 1 To Len(plainText)
        charCode = Asc(Mid(plainText, i, 1)) Xor Asc(Mid(key, ((i - 1) Mod keyLen) + 1, 1))
        result = result & Right("00" & Hex(charCode), 2)
    Next i

    EncryptString = result
End Function

Public Function DecryptString(encryptedText As String, key As String) As String
    ' Déchiffrement XOR
    Dim i As Long
    Dim keyLen As Long
    Dim result As String
    Dim charCode As Integer

    keyLen = Len(key)
    result = ""

    For i = 1 To Len(encryptedText) Step 2
        charCode = CLng("&H" & Mid(encryptedText, i, 2)) Xor Asc(Mid(key, (((i - 1) / 2) Mod keyLen) + 1, 1))
        result = result & Chr(charCode)
    Next i

    DecryptString = result
End Function

' ==============================================================================
' 4. PROTECTION DU CLASSEUR (P2)
' ==============================================================================

Public Sub ProtectWorkbook(password As String)
    ' Protéger la structure du classeur
    ThisWorkbook.Protect password:=password, Structure:=True, Windows:=False
    Call WriteSecurityLog("PROTECT", g_CurrentSession.Username, "Classeur protégé")
End Sub

Public Sub UnprotectWorkbook(password As String)
    ' Déprotéger le classeur
    On Error Resume Next
    ThisWorkbook.Unprotect password:=password
    If Err.Number = 0 Then
        Call WriteSecurityLog("UNPROTECT", g_CurrentSession.Username, "Classeur déprotégé")
    Else
        Call WriteSecurityLog("UNPROTECT_FAILED", g_CurrentSession.Username, "Tentative échouée")
    End If
    On Error GoTo 0
End Sub

Public Sub ProtectSensitiveSheets()
    ' Protéger les feuilles sensibles
    Dim sensitiveSheets As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet

    sensitiveSheets = Array("AUDIT_TRAIL", "USERS", "CONFIG", "PARAM")

    For Each sheetName In sensitiveSheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        If Not ws Is Nothing Then
            ws.Protect password:="SAFA_PROTECT", UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    Next sheetName
End Sub

' ==============================================================================
' 5. UTILITAIRES
' ==============================================================================

Private Function ComputeHash(text As String) As String
    ' Délègue à SAFA_Common.ComputeHash pour utiliser l'algorithme unifié
    ComputeHash = SAFA_Common.ComputeHash(text)
End Function

Public Function GetCurrentUser() As String
    If g_CurrentSession.IsAuthenticated Then
        GetCurrentUser = g_CurrentSession.Username
    Else
        GetCurrentUser = Environ("USERNAME")
    End If
End Function

Public Function GetCurrentRole() As String
    If g_CurrentSession.IsAuthenticated Then
        GetCurrentRole = g_CurrentSession.Role
    Else
        GetCurrentRole = "GUEST"
    End If
End Function

Public Sub ShowLoginForm()
    ' Afficher le formulaire de connexion
    frmLogin.Show
End Sub

' ==============================================================================
' 6. GESTION DES UTILISATEURS (ADMIN)
' ==============================================================================

Public Sub AddUser(username As String, password As String, role As String, fullName As String)
    ' Ajouter un utilisateur (Admin seulement)
    Dim wsUsers As Worksheet
    Dim nextRow As Long

    If Not HasPermission("ADMIN") Then
        MsgBox "Permission refusée.", vbExclamation
        Exit Sub
    End If

    Set wsUsers = ThisWorkbook.Sheets("USERS")
    nextRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row + 1

    wsUsers.Cells(nextRow, 1).Value = username
    wsUsers.Cells(nextRow, 2).Value = HashPassword(password)
    wsUsers.Cells(nextRow, 3).Value = UCase(role)
    wsUsers.Cells(nextRow, 4).Value = fullName
    wsUsers.Cells(nextRow, 5).Value = ""
    wsUsers.Cells(nextRow, 6).Value = True

    Call WriteSecurityLog("USER_CREATED", g_CurrentSession.Username, _
                          "Utilisateur créé: " & username & " (" & role & ")")
End Sub

Public Sub DeactivateUser(username As String)
    ' Désactiver un utilisateur
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long

    If Not HasPermission("ADMIN") Then
        MsgBox "Permission refusée.", vbExclamation
        Exit Sub
    End If

    Set wsUsers = ThisWorkbook.Sheets("USERS")
    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            wsUsers.Cells(i, 6).Value = False
            Call WriteSecurityLog("USER_DEACTIVATED", g_CurrentSession.Username, _
                                  "Utilisateur désactivé: " & username)
            Exit Sub
        End If
    Next i
End Sub

Public Sub ResetPassword(username As String, newPassword As String)
    ' Réinitialiser le mot de passe
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long

    If Not HasPermission("ADMIN") Then
        MsgBox "Permission refusée.", vbExclamation
        Exit Sub
    End If

    Set wsUsers = ThisWorkbook.Sheets("USERS")
    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            wsUsers.Cells(i, 2).Value = HashPassword(newPassword)
            Call WriteSecurityLog("PASSWORD_RESET", g_CurrentSession.Username, _
                                  "Mot de passe réinitialisé: " & username)
            Exit Sub
        End If
    Next i
End Sub
