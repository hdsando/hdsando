Attribute VB_Name = "Security_Enhanced"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SECURITY_ENHANCED v10.0
' ==============================================================================
' Description: Module de sécurité renforcée
' Corrections implémentées:
'   - SEC-001: Hash de mot de passe PBKDF2-HMAC-SHA256 reel (Crypto_Provider / .NET)
'   - SEC-002: Salt unique par utilisateur (RNG cryptographique)
'   - SEC-003: Force changement mot de passe au premier login
'   - SEC-004: Persistance du compteur de tentatives
'   - SEC-005: Session ID avec composant cryptographique
'   - SEC-010/011: Chiffrement AES-256-CBC reel (Crypto_Provider)
' Format des hash stockes: "pbkdf2$<iterations>$<salt>$<hex>" (ou "legacy$..." sans .NET)
' Les hash produits par les versions anterieures ne sont plus verifiables:
' recreer la feuille USERS (CreateExtendedUsersSheet) ou EnsureDefaultAdmin.
' ==============================================================================

' --- CONSTANTES SÉCURITÉ ---
Private Const HASH_ITERATIONS As Long = 10000   ' iterations du repli legacy (sans .NET)
Private Const SALT_LENGTH As Integer = 16
Private Const MIN_PASSWORD_LENGTH As Integer = 8
Private Const PASSWORD_EXPIRY_DAYS As Integer = 90
Private Const LOCKOUT_DURATION_MINUTES As Integer = 30

' --- COLONNES USERS ÉTENDUES ---
' A=Username, B=PasswordHash, C=Role, D=FullName, E=LastLogin, F=IsActive
' G=Salt, H=FailedAttempts, I=LockoutUntil, J=MustChangePassword, K=PasswordCreated, L=Email

' ==============================================================================
' 1. HASH SÉCURISÉ AVEC SALT (SEC-001, SEC-002)
' ==============================================================================

Public Function GenerateSecureSalt() As String
    ' Salt aleatoire (RNG cryptographique .NET si disponible): SALT_LENGTH octets en hexadecimal
    GenerateSecureSalt = Crypto_Provider.GenerateSaltHex(SALT_LENGTH)
End Function

Public Function HashPasswordSecure(password As String, salt As String) As String
    ' Hash de mot de passe auto-descriptif:
    '   pbkdf2$<iterations>$<salt>$<hex>   (PBKDF2-HMAC-SHA256 reel)
    '   legacy$<iterations>$<salt>$<hex>   (repli sans .NET, non cryptographique)
    If Len(password) = 0 Or Len(salt) = 0 Then
        HashPasswordSecure = ""
        Exit Function
    End If

    If Crypto_Provider.IsAvailable() Then
        HashPasswordSecure = "pbkdf2$" & Crypto_Provider.PBKDF2_ITERATIONS & "$" & salt & "$" & _
                             Crypto_Provider.Pbkdf2Sha256Hex(password, salt, Crypto_Provider.PBKDF2_ITERATIONS, 32)
    Else
        HashPasswordSecure = "legacy$" & HASH_ITERATIONS & "$" & salt & "$" & _
                             Crypto_Provider.Pbkdf2Sha256Hex(password, salt, HASH_ITERATIONS, 32)
    End If
End Function

Public Function VerifyPassword(password As String, storedHash As String, salt As String) As Boolean
    ' Verifie un mot de passe quel que soit le format stocke (pbkdf2$ / legacy$ / ancien hex brut)
    Dim parts() As String
    Dim iters As Long, storedSalt As String, digest As String

    VerifyPassword = False
    If Len(password) = 0 Or Len(storedHash) = 0 Then Exit Function

    If Left(storedHash, 7) = "pbkdf2$" Or Left(storedHash, 7) = "legacy$" Then
        parts = Split(storedHash, "$")
        If UBound(parts) < 3 Then Exit Function
        iters = SAFA_Common.SafeLong(parts(1))
        storedSalt = parts(2)
        digest = parts(3)
        If Left(storedHash, 7) = "pbkdf2$" And Not Crypto_Provider.IsAvailable() Then Exit Function
        VerifyPassword = (Crypto_Provider.Pbkdf2Sha256Hex(password, storedSalt, iters, 32) = digest)
    Else
        ' Ancien format (hex brut, algorithme des versions anterieures): non verifiable de facon sure
        VerifyPassword = False
    End If
End Function

Public Function IsHashUpToDate(storedHash As String) As Boolean
    ' True si le hash est deja au format PBKDF2 reel
    IsHashUpToDate = (Left(storedHash, 7) = "pbkdf2$")
End Function

' ==============================================================================
' 2. VALIDATION MOT DE PASSE (SEC-003)
' ==============================================================================

Public Function ValidatePasswordStrength(password As String) As String
    ' Retourne chaîne vide si OK, sinon message d'erreur
    Dim errors As String
    errors = ""

    ' Longueur minimale
    If Len(password) < MIN_PASSWORD_LENGTH Then
        errors = errors & "- Minimum " & MIN_PASSWORD_LENGTH & " caractères" & vbCrLf
    End If

    ' Au moins une majuscule
    If Not password Like "*[A-Z]*" Then
        errors = errors & "- Au moins une majuscule requise" & vbCrLf
    End If

    ' Au moins une minuscule
    If Not password Like "*[a-z]*" Then
        errors = errors & "- Au moins une minuscule requise" & vbCrLf
    End If

    ' Au moins un chiffre
    If Not password Like "*[0-9]*" Then
        errors = errors & "- Au moins un chiffre requis" & vbCrLf
    End If

    ' Au moins un caractère spécial
    If Not ContainsSpecialChar(password) Then
        errors = errors & "- Au moins un caractère spécial requis (!@#$%^&*)" & vbCrLf
    End If

    ' Pas de séquences triviales
    If InStr(1, password, "123", vbTextCompare) > 0 Or _
       InStr(1, password, "abc", vbTextCompare) > 0 Or _
       InStr(1, password, "password", vbTextCompare) > 0 Or _
       InStr(1, password, "admin", vbTextCompare) > 0 Then
        errors = errors & "- Séquences triviales interdites (123, abc, password, admin)" & vbCrLf
    End If

    ValidatePasswordStrength = errors
End Function

Private Function ContainsSpecialChar(s As String) As Boolean
    Dim specialChars As String
    Dim i As Integer

    specialChars = "!@#$%^&*()_+-=[]{}|;':"",./<>?"

    For i = 1 To Len(specialChars)
        If InStr(s, Mid(specialChars, i, 1)) > 0 Then
            ContainsSpecialChar = True
            Exit Function
        End If
    Next i

    ContainsSpecialChar = False
End Function

Public Function IsPasswordExpired(username As String) As Boolean
    ' Vérifie si le mot de passe a expiré
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim passwordCreated As Date

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        IsPasswordExpired = False
        Exit Function
    End If
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            passwordCreated = SAFA_Common.SafeDate(wsUsers.Cells(i, 11).Value)

            If passwordCreated = 0 Then
                ' Pas de date = mot de passe jamais changé = expiré
                IsPasswordExpired = True
            ElseIf DateDiff("d", passwordCreated, Now) > PASSWORD_EXPIRY_DAYS Then
                IsPasswordExpired = True
            Else
                IsPasswordExpired = False
            End If
            Exit Function
        End If
    Next i

    IsPasswordExpired = False
End Function

Public Function MustChangePassword(username As String) As Boolean
    ' Vérifie si l'utilisateur doit changer son mot de passe
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        MustChangePassword = False
        Exit Function
    End If
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            MustChangePassword = (wsUsers.Cells(i, 10).Value = True)
            Exit Function
        End If
    Next i

    MustChangePassword = False
End Function

' ==============================================================================
' 3. GESTION DES TENTATIVES ÉCHOUÉES (SEC-004)
' ==============================================================================

Public Function IsAccountLocked(username As String) As Boolean
    ' Vérifie si le compte est verrouillé
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim lockoutUntil As Date

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        IsAccountLocked = False
        Exit Function
    End If
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            lockoutUntil = SAFA_Common.SafeDate(wsUsers.Cells(i, 9).Value)

            If lockoutUntil > Now Then
                IsAccountLocked = True
            Else
                ' Débloquer si le délai est passé
                wsUsers.Cells(i, 9).Value = ""
                wsUsers.Cells(i, 8).Value = 0
                IsAccountLocked = False
            End If
            Exit Function
        End If
    Next i

    IsAccountLocked = False
End Function

Public Sub IncrementFailedAttempts(username As String)
    ' Incrémente le compteur de tentatives échouées (persistant)
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim currentAttempts As Long

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            currentAttempts = SAFA_Common.SafeLong(wsUsers.Cells(i, 8).Value) + 1
            wsUsers.Cells(i, 8).Value = currentAttempts

            ' Verrouiller après 3 tentatives
            If currentAttempts >= SAFA_Common.MAX_LOGIN_ATTEMPTS Then
                wsUsers.Cells(i, 9).Value = DateAdd("n", LOCKOUT_DURATION_MINUTES, Now)
                SAFA_Common.WriteAuditLog "SECURITY", "Compte verrouillé: " & username, _
                                          "Tentatives: " & currentAttempts, username
            End If
            Exit Sub
        End If
    Next i
End Sub

Public Sub ResetFailedAttempts(username As String)
    ' Réinitialise le compteur après connexion réussie
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then Exit Sub
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            wsUsers.Cells(i, 8).Value = 0
            wsUsers.Cells(i, 9).Value = ""
            Exit Sub
        End If
    Next i
End Sub

' ==============================================================================
' 4. SESSION ID SÉCURISÉ (SEC-005)
' ==============================================================================

Public Function GenerateSecureSessionID() As String
    ' Génère un ID de session avec composant cryptographique
    Dim timestamp As String
    Dim random1 As String
    Dim random2 As String
    Dim machineHash As String
    Dim combined As String

    ' Timestamp haute précision
    timestamp = Format(Now, "yyyymmddhhnnss") & Format(Timer * 1000, "000")

    ' Composants aléatoires
    random1 = GenerateSecureSalt()
    random2 = Right(GenerateSecureSalt(), 8)

    ' Hash machine (obscurcit l'identité)
    machineHash = Left(SAFA_Common.ComputeHash(Environ("COMPUTERNAME") & Environ("USERNAME")), 4)

    ' Combinaison
    combined = "SES_" & timestamp & "_" & random1 & "_" & random2 & "_" & machineHash

    ' Hash final pour uniformité
    GenerateSecureSessionID = "SES_" & SAFA_Common.ComputeHash(combined)
End Function

' ==============================================================================
' 5. CHIFFREMENT RENFORCÉ (SEC-010, SEC-011)
' ==============================================================================

Public Function EncryptAESLike(plainText As String, key As String) As String
    ' AES-256-CBC reel via Crypto_Provider (nom conserve pour compatibilite des appelants).
    ' Sortie prefixee "aes$" (ou "legacy$" si .NET indisponible).
    If Len(plainText) = 0 Or Len(key) = 0 Then
        EncryptAESLike = ""
        Exit Function
    End If
    EncryptAESLike = Crypto_Provider.AesEncrypt(plainText, key)
End Function

Public Function DecryptAESLike(encryptedText As String, key As String) As String
    ' Dechiffrement: route selon le prefixe ("aes$", "legacy$" ou ancien schema sans prefixe)
    If Len(encryptedText) = 0 Or Len(key) = 0 Then
        DecryptAESLike = ""
        Exit Function
    End If
    DecryptAESLike = Crypto_Provider.AesDecrypt(encryptedText, key)
End Function

Public Function IsRealCryptoAvailable() As Boolean
    ' Permet a l'interface d'avertir si la crypto reelle (.NET) est absente
    IsRealCryptoAvailable = Crypto_Provider.IsAvailable()
End Function

' ==============================================================================
' 6. AUTHENTIFICATION SÉCURISÉE AMÉLIORÉE
' ==============================================================================

Public Function AuthenticateSecure(username As String, password As String) As Integer
    ' Authentification avec toutes les vérifications de sécurité
    ' Retourne:
    '   0 = Échec
    '   1 = Succès
    '   2 = Succès mais doit changer mot de passe
    '   3 = Compte verrouillé
    '   4 = Compte inactif
    '   5 = Mot de passe expiré

    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim storedHash As String, salt As String, inputHash As String
    Dim isActive As Boolean

    ' Vérifier verrouillage
    If IsAccountLocked(username) Then
        SAFA_Common.WriteAuditLog "LOGIN_BLOCKED", "Tentative sur compte verrouillé", "", username
        AuthenticateSecure = 3
        Exit Function
    End If

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        Call CreateExtendedUsersSheet
        Set wsUsers = ThisWorkbook.Sheets("USERS")
    End If
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            ' Vérifier si actif
            isActive = (wsUsers.Cells(i, 6).Value = True)
            If Not isActive Then
                SAFA_Common.WriteAuditLog "LOGIN_FAILED", "Compte inactif", "", username
                AuthenticateSecure = 4
                Exit Function
            End If

            ' Récupérer hash et salt
            storedHash = SAFA_Common.SafeText(wsUsers.Cells(i, 2).Value)
            salt = SAFA_Common.SafeText(wsUsers.Cells(i, 7).Value)

            ' Vérifier mot de passe (tous formats: pbkdf2$ / legacy$)
            If VerifyPassword(password, storedHash, salt) Then
                ' Réinitialiser compteur
                Call ResetFailedAttempts(username)

                ' Mettre à jour dernière connexion
                wsUsers.Cells(i, 5).Value = Now

                ' Mise a niveau transparente vers PBKDF2 reel si le hash stocke est en format legacy
                If Not IsHashUpToDate(storedHash) And Crypto_Provider.IsAvailable() Then
                    Dim upSalt As String
                    upSalt = GenerateSecureSalt()
                    wsUsers.Cells(i, 2).Value = HashPasswordSecure(password, upSalt)
                    wsUsers.Cells(i, 7).Value = upSalt
                    SAFA_Common.WriteAuditLog "SECURITY", "Hash mot de passe mis a niveau (PBKDF2)", "", username
                End If

                SAFA_Common.WriteAuditLog "LOGIN_SUCCESS", "Connexion réussie", "", username

                ' Vérifier si doit changer mot de passe
                If MustChangePassword(username) Then
                    AuthenticateSecure = 2
                    Exit Function
                End If

                ' Vérifier expiration
                If IsPasswordExpired(username) Then
                    AuthenticateSecure = 5
                    Exit Function
                End If

                AuthenticateSecure = 1
                Exit Function
            Else
                ' Échec
                Call IncrementFailedAttempts(username)
                SAFA_Common.WriteAuditLog "LOGIN_FAILED", "Mot de passe incorrect", "", username
                AuthenticateSecure = 0
                Exit Function
            End If
        End If
    Next i

    ' Utilisateur non trouvé
    SAFA_Common.WriteAuditLog "LOGIN_FAILED", "Utilisateur inconnu", "", username
    AuthenticateSecure = 0
End Function

' ==============================================================================
' 7. CHANGEMENT DE MOT DE PASSE SÉCURISÉ
' ==============================================================================

Public Function ChangePasswordSecure(username As String, oldPassword As String, newPassword As String) As String
    ' Change le mot de passe avec validation
    ' Retourne chaîne vide si OK, sinon message d'erreur

    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim storedHash As String, salt As String
    Dim newSalt As String, newHash As String
    Dim validation As String

    ' Valider force du nouveau mot de passe
    validation = ValidatePasswordStrength(newPassword)
    If validation <> "" Then
        ChangePasswordSecure = "Nouveau mot de passe trop faible:" & vbCrLf & validation
        Exit Function
    End If

    ' Vérifier que nouveau != ancien
    If oldPassword = newPassword Then
        ChangePasswordSecure = "Le nouveau mot de passe doit être différent de l'ancien"
        Exit Function
    End If

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        ChangePasswordSecure = "Erreur système: feuille USERS introuvable"
        Exit Function
    End If
    On Error GoTo 0

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        If UCase(wsUsers.Cells(i, 1).Value) = UCase(username) Then
            ' Vérifier ancien mot de passe (tous formats)
            storedHash = SAFA_Common.SafeText(wsUsers.Cells(i, 2).Value)
            salt = SAFA_Common.SafeText(wsUsers.Cells(i, 7).Value)

            If Not VerifyPassword(oldPassword, storedHash, salt) Then
                ChangePasswordSecure = "Ancien mot de passe incorrect"
                Exit Function
            End If

            ' Générer nouveau salt et hash
            newSalt = GenerateSecureSalt()
            newHash = HashPasswordSecure(newPassword, newSalt)

            ' Mettre à jour
            wsUsers.Cells(i, 2).Value = newHash
            wsUsers.Cells(i, 7).Value = newSalt
            wsUsers.Cells(i, 10).Value = False  ' MustChangePassword = False
            wsUsers.Cells(i, 11).Value = Now    ' PasswordCreated

            SAFA_Common.WriteAuditLog "PASSWORD_CHANGED", "Mot de passe changé", "", username

            ChangePasswordSecure = ""
            Exit Function
        End If
    Next i

    ChangePasswordSecure = "Utilisateur non trouvé"
End Function

' ==============================================================================
' 8. CRÉATION FEUILLE USERS ÉTENDUE
' ==============================================================================

Public Sub CreateExtendedUsersSheet()
    ' Crée la feuille USERS avec colonnes de sécurité étendues
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("USERS").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "USERS"

    ' En-têtes étendues
    ws.Range("A1:L1").Value = Array("Username", "PasswordHash", "Role", "FullName", _
                                     "LastLogin", "IsActive", "Salt", "FailedAttempts", _
                                     "LockoutUntil", "MustChangePassword", "PasswordCreated", "Email")
    ws.Range("A1:L1").Font.Bold = True
    SAFA_Common.FormatHeader ws.Range("A1:L1")

    ' Admin par défaut avec nouveau système de sécurité
    Dim adminSalt As String
    Dim adminHash As String

    adminSalt = GenerateSecureSalt()
    adminHash = HashPasswordSecure("Admin@2024!", adminSalt)

    ws.Range("A2").Value = "admin"
    ws.Range("B2").Value = adminHash
    ws.Range("C2").Value = "ADMIN"
    ws.Range("D2").Value = "Administrateur Système"
    ws.Range("E2").Value = Now
    ws.Range("F2").Value = True
    ws.Range("G2").Value = adminSalt
    ws.Range("H2").Value = 0
    ws.Range("I2").Value = ""
    ws.Range("J2").Value = True  ' Doit changer au premier login
    ws.Range("K2").Value = Now
    ws.Range("L2").Value = ""

    ws.Columns("A:L").AutoFit
    ws.Visible = xlSheetVeryHidden

    SAFA_Common.WriteAuditLog "SYSTEM", "Feuille USERS créée avec sécurité renforcée", ""
End Sub

Public Sub EnsureDefaultAdmin()
    ' Garantit l'existence d'un compte admin avec un hash valide (PBKDF2 reel si .NET)
    ' Mot de passe initial: Admin@2024! - changement obligatoire a la premiere connexion
    Dim ws As Worksheet
    Dim lr As Long, i As Long, found As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("USERS")
    On Error GoTo 0

    If ws Is Nothing Then
        Call CreateExtendedUsersSheet
        Exit Sub
    End If

    ' Feuille presente mais layout ancien -> migrer les colonnes
    If SAFA_Common.SafeText(ws.Cells(1, 7).Value) <> "Salt" Then Call MigrateUsersToEnhancedSecurity

    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If UCase(SAFA_Common.SafeText(ws.Cells(i, 1).Value)) = "ADMIN" Then
            found = True
            ' Hash absent ou dans un format non verifiable -> reinitialiser
            Dim h As String
            h = SAFA_Common.SafeText(ws.Cells(i, 2).Value)
            If h = "" Or (Left(h, 7) <> "pbkdf2$" And Left(h, 7) <> "legacy$") Then
                Dim s As String
                s = GenerateSecureSalt()
                ws.Cells(i, 2).Value = HashPasswordSecure("Admin@2024!", s)
                ws.Cells(i, 7).Value = s
                ws.Cells(i, 6).Value = True
                ws.Cells(i, 8).Value = 0
                ws.Cells(i, 9).Value = ""
                ws.Cells(i, 10).Value = True
                ws.Cells(i, 11).Value = Now
                SAFA_Common.WriteAuditLog "SECURITY", "Compte admin reinitialise (format de hash invalide)", ""
            End If
            Exit For
        End If
    Next i

    If Not found Then
        Dim r As Long, adminSalt As String
        r = lr + 1
        adminSalt = GenerateSecureSalt()
        ws.Cells(r, 1).Value = "admin"
        ws.Cells(r, 2).Value = HashPasswordSecure("Admin@2024!", adminSalt)
        ws.Cells(r, 3).Value = "ADMIN"
        ws.Cells(r, 4).Value = "Administrateur Systeme"
        ws.Cells(r, 5).Value = Now
        ws.Cells(r, 6).Value = True
        ws.Cells(r, 7).Value = adminSalt
        ws.Cells(r, 8).Value = 0
        ws.Cells(r, 9).Value = ""
        ws.Cells(r, 10).Value = True
        ws.Cells(r, 11).Value = Now
        ws.Cells(r, 12).Value = ""
        SAFA_Common.WriteAuditLog "SECURITY", "Compte admin cree", ""
    End If
End Sub

' ==============================================================================
' 9. MIGRATION DES UTILISATEURS EXISTANTS
' ==============================================================================

Public Sub MigrateUsersToEnhancedSecurity()
    ' Migre les utilisateurs existants vers le nouveau système de sécurité
    Dim wsUsers As Worksheet
    Dim i As Long, lr As Long
    Dim currentSalt As String

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    If wsUsers Is Nothing Then
        Call CreateExtendedUsersSheet
        Exit Sub
    End If
    On Error GoTo 0

    ' Vérifier si migration nécessaire
    If wsUsers.Cells(1, 7).Value <> "Salt" Then
        ' Ajouter colonnes manquantes
        wsUsers.Cells(1, 7).Value = "Salt"
        wsUsers.Cells(1, 8).Value = "FailedAttempts"
        wsUsers.Cells(1, 9).Value = "LockoutUntil"
        wsUsers.Cells(1, 10).Value = "MustChangePassword"
        wsUsers.Cells(1, 11).Value = "PasswordCreated"
        wsUsers.Cells(1, 12).Value = "Email"
    End If

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lr
        currentSalt = SAFA_Common.SafeText(wsUsers.Cells(i, 7).Value)

        If currentSalt = "" Or currentSalt = "LEGACY" Then
            ' Utilisateur ancien systeme: son hash n'est plus verifiable -> l'admin doit
            ' redefinir le mot de passe (AdminResetPassword). On marque le compte.
            wsUsers.Cells(i, 7).Value = "LEGACY"
            wsUsers.Cells(i, 8).Value = 0
            wsUsers.Cells(i, 10).Value = True     ' Forcer changement
        End If
    Next i

    wsUsers.Columns("A:L").AutoFit

    SAFA_Common.WriteAuditLog "SYSTEM", "Migration sécurité utilisateurs effectuée", ""
End Sub

Public Function AdminResetPassword(username As String, newPassword As String) As String
    ' Reinitialisation par un administrateur (sans ancien mot de passe). Retourne "" si OK.
    Dim wsUsers As Worksheet, i As Long, lr As Long, validation As String, s As String

    validation = ValidatePasswordStrength(newPassword)
    If validation <> "" Then
        AdminResetPassword = "Mot de passe trop faible:" & vbCrLf & validation
        Exit Function
    End If

    On Error Resume Next
    Set wsUsers = ThisWorkbook.Sheets("USERS")
    On Error GoTo 0
    If wsUsers Is Nothing Then AdminResetPassword = "Feuille USERS introuvable": Exit Function

    lr = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If UCase(SAFA_Common.SafeText(wsUsers.Cells(i, 1).Value)) = UCase(username) Then
            s = GenerateSecureSalt()
            wsUsers.Cells(i, 2).Value = HashPasswordSecure(newPassword, s)
            wsUsers.Cells(i, 7).Value = s
            wsUsers.Cells(i, 8).Value = 0
            wsUsers.Cells(i, 9).Value = ""
            wsUsers.Cells(i, 10).Value = True
            wsUsers.Cells(i, 11).Value = Now
            SAFA_Common.WriteAuditLog "SECURITY", "Mot de passe reinitialise par admin", "", username
            AdminResetPassword = ""
            Exit Function
        End If
    Next i
    AdminResetPassword = "Utilisateur non trouve"
End Function

' ==============================================================================
' 10. VALIDATION DES ENTRÉES
' ==============================================================================

Public Function SanitizeInput(input As String, Optional maxLength As Long = 255) As String
    ' Nettoie et valide une entrée utilisateur
    Dim result As String
    Dim i As Long
    Dim c As String

    result = ""

    For i = 1 To Len(input)
        c = Mid(input, i, 1)

        ' Supprimer caractères de contrôle
        If Asc(c) >= 32 And Asc(c) < 127 Then
            ' Échapper caractères dangereux pour SQL/scripts
            Select Case c
                Case "'", """", "\", "<", ">", "&"
                    ' Remplacer par espace ou ignorer
                    result = result & " "
                Case Else
                    result = result & c
            End Select
        End If
    Next i

    ' Limiter longueur
    If Len(result) > maxLength Then
        result = Left(result, maxLength)
    End If

    SanitizeInput = Trim(result)
End Function

Public Function ValidateAccountNumber(acctNum As String) As Boolean
    ' Valide un numéro de compte (format alphanumérique)
    Dim i As Long
    Dim c As String

    If Len(acctNum) < 3 Or Len(acctNum) > 50 Then
        ValidateAccountNumber = False
        Exit Function
    End If

    For i = 1 To Len(acctNum)
        c = Mid(acctNum, i, 1)
        If Not (c Like "[A-Za-z0-9-_]") Then
            ValidateAccountNumber = False
            Exit Function
        End If
    Next i

    ValidateAccountNumber = True
End Function

Public Function ValidateAmount(amount As Variant) As Boolean
    ' Valide qu'un montant est numérique et raisonnable
    Dim numAmount As Double

    On Error Resume Next
    numAmount = CDbl(amount)
    If Err.Number <> 0 Then
        ValidateAmount = False
        Exit Function
    End If
    On Error GoTo 0

    ' Limites raisonnables
    If Abs(numAmount) > 999999999999# Then  ' 1 trillion
        ValidateAmount = False
    Else
        ValidateAmount = True
    End If
End Function
