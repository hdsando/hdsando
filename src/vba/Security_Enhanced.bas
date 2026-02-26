Attribute VB_Name = "Security_Enhanced"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: SECURITY_ENHANCED v10.0
' ==============================================================================
' Description: Module de sécurité renforcée
' Corrections implémentées:
'   - SEC-001: Hash PBKDF2-like avec itérations
'   - SEC-002: Salt unique par utilisateur
'   - SEC-003: Force changement mot de passe au premier login
'   - SEC-004: Persistance du compteur de tentatives
'   - SEC-005: Session ID avec composant cryptographique
'   - SEC-010/011: Chiffrement AES simulé (plus robuste que XOR)
' ==============================================================================

' --- CONSTANTES SÉCURITÉ ---
Private Const HASH_ITERATIONS As Long = 10000
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
    ' Génère un salt aléatoire de 16 caractères hexadécimaux
    Dim i As Integer
    Dim salt As String
    Dim charSet As String

    ' S'assurer que le générateur est initialisé
    Static initialized As Boolean
    If Not initialized Then
        Randomize Timer + CDbl(Now) * 86400
        initialized = True
    End If

    charSet = "0123456789ABCDEF"
    salt = ""

    For i = 1 To SALT_LENGTH
        salt = salt & Mid(charSet, Int(Rnd * 16) + 1, 1)
    Next i

    GenerateSecureSalt = salt
End Function

Public Function HashPasswordSecure(password As String, salt As String) As String
    ' Hash PBKDF2-like avec itérations multiples
    ' Plus résistant aux attaques par force brute que DJB2 simple
    Dim hash As String
    Dim i As Long
    Dim intermediate As String

    ' Validation des entrées
    If Len(password) = 0 Or Len(salt) = 0 Then
        HashPasswordSecure = ""
        Exit Function
    End If

    ' Première passe
    intermediate = password & salt

    ' Itérations pour renforcer le hash
    For i = 1 To HASH_ITERATIONS
        intermediate = SAFA_Common.ComputeHash(intermediate & salt & CStr(i))
    Next i

    ' Hash final
    HashPasswordSecure = SAFA_Common.ComputeHash(intermediate & "SAFA_FINAL")
End Function

Public Function VerifyPassword(password As String, storedHash As String, salt As String) As Boolean
    ' Vérifie si le mot de passe correspond au hash stocké
    Dim computedHash As String

    computedHash = HashPasswordSecure(password, salt)
    VerifyPassword = (computedHash = storedHash)
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
    ' Chiffrement par blocs avec substitution et permutation
    ' Plus robuste que XOR simple (bien que pas un vrai AES)
    Dim i As Long, j As Long
    Dim blockSize As Integer
    Dim keyExpanded As String
    Dim result As String
    Dim block As String
    Dim encryptedBlock As String

    If Len(plainText) = 0 Or Len(key) = 0 Then
        EncryptAESLike = ""
        Exit Function
    End If

    blockSize = 16

    ' Expansion de clé
    keyExpanded = ExpandKey(key, Len(plainText) + blockSize)

    ' Padding PKCS7
    Dim padLength As Integer
    padLength = blockSize - (Len(plainText) Mod blockSize)
    plainText = plainText & String(padLength, Chr(padLength))

    result = ""

    ' Chiffrement par blocs
    For i = 1 To Len(plainText) Step blockSize
        block = Mid(plainText, i, blockSize)
        encryptedBlock = ""

        For j = 1 To Len(block)
            Dim charCode As Integer
            Dim keyChar As Integer
            Dim encrypted As Integer

            charCode = Asc(Mid(block, j, 1))
            keyChar = Asc(Mid(keyExpanded, i + j - 1, 1))

            ' Substitution + XOR + rotation
            encrypted = ((charCode Xor keyChar) + j) Mod 256
            encrypted = ((encrypted * 7) + 13) Mod 256

            encryptedBlock = encryptedBlock & Right("00" & Hex(encrypted), 2)
        Next j

        result = result & encryptedBlock
    Next i

    ' Ajouter IV (vecteur d'initialisation) au début
    Dim iv As String
    iv = Left(SAFA_Common.ComputeHash(CStr(Now) & CStr(Timer)), 16)

    EncryptAESLike = iv & result
End Function

Public Function DecryptAESLike(encryptedText As String, key As String) As String
    ' Déchiffrement correspondant
    Dim i As Long, j As Long
    Dim blockSize As Integer
    Dim keyExpanded As String
    Dim result As String
    Dim iv As String
    Dim cipherText As String

    If Len(encryptedText) < 16 Then
        DecryptAESLike = ""
        Exit Function
    End If

    blockSize = 16

    ' Extraire IV
    iv = Left(encryptedText, 16)
    cipherText = Mid(encryptedText, 17)

    ' Expansion de clé
    keyExpanded = ExpandKey(key, Len(cipherText) / 2 + blockSize)

    result = ""

    ' Déchiffrement par blocs (2 caractères hex = 1 byte)
    Dim byteIndex As Long
    byteIndex = 1

    For i = 1 To Len(cipherText) Step (blockSize * 2)
        For j = 0 To blockSize - 1
            If i + j * 2 <= Len(cipherText) Then
                Dim hexByte As String
                Dim encrypted As Integer
                Dim keyChar As Integer
                Dim decrypted As Integer

                hexByte = Mid(cipherText, i + j * 2, 2)
                encrypted = CLng("&H" & hexByte)

                ' Rotation inverse
                decrypted = ((encrypted - 13) * 183) Mod 256  ' 183 est l'inverse de 7 mod 256
                If decrypted < 0 Then decrypted = decrypted + 256

                ' Soustraction position
                decrypted = (decrypted - (j + 1)) Mod 256
                If decrypted < 0 Then decrypted = decrypted + 256

                ' XOR avec clé
                keyChar = Asc(Mid(keyExpanded, byteIndex, 1))
                decrypted = decrypted Xor keyChar

                result = result & Chr(decrypted)
                byteIndex = byteIndex + 1
            End If
        Next j
    Next i

    ' Retirer padding PKCS7
    If Len(result) > 0 Then
        Dim lastChar As Integer
        lastChar = Asc(Right(result, 1))
        If lastChar > 0 And lastChar <= blockSize Then
            result = Left(result, Len(result) - lastChar)
        End If
    End If

    DecryptAESLike = result
End Function

Private Function ExpandKey(key As String, length As Long) As String
    ' Expansion de clé par hachage itératif
    Dim result As String
    Dim iteration As Long

    result = key
    iteration = 0

    Do While Len(result) < length
        iteration = iteration + 1
        result = result & SAFA_Common.ComputeHash(key & CStr(iteration))
    Loop

    ExpandKey = Left(result, length)
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

            ' Si pas de salt, utiliser l'ancien système (rétrocompatibilité)
            If salt = "" Then
                salt = "SAFA_SALT_2024"
                inputHash = SAFA_Common.ComputeHash(password & salt)
            Else
                inputHash = HashPasswordSecure(password, salt)
            End If

            ' Vérifier mot de passe
            If storedHash = inputHash Then
                ' Réinitialiser compteur
                Call ResetFailedAttempts(username)

                ' Mettre à jour dernière connexion
                wsUsers.Cells(i, 5).Value = Now

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
            ' Vérifier ancien mot de passe
            storedHash = SAFA_Common.SafeText(wsUsers.Cells(i, 2).Value)
            salt = SAFA_Common.SafeText(wsUsers.Cells(i, 7).Value)

            Dim oldHash As String
            If salt = "" Then
                salt = "SAFA_SALT_2024"
                oldHash = SAFA_Common.ComputeHash(oldPassword & salt)
            Else
                oldHash = HashPasswordSecure(oldPassword, salt)
            End If

            If oldHash <> storedHash Then
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

        If currentSalt = "" Then
            ' Utilisateur ancien système - marquer pour changement de mot de passe
            wsUsers.Cells(i, 7).Value = "LEGACY"  ' Marqueur spécial
            wsUsers.Cells(i, 8).Value = 0
            wsUsers.Cells(i, 10).Value = True     ' Forcer changement
        End If
    Next i

    wsUsers.Columns("A:L").AutoFit

    SAFA_Common.WriteAuditLog "SYSTEM", "Migration sécurité utilisateurs effectuée", ""
    MsgBox "Migration de sécurité terminée." & vbCrLf & _
           "Les utilisateurs existants devront changer leur mot de passe.", _
           vbInformation, "Migration Sécurité"
End Sub

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
