Attribute VB_Name = "Crypto_Provider"
Option Explicit

' ==============================================================================
' S.A.F.A - SYSTEM FOR AUTOMATED FINANCIAL AUDIT
' MODULE: CRYPTO_PROVIDER v10.0
' ==============================================================================
' Description: Cryptographie REELLE via les classes .NET exposees en COM
'   (disponibles sur tout Windows avec Excel, .NET Framework 4.x):
'     - SHA-256          : System.Security.Cryptography.SHA256Managed
'     - HMAC-SHA256      : System.Security.Cryptography.HMACSHA256
'     - PBKDF2-HMAC-SHA256 (RFC 8018) implemente sur le HMAC reel
'     - AES-256-CBC/PKCS7: System.Security.Cryptography.RijndaelManaged
'     - Aleatoire crypto : System.Security.Cryptography.RNGCryptoServiceProvider
'   Si .NET est indisponible (Mac, environnement bride), des fonctions de
'   repli "Legacy_*" (non cryptographiques) sont utilisees et IsAvailable()
'   renvoie False pour que les appelants puissent avertir l'utilisateur.
'   Chaque valeur stockee porte un prefixe indiquant l'algorithme utilise.
' ==============================================================================

Public Const PBKDF2_ITERATIONS As Long = 25000   ' ~1-2 s en VBA/COM par verification
Private Const AES_KDF_SALT As String = "SAFA.v10.AES.KDF.SALT"   ' salt fixe de derivation de cle (documente)
Private Const AES_KDF_ITER As Long = 5000

Private mChecked As Boolean
Private mAvailable As Boolean

' ==============================================================================
' DISPONIBILITE
' ==============================================================================

Public Function IsAvailable() As Boolean
    If Not mChecked Then
        Dim o As Object
        On Error Resume Next
        Set o = CreateObject("System.Security.Cryptography.SHA256Managed")
        mAvailable = (Err.Number = 0)
        If mAvailable Then mAvailable = Not (o Is Nothing)
        Err.Clear
        On Error GoTo 0
        mChecked = True
    End If
    IsAvailable = mAvailable
End Function

Public Function ProviderName() As String
    ' Nom de l'algorithme de hash effectivement utilise (enregistre dans AUDIT_TRAIL col. H)
    ProviderName = IIf(IsAvailable(), "sha256", "legacy")
End Function

Public Function HashHexWith(algo As String, text As String) As String
    ' Recalcule un hash avec un algorithme donne (pour verifier un journal ecrit ailleurs).
    ' Renvoie "" si l'algorithme demande n'est pas disponible sur ce poste.
    Select Case LCase(Trim(algo))
        Case "sha256"
            If IsAvailable() Then HashHexWith = Sha256Hex(text) Else HashHexWith = ""
        Case "legacy"
            HashHexWith = Legacy_Hash(text)
        Case Else
            HashHexWith = ""
    End Select
End Function

' ==============================================================================
' HASH
' ==============================================================================

' Hash generique utilise par S.A.F.A (journal d'audit, integrite): SHA-256 si possible
Public Function HashHex(text As String) As String
    If IsAvailable() Then
        HashHex = Sha256Hex(text)
    Else
        HashHex = Legacy_Hash(text)
    End If
End Function

Public Function Sha256Hex(text As String) As String
    On Error GoTo Fallback
    If Not IsAvailable() Then GoTo Fallback
    Dim sha As Object, b() As Byte, h() As Byte
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    b = Utf8Bytes(text)
    h = sha.ComputeHash_2(b)
    Sha256Hex = BytesToHex(h)
    Exit Function
Fallback:
    Sha256Hex = Legacy_Hash(text)
End Function

Public Function HmacSha256Hex(key As String, text As String) As String
    On Error GoTo Fallback
    If Not IsAvailable() Then GoTo Fallback
    Dim k() As Byte, d() As Byte
    k = Utf8Bytes(key): d = Utf8Bytes(text)
    HmacSha256Hex = BytesToHex(HmacSha256Bytes(k, d))
    Exit Function
Fallback:
    HmacSha256Hex = Legacy_Hash(key & "|" & text)
End Function

Private Function HmacSha256Bytes(key() As Byte, data() As Byte) As Byte()
    Dim h As Object
    Set h = CreateObject("System.Security.Cryptography.HMACSHA256")
    h.Key = key
    HmacSha256Bytes = h.ComputeHash_2(data)
End Function

' PBKDF2-HMAC-SHA256 (RFC 8018 section 5.2) construit sur le HMAC .NET reel
Public Function Pbkdf2Sha256Hex(password As String, salt As String, iterations As Long, dkLenBytes As Long) As String
    On Error GoTo Fallback
    If Not IsAvailable() Then GoTo Fallback
    If iterations < 1 Then iterations = 1
    If dkLenBytes < 1 Then dkLenBytes = 32

    Dim pw() As Byte, sb() As Byte
    pw = Utf8Bytes(password): sb = Utf8Bytes(salt)

    Dim hmac As Object
    Set hmac = CreateObject("System.Security.Cryptography.HMACSHA256")
    hmac.Key = pw

    Const HLEN As Long = 32
    Dim nBlocks As Long, saltLen As Long
    nBlocks = (dkLenBytes + HLEN - 1) \ HLEN
    saltLen = ArrLen(sb)

    Dim outB() As Byte
    ReDim outB(0 To nBlocks * HLEN - 1)

    Dim blk As Long, i As Long, j As Long
    Dim u() As Byte, t() As Byte, inp() As Byte

    For blk = 1 To nBlocks
        ReDim inp(0 To saltLen + 3)
        For j = 0 To saltLen - 1
            inp(j) = sb(j)
        Next j
        inp(saltLen) = (blk \ &H1000000) And &HFF
        inp(saltLen + 1) = (blk \ &H10000) And &HFF
        inp(saltLen + 2) = (blk \ &H100) And &HFF
        inp(saltLen + 3) = blk And &HFF

        u = hmac.ComputeHash_2(inp)
        t = u
        For i = 2 To iterations
            u = hmac.ComputeHash_2(u)
            For j = 0 To HLEN - 1
                t(j) = t(j) Xor u(j)
            Next j
        Next i
        For j = 0 To HLEN - 1
            outB((blk - 1) * HLEN + j) = t(j)
        Next j
    Next blk

    ReDim Preserve outB(0 To dkLenBytes - 1)
    Pbkdf2Sha256Hex = BytesToHex(outB)
    Exit Function
Fallback:
    Pbkdf2Sha256Hex = Legacy_Pbkdf(password, salt, iterations)
End Function

' ==============================================================================
' ALEATOIRE
' ==============================================================================

Public Function RandomBytes(nBytes As Long) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To nBytes - 1)
    On Error GoTo Fallback
    If Not IsAvailable() Then GoTo Fallback
    Dim rng As Object
    Set rng = CreateObject("System.Security.Cryptography.RNGCryptoServiceProvider")
    rng.GetBytes b
    RandomBytes = b
    Exit Function
Fallback:
    ' Repli NON cryptographique (signale par IsAvailable() = False)
    Randomize Timer + CDbl(Now) * 86400
    For i = 0 To nBytes - 1
        b(i) = Int(Rnd * 256)
    Next i
    RandomBytes = b
End Function

Public Function GenerateSaltHex(nBytes As Long) As String
    GenerateSaltHex = UCase(BytesToHex(RandomBytes(nBytes)))
End Function

Public Function GenerateSaltBase64(nBytes As Long) As String
    GenerateSaltBase64 = Base64Encode(RandomBytes(nBytes))
End Function

' ==============================================================================
' AES-256-CBC (cle derivee de la passphrase par PBKDF2, IV aleatoire par message)
' Format de sortie: "aes$" & base64(IV(16) || ciphertext)
' ==============================================================================

Public Function AesEncrypt(plainText As String, passphrase As String) As String
    On Error GoTo Fallback
    If Not IsAvailable() Then GoTo Fallback
    If Len(plainText) = 0 Then AesEncrypt = "": Exit Function

    Dim key() As Byte, iv() As Byte, data() As Byte, enc() As Byte
    key = HexToBytes(Pbkdf2Sha256Hex(passphrase, AES_KDF_SALT, AES_KDF_ITER, 32))
    iv = RandomBytes(16)
    data = Utf8Bytes(plainText)

    Dim aes As Object, tr As Object
    Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")
    aes.KeySize = 256
    aes.BlockSize = 128
    aes.Mode = 1        ' CBC
    aes.Padding = 2     ' PKCS7
    aes.Key = key
    aes.IV = iv
    Set tr = aes.CreateEncryptor()
    enc = tr.TransformFinalBlock(data, 0, ArrLen(data))

    AesEncrypt = "aes$" & Base64Encode(ConcatBytes(iv, enc))
    Exit Function
Fallback:
    AesEncrypt = "legacy$" & Legacy_Encrypt(plainText, passphrase)
End Function

Public Function AesDecrypt(cipherText As String, passphrase As String) As String
    On Error GoTo Fallback
    If Len(cipherText) = 0 Then AesDecrypt = "": Exit Function

    If Left(cipherText, 7) = "legacy$" Then
        AesDecrypt = Legacy_Decrypt(Mid(cipherText, 8), passphrase)
        Exit Function
    End If
    If Left(cipherText, 4) <> "aes$" Then
        ' Valeur sans prefixe: ancien schema
        AesDecrypt = Legacy_Decrypt(cipherText, passphrase)
        Exit Function
    End If
    If Not IsAvailable() Then
        AesDecrypt = ""     ' AES indechiffrable sans .NET
        Exit Function
    End If

    Dim all() As Byte, iv() As Byte, enc() As Byte, key() As Byte, dec() As Byte
    Dim i As Long, n As Long
    all = Base64Decode(Mid(cipherText, 5))
    n = ArrLen(all)
    If n <= 16 Then AesDecrypt = "": Exit Function

    ReDim iv(0 To 15)
    ReDim enc(0 To n - 17)
    For i = 0 To 15: iv(i) = all(i): Next i
    For i = 16 To n - 1: enc(i - 16) = all(i): Next i

    key = HexToBytes(Pbkdf2Sha256Hex(passphrase, AES_KDF_SALT, AES_KDF_ITER, 32))

    Dim aes As Object, tr As Object
    Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged")
    aes.KeySize = 256
    aes.BlockSize = 128
    aes.Mode = 1
    aes.Padding = 2
    aes.Key = key
    aes.IV = iv
    Set tr = aes.CreateDecryptor()
    dec = tr.TransformFinalBlock(enc, 0, ArrLen(enc))

    AesDecrypt = Utf8String(dec)
    Exit Function
Fallback:
    AesDecrypt = ""
End Function

' ==============================================================================
' AUTO-TEST (vecteurs de test officiels)
' ==============================================================================

Public Function SelfTest() As String
    Dim errs As String, v As String, rt As String, plain As String

    If Not IsAvailable() Then
        SelfTest = "NON DISPONIBLE: .NET COM inaccessible - repli legacy actif (non cryptographique)"
        Exit Function
    End If

    On Error GoTo ErrHandler

    ' SHA-256("abc") - FIPS 180-2
    v = Sha256Hex("abc")
    If v <> "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad" Then errs = errs & "SHA256 KO (" & v & "); "

    ' HMAC-SHA256 - RFC 4231 test case 2
    v = HmacSha256Hex("Jefe", "what do ya want for nothing?")
    If v <> "5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843" Then errs = errs & "HMAC KO (" & v & "); "

    ' PBKDF2-HMAC-SHA256("password","salt",1,32) - vecteur de reference
    v = Pbkdf2Sha256Hex("password", "salt", 1, 32)
    If v <> "120fb6cffcf8b32c43e7225256c4f837a86548c92ccc35480805987cb70be17b" Then errs = errs & "PBKDF2 KO (" & v & "); "

    ' PBKDF2 2 iterations
    v = Pbkdf2Sha256Hex("password", "salt", 2, 32)
    If v <> "ae4d0c95af6b46d32d0adff928f06dd02a303f8ef3c251dfd6e2d85a95474c43" Then errs = errs & "PBKDF2(2) KO; "

    ' AES aller-retour
    plain = "S.A.F.A - test AES 256 CBC - caracteres speciaux: XAF 1 234 567"
    rt = AesDecrypt(AesEncrypt(plain, "Passphrase!2024"), "Passphrase!2024")
    If rt <> plain Then errs = errs & "AES round-trip KO; "

    ' AES mauvaise cle => pas de resultat identique
    rt = AesDecrypt(AesEncrypt(plain, "Passphrase!2024"), "autre")
    If rt = plain Then errs = errs & "AES accepte une mauvaise cle; "

    If errs = "" Then SelfTest = "OK" Else SelfTest = "ECHEC: " & errs
    Exit Function
ErrHandler:
    SelfTest = "ERREUR " & Err.Number & ": " & Err.Description
End Function

' ==============================================================================
' HELPERS BYTES / ENCODAGE
' ==============================================================================

Private Function Utf8Bytes(s As String) As Byte()
    Dim enc As Object
    Set enc = CreateObject("System.Text.UTF8Encoding")
    Utf8Bytes = enc.GetBytes_4(s)
End Function

Private Function Utf8String(b() As Byte) As String
    Dim enc As Object
    Set enc = CreateObject("System.Text.UTF8Encoding")
    Utf8String = enc.GetString(b)
End Function

Private Function ArrLen(b() As Byte) As Long
    On Error Resume Next
    ArrLen = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Then ArrLen = 0
    Err.Clear
End Function

Private Function ConcatBytes(a() As Byte, b() As Byte) As Byte()
    Dim r() As Byte, i As Long, na As Long, nb As Long
    na = ArrLen(a): nb = ArrLen(b)
    ReDim r(0 To na + nb - 1)
    For i = 0 To na - 1: r(i) = a(LBound(a) + i): Next i
    For i = 0 To nb - 1: r(na + i) = b(LBound(b) + i): Next i
    ConcatBytes = r
End Function

Private Function BytesToHex(b() As Byte) As String
    Dim i As Long, s As String
    If ArrLen(b) = 0 Then BytesToHex = "": Exit Function
    For i = LBound(b) To UBound(b)
        s = s & Right("0" & Hex(b(i)), 2)
    Next i
    BytesToHex = LCase(s)
End Function

Private Function HexToBytes(h As String) As Byte()
    Dim b() As Byte, i As Long, n As Long
    n = Len(h) \ 2
    If n = 0 Then HexToBytes = b: Exit Function
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = CByte("&H" & Mid(h, i * 2 + 1, 2))
    Next i
    HexToBytes = b
End Function

Private Function Base64Encode(b() As Byte) As String
    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = b
    Base64Encode = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

Private Function Base64Decode(s As String) As Byte()
    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = s
    Base64Decode = node.nodeTypedValue
End Function

' ==============================================================================
' REPLI LEGACY (NON CRYPTOGRAPHIQUE) - utilise uniquement si .NET est absent
' ==============================================================================

' FNV-1a 32 bits applique 2 fois avec des graines differentes -> 16 hex
Private Function Legacy_Hash(text As String) As String
    Dim h1 As Double, h2 As Double, i As Long, c As Long
    h1 = 2166136261#: h2 = 5381
    For i = 1 To Len(text)
        c = AscW(Mid(text, i, 1)) And &HFFFF&
        h1 = FMod((h1 Xor c) * 16777619#, 4294967296#)
        h2 = FMod(h2 * 33 + c, 4294967296#)
    Next i
    Legacy_Hash = Right("00000000" & Hex(CDbl2Long(h1)), 8) & Right("00000000" & Hex(CDbl2Long(h2)), 8)
    Legacy_Hash = LCase(Legacy_Hash)
End Function

Private Function FMod(x As Double, m As Double) As Double
    FMod = x - m * Int(x / m)
End Function

Private Function CDbl2Long(x As Double) As Long
    If x >= 2147483648# Then CDbl2Long = CLng(x - 4294967296#) Else CDbl2Long = CLng(x)
End Function

Private Function Legacy_Pbkdf(password As String, salt As String, iterations As Long) As String
    Dim i As Long, s As String
    s = password & salt
    For i = 1 To iterations
        s = Legacy_Hash(s & salt & CStr(i))
    Next i
    Legacy_Pbkdf = Legacy_Hash(s & "SAFA_FINAL")
End Function

Private Function Legacy_ExpandKey(key As String, length As Long) As String
    Dim result As String, iteration As Long
    result = key
    Do While Len(result) < length
        iteration = iteration + 1
        result = result & Legacy_Hash(key & CStr(iteration))
    Loop
    Legacy_ExpandKey = Left(result, length)
End Function

' Chiffrement par substitution/permutation (ancien "AES-like") - repli uniquement
Private Function Legacy_Encrypt(plainText As String, key As String) As String
    Dim i As Long, j As Long, blockSize As Integer
    Dim keyExpanded As String, result As String, block As String, encryptedBlock As String
    Dim charCode As Integer, keyChar As Integer, encrypted As Integer, padLength As Integer
    If Len(plainText) = 0 Or Len(key) = 0 Then Legacy_Encrypt = "": Exit Function
    blockSize = 16
    keyExpanded = Legacy_ExpandKey(key, Len(plainText) + blockSize)
    padLength = blockSize - (Len(plainText) Mod blockSize)
    plainText = plainText & String(padLength, Chr(padLength))
    For i = 1 To Len(plainText) Step blockSize
        block = Mid(plainText, i, blockSize)
        encryptedBlock = ""
        For j = 1 To Len(block)
            charCode = Asc(Mid(block, j, 1))
            keyChar = Asc(Mid(keyExpanded, i + j - 1, 1))
            encrypted = ((charCode Xor keyChar) + j) Mod 256
            encrypted = ((encrypted * 7) + 13) Mod 256
            encryptedBlock = encryptedBlock & Right("00" & Hex(encrypted), 2)
        Next j
        result = result & encryptedBlock
    Next i
    Legacy_Encrypt = Left(Legacy_Hash(CStr(Now) & CStr(Timer)), 16) & result
End Function

Private Function Legacy_Decrypt(encryptedText As String, key As String) As String
    Dim i As Long, j As Long, blockSize As Integer
    Dim keyExpanded As String, result As String, cipherText As String
    Dim byteIndex As Long, hexByte As String, encrypted As Integer, keyChar As Integer, decrypted As Integer
    If Len(encryptedText) < 16 Then Legacy_Decrypt = "": Exit Function
    blockSize = 16
    cipherText = Mid(encryptedText, 17)
    keyExpanded = Legacy_ExpandKey(key, Len(cipherText) / 2 + blockSize)
    byteIndex = 1
    For i = 1 To Len(cipherText) Step (blockSize * 2)
        For j = 0 To blockSize - 1
            If i + j * 2 <= Len(cipherText) Then
                hexByte = Mid(cipherText, i + j * 2, 2)
                encrypted = CLng("&H" & hexByte)
                decrypted = ((encrypted - 13) * 183) Mod 256
                If decrypted < 0 Then decrypted = decrypted + 256
                decrypted = (decrypted - (j + 1)) Mod 256
                If decrypted < 0 Then decrypted = decrypted + 256
                keyChar = Asc(Mid(keyExpanded, byteIndex, 1))
                decrypted = decrypted Xor keyChar
                result = result & Chr(decrypted)
                byteIndex = byteIndex + 1
            End If
        Next j
    Next i
    If Len(result) > 0 Then
        Dim lastChar As Integer
        lastChar = Asc(Right(result, 1))
        If lastChar > 0 And lastChar <= blockSize Then result = Left(result, Len(result) - lastChar)
    End If
    Legacy_Decrypt = result
End Function
