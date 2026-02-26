VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "S.A.F.A v10.0 - Authentification"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' FORMULAIRE: frmLogin
' Description: Interface d'authentification utilisateur
' Version: 10.0
' Auteur: S.A.F.A Team
' Date: Decembre 2024
'===============================================================================
Option Explicit

' ===== CONSTANTES =====
Private Const FORM_NAME As String = "frmLogin"
Private Const MAX_LOGIN_ATTEMPTS As Integer = 3

' ===== VARIABLES =====
Private mLoginAttempts As Integer
Private mAuthenticated As Boolean

'===============================================================================
' EVENEMENT: UserForm_Initialize
' Description: Initialisation du formulaire
'===============================================================================
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Configuration du formulaire
    Me.Caption = "S.A.F.A v10.0 - Authentification"
    Me.Width = 360
    Me.Height = 300

    ' Creer les controles dynamiquement
    CreateControls

    ' Initialiser les variables
    mLoginAttempts = 0
    mAuthenticated = False

    ' Focus sur le champ utilisateur
    Me.txtUsername.SetFocus
End Sub

'===============================================================================
' PROCEDURE: CreateControls
' Description: Creation des controles du formulaire
'===============================================================================
Private Sub CreateControls()
    On Error Resume Next

    ' Frame principal
    Dim frm As MSForms.Frame
    Set frm = Me.Controls.Add("Forms.Frame.1", "frmMain", True)
    With frm
        .Caption = ""
        .Left = 10
        .Top = 10
        .Width = Me.InsideWidth - 20
        .Height = Me.InsideHeight - 60
        .BorderStyle = fmBorderStyleNone
        .BackColor = &HFFFFFF
    End With

    ' Logo/Titre
    Dim lblLogo As MSForms.Label
    Set lblLogo = Me.Controls.Add("Forms.Label.1", "lblLogo", True)
    With lblLogo
        .Caption = Chr(127974) & " S.A.F.A v10.0"
        .Left = 50
        .Top = 20
        .Width = 260
        .Height = 30
        .Font.Size = 16
        .Font.Bold = True
        .ForeColor = &H663300
        .TextAlign = fmTextAlignCenter
    End With

    ' Sous-titre
    Dim lblSubtitle As MSForms.Label
    Set lblSubtitle = Me.Controls.Add("Forms.Label.1", "lblSubtitle", True)
    With lblSubtitle
        .Caption = "System for Automated Financial Audit"
        .Left = 50
        .Top = 50
        .Width = 260
        .Height = 15
        .Font.Size = 8
        .ForeColor = &H808080
        .TextAlign = fmTextAlignCenter
    End With

    ' Separateur
    Dim lblSep As MSForms.Label
    Set lblSep = Me.Controls.Add("Forms.Label.1", "lblSep", True)
    With lblSep
        .Caption = ""
        .Left = 30
        .Top = 75
        .Width = 300
        .Height = 1
        .BackColor = &HC0C0C0
    End With

    ' Label Utilisateur
    Dim lblUser As MSForms.Label
    Set lblUser = Me.Controls.Add("Forms.Label.1", "lblUser", True)
    With lblUser
        .Caption = "Utilisateur:"
        .Left = 30
        .Top = 95
        .Width = 80
        .Height = 15
        .Font.Size = 9
    End With

    ' TextBox Utilisateur
    Dim txtUser As MSForms.TextBox
    Set txtUser = Me.Controls.Add("Forms.TextBox.1", "txtUsername", True)
    With txtUser
        .Left = 30
        .Top = 115
        .Width = 300
        .Height = 24
        .Font.Size = 10
    End With

    ' Label Mot de passe
    Dim lblPwd As MSForms.Label
    Set lblPwd = Me.Controls.Add("Forms.Label.1", "lblPwd", True)
    With lblPwd
        .Caption = "Mot de passe:"
        .Left = 30
        .Top = 150
        .Width = 80
        .Height = 15
        .Font.Size = 9
    End With

    ' TextBox Mot de passe
    Dim txtPwd As MSForms.TextBox
    Set txtPwd = Me.Controls.Add("Forms.TextBox.1", "txtPassword", True)
    With txtPwd
        .Left = 30
        .Top = 170
        .Width = 300
        .Height = 24
        .Font.Size = 10
        .PasswordChar = "*"
    End With

    ' Bouton Connexion
    Dim btnLogin As MSForms.CommandButton
    Set btnLogin = Me.Controls.Add("Forms.CommandButton.1", "btnLogin", True)
    With btnLogin
        .Caption = "Se Connecter"
        .Left = 100
        .Top = 220
        .Width = 160
        .Height = 30
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = &H663300
        .ForeColor = &HFFFFFF
    End With

    ' Label Status
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Caption = ""
        .Left = 30
        .Top = 260
        .Width = 300
        .Height = 15
        .Font.Size = 8
        .ForeColor = &HFF&
        .TextAlign = fmTextAlignCenter
    End With
End Sub

'===============================================================================
' EVENEMENT: btnLogin_Click
' Description: Tentative de connexion
'===============================================================================
Private Sub btnLogin_Click()
    On Error GoTo ErrorHandler

    Dim username As String
    Dim password As String

    username = Trim(Me.txtUsername.Value)
    password = Me.txtPassword.Value

    ' Validation des champs
    If username = "" Then
        ShowStatus "Veuillez entrer un nom d'utilisateur", True
        Me.txtUsername.SetFocus
        Exit Sub
    End If

    If password = "" Then
        ShowStatus "Veuillez entrer un mot de passe", True
        Me.txtPassword.SetFocus
        Exit Sub
    End If

    ' Incrementer les tentatives
    mLoginAttempts = mLoginAttempts + 1

    ' Verifier les credentials
    If Security_Module.Authenticate(username, password) Then
        mAuthenticated = True
        ShowStatus "Connexion reussie!", False

        ' Petit delai pour afficher le message
        Application.Wait Now + TimeValue("00:00:01")

        Me.Hide
    Else
        ShowStatus "Identifiants incorrects (" & mLoginAttempts & "/" & MAX_LOGIN_ATTEMPTS & ")", True
        Me.txtPassword.Value = ""
        Me.txtPassword.SetFocus

        ' Verifier le nombre de tentatives
        If mLoginAttempts >= MAX_LOGIN_ATTEMPTS Then
            MsgBox "Nombre maximum de tentatives atteint." & vbNewLine & _
                   "L'application va se fermer.", vbCritical, "Acces Refuse"
            mAuthenticated = False
            Me.Hide
        End If
    End If

    Exit Sub

ErrorHandler:
    ShowStatus "Erreur: " & Err.Description, True
End Sub

'===============================================================================
' EVENEMENT: txtPassword_KeyDown
' Description: Connexion sur Enter dans le champ mot de passe
'===============================================================================
Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Enter
        btnLogin_Click
    End If
End Sub

'===============================================================================
' EVENEMENT: txtUsername_KeyDown
' Description: Passer au mot de passe sur Enter
'===============================================================================
Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Enter
        Me.txtPassword.SetFocus
    End If
End Sub

'===============================================================================
' PROCEDURE: ShowStatus
' Description: Affiche un message de statut
'===============================================================================
Private Sub ShowStatus(message As String, isError As Boolean)
    Me.lblStatus.Caption = message
    If isError Then
        Me.lblStatus.ForeColor = &HFF& ' Rouge
    Else
        Me.lblStatus.ForeColor = &HC000& ' Vert
    End If
End Sub

'===============================================================================
' PROPRIETE: IsAuthenticated
' Description: Retourne True si l'utilisateur est authentifie
'===============================================================================
Public Property Get IsAuthenticated() As Boolean
    IsAuthenticated = mAuthenticated
End Property

'===============================================================================
' EVENEMENT: UserForm_QueryClose
' Description: Gestion de la fermeture du formulaire
'===============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' L'utilisateur a clique sur X
        mAuthenticated = False
    End If
End Sub
