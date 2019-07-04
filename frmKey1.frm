VERSION 5.00
Begin VB.Form frmKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Licence register"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   1332
   End
   Begin VB.TextBox TxtLicence 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgLogo 
      Height          =   1215
      Index           =   0
      Left            =   4560
      Picture         =   "frmKey.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label LblTitre 
      Caption         =   "Please, register your licence"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label LblLicence 
      Alignment       =   1  'Right Justify
      Caption         =   "Licence :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label LblUser 
      Alignment       =   1  'Right Justify
      Caption         =   "User :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'intialisation
    Me.Caption = Titre
    Me.LblTitre.Caption = MSG
    Me.LblUser.Caption = LBUSER
    Me.LblLicence.Caption = LBLICENCE
    Me.cmdOK.Caption = BTNOK
    Me.cmdCancel.Caption = BTNCANCEL

End Sub

'l'utilisateur clique sur annuler
Private Sub cmdCancel_Click()
    Dim licencevalide As Boolean
    'appel de la méthode
    licencevalide = fin
End Sub

'l'utilisateur clique sur OK
Private Sub cmdOK_Click()
    Dim licencevalide As Boolean
    'appel de la méthode
    If VerifLicence("CERTU", TxtUser.Text, TxtLicence.Text) Then
        MsgBox MSGPWDVALID, vbInformation
    End If
End Sub

