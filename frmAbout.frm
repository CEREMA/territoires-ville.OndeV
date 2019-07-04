VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de OndeV"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "À propos de Talon"
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   3840
      Top             =   1920
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   6
      Top             =   2760
      Width           =   6615
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   1500
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   1440
      ScaleMode       =   0  'User
      ScaleWidth      =   1380
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
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
      Height          =   345
      Left            =   3120
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4080
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   4800
      Picture         =   "frmAbout.frx":1749
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1920
   End
   Begin VB.Label NumLicence 
      AutoSize        =   -1  'True
      Caption         =   "Licence N° : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label lblDescription 
      Caption         =   "Calcul et Tracé assisté d'ondes vertes Développé par le groupe ITS du CERTU"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "OndeV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   4
      Tag             =   "Titre de l'application"
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   240
      Width           =   1140
   End
   Begin VB.Label WarningLabel 
      Caption         =   "Avertissement: Logiciel protégé.                 Toute reproduction interdite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   240
      TabIndex        =   2
      Tag             =   "Avertissement: ..."
      Top             =   1920
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SRCCOPY = &HCC0020
Const ShowText$ = "Frank TRIFILETTI"
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nheight As Integer, ByVal hSrcDC As Long, ByVal Xsrc As Integer, ByVal Ysrc As Integer, ByVal dwRop As Long) As Integer
Dim ShowIt%, monIndMsg%
Dim monTabString(9) As String

Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub Form_Load()
    'Affichage du numéro de licence
    If maDemoVersion Then NumeroLicence = "DEMO"
    NumLicence.Caption = LBLICENCE + NumeroLicence
     
    'Traitement permettant de lister la boucle des intervenants
    unDecalage = "     "
    WarningLabel.Caption = "Avertissement :" + Chr(13) + unDecalage + "Logiciel protégé"
    WarningLabel.Caption = WarningLabel.Caption + Chr(13) + unDecalage + "Toute reproduction interdite"
    WarningLabel.Font.Bold = True
    'Initialisation de l'indice des messages listant les participants
    monIndMsg% = 0
    'Initialisation des noms des participants
    monTabString(0) = "Production du cahier des charges"
    monTabString(1) = "    CERTU / Département SYSTEMES / Groupe Transport"
    monTabString(2) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
    monTabString(3) = "Réalisation du développement du logiciel"
    monTabString(4) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
    monTabString(5) = "Diffusion et Assistance au logiciel"
    monTabString(6) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim uneString As String
    
    If (ShowIt% Mod 20 = 0) Then
        Picture1.CurrentX = 20
        Picture1.CurrentY = Picture1.ScaleHeight - 20
        'Affichage du participant d'indice monIndMsg%
        Picture1.Print monTabString(monIndMsg% Mod 7)
        ShowIt% = 1
        If monIndMsg% = 7 Then
            'Pour éviter un débordement de capacité des entiers
            monIndMsg% = 1
        Else
            'Permettra l'affichage du message suivant
            monIndMsg% = monIndMsg% + 1
        End If
    Else
        i = BitBlt(Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight - 1, Picture1.hdc, 0, 1, SRCCOPY)
        ShowIt% = ShowIt% + 1
    End If
End Sub
