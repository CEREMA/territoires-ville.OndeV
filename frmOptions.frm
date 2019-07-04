VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options d'affichage et d'impression"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CheckBox CheckSaveDefaut 
      Caption         =   "Conserver en valeurs par défaut"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Frame FrameOptionGen 
      Caption         =   "Options générales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   3615
      Begin VB.PictureBox ColorTitreEchelle 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   27
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox ColorArretTC 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   25
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox ColorNomCarf 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox ColorLigneVert 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox ColorPtRef 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox CheckLigneVert 
         Caption         =   "Ligne verticale toutes les 10 secondes"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Titre de l'étude et échelle"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nom d'arrêt TC"
         Height          =   195
         Left            =   1200
         TabIndex        =   26
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nom de carrefour"
         Height          =   195
         Left            =   1080
         TabIndex        =   24
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Point de référence"
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame FrameSensD 
      Caption         =   "Sens descendant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   3615
      Begin VB.CheckBox CheckBandeD 
         Caption         =   "Bande passante commune"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox CheckBandeDInter 
         Caption         =   "Bande passante inter-carrefours"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2655
      End
      Begin VB.PictureBox ColorBD 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox ColorBDInter 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame FrameSensM 
      Caption         =   "Sens montant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox ColorBMInter 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox ColorBM 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   3240
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox CheckBandeMInter 
         Caption         =   "Bande passante inter-carrefours"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox CheckBandeM 
         Caption         =   "Bande passante commune"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Tag             =   "Annuler"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   7
         Tag             =   "Exemple 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   6
         Tag             =   "Exemple 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   4
         Tag             =   "Exemple 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Unload Me
    
    'Affichage de la form fille active pour éviter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    'si on n'est pas en plein écran et que l'on ne vient pas de la
    'fenetre impression
    If monPleinEcranVisible = False And monCallOptionByPrint = False Then
        frmMain.ActiveForm.Show
    End If
End Sub


Private Sub cmdOK_Click()
    'Stockage des nouvelles options d'affichage
    StockerOptionsAffImp
    
    'Redessin du graphique de l'onde verte
    MettreAJourDessin
        
    'Sauvegarde si demandé par l'utilisateur
    'des couleurs et options par défaut dans de la base
    'de registre à la place du fichier
    'OndeV.ini (fait à partir de la version 1.00.0002)
    If CheckSaveDefaut.Value = 1 Then
        SauverOptionsAffImp False
    End If
    
    'Fermeture de la fenetre
    Unload Me
    
    'Affichage de la form fille active pour éviter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    'si on n'est pas en plein écran et que l'on ne vient pas de la
    'fenetre impression
    If monPleinEcranVisible = False And monCallOptionByPrint = False Then
        frmMain.ActiveForm.Show
    End If
End Sub


Private Sub ColorArretTC_Click()
    frmMain.ChangerCouleurPicBox ColorArretTC
End Sub

Private Sub ColorBD_Click()
    frmMain.ChangerCouleurPicBox ColorBD
End Sub

Private Sub ColorBDInter_Click()
    frmMain.ChangerCouleurPicBox ColorBDInter
End Sub

Private Sub ColorBM_Click()
    frmMain.ChangerCouleurPicBox ColorBM
End Sub


Private Sub ColorBMInter_Click()
    frmMain.ChangerCouleurPicBox ColorBMInter
End Sub

Private Sub ColorLigneVert_Click()
    frmMain.ChangerCouleurPicBox ColorLigneVert
End Sub

Private Sub ColorNomCarf_Click()
    frmMain.ChangerCouleurPicBox ColorNomCarf
End Sub

Private Sub ColorPtRef_Click()
    frmMain.ChangerCouleurPicBox ColorPtRef
End Sub

Private Sub ColorTitreEchelle_Click()
    frmMain.ChangerCouleurPicBox ColorTitreEchelle
End Sub

Private Sub Form_Load()
    'Index pour l'aide
    HelpContextID = IDhlp_WinAffichageOptions
    
    'Centrage de la fenêtre à l'écran
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'Chargement des options d'affichage et d'impression
    AfficherOptionsAffImp monSite
End Sub

Public Sub AfficherOptionsAffImp(unSite As Form)
    'Affichage des valeurs des options d'affichage et d'impression
    'dans la fenêtre Options d'affichage et d'impression
    
    ColorBD.BackColor = unSite.mesOptionsAffImp.maCoulBandComD
    ColorBM.BackColor = unSite.mesOptionsAffImp.maCoulBandComM
    ColorBDInter.BackColor = unSite.mesOptionsAffImp.maCoulBandInterCarfD
    ColorBMInter.BackColor = unSite.mesOptionsAffImp.maCoulBandInterCarfM
    ColorLigneVert.BackColor = unSite.mesOptionsAffImp.maCoulLigne
    ColorArretTC.BackColor = unSite.mesOptionsAffImp.maCoulNomArret
    ColorNomCarf.BackColor = unSite.mesOptionsAffImp.maCoulNomCarf
    ColorPtRef.BackColor = unSite.mesOptionsAffImp.maCoulPtRef
    ColorTitreEchelle.BackColor = unSite.mesOptionsAffImp.maCoulTitreEch
    
    If unSite.mesOptionsAffImp.maVisuBandComD Then
        CheckBandeD.Value = 1
    Else
        CheckBandeD.Value = 0
    End If
    
    If unSite.mesOptionsAffImp.maVisuBandComM Then
        CheckBandeM.Value = 1
    Else
        CheckBandeM.Value = 0
    End If
    
    If unSite.mesOptionsAffImp.maVisuBandInterCarfD Then
        CheckBandeDInter.Value = 1
    Else
        CheckBandeDInter.Value = 0
    End If
    
    If unSite.mesOptionsAffImp.maVisuBandInterCarfM Then
        CheckBandeMInter.Value = 1
    Else
        CheckBandeMInter.Value = 0
    End If
    
    If unSite.mesOptionsAffImp.maVisuLigne Then
        CheckLigneVert.Value = 1
    Else
        CheckLigneVert.Value = 0
    End If
End Sub

Private Sub StockerOptionsAffImp()
    'Stockage des valeurs des options d'affichage et d'impression dans
    'l'instance Options d'affichage et d'impression de la fenêtre fille
    
    monSite.mesOptionsAffImp.maCoulBandComD = ColorBD.BackColor
    monSite.mesOptionsAffImp.maCoulBandComM = ColorBM.BackColor
    monSite.mesOptionsAffImp.maCoulBandInterCarfD = ColorBDInter.BackColor
    monSite.mesOptionsAffImp.maCoulBandInterCarfM = ColorBMInter.BackColor
    monSite.mesOptionsAffImp.maCoulLigne = ColorLigneVert.BackColor
    monSite.mesOptionsAffImp.maCoulNomArret = ColorArretTC.BackColor
    monSite.mesOptionsAffImp.maCoulNomCarf = ColorNomCarf.BackColor
    monSite.mesOptionsAffImp.maCoulPtRef = ColorPtRef.BackColor
    monSite.mesOptionsAffImp.maCoulTitreEch = ColorTitreEchelle.BackColor
    
    If CheckBandeD.Value = 1 Then
        monSite.mesOptionsAffImp.maVisuBandComD = True
    Else
        monSite.mesOptionsAffImp.maVisuBandComD = False
    End If
    
    If CheckBandeM.Value = 1 Then
        monSite.mesOptionsAffImp.maVisuBandComM = True
    Else
        monSite.mesOptionsAffImp.maVisuBandComM = False
    End If
        
    If CheckBandeDInter.Value = 1 Then
        monSite.mesOptionsAffImp.maVisuBandInterCarfD = True
    Else
        monSite.mesOptionsAffImp.maVisuBandInterCarfD = False
    End If
    
    If CheckBandeMInter.Value = 1 Then
        monSite.mesOptionsAffImp.maVisuBandInterCarfM = True
    Else
        monSite.mesOptionsAffImp.maVisuBandInterCarfM = False
    End If
    
    If CheckLigneVert.Value = 1 Then
        monSite.mesOptionsAffImp.maVisuLigne = True
    Else
        monSite.mesOptionsAffImp.maVisuLigne = False
    End If
End Sub


