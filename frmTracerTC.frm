VERSION 5.00
Begin VB.Form frmTracerTC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix des Transports Collectifs pour tracer leur progression"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmTracerTC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BoutonAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton BoutonOK 
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
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton BoutonEnlev 
      Caption         =   "Enlever"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "frmTracerTC.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton BoutonAjout 
      Caption         =   "Ajouter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "frmTracerTC.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox ListSelTC 
      Height          =   2790
      Left            =   3600
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox ListToutTC 
      Height          =   2790
      ItemData        =   "frmTracerTC.frx":0B8E
      Left            =   120
      List            =   "frmTracerTC.frx":0B90
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Liste des TC � tracer"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Liste des TC disponibles"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frmTracerTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BoutonAjout_Click()
    'Parcours de la liste des TC disponibles
    'et ajout de ceux s�lectionn�s dans la liste
    'des TC � tracer et suppression des s�lectionn�s
    'dans la liste des TC disponibles
    For i = ListToutTC.ListCount - 1 To 0 Step -1
        If ListToutTC.Selected(i) Then
            ListSelTC.AddItem ListToutTC.List(i)
            ListToutTC.RemoveItem i
        End If
    Next i
End Sub

Private Sub BoutonAnnuler_Click()
    Unload Me
End Sub

Private Sub BoutonEnlev_Click()
    'Parcours de la liste des TC � tracer
    'et suppression de ceux s�lectionn�s dans la liste
    'des TC � tracer et ajout dans la liste des TC disponibles
    For i = ListSelTC.ListCount - 1 To 0 Step -1
        If ListSelTC.Selected(i) Then
            ListToutTC.AddItem ListSelTC.List(i)
            ListSelTC.RemoveItem i
        End If
    Next i
End Sub

Private Sub BoutonOK_Click()
    Dim unePosTC As Integer
    Dim unTC As TC, unePhase As PhaseTabMarche
    
    'On vide la liste des TC util
    ViderCollection monSite.mesTCutil
    
    'Alimentation de la liste des TC util
    For i = 0 To ListSelTC.ListCount - 1
        unePosTC = TrouverTCParNom(monSite, ListSelTC.List(i))
        'Stockage dans les TC utilis�s
        monSite.mesTCutil.Add monSite.mesTC(unePosTC)
    Next i
        
    'Redessin du graphique de l'onde verte
    MettreAJourDessin
       
    'Fermeture fen�tre
    Unload Me
End Sub

Private Sub Form_Load()
    'Alimentation des TC disponibles
    'Les TC dont on n'a pas encore tracer la progression
    Dim i As Integer
        
    'Index pour l'aide
    HelpContextID = IDhlp_WinTracerTC
    
    For i = 1 To monSite.mesTC.Count
        unNomTC = monSite.mesTC(i).monNom
        If EstTCUtil(i) = False Then
            'Cas du TC ne faisant pas partie des TC � tracer
            ListToutTC.AddItem unNomTC
        End If
    Next i
    
    'Alimentation des TC � tracer
    For i = 1 To monSite.mesTCutil.Count
        ListSelTC.AddItem monSite.mesTCutil(i).monNom
    Next i
    
    'Centrage de la fen�tre � l'�cran
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub

Private Sub ListSelTC_DblClick()
    'Ajout de l'item double cliqu� dans la liste TC disponibles
    ListToutTC.AddItem ListSelTC.List(ListSelTC.ListIndex)
    'Suppression dans la liste des TC � tracer
    ListSelTC.RemoveItem ListSelTC.ListIndex
End Sub

Private Sub ListToutTC_DblClick()
    'Ajout de l'item double cliqu� dans la liste des TC � tracer
    ListSelTC.AddItem ListToutTC.List(ListToutTC.ListIndex)
    'Suppression dans la liste des TC disponibles
    ListToutTC.RemoveItem ListToutTC.ListIndex
End Sub
