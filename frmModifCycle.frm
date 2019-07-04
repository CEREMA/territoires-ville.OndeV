VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmModifCycle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modification de la dur�e du cycle"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmModifCycle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Dur�eCycle 
      Height          =   300
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "999"
      Top             =   1920
      Width           =   375
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   405
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   714
      _Version        =   327680
      Value           =   50
      BuddyControl    =   "Dur�eCycle"
      BuddyDispid     =   196609
      OrigLeft        =   2640
      OrigTop         =   720
      OrigRight       =   2880
      OrigBottom      =   1095
      Max             =   150
      Min             =   20
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton CancelBouton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton OKBouton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "secondes"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmModifCycle.frx":0442
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmModifCycle.frx":04EB
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dur�e du cycle : "
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1485
   End
End
Attribute VB_Name = "frmModifCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsCorrig� As Boolean
Public maColFeux As Collection

Private Sub CancelBouton_Click()
    Unload Me
End Sub



Private Sub Dur�eCycle_KeyUp(KeyCode As Integer, Shift As Integer)
    Call VerifSaisieEntierPositif(KeyCode, Dur�eCycle, 50)
End Sub

Private Sub Form_Load()
    'R�cup�ration de la dur�e du cycle existant du site courant
    Dur�eCycle.Text = monSite.maDur�eDeCycle
    UpDown1.Value = Str(Dur�eCycle.Text)
    'Centrage de UpDown1 par rapport � Dur�eCycle
    UpDown1.Top = UpDown1.Top - (UpDown1.Height - Dur�eCycle.Height) / 2
End Sub

Private Sub OKBouton_Click()
    Dim uneNewDureeCycle As Integer
    Dim unCarf As Carrefour
    
    'Initialisation
    IsCorrig� = False
    Set maColFeux = New Collection
    
    If Val(Dur�eCycle.Text) > UpDown1.Max Or Val(Dur�eCycle.Text) < UpDown1.Min Then
        'Test des valeurs min et max du cycle
        uneChaine = "La dur�e du cycle doit �tre comprise entre " + Str(UpDown1.Min)
        uneChaine = uneChaine + " et " + Str(UpDown1.Max)
        MsgBox uneChaine, vbCritical, "Message d'erreur de OndeV"
        Dur�eCycle.SetFocus
    Else
        'Recherche des feux dont la dur�e de vert > nouvelle dur�e du cycle
        uneNewDureeCycle = Val(frmModifCycle.Dur�eCycle.Text)
        For i = 1 To monSite.mesCarrefours.Count
            Set unCarf = monSite.mesCarrefours(i)
            For j = 1 To unCarf.mesFeux.Count
                If unCarf.mesFeux(j).maDur�eDeVert > uneNewDureeCycle Then
                    maColFeux.Add unCarf.mesFeux(j)
                End If
            Next j
        Next i
        If maColFeux.Count > 0 Then
            'Cas o� il y a des corrections � faire
            'Affichage de la fen�tre de correction en proposant des valeurs
            frmCorrigFeux.Show vbModal
        End If
        
        If IsCorrig� Or maColFeux.Count = 0 Then
            'Cas o� l'utilisateur valide la correction ==> click sur OK
            'Le click sur Annuler redonne la main � cette fen�tre
            
            'Modification de la dur�e du cycle du site courant
            monSite.maDur�eDeCycle = Val(Dur�eCycle.Text)
            'Modification des affichages dans la fenetre du site courant
            monSite.Dur�eCycle.Text = monSite.maDur�eDeCycle
            'Indication d'une modification dans les donn�es Carrefour
            monSite.maModifDataCarf = True
        
            'Recalcul des ondes vertes et redessin
            'si l'onglet courant est l'un des trois derniers onglets :
            'R�sultats D�calages, Dessin onde verte et Fiche R�sultats
            unTabPred = monSite.TabFeux.Tab
            If unTabPred > 2 Then
                monSite.TabFeux.Tab = 1
                monSite.TabFeux.Tab = unTabPred
                '==> Red�clenchement des calculs et/ou du dessin
            End If
        End If
        
        IsCorrig� = False
        ViderCollection maColFeux
        Set maColFeux = Nothing
        Unload Me
    End If
End Sub

