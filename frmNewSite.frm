VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewSite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Création d'un plan de feux par défaut"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmNewSite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox DuréeCycle 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2655
      TabIndex        =   6
      Top             =   720
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      Value           =   50
      BuddyControl    =   "DuréeCycle"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton OKBouton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox NomPlanDeFeux 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Plan 1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Durée du cycle : "
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
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nom du plan de feux : "
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
      TabIndex        =   0
      Top             =   240
      Width           =   1950
   End
End
Attribute VB_Name = "frmNewSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelBouton_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    If IsEtatCreation Then
        'Valeur par défaut à la création
        DuréeCycle.Text = Str(50)
        UpDown1.Value = 50
    Else
        'Récupération des valeurs existantes du plan de feux courant
        DuréeCycle.Text = monSite.monPlanDeFeuxCourant.maDuréeDeCycle
        NomPlanDeFeux.Text = monSite.monPlanDeFeuxCourant.monNom
    End If
End Sub

Private Sub OKBouton_Click()
    If Val(DuréeCycle.Text) > UpDown1.Max Or Val(DuréeCycle.Text) < UpDown1.Min Then
        'Test des valeurs min et max du cycle
        uneChaine = "La durée du cycle doit être comprise entre " + Str(UpDown1.Min)
        uneChaine = uneChaine + " et " + Str(UpDown1.Max)
        MsgBox uneChaine, vbCritical, "Message d'erreur de Talon"
        DuréeCycle.SetFocus
    ElseIf IsEtatCreation Then
        'Création d'une nouvelle fenetre fille
        Dim frmD As frmDocument
        
        monDocumentCount = monDocumentCount + 1
        Set frmD = New frmDocument
        Set frmD.monPlanDeFeuxCourant = frmD.mesPlansDeFeux.Add(NomPlanDeFeux.Text, Val(DuréeCycle.Text))
        'Création du premier carrefour
        Set frmD.monCarrefourCourant = frmD.monPlanDeFeuxCourant.mesCarrefours.Add("Carrefour 1")
        'Modification du titre de la fenetre Site
        frmD.Caption = "Site : Sans Nom " & monDocumentCount
        'Stockage de la fenetre du site courant
        Set monSite = frmD
        'Activation du menu Plans de feux
        fMainForm.mnuPlanFeux.Enabled = True
        'Affichage de la fenetre Site
        Unload Me
        frmD.Show
    Else
        'Modification des attributs du plan de feux courant du site courant
        monSite.monPlanDeFeuxCourant.maDuréeDeCycle = Val(DuréeCycle.Text)
        monSite.monPlanDeFeuxCourant.monNom = NomPlanDeFeux.Text
        'Modification des affichages dans la fenetre du site courant
        monSite.DuréeCycle.Caption = monSite.monPlanDeFeuxCourant.maDuréeDeCycle
        monSite.NomPlanDeFeux.Caption = monSite.monPlanDeFeuxCourant.monNom
        Unload Me
    End If
End Sub
