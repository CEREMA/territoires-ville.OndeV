VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmPropsObjPick 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propri�t�s du dernier objet s�lectionn�"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmPropsObjPick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread TabInfoFeux 
      Height          =   2415
      Left            =   840
      OleObjectBlob   =   "frmPropsObjPick.frx":0442
      TabIndex        =   5
      Top             =   1560
      Width           =   6855
   End
   Begin VB.CommandButton BoutonFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelBD 
      AutoSize        =   -1  'True
      Caption         =   "Bande passante descendante :"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   2220
   End
   Begin VB.Label LabelValBD 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "145.55"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label LabelSecBD 
      AutoSize        =   -1  'True
      Caption         =   "secondes"
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label LabelBM 
      AutoSize        =   -1  'True
      Caption         =   "Bande passante montante :"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1950
   End
   Begin VB.Label LabelValBM 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "145.55"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   600
      Width           =   675
   End
   Begin VB.Label LabelSecBM 
      AutoSize        =   -1  'True
      Caption         =   "secondes"
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   690
   End
   Begin VB.Label LabelFeux 
      AutoSize        =   -1  'True
      Caption         =   "Feux"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label3 
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
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
   Begin VB.Label LabelDecal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "555"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "D�calage :"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   945
   End
   Begin VB.Label LabelNomCarf 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MMMMMWWWWWMMMMM "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label LabelCarf 
      AutoSize        =   -1  'True
      Caption         =   "Carrefour : "
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
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPropsObjPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BoutonFermer_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim unObjPick As Object
    Dim unCarf As Carrefour, unFeu As Feu
    Dim uneCol As New Collection
    Dim unY As Long, unYMax As Long
    Dim unDebVert As Single, unFinVert As Single
    
    Set unObjPick = DonnerObjPick
    Set unCarf = monSite.mesCarrefours(unObjPick.monIndCarf)
    
    'Remplissage du nom et du d�calage du carrefour
    LabelNomCarf.Caption = unCarf.monNom
    LabelDecal.Caption = CIntCorrig�(unCarf.monDecModif)

    'Remplissage des bandes passantes avec les couleurs des sens
    LabelValBM.Caption = Format(monSite.maBandeModifM, "Fixed")
    LabelValBM.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
    LabelBM.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
    LabelSecBM.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
    LabelValBD.Caption = Format(monSite.maBandeModifD, "Fixed")
    LabelValBD.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
    LabelBD.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
    LabelSecBD.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
    
    'Retaillage du Spread TabInfoFeux et de la fen�tre
    uneRowWidth = TabInfoFeux.RowHeight(1)
    TabInfoFeux.MaxRows = unCarf.mesFeux.Count
    TabInfoFeux.Height = uneRowWidth * (unCarf.mesFeux.Count + 2)
    '+ 2 pour la ligne d'ent�te et un d�calage par rapport au bord
    'du bas de fen�tre
    Height = (Height - ScaleHeight) + TabInfoFeux.Top + TabInfoFeux.Height
    
    'Tri des feux par ordre croissant de leur Y
    'Rangement dans une collection uneCol rang�s par ordre
    'croissant au fur et � mesure
    uneCol.Add unCarf.mesFeux(1) 'Insertion du premier feux
    For i = 2 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        For j = uneCol.Count To 1 Step -1
            If unFeu.monOrdonn�e > uneCol(j).monOrdonn�e Then
                'Insertion �pr�s le j �me �l�ment
                uneCol.Add unFeu, , , j
                Exit For
            End If
            'Insertion en t�te s'il n'a pas �t� ajout�
            'avant, donc c'est le plus petit Y
            If j = 1 Then uneCol.Add unFeu, , 1
        Next j
    Next i
    
    'Positionnement en X des ext�mit�s des plages de vert
    '� dessiner du feu en face de sa ligne de propri�t�s
    unX1 = LabelFeux.Left
    unX2 = TabInfoFeux.Left - LabelFeux.Left
    unY = TabInfoFeux.Top + uneRowWidth * 0.5 'Initialisation
        
    'Remplissage du spread les feux d'Y le plus haut seront mis
    'en dans les lignes les plus hautes du tableau
    For i = uneCol.Count To 1 Step -1
        Set unFeu = uneCol(i)
        'Calcul du d�but et de fin de vert de chaque feu
        unDebVert = unCarf.monDecModif + unFeu.maPositionPointRef
        unFinVert = unCarf.monDecModif + unFeu.maPositionPointRef + unFeu.maDur�eDeVert
        'On les ram�ne modulo cycle
        unDebVert = ModuloZeroCycle(unDebVert, monSite.maDur�eDeCycle)
        unFinVert = ModuloZeroCycle(unFinVert, monSite.maDur�eDeCycle)
        
        'Remplissage de la ligne de propri�t�s du feu
        TabInfoFeux.Row = uneCol.Count + 1 - i
        TabInfoFeux.Col = 1
        TabInfoFeux.Text = unFeu.monOrdonn�e
        TabInfoFeux.Col = 2
        TabInfoFeux.Text = unDebVert
        TabInfoFeux.Col = 3
        TabInfoFeux.Text = unFinVert
        TabInfoFeux.Col = 4
        TabInfoFeux.Text = unFeu.maDur�eDeVert
        TabInfoFeux.Col = 5
        TabInfoFeux.Text = -unFeu.maPositionPointRef
        'Dessin de la plage de vert avec la couleur
        'du sens en face de sa ligne de propri�t�s
        'gr�ce � la valeur de uneRowWidth
        unY = unY + uneRowWidth
        If unFeu.monSensMontant Then
            uneCouleur = monSite.mesOptionsAffImp.maCoulBandComM
        Else
            uneCouleur = monSite.mesOptionsAffImp.maCoulBandComD
        End If
        DrawWidth = 3
        Line (unX1, unY)-(unX2, unY), uneCouleur
    Next i
    
    'Centrage � l'�cran de cette fen�tre
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'Lib�ration de m�moire
    Set uneCol = Nothing
End Sub
