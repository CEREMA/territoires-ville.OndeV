VERSION 5.00
Begin VB.Form frmPleinEcran 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7230
   Icon            =   "PleinEcran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameCarfTC 
      Caption         =   "Arrêts TC ------------------------ Carrefours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3735
      Begin VB.Line LigneArret 
         Index           =   0
         Visible         =   0   'False
         X1              =   360
         X2              =   1800
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line OyD 
         X1              =   3480
         X2              =   3540
         Y1              =   360
         Y2              =   600
      End
      Begin VB.Line OyG 
         X1              =   3480
         X2              =   3420
         Y1              =   360
         Y2              =   600
      End
      Begin VB.Line AxeY 
         X1              =   3480
         X2              =   3480
         Y1              =   360
         Y2              =   2880
      End
      Begin VB.Label NomCarfTC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WWWWWWWWWWWWWWWWWWWW"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   3300
      End
   End
   Begin VB.Label InfoModif 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "InfoModif"
      Height          =   195
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Shape PoigneeDroite 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   5160
      Top             =   1560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape PoigneeGauche 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   4200
      Top             =   1560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Line PlageVert 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   5280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Image PtRef 
      Height          =   180
      Left            =   4440
      Picture         =   "PleinEcran.frx":0442
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label DureeCycle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LabelTemps 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t en secondes"
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Line OtB 
      X1              =   6000
      X2              =   5760
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line OtH 
      X1              =   6000
      X2              =   5760
      Y1              =   2880
      Y2              =   2760
   End
   Begin VB.Line AxeT 
      X1              =   3840
      X2              =   6000
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Menu mnuFenetre 
      Caption         =   "&Fenêtre"
      Begin VB.Menu mnuAnnulerModif 
         Caption         =   "&Annuler la dernière modification graphique"
      End
      Begin VB.Menu mnuDessinerTpsTC 
         Caption         =   "&Dessiner Temps parcours TC = F (Instant départ TC) ..."
      End
      Begin VB.Menu mnuInterCarfVP 
         Caption         =   "&Montrer les bandes inter-carrefours voitures"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options d'affichage..."
      End
      Begin VB.Menu mnuFindBandes 
         Caption         =   "&Rechercher bandes passantes suivant les vitesses..."
      End
      Begin VB.Menu mnuTracerTC 
         Caption         =   "&Tracer les progressions des Transports Collectifs..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProps 
         Caption         =   "&Propriétés..."
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFermerFenetre 
         Caption         =   "&Fermer"
      End
   End
End
Attribute VB_Name = "frmPleinEcran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable disant si la form a déjà été chargée
Private monIsLoaded As Boolean
'variable stockant le nombre de durée de cycle à afficher
Public monNbCycle As Integer

Private Sub Form_Activate()
    'Affichage des noms de tous les carrefours au milieu de ses feux et
    'de tous les arrêts TC avec un niveau de zoom correspondant à l'écart
    'entre le minimun des Y des feux et le maximun des Y des feux
    Dim unYreel As Long, unePos As Long
    Dim unCarf As Carrefour, unInd As Integer
    Dim unTC As TC
    Dim uneLongEcranAxeY As Long, uneLongueurReelAxeY As Long
    Dim unControl As Object
    
    'Aide contextuelle
    HelpContextID = IDhlp_OngletDesOnde
    
    'Affichage du site en titre de fenêtre
    Caption = monSite.Caption
    
    'Indication d'ouverture de la fenêtre plein écran
    monPleinEcranVisible = True
    
    If monIsLoaded = False Then
        'Cas où le Form_Activate n'est pas consécutif à un Form_Load
        'on ne fait rien sinon on recharge des controles existants
        '==> Plantage
        Exit Sub
    Else
        monIsLoaded = False
    End If
    
    'Dessin de l'onde verte et des plages de vert de tous les feux
    'et des progressions de TC éventuelles
    'Dans DessinerTout appelé avec False en dernier paramètre,
    'monSite.monYMaxFeuUtil = le Y maximun des feux des carrefours
    'utilisés dans le calcul de l'onde est calculé pour utilisation
    'dans le reste du code de cette procédure
    DessinerTout Me, AxeT.X1, FrameCarfTC.Top + AxeY.Y2, AxeT.X2 - AxeT.X1, AxeY.Y2 - AxeY.Y1, False
    
    'Calcul de la longueur réelle de l'englobant en Y
    'de tous carrefours utilisés dans le calcul de l'onde
    uneLongueurReelAxeY = monSite.monYMaxFeuUtil - monSite.monYMinFeuUtil
    
    'Calcul de la longueur écran de l'axe des ordonnées
    uneLongEcranAxeY = AxeY.Y2 - AxeY.Y1
    
    'Redessin de tous les nom de carrefours au bon zoom
    For i = 1 To monSite.mesCarrefours.Count
        Set unCarf = monSite.mesCarrefours(i)
        If unCarf.monDecCalcul <> -99 Then
            unInd = unInd + 1
            'Calcul du Y carrefour = barycentre des Y de ses Feux
            unYreel = DonnerYCarrefour(unCarf)
            'Distance par rapport au Y max des feux des carrefours
            'utilisés pour le calcul de l'onde
            '(zoom calculé à partir de ce point)
            unYreel = monSite.monYMaxFeuUtil - unYreel
            'Conversion du Yréel en Y écran dans la FrameCarfTC
            unePos = ConvertirReelEnEcran(unYreel, uneLongueurReelAxeY, uneLongEcranAxeY)
            'Création du label affichant le nom du carrefour en un Y écran
            'correspondant au Y réel calculé avant
            Load NomCarfTC(unInd)
            NomCarfTC(unInd).ForeColor = monSite.mesOptionsAffImp.maCoulNomCarf
            NomCarfTC(unInd).Caption = unCarf.monNom
            NomCarfTC(unInd).Top = unePos + AxeY.Y1 - NomCarfTC(unInd).Height / 2
            'Cadrage à droite de l'axe des Y des noms de carrefours
            NomCarfTC(unInd).Left = AxeY.X1 - NomCarfTC(unInd).Width - 120
            NomCarfTC(unInd).Visible = True
        End If
    Next i
    
    'Redessin de tous les arrêts TC au bon zoom
    For i = 1 To monSite.mesTC.Count
        Set unTC = monSite.mesTC(i)
        For j = 1 To unTC.mesArrets.Count
            unYreel = monSite.monYMaxFeuUtil - unTC.mesArrets(j).monOrdonnee
            'Conversion du Yréel en Y écran dans la FrameVisuCarf
            unePos = ConvertirReelEnEcran(unYreel, uneLongueurReelAxeY, uneLongEcranAxeY)
            'Création en Y écran du label affichant le nom de l'arrêt TC
            'correspondant au Y réel calculé avant
            unInd = NomCarfTC.Count
            Load NomCarfTC(unInd)
            NomCarfTC(unInd).ForeColor = monSite.mesOptionsAffImp.maCoulNomArret
            NomCarfTC(unInd).Caption = unTC.mesArrets(j).monLibelle + " (Y = " + Format(unTC.mesArrets(j).monOrdonnee) + " m)"
            NomCarfTC(unInd).Top = unePos + AxeY.Y1 - NomCarfTC(unInd).Height
            'Cadrage à gauche de la frame FrameCarfTC des noms d'arrêts
            NomCarfTC(unInd).Left = 60
            NomCarfTC(unInd).Visible = True
            'Dessin d'une ligne de l'arrêt jusqu'à l'axe des Y
            unInd2 = LigneArret.Count
            Load LigneArret(unInd2)
            LigneArret(unInd2).BorderColor = monSite.mesOptionsAffImp.maCoulNomArret
            LigneArret(unInd2).X1 = NomCarfTC(unInd).Left + NomCarfTC(unInd).Width
            LigneArret(unInd2).X2 = AxeY.X1
            LigneArret(unInd2).Y1 = unePos + AxeY.Y1
            LigneArret(unInd2).Y2 = LigneArret(unInd2).Y1
            LigneArret(unInd2).Visible = True
        Next j
    Next i
End Sub


Private Sub Form_Load()
    'Indication que l'on est passé dans le Form_Load
    'Important pour le Form_Activate
    monIsLoaded = True
    
    'Mise en actif du menu permettant d'afficher les bandes
    'inter-carrefours voitures si onde cadrée par un TC montant
    'et/ou descendant sinon il est mis en inactif
    mnuInterCarfVP.Enabled = (monSite.monTypeOnde = OndeTC)
    mnuInterCarfVP.Checked = monSite.monDessinInterCarfVP
    
    'Mise en cohérence du menu Annuler la dernière modification
    'de l'onglet Dessin d'onde verte et fenêtre Plein Ecran
    mnuAnnulerModif.Enabled = frmMain.mnuGraphicOndeAnnul.Enabled
    
    'Cadrage en haut à gauche
    Left = 0
    Top = 0
    'Agrandissment de la frame FrameCarfTC
    'On utilise Screen.height car la fenetre PleinEcran n'a pas
    'encore sa taille plein écran lors du passage dans ce load
    FrameCarfTC.Height = Screen.Height - (Height - ScaleHeight)
    'Agrandissement de l'axe des Y
    AxeY.Y2 = FrameCarfTC.Height - 180
    'Positionnement de l'axe de Temps
    AxeT.Y1 = FrameCarfTC.Top + 120
    AxeT.X1 = FrameCarfTC.Left + FrameCarfTC.Width
    AxeT.Y2 = AxeT.Y1
    AxeT.X2 = FrameCarfTC.Left + Screen.Width - (Width - ScaleWidth)
    'Positionnement de la fléche sur l'axe des temps
    OtH.X2 = AxeT.X2
    OtH.Y2 = AxeT.Y2
    OtH.X1 = OtH.X2 - 120
    OtH.Y1 = OtH.Y2 - 60
    OtB.X2 = AxeT.X2
    OtB.Y2 = AxeT.Y2
    OtB.X1 = OtH.X2 - 120
    OtB.Y1 = OtB.Y2 + 60
    'Positionnement du label LabelTemps
    LabelTemps.Top = AxeT.Y1
    LabelTemps.Left = OtH.X1 - LabelTemps.Width - 120
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Sélection graphique si bouton gauche appuyé
        'Pas de sélection multiple
        SelectionGraphique Me, X, Y
        'Indication que le bouton souris gauche enfoncé
        Tag = "DownBtnG"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Cas d'un MouseMove avec bouton gauche enfoncé ==> Modif Interactive
        ModifierSelection Me, X, Y
    Else
        'Changement du pointeur souris en croix si on passe
        'sur les poignées de sélection si elles sont visibles
        ChangerPointeurSouris Me, X, Y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Mise à jour après les modifications graphiques interactives
        MettreAJourSelection Me, X
        'Indication que le bouton souris gauche n'est plus enfoncé
        Tag = ""
    ElseIf Button = 2 And Tag = "" Then
        'Bouton droit relaché et bouton gauche pas enfoncé
        '==> Affichage du menu contextuel Fenetre
        PopupMenu mnuFenetre, vbPopupMenuRightButton
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Indication de fermeture de la fenêtre plein écran
    monPleinEcranVisible = False
    
    If monFermerParMereMDI = False Then
        'Cas d'une fermeture autre que celle déclenchée
        'par la fermeture de la fenêtre mère MDI
        '==> Fin de l'appli
        
        'Mise en cohérence du menu Annuler la dernière modification
        'de l'onglet Dessin d'onde verte
        frmMain.mnuGraphicOndeAnnul.Enabled = mnuAnnulerModif.Enabled
        
        'Redessiner le graphique de l'onde verte dans l'onglet
        'Graphique Onde Verte pour retrouver le bon repère
        With monSite
            .ZoneDessin.Cls
            unEspacement = 120 'même valeur que dans AffichageOngletVisu
            DessinerTout .ZoneDessin, .AxeTemps.X1, .AxeTemps.Y1 - unEspacement / 4, .AxeTemps.X2 - .AxeTemps.X1, .AxeOrdonnée.Y2 - .AxeOrdonnée.Y1
            'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
            'niveau que le min des Y
        End With
    End If
        
    'Désinhibition de la fenetre mère MDI, elle avait été inhibée
    'à l'ouverture de frmPleinEcran
    frmMain.Enabled = True
    
    'Affichage de la fenêtre du site étudié pour remise au premier plan
    'et ainsi éviter qu'une autre fenêtre Windows vienne au 1er plan
    'Bug affichage windows à priori
    monSite.Show
End Sub

Private Sub mnuAnnulerModif_Click()
    'Annulation de la dernière modification
    'dans le graphique d'onde verte
    AnnulerLastModifGraphic Me
    
    'Remise en grisée après l'utilisation du menu Annuler
    mnuAnnulerModif.Enabled = False
End Sub

Private Sub mnuDessinerTpsTC_Click()
    frmInfoTpsTC.Show vbModal
End Sub

Private Sub mnuFermerFenetre_Click()
    Unload Me
End Sub

Public Sub AfficherDureeCycle(unX)
    'Affichage d'un label contenant 0/Durée du cycle
    'sur chaque trait de cycle
    monNbCycle = monNbCycle + 1
    Load DureeCycle(monNbCycle)
    DureeCycle(monNbCycle).Caption = "0 - " + Format(monSite.maDuréeDeCycle)
    DureeCycle(monNbCycle).Top = LabelTemps.Top - DureeCycle(monNbCycle).Height
    DureeCycle(monNbCycle).Left = unX - DureeCycle(monNbCycle).Width / 2
    DureeCycle(monNbCycle).Visible = True
End Sub

Private Sub mnuFindBandes_Click()
    frmInfoVitBande.Show vbModal
End Sub

Private Sub mnuInterCarfVP_Click()
    'Redessin avec affichage des bandes inter-carrefours voitures
    'Cela ne se produit que si on est en onde cadrée par un TC
    mnuInterCarfVP.Checked = Not monSite.monDessinInterCarfVP
    monSite.monDessinInterCarfVP = mnuInterCarfVP.Checked
    MettreAJourDessin
    frmMain.mnuGraphicOndeInterCarfVP.Checked = monSite.monDessinInterCarfVP
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuProps_Click()
    'Affichage des propriétés du dernier objet sélectionné graphiquement
    AfficherPropsObjPick
End Sub

Private Sub mnuTracerTC_Click()
    frmTracerTC.Show vbModal
End Sub


