VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmInfoVitBande 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recherche de bandes passantes suivant les vitesses"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmInfoVitBande.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChartVitBandes 
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "frmInfoVitBande.frx":0442
      TabIndex        =   9
      Top             =   1920
      Width           =   10575
   End
   Begin FPSpread.vaSpread TabInfoVit 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9375
      _Version        =   131077
      _ExtentX        =   16536
      _ExtentY        =   1508
      _StockProps     =   64
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   1
      RowHeaderDisplay=   2
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmInfoVitBande.frx":2798
      UnitType        =   2
      UserResize      =   0
      VisibleCols     =   5
      VisibleRows     =   3
   End
   Begin VB.CommandButton BoutonImp 
      Caption         =   "Imprimer..."
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame FrameSeparateur 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   10575
   End
   Begin VB.CommandButton BoutonCalculer 
      Caption         =   "Calculer"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton BoutonFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   375
      Left            =   9600
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LabeLCarfDebFin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calcul entre les carrefours "
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
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label LabelNomFichier 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du fichier : "
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
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label LabelTitre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titre de l'étude : "
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
      Top             =   600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LabelDureeCycle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1485
   End
End
Attribute VB_Name = "frmInfoVitBande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BoutonCalculer_Click()
    Dim uneVDeb As Integer, uneVFin As Integer
    Dim unPas As Integer
    Dim unNomCarfDeb As String, unNomCarfFin As String
    
    'Changement du pointeur souris en sablier
    MousePointer = vbHourglass
    
    'Stockage des paramètres du site en cours qui vont modifiés
    'lors du calcul
    maSaveTypeOnde = monSite.monTypeOnde
    maSaveTypeVit = monSite.monTypeVit
    maSaveVitSensM = monSite.maVitSensM
    maSaveVitSensD = monSite.maVitSensD
    
    TabInfoVit.Row = 1
    'Récup de la vitesse début de recherche
    TabInfoVit.Col = 1
    uneVDeb = Val(TabInfoVit.Text)
    'Récup de la vitesse fin de recherche
    TabInfoVit.Col = 2
    uneVFin = Val(TabInfoVit.Text)
    'Récup du pas de recherche
    TabInfoVit.Col = 3
    unPas = Val(TabInfoVit.Text)
    
    'Récupération des noms des carrefours début et fin
    TabInfoVit.Row = 1
    TabInfoVit.Col = 4
    unNomCarfDeb = TabInfoVit.Text
    TabInfoVit.Row = 1
    TabInfoVit.Col = 5
    unNomCarfFin = TabInfoVit.Text
    
    'Test des cas de saisies erronées
    uneSaisieErronée = True
    If uneVDeb = 0 Or uneVFin = 0 Or unPas = 0 Then
        unMsg = "Les vitesses début, fin et le pas de recherche doivent être non nulles"
    ElseIf uneVDeb >= uneVFin Then
        unMsg = "La vitesse de début doit être inférieure à la vitesse de fin"
    ElseIf unNomCarfDeb = unNomCarfFin Then
        unMsg = "Les carrefours début et fin doivent être différents"
    Else
        uneSaisieErronée = False
    End If
    
    If uneSaisieErronée Then
        'Affichage du message d'erreur
        MsgBox unMsg, vbCritical
    Else
        'Calcul et Affichage des bandes passantes pour les différentes vitesses
        CalculerEtAfficherBandes uneVDeb, uneVFin, unPas
        
        'Affichage des controls se trouvant sous FrameSeparateur
        Height = Height - ScaleHeight + MSChartVitBandes.Top + MSChartVitBandes.Height
            
        'Restauration des paramètres du site en cours qui ont été
        'modifiés lors du calcul
        monSite.monTypeOnde = maSaveTypeOnde
        monSite.monTypeVit = maSaveTypeVit
        monSite.maVitSensM = maSaveVitSensM
        monSite.maVitSensD = maSaveVitSensD
    End If
    
    'Restauration du pointeur souris
    MousePointer = vbDefault
End Sub

Private Sub BoutonFermer_Click()
    'Fermeture
    Unload Me
    
    'Affichage de la form fille active pour éviter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    'si on n'est pas en plein écran
    If monPleinEcranVisible = False Then
        frmMain.ActiveForm.Show
    End If
End Sub

Private Sub BoutonImp_Click()
    ' Active la routine de gestion d'erreur.
    On Error GoTo CancelPress
    'Affichage de la fenetre de configuration d'imprimante
    frmMain.dlgCommonDialog.CancelError = True
    frmMain.dlgCommonDialog.flags = cdlPDPrintSetup
    frmMain.dlgCommonDialog.ShowPrinter
    
    'Impression sur l'imprimante courante de la fenêtre sans les
    'boutons et le spread mais avec le nom de fichier, le titre
    'de l'étude et la durée du cycle
    
    'Modif de l'affichage de la fenêtre
    BoutonFermer.Visible = False
    TabInfoVit.Visible = False
    BoutonCalculer.Visible = False
    BoutonImp.Visible = False
    LabelTitre.Visible = True
    LabelNomFichier.Visible = True
    TabInfoVit.Row = 1
    TabInfoVit.Col = 4
    LabeLCarfDebFin.Caption = LabeLCarfDebFin.Caption + TabInfoVit.Text
    TabInfoVit.Row = 1
    TabInfoVit.Col = 5
    LabeLCarfDebFin.Caption = LabeLCarfDebFin.Caption + " et " + TabInfoVit.Text
    LabeLCarfDebFin.Visible = True
    
    'Modif du fond d'écran de la fenêtre en gris clair
    unSaveFond = BackColor
    BackColor = RGB(230, 230, 230)
    
    'Impression
    PrintForm
    
    'Restauration de l'affichage de la fenêtre
    BoutonFermer.Visible = True
    TabInfoVit.Visible = True
    BoutonCalculer.Visible = True
    BoutonImp.Visible = True
    LabelTitre.Visible = False
    LabelNomFichier.Visible = False
    LabeLCarfDebFin.Visible = False
    
    'Restauration du fond d'écran de la fenêtre
    BackColor = unSaveFond
    
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub 'Pour éviter le traitement des erreurs s'il n'y a pas eu
    
    'Gestion des erreurs
CancelPress:
    
    Select Case Err.Number
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim unCarf As Carrefour
    
    HelpContextID = IDhlp_OngletDesOnde
    LabelTitre.Caption = LabelTitre.Caption + monSite.TitreEtude.Text
    LabelNomFichier.Caption = LabelNomFichier.Caption + Mid(monSite.Caption, Len("Site : ") + 1)
    
    'Index pour l'aide
    HelpContextID = IDhlp_WinFindBande
        
    'Centrage de la fenetre
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'Masquage des controls se trouvant sous FrameSeparateur
    'Il apparaitront lors du click sur le bouton calculer
    Height = Height - ScaleHeight + FrameSeparateur.Top
            
    'Affichage de la durée du cycle
    LabelDureeCycle.Caption = "Durée du cycle : " + Format(monSite.maDuréeDeCycle) + " secondes"

    'On vide les listes des carrefours début et fin
    TabInfoVit.Row = 1
    TabInfoVit.Col = 4
    TabInfoVit.Action = 26 ' = SS_ACTION_COMBO_CLEAR
    TabInfoVit.Row = 1
    TabInfoVit.Col = 5
    TabInfoVit.Action = 26 ' = SS_ACTION_COMBO_CLEAR
    
    'Remplissage des listes de carrefours début et fin utilisés
    'dans le calcul de l'onde verte
    unYMin = 300000
    unYMax = -300000
    unNbCarfUtil = 0
    For i = 1 To monSite.mesCarrefours.Count
        Set unCarf = monSite.mesCarrefours(i)
            If unCarf.monDecCalcul <> -99 Then
            unNbCarfUtil = unNbCarfUtil + 1
            TabInfoVit.Row = 1
            TabInfoVit.Col = 4
            TabInfoVit.TypeComboBoxIndex = unNbCarfUtil - 1
            TabInfoVit.TypeComboBoxString = unCarf.monNom
            TabInfoVit.Row = 1
            TabInfoVit.Col = 5
            TabInfoVit.TypeComboBoxIndex = unNbCarfUtil - 1
            TabInfoVit.TypeComboBoxString = unCarf.monNom
            'Recherche du carrefour dY min et d'Y max
            unY = DonnerYCarrefour(unCarf)
            If unY > unYMax Then
                unYMax = unY
                unIndCarfMax = unNbCarfUtil
            End If
            If unY < unYMin Then
                unYMin = unY
                unIndCarfMin = unNbCarfUtil
            End If
        End If
    Next i
    
    'Remplissage des carrefours début et fin
    TabInfoVit.Row = 1
    TabInfoVit.Col = 4
    TabInfoVit.TypeComboBoxIndex = unIndCarfMin - 1
    TabInfoVit.Text = TabInfoVit.TypeComboBoxString
    TabInfoVit.Row = 1
    TabInfoVit.Col = 5
    TabInfoVit.TypeComboBoxIndex = unIndCarfMax - 1
    TabInfoVit.Text = TabInfoVit.TypeComboBoxString
End Sub


Private Sub TabInfoVit_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    'Cas d'une saisie
    'Masquage des controls se trouvant sous FrameSeparateur
    Height = Height - ScaleHeight + FrameSeparateur.Top
End Sub

Private Sub CalculerEtAfficherBandes(uneVDeb As Integer, uneVFin As Integer, unPas As Integer)
    'Calcul et affichage des bandes passantes montantes et descendantes
    'suivant chaque vitesse montante et descendante.
    'Et affichage dans le contrôle MSChart MSChartVitBandes
    
    Dim unBlue As Long, unGreen As Long, unRed As Long
    Dim unRowLabelCount As Integer
    Dim unColumnLabelCount As Integer
    Dim unRowCount As Integer, unColumnCount As Integer
    Dim unTabFeuM() As Feu, unTabFeuD() As Feu
    Dim unCarf As Carrefour, uneV As Single
    Dim unYDeb As Integer, unYFin As Integer, unYTmp As Integer
    
    'Détermination des Y des carrefours début et fin
    'pour n'utiliser que les carrefours entre ces deux Y
    TabInfoVit.Row = 1
    TabInfoVit.Col = 4
    unYDeb = DonnerYCarrefour(monSite.mesCarrefours(TrouverCarfParNom(TabInfoVit.Text)))
    TabInfoVit.Row = 1
    TabInfoVit.Col = 5
    unYFin = DonnerYCarrefour(monSite.mesCarrefours(TrouverCarfParNom(TabInfoVit.Text)))
    'Tri pour avoir YDeb > YFin
    If unYDeb > unYFin Then
        unYTmp = unYFin
        unYFin = unYDeb
        unYDeb = unYTmp
    End If
    
    'Remplissage des tableaux de feux montants et descendants
    unNbFeuxM = 0
    unNbFeuxD = 0
    unNbCarf = monSite.mesCarrefours.Count
    For i = 1 To unNbCarf
        Set unCarf = monSite.mesCarrefours(i)
        unYTmp = DonnerYCarrefour(unCarf)
        If unCarf.monDecCalcul <> -99 And unYTmp >= unYDeb And unYTmp <= unYFin Then
            unNbFeux = unCarf.mesFeux.Count
            For j = 1 To unNbFeux
                If unCarf.mesFeux(j).monSensMontant Then
                    'Retaillage dynamique du tableau des feux montants
                    unNbFeuxM = unNbFeuxM + 1
                    ReDim Preserve unTabFeuM(1 To unNbFeuxM)
                    'Stockage du feu montant
                    Set unTabFeuM(unNbFeuxM) = unCarf.mesFeux(j)
                Else
                    'Retaillage dynamique du tableau des feux descendants
                    unNbFeuxD = unNbFeuxD + 1
                    ReDim Preserve unTabFeuD(1 To unNbFeuxD)
                    'Stockage du feu descendant
                    Set unTabFeuD(unNbFeuxD) = unCarf.mesFeux(j)
                End If
            Next j
        End If
    Next i
    
    'Remplissage et configuration du graphique à barres 2D
    MSChartVitBandes.chartType = VtChChartType2dBar
    With MSChartVitBandes.DataGrid
        ' Paramètre le graphique à l'aide de méthodes.
        unRowLabelCount = 1
        unColumnLabelCount = 1
        unRowCount = Int((uneVFin - uneVDeb) / unPas) + 1
        unColumnCount = 2
        .SetSize unRowLabelCount, unColumnLabelCount, unRowCount, unColumnCount

        ' Insère des données
        '.RandomDataFill
        For i = 1 To unRowCount
            MSChartVitBandes.Row = i
            'Calcul d'une vitesse montante et descendante
            uneV = uneVDeb + (i - 1) * unPas
            
            'Vérification du passage à tous les verts montants
            'avec la vitesse choisi dans les 2 sens.
            'La valeur retournée est la bande passante trouvée dans
            'le sens considéré
            MSChartVitBandes.Column = 1
            If unNbFeuxM <> 0 Then
                uneBande = VerifierVitessePasseToutVert(monSite, uneV, unTabFeuM, True)
            Else
                uneBande = 0
            End If
            MSChartVitBandes.Data = uneBande
                        
            MSChartVitBandes.Column = 2
            If unNbFeuxD <> 0 Then
                uneBande = VerifierVitessePasseToutVert(monSite, uneV, unTabFeuD, False)
            Else
                uneBande = 0
            End If
            MSChartVitBandes.Data = uneBande
        Next i
        
        'Positionnement de la légende
        MSChartVitBandes.Legend.Location.LocationType = VtChLocationTypeTop
        MSChartVitBandes.Legend.VtFont.Name = "Arial"
        MSChartVitBandes.Legend.VtFont.Size = 8
            
        'Affichage des axes
        For unAxeID = VtChAxisIdX To VtChAxisIdY
            With MSChartVitBandes.Plot.Axis(unAxeID, 1).AxisTitle
                .VtFont.Size = 12
                .VtFont.Name = "Arial"
                .VtFont.Style = VtFontStyleBold
                .Visible = True
                Select Case unAxeID
                    Case 0
                        .Text = "Vitesse montante et descendante"
                    Case 1
                        .Text = "Largeur de bandes passantes"
                End Select
            End With
        Next
        
        ' Ajoute ensuite des étiquettes au premier niveau.
        labelIndex = 1
        unColumn = 1
        .ColumnLabel(unColumn, labelIndex) = "Sens M"
        MSChartVitBandes.Plot.SeriesCollection(unColumn).LegendText = "Montante"
        With MSChartVitBandes.Plot.SeriesCollection.Item(unColumn).DataPoints(-1)
            ' Attribue la couleur bleu au point de données.
            .Brush.Style = VtBrushStyleSolid
            ' Associe la couleur bleu.
            unBlue = monSite.mesOptionsAffImp.maCoulBandComM \ CarreDe256
            unGreen = (monSite.mesOptionsAffImp.maCoulBandComM Mod CarreDe256) \ 256
            unRed = monSite.mesOptionsAffImp.maCoulBandComM - unBlue * CarreDe256 - unGreen * 256
            .Brush.FillColor.Set unRed, unGreen, unBlue
        End With
        
        unColumn = 2
        .ColumnLabel(unColumn, labelIndex) = "Sens D"
        MSChartVitBandes.Plot.SeriesCollection(unColumn).LegendText = "Descendante"
        With MSChartVitBandes.Plot.SeriesCollection.Item(unColumn).DataPoints(-1)
            ' Attribue la couleur bleu au point de données.
            .Brush.Style = VtBrushStyleSolid
            ' Associe la couleur bleu.
            unBlue = monSite.mesOptionsAffImp.maCoulBandComD \ CarreDe256
            unGreen = (monSite.mesOptionsAffImp.maCoulBandComD Mod CarreDe256) \ 256
            unRed = monSite.mesOptionsAffImp.maCoulBandComD - unBlue * CarreDe256 - unGreen * 256
            .Brush.FillColor.Set unRed, unGreen, unBlue
        End With
        
        'Affichage des valeurs des vitesses horizontalement
        For i = 1 To unRowCount
            unRow = i
            .RowLabel(unRow, labelIndex) = Format(uneVDeb + (i - 1) * unPas)
        Next i
        
        'Suppression du pivotement automatique des labels
        '==> on les garde horizontale
        MSChartVitBandes.Plot.Axis(VtChAxisIdX, 1).Labels(1).Auto = False
        'Choix de la taille suivant le nombre de graduations en X,
        'donc de vitesses
        If unRowCount > 26 Then
            uneSize = 6
        Else
            uneSize = 8
        End If
        MSChartVitBandes.Plot.Axis(VtChAxisIdX, 1).Labels(1).VtFont.Size = uneSize
    End With
End Sub
