VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmInfoTpsTC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dessin du Temps parcours TC en fonction de l'instant de départ"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmInfoTpsTC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChartTempsTC 
      Height          =   6255
      Left            =   120
      OleObjectBlob   =   "frmInfoTpsTC.frx":0442
      TabIndex        =   9
      Top             =   1800
      Width           =   10575
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton BoutonDessiner 
      Caption         =   "Dessiner"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton BoutonPrint 
      Caption         =   "Imprimer..."
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton BoutonFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Default         =   -1  'True
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox ComboTC 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label LabelTmpMoyen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Temps de parcours moyen = "
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
      Top             =   1440
      Width           =   2460
   End
   Begin VB.Label LabelInfoVert 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plage de vert du départ du TC : Inconnue"
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
      Top             =   1080
      Width           =   3570
   End
   Begin VB.Label DureeCycle 
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
      TabIndex        =   5
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Affichage du temps de parcours du TC : "
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
      Top             =   600
      Width           =   3465
   End
End
Attribute VB_Name = "frmInfoTpsTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant les temps de parcours TC calculés
Private maColTpsTC As New Collection
'Variable stockant le dernier TC dont on a fait le dessin
Private monLastTC As TC

Private Sub BoutonDessiner_Click()
    'Dessin de Tparcours du TC choisi = F (son Tdépart)
    Dim unTParcours As Single, unTmpMoyen As Single
    Dim unePhase As PhaseTabMarche
    'Variables stockant les temps de parcours min et max
    Dim unTpsMin As Single
    Dim unTpsMax As Single
    Dim unIndFeu As Integer
    
    unTmpMoyen = 0

    'Stockage en tant que dernier TC ayant servi au dessin
    Set monLastTC = monSite.mesTC(ComboTC.ListIndex)
    
    'Modif pointeur souris en sablier
    MousePointer = vbHourglass
    
    'Initialisation des temps min et max
    unTpsMin = 1000000#
    unTpsMax = -1000000#
    
    'Détermination de la plage pendant laquelle le TC
    'démarre au vert
    If DonnerYCarrefour(monLastTC.monCarfDep) >= DonnerYCarrefour(monLastTC.monCarfArr) Then
        'Cas d'un TC descendant
        unYDeb = DonnerYMaxCarfSens(monLastTC.monCarfDep, False, unIndFeu)
    Else
        'Cas d'un TC montant
        unYDeb = DonnerYMinCarfSens(monLastTC.monCarfDep, True, unIndFeu)
    End If
    uneDureeVert = monLastTC.monCarfDep.mesFeux(unIndFeu).maDuréeDeVert
    LabelInfoVert = "Plage de vert du départ du TC : Entre 0 et " + Format(uneDureeVert) + " secondes"
    
    'Initialisation et affichage de la progress bar
    ProgressBar1.Min = 0
    ProgressBar1.Max = monSite.maDuréeDeCycle - 1
    ProgressBar1.Visible = True
    
    'On vide les anciens temps de parcours calculés
    For i = 1 To maColTpsTC.Count
        maColTpsTC.Remove 1
    Next i
    
    'Sauvegarde de l'instant de départ du TC servant au dessin
    unTDepSave = monLastTC.monTDep
    
    'Calcul des temps de parcours par pas de 1 seconde
    'entre 0 et durée de cycle - 1 seconde
    For i = 0 To monSite.maDuréeDeCycle - 1
        'Modif de l'instant de départ
        monLastTC.monTDep = i
        
        'Calcul du tableau de marche de progresssion
        monLastTC.CalculerTableauMarcheProg
        
        'Calcul du temps de parcours (début dernière phase
        '+ durée dernière phase)
        Set unePhase = monLastTC.mesPhasesTMProg(monLastTC.mesPhasesTMProg.Count)
        unTParcours = unePhase.monTDeb + unePhase.maDureePhase - monLastTC.mesPhasesTMProg(1).monTDeb
        
        'Stockage du temps min et max
        If unTParcours > unTpsMax Then unTpsMax = unTParcours
        If unTParcours < unTpsMin Then unTpsMin = unTParcours
        
        'Ajout à la collection des temps de parcours
        maColTpsTC.Add unTParcours
        
        'Calcul du cumul des temps de parcours pour calculer le temps moyen
        unTmpMoyen = unTmpMoyen + unTParcours
        
        'Mise à jour de la barre de progression
        ProgressBar1.Value = i
    Next i
        
    'Calcul et affichage du temps de parcours moyen du TC
    unTmpMoyen = CInt(unTmpMoyen / monSite.maDuréeDeCycle)
    LabelTmpMoyen = "Temps moyen de parcours = " + Format(unTmpMoyen) + " secondes"
    
    'Restauration de l'instant de départ du TC servant au dessin
    monLastTC.monTDep = unTDepSave
    
    'Remplissage et Affichage du control MSChart
    RemplirGraphique maColTpsTC, unTpsMin, unTpsMax
    Height = MSChartTempsTC.Top + MSChartTempsTC.Height + 120
    
    'Initialisation et affichage de la progress bar
    ProgressBar1.Visible = False
    
    'Restauration pointeur souris
    MousePointer = vbDefault
End Sub

Private Sub BoutonFermer_Click()
    Unload Me
    
    'Libération de la mémoire de la collection des temps de
    'parcours calculés
    For i = 1 To maColTpsTC.Count
        maColTpsTC.Remove 1
    Next i
    Set maColTpsTC = Nothing
    
    'Affichage de la form fille active pour éviter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    'si on n'est pas en plein écran
    If monPleinEcranVisible = False Then
        frmMain.ActiveForm.Show
    End If
End Sub


Private Sub BoutonPrint_Click()
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
    BoutonDessiner.Visible = False
    BoutonPrint.Visible = False
    
    'Modif du fond d'écran de la fenêtre en gris clair
    unSaveFond = BackColor
    BackColor = RGB(230, 230, 230)
    
    'Impression
    PrintForm
    
    'Restauration de l'affichage de la fenêtre
    BoutonFermer.Visible = True
    BoutonDessiner.Visible = True
    BoutonPrint.Visible = True
    
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

Private Sub ComboTC_Click()
    If ComboTC.ListIndex = 0 Then
        BoutonDessiner.Enabled = False
        'Masquage du control MSChart
        Height = Height - ScaleHeight + BoutonPrint.Top
    Else
        BoutonDessiner.Enabled = True
        If monLastTC Is monSite.mesTC(ComboTC.ListIndex) Then
            'Affichage du control MSChart si on choisi le dernier
            'TC ayant servi au dessin
            Height = MSChartTempsTC.Top + MSChartTempsTC.Height + 120
        Else
            'Masquage du control MSChart si on choisi un TC
            'différent du dernier ayant servi au dessin
            Height = Height - ScaleHeight + BoutonPrint.Top
        End If
    End If
End Sub

Private Sub Form_Activate()
    BoutonFermer.SetFocus
End Sub

Private Sub Form_Load()
    Set monLastTC = Nothing
    DureeCycle.Caption = DureeCycle.Caption + Format(monSite.maDuréeDeCycle) + " secondes"
    
    'Index pour l'aide
    HelpContextID = IDhlp_WinTempsTC
    
    'Remplissage de la liste des TC
    ComboTC.AddItem "" 'TC vide en tête
    For i = 1 To monSite.mesTC.Count
        ComboTC.AddItem monSite.mesTC(i).monNom
    Next i
        
    'Centrage de la fenêtre
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'Ouverture sur le TC vide
    ComboTC.ListIndex = 0
    'Cette instruction déclenche le click event de comboTC
    '==> Recalcul du Height de la form
End Sub

Public Sub RemplirGraphique(uneCol As Collection, unTpsMin As Single, unTpsMax As Single)
    Dim unMax As Integer, unMin As Integer
    
    'Remplissage et configuration du graphique à barres 2D
    MSChartTempsTC.chartType = VtChChartType2dBar
    
    'Affichage de l'axe des Y entre les temps de parcours min et max
    With MSChartTempsTC.Plot.Axis(VtChAxisIdY, 1).ValueScale
        .Auto = False
        unMax = Int(unTpsMax) - Int(unTpsMax) Mod 10 + 10 'arrondi à la dizaine sup
        unMin = Int(unTpsMin) - Int(unTpsMin) Mod 10      'arrondi à la dizaine inf
        .Maximum = unMax
        .Minimum = unMin
        'Une division toutes les 10 secondes
        .MajorDivision = (unMax - unMin) / 10
    End With
        
    With MSChartTempsTC.DataGrid
        ' Paramètre le graphique à l'aide de méthodes.
        unRowLabelCount = 1
        unColumnLabelCount = 1
        unRowCount = monSite.maDuréeDeCycle
        unColumnCount = 1
        .SetSize unRowLabelCount, unColumnLabelCount, unRowCount, unColumnCount

        ' Insère des données
        '.RandomDataFill
        For i = 1 To uneCol.Count
            MSChartTempsTC.Row = i
            MSChartTempsTC.Column = 1
            MSChartTempsTC.Data = uneCol(i)
        Next i
        
        'Positionnement de la légende
        MSChartTempsTC.Legend.Location.LocationType = VtChLocationTypeTop
        MSChartTempsTC.Legend.VtFont.Name = "Arial"
        MSChartTempsTC.Legend.VtFont.Size = 8
            
        'Affichage des axes
        For unAxeID = VtChAxisIdX To VtChAxisIdY
            With MSChartTempsTC.Plot.Axis(unAxeID, 1).AxisTitle
                .VtFont.Size = 12
                .VtFont.Name = "Arial"
                .VtFont.Style = VtFontStyleBold
                .Visible = True
                Select Case unAxeID
                    Case 0
                        .Text = "Instant de départ du TC en secondes"
                    Case 1
                        .Text = "Temps de parcours du TC en secondes"
                End Select
            End With
        Next
        
        ' Ajoute la couleur des étiquettes au premier niveau.
        With MSChartTempsTC.Plot.SeriesCollection.Item(1).DataPoints(-1)
            ' Attribue le style au point de données.
            .Brush.Style = VtBrushStyleSolid
            ' Associe la couleur du TC.
            unBlue = monLastTC.maCouleur \ CarreDe256
            unGreen = (monLastTC.maCouleur Mod CarreDe256) \ 256
            unRed = monLastTC.maCouleur - unBlue * CarreDe256 - unGreen * 256
            .Brush.FillColor.Set unRed, unGreen, unBlue
        End With
                
        'Affichage des valeurs des instants de départ horizontalement
        For i = 1 To unRowCount
            unRow = i
            .RowLabel(unRow, 1) = Format(i - 1)
        Next i
        
        'Suppression du pivotement automatique des labels
        '==> on les garde horizontale
        MSChartTempsTC.Plot.Axis(VtChAxisIdX, 1).Labels(1).Auto = False
        uneSize = 6
        MSChartTempsTC.Plot.Axis(VtChAxisIdX, 1).Labels(1).VtFont.Size = uneSize
    End With
End Sub
