VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmImprimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimer"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmImprimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread TabFicheDataTC 
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   6840
      Visible         =   0   'False
      Width           =   9735
      _Version        =   131077
      _ExtentX        =   17171
      _ExtentY        =   2143
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   7
      MaxRows         =   10
      OperationMode   =   1
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmImprimer.frx":0442
      UnitType        =   2
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin FPSpread.vaSpread TabFicheDataCarf 
      Height          =   1815
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   9975
      _Version        =   131077
      _ExtentX        =   17595
      _ExtentY        =   3201
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   10
      MaxRows         =   10
      OperationMode   =   1
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmImprimer.frx":0F6F
      UnitType        =   2
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin FPSpread.vaSpread TabFicheRes 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   9495
      _Version        =   131077
      _ExtentX        =   16748
      _ExtentY        =   873
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   7
      MaxRows         =   10
      OperationMode   =   1
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmImprimer.frx":1BB6
      UnitType        =   2
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin VB.CommandButton BoutonConfig 
      Caption         =   "Configurer l'imprimante..."
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Tag             =   "Annuler"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton BoutonOptions 
      Caption         =   "Options d'affichage..."
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame FrameEchelle 
      Caption         =   "Echelle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   3495
      Begin VB.ComboBox ComboEchOrd 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmImprimer.frx":263F
         Left            =   1200
         List            =   "frmImprimer.frx":264F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox ComboEchTmp 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmImprimer.frx":2681
         Left            =   1200
         List            =   "frmImprimer.frx":268B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox CheckAjuster 
         Caption         =   "Ajuster au format papier"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label LabelEchOrd 
         AutoSize        =   -1  'True
         Caption         =   "En ordonn�e : "
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label LabelEchTmp 
         AutoSize        =   -1  'True
         Caption         =   "En temps : "
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.Frame FrameImprimer 
      Caption         =   "Imprimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3720
      TabIndex        =   18
      Top             =   600
      Width           =   2895
      Begin VB.CheckBox CheckNomFichier 
         Caption         =   "Nom du fichier"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox CheckTitre 
         Caption         =   "Titre de l'�tude"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox CheckImpDes 
         Caption         =   "Graphique Onde Verte"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox CheckImpRes 
         Caption         =   "Fiche R�sultats"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox CheckImpTC 
         Caption         =   "Donn�es Transports Collectifs"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CheckBox CheckImpCarf 
         Caption         =   "Donn�es Carrefours"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.Frame FrameParamGen 
      Caption         =   "Param�tres g�n�raux"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   3495
      Begin VB.TextBox NbSecondes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "10"
         Top             =   1395
         Width           =   285
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2400
         TabIndex        =   5
         Top             =   1353
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "NbSecondes"
         BuddyDispid     =   196628
         OrigLeft        =   2450
         OrigTop         =   1335
         OrigRight       =   2690
         OrigBottom      =   1710
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox ComboTailleLigne 
         Height          =   315
         ItemData        =   "frmImprimer.frx":26A9
         Left            =   2100
         List            =   "frmImprimer.frx":26BC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "secondes"
         Height          =   195
         Left            =   2700
         TabIndex        =   28
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lignes de rappel toutes les "
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Orientation : "
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Epaisseur de ligne : "
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Image ImagePortrait 
         Height          =   600
         Left            =   2100
         Picture         =   "frmImprimer.frx":26CF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image ImagePaysage 
         Height          =   495
         Left            =   2100
         Picture         =   "frmImprimer.frx":2E11
         Stretch         =   -1  'True
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Tag             =   "Annuler"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label NomImp 
      AutoSize        =   -1  'True
      Caption         =   "Imprimante courante :  "
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
      TabIndex        =   16
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmImprimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const TailleCellule = 210   ' en twips


Private Sub BoutonConfig_Click()
    If PlateformeNT Then
        'Affichage de la fenetre de configuration d'imprimante
        'par appel de celle de la dll, comdlg32.dll
        'car comdlg32.ocx est bugg� en NT
        '(orientation, taille papier inchangeable)
        'ShowPrinter fonction d�fini dans ModulePrintAPI.bas
        ShowPrinter Me, PD_PRINTSETUP
    Else
        'Sous plateforme non NT ( = 95 ou 98)
        ' Active la routine de gestion d'erreur.
        On Error GoTo CancelPress
        'Affichage de la fenetre de configuration d'imprimante
        frmMain.dlgCommonDialog.CancelError = True
        frmMain.dlgCommonDialog.flags = cdlPDPrintSetup
        frmMain.dlgCommonDialog.ShowPrinter
    End If
    
    'Mise � jour du nom de l'imprimante courante
    NomImp.Caption = "Imprimante courante : " + Printer.DeviceName
    'Mise � jour de l'orientation
    If Printer.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
    End If
    
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub 'Pour �viter le traitement des erreurs s'il n'y a pas eu
    
    'Gestion des erreurs sous Plateforme non NT
CancelPress:
    Select Case Err.Number
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BoutonOptions_Click()
    'Affichage de la fenetre d'options d'affichage
    'la variable public monCallOptionByPrint permet de signaler
    'que frmOptions est appel� par frmImprimer
    'monCallOptionByPrint est utilis� dans le code de frmOptions
    monCallOptionByPrint = True
    frmOptions.Show vbModal
    'Restauration
    monCallOptionByPrint = False
End Sub

Private Sub CheckAjuster_Click()
    'Si Ajuster � la taille du papier de l'imprimante
    'est vrai ==> Inhibition des �chelles Abscisse et ordonn�e
    'sinon activation de ces �chelles
    If CheckAjuster.Value = 0 Then
        'Cas o� ajuster est non coch�e
        unNonAjustement = True
    Else
        'Cas o� ajuster est coch�e
        unNonAjustement = False
    End If
    
    LabelEchTmp.Enabled = unNonAjustement
    LabelEchOrd.Enabled = unNonAjustement
    ComboEchTmp.Enabled = unNonAjustement
    ComboEchOrd.Enabled = unNonAjustement
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
    'Affichage de la form active pour �viter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    frmMain.ActiveForm.Show
End Sub

Private Sub cmdOK_Click()
    Dim unYreel As Long, unePos As Long
    Dim unCarf As Carrefour, unInd As Integer
    Dim unTC As TC, unHeader As String
    Dim uneLongImpAxeY As Long, uneLongueurReelAxeY As Long
    Dim unControl As Object
    Dim unX0 As Long, unY0 As Long, uneHt As Long, uneLg As Long
    Dim uneNewHt As Long, uneNewLg As Long
       
    'Toutes les valeurs de currentX, currentY ou des X et de Y
    'des m�thodes Line sont en twips
        
    If PlateformeNT Then
        'Stockage des param�tres de config d'imprimante
        'pour ce logiciel sous NT seulement
        With Printer
            SaveSetting "OndeV", "PrintOptions", "Orientation", .Orientation
            SaveSetting "OndeV", "PrintOptions", "Copies", .Copies
            SaveSetting "OndeV", "PrintOptions", "Duplex", .Duplex
            SaveSetting "OndeV", "PrintOptions", "PaperSize", .PaperSize
            SaveSetting "OndeV", "PrintOptions", "ColorMode", .ColorMode
            SaveSetting "OndeV", "PrintOptions", "Zoom", .Zoom
            SaveSetting "OndeV", "PrintOptions", "PrintQuality", .PrintQuality
        End With
    End If
    
    'Stockage pour les autres fois de l'�paisseur de trait d'impression
    monSite.mesOptionsAffImp.monEpaisseurLigne = Val(ComboTailleLigne.Text)
    
    'Stockage pour les autres fois de l'intervalle en secondes
    'entre chaque ligne de rappel
    monSite.mesOptionsAffImp.monNbSecondesRappel = Val(NbSecondes.Text)
    
    'Calcul de l'�tat de coh�rence entre les donn�es et les r�sultats
    'des calculs dans l'�tude en cours
    If monSite.TabFeux.Tab > 2 Then
        unEtatIncoherenceDataCalc = False
    Else
        unEtatIncoherenceDataCalc = monSite.maModifDataCarf Or monSite.maModifDataOndeTC = True Or monSite.maModifDataOnde
    End If
    
    'Test si une ou plusieurs donn�es du calcul d'onde ont
    'chang� ou si incoh�rence entre donn�es et r�sultats
    '==> Pas d'impression des r�sultats et du dessin
    'd'onde verte tant qu'il y a incoh�rence
    If monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul Or unEtatIncoherenceDataCalc Or monSite.maCoherenceDataCalc = CalculImpossible Then
        If monSite.maCoherenceDataCalc = CalculImpossible Then
            unMsgMilieu = unMsgMilieu + "Raison : Le calcul d'onde verte est impossible avec les donn�es de ce site."
            If monSite.monTypeOnde = 3 And monSite.monTCM = 0 And monSite.monTCD = 0 Then
                unMsgMilieu = unMsgMilieu + Chr(13) + "En effet, dans l'onglet Cadrage Onde Verte, aucun TC montant et/ou descendant n'ont �t� choisis." + Chr(13) + Chr(13) + "Calcul d'onde verte prenant en compte les TC impossible"
            End If
        ElseIf monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul Or unEtatIncoherenceDataCalc Then
            monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul
            unMsgMilieu = unMsgMilieu + "Raison : une ou plusieurs donn�es du calcul d'onde verte ont chang�." + Chr(13)
            unMsgMilieu = unMsgMilieu + "Ces donn�es sont incoh�rentes avec les r�sultats des calculs pr�c�dant ces changements."
        End If
        
        If CheckImpRes.Value = 1 Or CheckImpDes.Value = 1 Then
            CheckImpRes.Value = 0
            CheckImpDes.Value = 0
            unMsg = "Impression des �l�ments s�lectionn�s sauf de la Fiche R�sultats et du Graphique onde verte" + Chr(13) + Chr(13)
            unMsg = unMsg + unMsgMilieu + Chr(13) + Chr(13)
            unMsg = unMsg + "Vous pouvez recalculer les ondes vertes en s�lectionnant l'un des 3 onglets suivants :" + Chr(13)
            unMsg = unMsg + "     - R�sultat d�calages" + Chr(13)
            unMsg = unMsg + "     - Dessin onde verte" + Chr(13)
            unMsg = unMsg + "     - Fiche R�sultats"
            unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
            If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
                'Annulation de l'impression
                cmdCancel_Click
                Exit Sub
            End If
        End If
    End If
        
    'Cadre du dessin total (2 cm de marge gauche et droite
    ' et 2 cm en haut et en bas)
    uneMargeAutour = 2 * unTwipToCm
       
    'Impression des fiches choisies par l'utilisateur
    For K = 1 To 4
        If K = 1 And CheckImpCarf.Value = 1 Then
            'Initialisation de l'ent�te
            unHeader = InitialiserEntete
            'Impression des donn�es des carrefours
            ImprimerFicheDonneesCarrefour unHeader
        ElseIf K = 2 And CheckImpTC.Value = 1 And monSite.mesTC.Count > 0 Then
            'Initialisation de l'ent�te
            unHeader = InitialiserEntete
            'Impression des donn�es des TC
            ImprimerFicheDonneesTC unHeader
        ElseIf K = 3 And CheckImpRes.Value = 1 Then
            'Initialisation de l'ent�te
            unHeader = InitialiserEntete
            'Impression de la Fiche R�sultats
            ImprimerFicheResultats unHeader
        ElseIf K = 4 And CheckImpDes.Value = 1 Then
            'Impression du dessin des ondes vertes
            
            'Initialisation de la fonte et de sa taille
            Printer.Font.Name = monSite.Font.Name
            Printer.Font.Size = 8
            
            'Affectation de la couleur � noir
            Printer.ForeColor = 0
            
            'Dessin du cadre englobant total de l'imprimante
            'Il sera affich� autour du dessin des ondes vertes
            unDec = 10
            Printer.CurrentX = 0
            Printer.CurrentY = 0
            Printer.Line (0, 0)-(0, Printer.ScaleHeight)
            Printer.Line (0, 0)-(Printer.ScaleWidth, 0)
            Printer.Line (Printer.ScaleWidth - unDec, 0)-(Printer.ScaleWidth - unDec, Printer.ScaleHeight - unDec)
            Printer.Line (0, Printer.ScaleHeight - unDec)-(Printer.ScaleWidth - unDec, Printer.ScaleHeight - unDec)
            
            'Initialisation du d�calage entre les lignes de texte
            unDecalageTexte = 0
            
            'Utilisation de l'�paisseur de trait choisi
            Printer.DrawWidth = monSite.mesOptionsAffImp.monEpaisseurLigne
            
            'Mise en gras des textes
            Printer.FontBold = True
            
            'Affichage du Titre en haut
            If CheckTitre.Value Then
                unDecalageTexte = Printer.TextHeight(monSite.TitreEtude.Text)
                Printer.CurrentX = uneMargeAutour
                Printer.CurrentY = uneMargeAutour - unDecalageTexte
                Printer.ForeColor = monSite.mesOptionsAffImp.maCoulTitreEch
                Printer.Print monSite.TitreEtude.Text
            End If
            
            'Affichage du nom de fichier
            If CheckNomFichier.Value Then
                uneString = Mid(monSite.Caption, Len("Site : ") + 1)
                Printer.CurrentX = uneMargeAutour
                Printer.CurrentY = uneMargeAutour + unDecalageTexte - Printer.TextHeight(uneString)
                Printer.ForeColor = monSite.mesOptionsAffImp.maCoulTitreEch
                Printer.Print uneString
                unDecalageTexte = unDecalageTexte + Printer.TextHeight(uneString)
            End If
               
            'Cadre du dessin d'onde verte (2 cm de marge gauche et droite
            ' et 2 cm en haut et en bas)
            unX0 = uneMargeAutour + 5 * unTwipToCm
            unY0 = Printer.ScaleHeight - uneMargeAutour
            uneLg = Printer.ScaleWidth - 5 * unTwipToCm - 2 * uneMargeAutour
            uneHt = Printer.ScaleHeight - 3 * unTwipToCm - 2 * uneMargeAutour
            
            'Restauration de la couleur � noir
            Printer.ForeColor = 0
            
            'Affichage de la dur�e du cycle � cot� des bandes passantes
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = "Cycle de " + Format(monSite.maDur�eDeCycle) + " secondes"
            Printer.CurrentX = unX0 + uneLg - Printer.TextWidth(uneString)
            Printer.Print uneString
            
            'Affichage des noms de TC utilis�s avec leur couleur
            'pour faire une l�gende en dessous de la valeur du cycle
            unDec = 0
            If monSite.mesTCutil.Count > 0 Then
                For i = 1 To monSite.mesTCutil.Count
                    uneString = monSite.mesTCutil(i).monNom + " "
                    unDec = unDec + Printer.TextWidth(uneString)
                    Printer.CurrentY = uneMargeAutour + 2 * unDecalageTexte
                    'Cadrage sur le bord droit de la feuille
                    Printer.CurrentX = unX0 + uneLg - unDec
                    Printer.ForeColor = monSite.mesTCutil(i).maCouleur
                    Printer.Print uneString
                Next i
                'Affichage devant les noms de TC de "TC : " en noir
                Printer.ForeColor = 0
                Printer.CurrentY = uneMargeAutour + 2 * unDecalageTexte
                Printer.CurrentX = unX0 + uneLg - unDec - Printer.TextWidth("TC : ")
                Printer.Print "TC : "
            End If
            
            'Affichage des bandes passantes
            Printer.CurrentX = uneMargeAutour
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = "Bandes passantes en secondes :"
            Printer.Print uneString
            unDecalageTexte = unDecalageTexte + Printer.TextHeight(uneString)
            
            Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
            Printer.CurrentX = uneMargeAutour
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = "   Montante"
            Printer.Print uneString
            Printer.CurrentX = uneMargeAutour + Printer.TextWidth("   Descendante  ")
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = Format(monSite.maBandeModifM)
            Printer.Print uneString
            unDecalageTexte = unDecalageTexte + Printer.TextHeight(uneString)
            
            Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
            Printer.CurrentX = uneMargeAutour
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = "   Descendante"
            Printer.Print uneString
            Printer.CurrentX = uneMargeAutour + Printer.TextWidth("   Descendante  ")
            Printer.CurrentY = uneMargeAutour + unDecalageTexte
            uneString = Format(monSite.maBandeModifD)
            Printer.Print uneString
            
            'Restauration de la couleur � noir
            Printer.ForeColor = 0
            
            'Calcul des �chelles suivant le choix de l'utilisateur
            'Pour cela il faut dabord lancer DessinerTout qui calcule les champs
            'monTmpTotal et monDYTotal de monSite si l'onglet Dessin onde
            'verte n'est pas actif, car ces champs ont d�j� �t� calcul�s
            If monSite.TabFeux.Tab <> 4 Then
                DessinerTout monSite.ZoneDessin, unX0, unY0, uneLg, uneHt, True
            End If
            
            If CheckAjuster.Value Then
                'Cas o� on veut ajuster le dessin dans son cadre
                'Calcul des �chelles
                uneEchT = unTwipToCm * monSite.monTmpTotal / uneLg
                uneEchY = unTwipToCm * monSite.monDYTotal / uneHt
                'Affectation de la hauteur et du cadre du dessin d'onde verte
                uneNewLg = uneLg
                uneNewHt = uneHt
            Else
                'Cas o� l'�chelle est fix�e par l'utilisateur
                'Affectation des nouvelles �chelles
                uneEchT = DonnerEchelleT
                uneEchY = DonnerEchelleY / 100
                'Affectation de la hauteur et du cadre du dessin d'onde verte
                uneNewLg = unTwipToCm * monSite.monTmpTotal / uneEchT
                uneNewHt = unTwipToCm * monSite.monDYTotal / uneEchY
                'Avertissement si les �chelles choisis g�n�re un dessin d'onde
                'verte et de progression TC en dehors du cadre affect� � ce dessin
                If uneNewHt > uneHt Or uneNewLg > uneLg Then
                    unMsg = "Avec votre choix d'�chelle, le dessin va d�passer le cadre qui lui est destin�."
                    unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
                    If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
                        'Cas de confirmation n�gative ==> Annulation
                        cmdCancel_Click
                        Exit Sub
                    End If
                End If
            End If
            
            'Affichage des �chelles
            Printer.ForeColor = monSite.mesOptionsAffImp.maCoulTitreEch
            uneString = "1 cm = " + Format(uneEchT, "###0.0") + " s"
            unDebEchT = unX0 + uneLg - Printer.TextWidth(uneString) / 2
            Printer.CurrentX = unDebEchT
            Printer.CurrentY = unY0 + Printer.TextHeight(uneString) * 3
            Printer.Print uneString
            uneString = "1 cm = " + Format(uneEchY, "###0.0") + " m"
            Printer.CurrentX = unX0 - Printer.TextWidth(uneString) / 2
            unDec = Printer.TextHeight(uneString)
            Printer.CurrentY = unY0 - uneHt - 3 * unDec
            Printer.Print uneString
            
            'Dessin de l'axe des Y en noir
            Printer.ForeColor = 0
            Printer.Line (unX0 - unDec, unY0 + 3.5 * unDec)-(unX0 - unDec, unY0 - uneHt - 2 * unDec)
            Printer.Line (unX0 - unDec, unY0 - uneHt - 2 * unDec)-(unX0 - unDec / 2, unY0 - uneHt - unDec)
            Printer.Line (unX0 - unDec, unY0 - uneHt - 2 * unDec)-(unX0 - unDec * 1.5, unY0 - uneHt - unDec)
            
            'Dessin de l'axe des temps en noir
            Printer.Line (unX0 - unDec, unY0 + 3.5 * unDec)-(unDebEchT - 60, unY0 + 3.5 * unDec)
            Printer.Line (unDebEchT - 60, unY0 + 3.5 * unDec)-(unDebEchT - 60 - unDec, unY0 + 3 * unDec)
            Printer.Line (unDebEchT - 60, unY0 + 3.5 * unDec)-(unDebEchT - 60 - unDec, unY0 + 4 * unDec)
            
            'Dessin de l'onde verte et des plages de vert de tous les feux
            'et des progressions de TC �ventuelles
            'Le False dit qu'on ne dessine pas dans l'onglet Dessin Onde Verte
            DessinerTout Printer, unX0, unY0, uneNewLg, uneNewHt, False
            
            'Calcul de la longueur r�elle de l'englobant en Y
            'de tous carrefours utilis�s dans le calcul de l'onde
            uneLongueurReelAxeY = monSite.monYMaxFeuUtil - monSite.monYMinFeuUtil
            
            'Calcul de la longueur imprimante de l'axe des ordonn�es
            uneLongImpAxeY = uneNewHt
            
            'Impression des nom de carrefours au bon zoom
            For i = 1 To monSite.mesCarrefours.Count
                Set unCarf = monSite.mesCarrefours(i)
                If unCarf.monDecCalcul <> -99 Then
                    'Calcul du Y carrefour = barycentre des Y de ses Feux
                    unYreel = DonnerYCarrefour(unCarf)
                    'Distance par rapport au Y max des feux des carrefours
                    'utilis�s pour le calcul de l'onde
                    '(zoom calcul� � partir de ce point)
                    unYreel = monSite.monYMaxFeuUtil - unYreel
                    
                    'Conversion du Yr�el du carrefour en Y imprimante
                    unePos = ConvertirReelEnEcran(unYreel, uneLongueurReelAxeY, uneLongImpAxeY)
                    unYimp = unePos + unY0 - uneLongImpAxeY - Printer.TextHeight(unCarf.monNom) / 2
                    
                    'Affichage de la vitesse montante au carrefour avec une fl�che
                    'en dessous du Y imprimante du Y r�el du carrefour
                    'si le carrefour a des feux montants
                    If unCarf.monCarfRed.HasFeuMontant Then
                        'Positionnement de la couleur d'avant plan
                        Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
                        'Texte correspondant � 3 chiffres
                        uneString = "000"
                        'Impression d'une fl�che montante
                        unX1 = unX0 - unDec - 120
                        unY1 = unYimp + Printer.TextHeight(uneString) * 2
                        unY2 = unYimp + Printer.TextHeight(uneString)
                        Printer.Line (unX1, unY1)-(unX1, unY2)
                        Printer.Line (unX1, unY2)-(unX1 - 30, unY2 + 120)
                        Printer.Line (unX1, unY2)-(unX1 + 30, unY2 + 120)
                        'Stockage du d�calage en X de l'impression de la vitesse
                        unX1 = unX1 - 60 - Printer.TextWidth(uneString)
                        'Impression de la vitesse montante en km/h
                        uneString = Format(CInt(unCarf.DonnerVitSens(True) * 3.6))
                        Printer.CurrentX = unX1
                        Printer.CurrentY = unYimp + Printer.TextHeight(uneString)
                        Printer.Print uneString
                    End If
                    
                    'Affichage de la vitesse descendante au carrefour avec une fl�che
                    'au dessus du Y imprimante du Y r�el du carrefour
                    'si le carrefour a des feux descendants
                    If unCarf.monCarfRed.HasFeuDescendant Then
                        'Positionnement de la couleur d'avant plan
                        Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
                        'Texte correspondant � 3 chiffres
                        uneString = "000"
                        'Impression d'une fl�che descendante
                        unX1 = unX0 - unDec - 120
                        unY1 = unYimp - Printer.TextHeight(uneString)
                        unY2 = unYimp
                        Printer.Line (unX1, unY1)-(unX1, unY2)
                        Printer.Line (unX1, unY2)-(unX1 - 30, unY2 - 120)
                        Printer.Line (unX1, unY2)-(unX1 + 30, unY2 - 120)
                        'Stockage du d�calage en X de l'impression de la vitesse
                        unX1 = unX1 - 60 - Printer.TextWidth(uneString)
                        'Impression de la vitesse descendante en km/h en > 0
                        uneString = Format(CInt(-unCarf.DonnerVitSens(False) * 3.6))
                        Printer.CurrentX = unX1
                        Printer.CurrentY = unYimp - Printer.TextHeight(uneString)
                        Printer.Print uneString
                    End If
                    
                    'Affichage du nom du carrefour en un Y imprimante
                    'correspondant au Y r�el calcul� avant
                    Printer.ForeColor = monSite.mesOptionsAffImp.maCoulNomCarf
                    Printer.CurrentY = unYimp
                    'Cadrage � droite de l'axe des Y des noms de carrefours
                    ' � cot� des vitesses montantes et descendantes
                    Printer.CurrentX = unX1 - Printer.TextWidth(unCarf.monNom) - 60
                    Printer.Print unCarf.monNom
                    
                    'Affichage des Y des feux du carrefours � droite
                    'du dessin ==> On �crit dans la marge droite
                    Printer.Font.Size = 5 'Affichage en petit
                    For j = 1 To unCarf.mesFeux.Count
                        'Distance par rapport au Y max des feux des carrefours
                        'utilis�s pour le calcul de l'onde
                        '(zoom calcul� � partir de ce point)
                        unYreel = monSite.monYMaxFeuUtil - unCarf.mesFeux(j).monOrdonn�e
                        
                        'Conversion du Yr�el du feu en Y imprimante
                        uneString = Format(unCarf.mesFeux(j).monOrdonn�e) + " m"
                        unePos = ConvertirReelEnEcran(unYreel, uneLongueurReelAxeY, uneLongImpAxeY)
                        unYimp = unePos + unY0 - uneLongImpAxeY - Printer.TextHeight(unString) / 2
                        
                        'Impression dans la marge de droite avec la
                        'couleur du sens du feu
                        Printer.CurrentX = unX0 + uneLg + 60
                        Printer.CurrentY = unYimp
                        If unCarf.mesFeux(j).monSensMontant Then
                            Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComM
                        Else
                            Printer.ForeColor = monSite.mesOptionsAffImp.maCoulBandComD
                        End If
                        Printer.Print uneString
                    Next j
                    'Restauration de la taille de la fonte
                    Printer.Font.Size = 8
                End If
            Next i
            
            'Impression des arr�ts TC au bon zoom
            For i = 1 To monSite.mesTC.Count
                Set unTC = monSite.mesTC(i)
                For j = 1 To unTC.mesArrets.Count
                    unYreel = monSite.monYMaxFeuUtil - unTC.mesArrets(j).monOrdonnee
                    'Conversion du Yr�el en Y �cran dans la FrameVisuCarf
                    unePos = ConvertirReelEnEcran(unYreel, uneLongueurReelAxeY, uneLongImpAxeY)
                    'Affichage en Y imprimnte du nom de l'arr�t TC
                    'correspondant au Y r�el calcul� avant
                    Printer.ForeColor = monSite.mesOptionsAffImp.maCoulNomArret
                    Printer.CurrentY = unePos + unY0 - uneLongImpAxeY - Printer.TextHeight(unTC.mesArrets(j).monLibelle)
                    'Cadrage � gauche de la frame FrameCarfTC des noms d'arr�ts
                    Printer.CurrentX = uneMargeAutour
                    uneString = unTC.mesArrets(j).monLibelle + " (Y = " + Format(unTC.mesArrets(j).monOrdonnee) + " m)"
                    Printer.Print uneString
                    'Dessin d'une ligne de l'arr�t jusqu'� l'axe des Y
                    unX1 = uneMargeAutour + Printer.TextWidth(uneString)
                    unX2 = unX0 - unDec 'D�but de l'axe des Y cf plus haut
                    unY1 = unePos + unY0 - uneLongImpAxeY
                    unY2 = unY1
                    Printer.Line (unX1, unY1)-(unX2, unY2), monSite.mesOptionsAffImp.maCoulNomArret
                Next j
            Next i
            
            'Restauration de la couleur � noir
            Printer.ForeColor = 0
            
            'Signature OndeV en bas
            If maDemoVersion Then
                uneString = "OndeV version 1.0 DEMO"
            Else
                uneString = "OndeV version 1.0"
            End If
            Printer.CurrentX = uneMargeAutour
            Printer.CurrentY = unY0 + Printer.TextHeight(uneString) * 3
            Printer.Print uneString
            
            'Envoi � l'imprimante
            Printer.EndDoc
        End If
    Next K
    
    'Fermeture de la fen�tre d'impression
    Unload Me
        
    'Affichage de la form active pour �viter l'apparition
    'en premier d'une autre fenetre windows (exemple un explorer)
    frmMain.ActiveForm.Show
End Sub



Private Sub Form_Load()
    'Index pour l'aide
    HelpContextID = IDhlp_PrintSite
    
    'Retaillage de la fen�tre d'impression pour cacher les spreads
    'invisible servant � l'impression des donn�es et fiches
    Width = cmdOK.Left + cmdOK.Width + NomImp.Left + Width - ScaleWidth
    Height = FrameImprimer.Top + FrameImprimer.Height + NomImp.Top + Height - ScaleHeight
    
    'Affichage de l'imprimante par d�faut
    NomImp.Caption = "Imprimante courante : " + Printer.DeviceName
    
    If PlateformeNT Then
        'R�cup des param�tres de config d'imprimante
        'pour ce logiciel sous NT seulement
        On Error Resume Next
        With Printer
            .Orientation = GetSetting("OndeV", "PrintOptions", "Orientation", .Orientation)
            .Copies = GetSetting("OndeV", "PrintOptions", "Copies", .Copies)
            .Duplex = GetSetting("OndeV", "PrintOptions", "Duplex", .Duplex)
            .PaperSize = GetSetting("OndeV", "PrintOptions", "PaperSize", .PaperSize)
            .ColorMode = GetSetting("OndeV", "PrintOptions", "ColorMode", .ColorMode)
            .Zoom = GetSetting("OndeV", "PrintOptions", "Zoom", .Zoom)
            .PrintQuality = GetSetting("OndeV", "PrintOptions", "PrintQuality", .PrintQuality)
        End With
        On Error GoTo 0
    End If
    
    'Mise � jour de l'orientation
    If Printer.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
    End If
    
    'Affichage des premiers �l�ments dans les combobox
    ComboEchTmp.ListIndex = 0
    ComboEchOrd.ListIndex = 0
    
    'Affichage de l'�paisseur de trait stock�e
    ComboTailleLigne.Text = Format(monSite.mesOptionsAffImp.monEpaisseurLigne)
    'Affichage de l'intervalle en secondes entre chaque ligne de rappel stock�
    NbSecondes.Text = Format(monSite.mesOptionsAffImp.monNbSecondesRappel)
End Sub


Public Function DonnerEchelleT() As Integer
    'Retourne l'�quivalent en cm r�el d'un cm �cran
    'ou imprimante pour l'axe des Temps
    Select Case ComboEchTmp.ListIndex
    Case 0
        DonnerEchelleT = 10 '1 cm repr�sentera 10 secondes
    Case 1
        DonnerEchelleT = 20 '1 cm repr�sentera 20 secondes
    Case Else
        MsgBox "Erreur de programmation de OndeV dans DonnerEchelleT", vbCritical
    End Select
End Function

Public Function DonnerEchelleY() As Integer
    'Retourne l'�quivalent en cm r�el d'un cm �cran
    'ou imprimante pour l'axe des Ordonn�es
    Select Case ComboEchOrd.ListIndex
    Case 0
        DonnerEchelleY = 2000 '1 cm repr�sentera 2000 cm
    Case 1
        DonnerEchelleY = 5000 '1 cm repr�sentera 5000 cm
    Case 2
        DonnerEchelleY = 10000 '1 cm repr�sentera 10000 cm
    Case 3
        DonnerEchelleY = 20000 '1 cm repr�sentera 20000 cm
    Case Else
        MsgBox "Erreur de programmation de OndeV dans DonnerEchelleY", vbCritical
    End Select
End Function


Public Function InitialiserEntete() As String
    Dim unHeader As String
    
    'Initialisation de l'ent�te des pages d'impression de spread
    unHeader = "/fb1"
    
    'Affichage du Titre de l'�tude en ent�te
    If CheckTitre.Value Then
        unHeader = unHeader + monSite.TitreEtude.Text
    End If
    
    'Affichage du nom de fichier en ent�te
    If CheckNomFichier.Value Then
        If unHeader <> "/fb1" Then unHeader = unHeader + "/n"
        unHeader = unHeader + Mid(monSite.Caption, Len("Site : ") + 1)
    End If
    
    InitialiserEntete = unHeader
End Function

Public Sub ImprimerFicheResultats(unHeader As String)
    Dim unTitreFiche As String, unNomFiche As String
    
    'Impression de la Fiche R�sultats
    
    'Remplissage des spread de l'onglet Fiche R�sultats
    If RemplirFicheResultPourImp Then
    
        'Affectation d'un nom de fiche
        unNomFiche = "R�sultats"
        
        'Affectation d'un titre de fiche
        unTitreFiche = "R�sultats du calcul d'onde verte "
        If EstModifierManuel Then
            'Cas d'une modification manuelle des d�calages
            unTitreFiche = unTitreFiche + "avec d�calages modifi�s manuellement"
        Else
            If monSite.monTypeOnde = OndeDouble Then
                unTitreFiche = unTitreFiche + "� double sens"
            ElseIf monSite.monTypeOnde = OndeSensM Then
                unTitreFiche = unTitreFiche + "� sens privil�gi� montant"
            ElseIf monSite.monTypeOnde = OndeSensD Then
                unTitreFiche = unTitreFiche + "� sens privil�gi� descendant"
            ElseIf monSite.monTypeOnde = OndeTC Then
                unTitreFiche = unTitreFiche + "prenant en compte les TC"
            End If
        End If
        
        'Affectation des options d'impression du spread TabFicheRes
        ConfigurerSpreadToPrint TabFicheRes, unHeader, unNomFiche, unTitreFiche
                
        'Affectation de la couleur blanches aux lignes autres
        'que les ent�tes (celles lock�es)
        TabFicheRes.LockBackColor = RGB(255, 255, 255)
        'Affectation de la couleur des lignes d'ent�tes (non lock�es)
        TabFicheRes.BackColor = RGB(220, 220, 220) 'Gris clair
        
        With monSite
            'Remplissage � partir des spread de l'onglet Fiche R�sultats
            TabFicheRes.MaxRows = 5 + monSite.mesCarrefours.Count + monSite.mesTCutil.Count
            
            'Remplissage de la ligne rempla�ant l'ent�te de
            'la partie issue de TabFicheOnde
            RemplirLigneEnteteFicheOnde
            
            'Remplissage de la partie issue de TabFicheOnde
            For i = 1 To 2
                TabFicheRes.Row = i + 1
                TabFicheRes.RowHeight(i + 1) = TailleCellule
                monSite.TabFicheOnde.Row = i
                For j = 1 To monSite.TabFicheOnde.MaxCols
                    monSite.TabFicheOnde.Col = j
                    TabFicheRes.Col = j + 1
                    TabFicheRes.ForeColor = monSite.TabFicheOnde.ForeColor
                    TabFicheRes.Text = monSite.TabFicheOnde.Text
                Next j
                '7 �me colonne vide
                TabFicheRes.Col = 7
                TabFicheRes.Text = ""
            Next i

            'Stockage de la 1�re ligne � remplir issue de TabFicheCarf
            uneLigneDebut = 4
            
            'Remplissage de la ligne rempla�ant l'ent�te de
            'la partie issue de TabFicheCarf
            RemplirLigneEnteteFicheCarf uneLigneDebut
            
            'Remplissage de la partie issue de TabFicheCarf
            For i = 1 To monSite.mesCarrefours.Count
                TabFicheRes.Row = uneLigneDebut + i
                TabFicheRes.RowHeight(uneLigneDebut + i) = TailleCellule
                monSite.TabFicheCarf.Row = i
                For j = 1 To monSite.TabFicheCarf.MaxCols
                    monSite.TabFicheCarf.Col = j
                    TabFicheRes.Col = j
                    TabFicheRes.ForeColor = monSite.TabFicheCarf.ForeColor
                    TabFicheRes.Text = monSite.TabFicheCarf.Text
                Next j
            Next i
            
            'Stockage de la 1�re ligne � remplir issue de TabFicheTC
            uneLigneDebut = uneLigneDebut + monSite.mesCarrefours.Count + 1
            
            'Remplissage de la ligne rempla�ant l'ent�te de
            'la partie issue de TabFicheTC
            RemplirLigneEnteteFicheTC uneLigneDebut
            
            'Remplissage de la partie issue de TabFicheTC
            For i = 1 To monSite.mesTCutil.Count
                TabFicheRes.Row = uneLigneDebut + i
                TabFicheRes.RowHeight(uneLigneDebut + i) = TailleCellule
                monSite.TabFicheTC.Row = i
                For j = 1 To monSite.TabFicheTC.MaxCols
                    monSite.TabFicheTC.Col = j
                    TabFicheRes.Col = j
                    TabFicheRes.ForeColor = monSite.TabFicheTC.ForeColor
                    TabFicheRes.Text = monSite.TabFicheTC.Text
                Next j
            Next i
            
            'Impression du spread
            TabFicheRes.Action = 13 ' = SS_ACTION_PRINT
        End With
    Else
        MsgBox "Pas d'impression de la fiche de r�sultats car le calcul d'onde verte est impossible", vbCritical
    End If
End Sub

Public Sub RemplirLigneEnteteFicheOnde()
    'Remplissage de la ligne rempla�ant l'ent�te de
    'la partie issue de TabFicheOnde
    
    'Modif de la hauteur de la ligne correspondant
    'aux ent�tes des spreads de la Fiche R�sultats
    TabFicheRes.RowHeight(1) = TailleCellule * 2
    
    'S�lection de la ligne d'ent�te pour fond gris + fonte grasse
    TabFicheRes.Row = 1
    TabFicheRes.Col = -1
    TabFicheRes.Lock = False
    TabFicheRes.ForeColor = monSite.TabFicheOnde.ShadowText
    TabFicheRes.FontBold = True
    
    'Remplissage 1�re ligne
    TabFicheRes.Row = 1
    TabFicheRes.Col = 1
    TabFicheRes.Text = "Sens de parcours"
    TabFicheRes.Col = 2
    TabFicheRes.Text = "Temps de parcours (s)"
    TabFicheRes.Col = 3
    TabFicheRes.Text = "Bande passante (s)"
    TabFicheRes.Col = 4
    TabFicheRes.Text = "Vitesse max (km/h)"
    TabFicheRes.Col = 5
    TabFicheRes.Text = "Poids"
    TabFicheRes.Col = 6
    TabFicheRes.Text = "TC pris en compte"
    TabFicheRes.Col = 7
    TabFicheRes.Text = ""
    
    TabFicheRes.Row = 2
    TabFicheRes.Col = 1
    TabFicheRes.Text = "MONTANT"
    TabFicheRes.Row = 3
    TabFicheRes.Col = 1
    TabFicheRes.Text = "DESCENDANT"
End Sub

Public Sub RemplirLigneEnteteFicheCarf(uneLigneDebut)
    'Remplissage de la ligne rempla�ant l'ent�te de
    'la partie issue de TabFicheCarf
    
    'Modif de la hauteur de la ligne correspondant
    'aux ent�tes des spreads de la Fiche R�sultats
    TabFicheRes.RowHeight(uneLigneDebut) = TailleCellule * 2
    
    'S�lection de la ligne d'ent�te pour fond gris + fonte grasse
    TabFicheRes.Row = uneLigneDebut
    TabFicheRes.Col = -1
    TabFicheRes.Lock = False
    TabFicheRes.ForeColor = monSite.TabFicheOnde.ShadowText
    TabFicheRes.FontBold = True
    
    'Remplissage de la ligne d'ent�te
    TabFicheRes.Row = uneLigneDebut
    TabFicheRes.Col = 1
    TabFicheRes.Text = "Carrefour"
    TabFicheRes.Col = 2
    TabFicheRes.Text = "D�calages (s)"
    TabFicheRes.Col = 3
    TabFicheRes.Text = "R Capacit� Mont (%)"
    TabFicheRes.Col = 4
    TabFicheRes.Text = "R Capacit� Desc (%)"
    TabFicheRes.Col = 5
    TabFicheRes.Text = "Vitesse Mon (km/h)"
    TabFicheRes.Col = 6
    TabFicheRes.Text = "Vitesse Des (km/h)"
    TabFicheRes.Col = 7
    TabFicheRes.Text = "D�calage ouverture (s)"
End Sub

Public Sub RemplirLigneEnteteFicheTC(uneLigneDebut)
    'Remplissage de la ligne rempla�ant l'ent�te de
    'la partie issue de TabFicheTC
    
    'Modif de la hauteur de la ligne correspondant
    'aux ent�tes des spreads de la Fiche R�sultats
    TabFicheRes.RowHeight(uneLigneDebut) = TailleCellule * 2
    
    'S�lection de la ligne d'ent�te pour fond gris + fonte grasse
    TabFicheRes.Row = uneLigneDebut
    TabFicheRes.Col = -1
    TabFicheRes.Lock = False
    TabFicheRes.ForeColor = monSite.TabFicheOnde.ShadowText
    TabFicheRes.FontBold = True
    
    'Remplissage de la ligne d'ent�te
    TabFicheRes.Row = uneLigneDebut
    TabFicheRes.Col = 1
    TabFicheRes.Text = "Transport Collectif"
    
    If monSite.mesTCutil.Count > 0 Then
        'Cas o� il y a des TC utilis�s
        TabFicheRes.Col = 2
        TabFicheRes.Text = "Sens de parcours"
        TabFicheRes.Col = 3
        TabFicheRes.Text = "Instant de d�part (s)"
        TabFicheRes.Col = 4
        TabFicheRes.Text = "Nb d'arr�ts aux feux"
        TabFicheRes.Col = 5
        TabFicheRes.Text = "Temps d'arr�t aux feux (s)"
        TabFicheRes.Col = 6
        TabFicheRes.Text = "Temps de parcours (s)"
        TabFicheRes.Col = 7
        TabFicheRes.Text = "Vit moyenne (km/h)"
    Else
        'Cas o� il n'y a pas de TC utilis�s
        TabFicheRes.Col = 2
        TabFicheRes.Text = "Aucun r�sultat"
    End If
End Sub

Public Sub ImprimerFicheDonneesCarrefour(unHeader As String)
    'Impression de la Fiche des donn�es carrefours
    Dim unTitreFiche As String, unNomFiche As String
    Dim unFeu As Feu, unCarf As Carrefour
    
    'Affectation d'un nom de fiche
    unNomFiche = "Donn�es des carrefours"
    
    'Affectation d'un titre de fiche
    unTitreFiche = "Donn�es des carrefours"
    
    'Affectation des options d'impression du spread TabFicheRes
    'avec affichage des ent�tes de colonnes
    ConfigurerSpreadToPrint TabFicheDataCarf, unHeader, unNomFiche, unTitreFiche
    TabFicheDataCarf.PrintColHeaders = True
    
    'Affectation de la couleur des lignes d'ent�tes (non lock�es)
    TabFicheDataCarf.ShadowColor = RGB(220, 220, 220) 'Gris clair
    
    'S�lection de la ligne d'ent�te pour mettre une fonte grasse
    TabFicheDataCarf.Row = 0
    TabFicheDataCarf.Col = -1
    TabFicheDataCarf.FontBold = True
    
    With monSite
        'Remplissage � partir des spread de l'onglet Carrefour
        unNbLignes = 0
        For i = 1 To monSite.mesCarrefours.Count
            unNbLignes = unNbLignes + monSite.mesCarrefours(i).mesFeux.Count
        Next i
        TabFicheDataCarf.MaxRows = unNbLignes
        
        'Remplissage du spread TabFicheDataCarf
        unNbFeuxPred = 0
        For i = 1 To monSite.mesCarrefours.Count
            Set unCarf = monSite.mesCarrefours(i)
            TabFicheDataCarf.Row = unNbFeuxPred + 1
            TabFicheDataCarf.Col = 1
            TabFicheDataCarf.Text = unCarf.monNom
            
            'Remplissage � partir de tous les feux
            For K = 1 To monSite.mesCarrefours(i).mesFeux.Count
                Set unFeu = unCarf.mesFeux(K)
                TabFicheDataCarf.Row = unNbFeuxPred + K
                
               'Affichage du num�ro de feu
                TabFicheDataCarf.Col = 2
                TabFicheDataCarf.Text = Format(K)
            
                'Affichage du sens de parcours du feu
                TabFicheDataCarf.Col = 3
                If unFeu.monSensMontant Then
                    TabFicheDataCarf.Text = "Montant"
                Else
                    TabFicheDataCarf.Text = "Descendant"
                End If
                
                 'Affichage de l'ordonn�e du feu
                TabFicheDataCarf.Col = 4
                TabFicheDataCarf.Text = Format(unFeu.monOrdonn�e)
               
                 'Affichage de la dur�e de vert du feu
                TabFicheDataCarf.Col = 5
                TabFicheDataCarf.Text = Format(unFeu.maDur�eDeVert)
               
                 'Affichage de la position du point de r�f�rence du feu
                 'stock�e avec un signe oppos� � la saisie en interne
                TabFicheDataCarf.Col = 6
                TabFicheDataCarf.Text = Format(-unFeu.maPositionPointRef)
               
               'Remplissage des demandes M et D
                TabFicheDataCarf.Col = 7
                TabFicheDataCarf.Text = Format(unCarf.maDemandeM)
                TabFicheDataCarf.Col = 8
                TabFicheDataCarf.Text = Format(unCarf.maDemandeD)
                
                'Remplissage des d�bits de saturation M et D
                TabFicheDataCarf.Col = 9
                TabFicheDataCarf.Text = Format(unCarf.monDebSatM)
                TabFicheDataCarf.Col = 10
                TabFicheDataCarf.Text = Format(unCarf.monDebSatD)
            Next K
            
            'Stockage du nombre de feux affich�s
            unNbFeuxPred = unNbFeuxPred + unCarf.mesFeux.Count
        Next i
    End With
    
    'Impression du spread
    TabFicheDataCarf.Action = 13 ' = SS_ACTION_PRINT
End Sub

Public Sub ImprimerFicheDonneesTC(unHeader As String)
    'Impression de la Fiche des donn�es TC
    Dim unNomFiche As String, unTitreFiche As String
    Dim unTC As TC, unArret As ArretTC
    
    'Affectation d'un nom de fiche
    unNomFiche = "Donn�es des Transports Collectifs"
    
    'Affectation d'un titre de fiche
    unTitreFiche = "Donn�es des Transports Collectifs"
    
    'Affectation des options d'impression du spread TabFicheRes
    ConfigurerSpreadToPrint TabFicheDataTC, unHeader, unNomFiche, unTitreFiche
            
    'Affectation de la couleur blanches aux lignes autres
    'que les ent�tes (celles lock�es)
    TabFicheDataTC.LockBackColor = RGB(255, 255, 255)
    'Affectation de la couleur des lignes d'ent�tes (non lock�es)
    TabFicheDataTC.BackColor = RGB(220, 220, 220) 'Gris clair
    
    'Calcul du nombre de lignes avec un ent�te pour chaque TC
    'et un ent�te avant la liste des arr�ts de chaque TC
    unNbRows = 0
    For i = 1 To monSite.mesTC.Count
        unNbRows = unNbRows + 3 + monSite.mesTC(i).mesArrets.Count
    Next i
    TabFicheDataTC.MaxRows = unNbRows
    
    'Remplissage du spread TabFicheDataTC
    uneLigneDebut = 0
    For i = 1 To monSite.mesTC.Count
        Set unTC = monSite.mesTC(i)
        'Stockage de la ligne de d�but de remplissage des
        'donn�es g�n�rales du TC
        uneLigneDebTC = uneLigneDebut + 1
        '(i-1)*2 car il y a 2 lignes d'ent�te pour chaque TC
        
        'Remplissage de la ligne d'ent�te du TC
        RemplirLigneEnteteTC uneLigneDebTC
        
        'Remplissage de la ligne contenant les donn�es g�n�rales du TC
        TabFicheDataTC.Row = uneLigneDebTC + 1
        TabFicheDataTC.RowHeight(uneLigneDebTC + 1) = TailleCellule

        'Affichage du nom
        TabFicheDataTC.Col = 1
        TabFicheDataTC.Text = Format(unTC.monNom)
        
        'Affichage de l'instant de d�part
        TabFicheDataTC.Col = 2
        TabFicheDataTC.Text = Format(unTC.monTDep)
        
        'Affichage de la distance Acc�l�ration/Freinage
        TabFicheDataTC.Col = 3
        TabFicheDataTC.Text = Format(unTC.maDistAccFrein)
        
        'Affichage de la dur�e Acc�l�ration/Freinage
        TabFicheDataTC.Col = 4
        TabFicheDataTC.Text = Format(unTC.maDureeAccFrein)
        
        'Affichage du nom du carrefour de d�part
        TabFicheDataTC.Col = 5
        TabFicheDataTC.Text = Format(unTC.monCarfDep.monNom)
        
        'Affichage du nom du carrefour d'arriv�e
        TabFicheDataTC.Col = 6
        TabFicheDataTC.Text = Format(unTC.monCarfArr.monNom)
        
        'Affichage de la couleur de repr�sentation
        TabFicheDataTC.Col = 7
        TabFicheDataTC.Lock = False
        TabFicheDataTC.BackColor = unTC.maCouleur
        
        'Stockage de la ligne de d�but de remplissage des
        'donn�es des arr�ts TC
        uneLigneDebut = uneLigneDebTC + 2
        
        'Remplissage de la ligne d'ent�te des donn�es des arr�ts du TC
        RemplirLigneEnteteArret uneLigneDebut
        
        'Remplissage des donn�es des arr�ts du TC
        For j = 1 To unTC.mesArrets.Count
            Set unArret = unTC.mesArrets(j)
            TabFicheDataTC.Row = uneLigneDebut + j
            TabFicheDataTC.RowHeight(uneLigneDebut + j) = TailleCellule
            
            'Affichage du num�ro de l'arr�t
            TabFicheDataTC.Col = 2
            TabFicheDataTC.Text = Format(j)
            
            'Affichage de l'ordonn�e de l'arr�t
            TabFicheDataTC.Col = 3
            TabFicheDataTC.Text = Format(unArret.monOrdonnee)
            
            'Affichage de la vitesse de marche de l'arr�t
            TabFicheDataTC.Col = 4
            TabFicheDataTC.Text = Format(unArret.maVitesseMarche)
            
            'Affichage du temps d'arr�t � cet arr�t
            TabFicheDataTC.Col = 5
            TabFicheDataTC.Text = Format(unArret.monTempsArret)
            
            'Affichage du libell� de l'arr�t
            TabFicheDataTC.Col = 6
            TabFicheDataTC.Text = Format(unArret.monLibelle)
        Next j
        
        'Stockage du nombre d'arr�ts total
        uneLigneDebut = uneLigneDebut + unTC.mesArrets.Count
    Next i
    
    'Impression du spread
    TabFicheDataTC.Action = 13 ' = SS_ACTION_PRINT
End Sub



Public Sub RemplirLigneEnteteTC(uneLigneDebut)
    'Remplissage de la ligne d'ent�te des
    'donn�es g�n�rales d'un TC pour imprimer les donn�es TC
    
    'Modif de la hauteur de la ligne correspondant
    'aux ent�tes
    TabFicheDataTC.RowHeight(uneLigneDebut) = TailleCellule * 3
    
    'S�lection de la ligne d'ent�te pour fond gris + fonte grasse
    TabFicheDataTC.Row = uneLigneDebut
    TabFicheDataTC.Col = -1
    TabFicheDataTC.Lock = False
    TabFicheDataTC.ForeColor = 0 'Noir
    TabFicheDataTC.FontBold = True
    
    'Remplissage de la ligne d'ent�te
    TabFicheDataTC.Row = uneLigneDebut
    TabFicheDataTC.Col = 1
    TabFicheDataTC.Text = "TC"
    
    TabFicheDataTC.Col = 2
    TabFicheDataTC.Text = "Instant d�part (s)"
    TabFicheDataTC.Col = 3
    TabFicheDataTC.Text = "Distance Acc�l. + Frein (m)"
    TabFicheDataTC.Col = 4
    TabFicheDataTC.Text = "Dur�e Acc�l. + Frein (s)"
    TabFicheDataTC.Col = 5
    TabFicheDataTC.Text = "Carrefour de d�part"
    TabFicheDataTC.Col = 6
    TabFicheDataTC.Text = "Carrefour d'arriv�e"
    TabFicheDataTC.Col = 7
    TabFicheDataTC.Text = "Couleur"
End Sub

Public Sub RemplirLigneEnteteArret(uneLigneDebut)
    'Remplissage de la ligne d'ent�te des donn�es des
    'arr�ts d'un TC pour imprimer les donn�es TC
    
    'Modif de la hauteur de la ligne correspondant
    'aux ent�tes
    TabFicheDataTC.RowHeight(uneLigneDebut) = TailleCellule
    
    'S�lection de la ligne d'ent�te pour fond gris + fonte grasse
    TabFicheDataTC.Row = uneLigneDebut
    TabFicheDataTC.Col = -1
    TabFicheDataTC.Lock = False
    TabFicheDataTC.ForeColor = 0 'Noir
    TabFicheDataTC.FontBold = True
        
    'Positionnement sur la ligne d'ent�te des arr�ts
    TabFicheDataTC.Row = uneLigneDebut
    
    'Mise en blanc de la colonne 1 en la lockant
    TabFicheDataTC.Col = 1
    TabFicheDataTC.Lock = True
    TabFicheDataTC.Text = ""
    
    'Remplissage de la ligne d'ent�te
    TabFicheDataTC.Col = 2
    TabFicheDataTC.Text = "Arr�t"
    TabFicheDataTC.Col = 3
    TabFicheDataTC.Text = "Y (m)"
    TabFicheDataTC.Col = 4
    TabFicheDataTC.Text = "V (km/h)"
    TabFicheDataTC.Col = 5
    TabFicheDataTC.Text = "Temps d'arr�t (s)"
    TabFicheDataTC.Col = 6
    TabFicheDataTC.Text = "Libell�"
    
    'Mise en blanc de la colonne 7 en la lockant
    TabFicheDataTC.Col = 7
    TabFicheDataTC.Lock = True
    TabFicheDataTC.Text = ""
End Sub


Private Sub NbSecondes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If NbSecondes + UpDown1.Increment > UpDown1.Max Then
            NbSecondes = UpDown1.Min
        Else
            NbSecondes = NbSecondes + UpDown1.Increment
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        If NbSecondes - UpDown1.Increment < UpDown1.Min Then
            NbSecondes = UpDown1.Max
        Else
            NbSecondes = NbSecondes - UpDown1.Increment
        End If
    End If
End Sub

