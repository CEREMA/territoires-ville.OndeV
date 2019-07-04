VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000009&
   Caption         =   "OndeV"
   ClientHeight    =   4020
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nouveau"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Ouvrir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Enregistrer"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3750
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6165
            Text            =   "OndeV version 1.0"
            TextSave        =   "OndeV version 1.0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "28/09/2005"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:19"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":118A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":182E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2224
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2576
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":28C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2C1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&Site"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Fermer"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Enre&gistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Enregistrer &sous..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Enregistrer &tout"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Exporter vers fichier texte..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSite1 
         Caption         =   "&1 Site1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSite2 
         Caption         =   "&2 Site2"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSite3 
         Caption         =   "&3 Site3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSite4 
         Caption         =   "&4 Site4"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Barre d'&outils"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barre d'&état"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuGraphicOnde 
      Caption         =   "&Graphique onde verte"
      Enabled         =   0   'False
      Begin VB.Menu mnuGraphicOndeAnnul 
         Caption         =   "&Annuler la dernière modification graphique"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGraphicOndeTempsTC 
         Caption         =   "&Dessiner Temps parcours TC = F (Instant départ TC) ..."
      End
      Begin VB.Menu mnuGraphicOndeInterCarfVP 
         Caption         =   "&Montrer les bandes inter-carrefours voitures en onde TC"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGraphicOndeOptions 
         Caption         =   "&Options d'affichage..."
      End
      Begin VB.Menu mnuGraphicOndeFindBandes 
         Caption         =   "&Rechercher bandes passantes suivant les vitesses..."
      End
      Begin VB.Menu mnuGraphicOndeTracerTC 
         Caption         =   "&Tracer les progressions des Transports Collectifs..."
      End
      Begin VB.Menu mnuGraphicOndePleinEcran 
         Caption         =   "&Visualiser en plein écran..."
      End
      Begin VB.Menu mnuGraphicOndeSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraphicOndeProp 
         Caption         =   "&Propriétés..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Fenêtre"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Réorganiser les icônes"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Sommaire de l'aide"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "A&ide sur..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "À &propos de OndeV..."
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence"
      End
   End
   Begin VB.Menu mnuObjGraphic 
      Caption         =   "&PopupObjetGraphic"
      Visible         =   0   'False
      Begin VB.Menu mnuObjGraphicNew 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuObjGraphicDel 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu mnuObjGraphicRen 
         Caption         =   "&Renommer..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private monTypeObjetSelect As String

Private Sub MDIForm_Load()
    Dim uneString As String
    
    'Mise à jour de l'ihm
    Call InitQlm
    
    'Affectation du fichier d'aide
    App.HelpFile = GetAppPath() + "OndeV.chm"
    dlgCommonDialog.HelpFile = App.HelpFile
    
    'Index des aides pour les items de menus
    mnuFileNew.HelpContextID = IDhlp_NewSite
    mnuFileOpen.HelpContextID = IDhlp_OpenSite
    mnuFileSave.HelpContextID = IDhlp_Save
    mnuFileSaveAs.HelpContextID = IDhlp_SaveAs
    mnuFileSaveAll.HelpContextID = IDhlp_SaveAll
    mnuFilePrint.HelpContextID = IDhlp_PrintSite
    mnuViewOptions.HelpContextID = IDhlp_MenuAffichageOptions
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    'On masque les boutons dans la toolbar permettant l'impression et la
    'sauvegarde car il n'y a pas de fenêtre fille ouverte
    '==> Impression et sauvegarde impossible
    Me.tbToolBar.Buttons("Print").Visible = False
    Me.tbToolBar.Buttons("Save").Visible = False
    
    'Lecture dans la base de registre pour récupérer
    'les derniers fichiers ouverts, maximum 4
    uneString = Trim(GetSetting(App.Title, "Recent Files", "File1", ""))
    mnuFileSite1.Visible = (uneString <> "")
    mnuFileSite1.Caption = "&1 " + uneString
    
    'Dés qu'il y a un fichier récent on met le séparateur
    mnuFileBar3.Visible = (uneString <> "")
    
    uneString = Trim(GetSetting(App.Title, "Recent Files", "File2", ""))
    mnuFileSite2.Visible = (uneString <> "")
    mnuFileSite2.Caption = "&2 " + uneString
    
    uneString = Trim(GetSetting(App.Title, "Recent Files", "File3", ""))
    mnuFileSite3.Visible = (uneString <> "")
    mnuFileSite3.Caption = "&3 " + uneString
    
    uneString = Trim(GetSetting(App.Title, "Recent Files", "File4", ""))
    mnuFileSite4.Visible = (uneString <> "")
    mnuFileSite4.Caption = "&4 " + uneString
End Sub


Private Sub LoadNewDoc()
    'Création d'une nouvelle fenetre fille
    Dim frmD As frmDocument
    
    monDocumentCount = monDocumentCount + 1
    Set frmD = New frmDocument
    DoEvents
    
    'Initialisation de la nouvelle fenetre
    
    'Initialisation des Ymin et Ymax de l'englobant des feux
    frmD.monYMinFeu = -LongueurAxeY / 2
    frmD.monYMaxFeu = LongueurAxeY / 2
    'Initialisation des bandes TC pour l'onde TC
    frmD.maBandeTCM = 15
    frmD.maBandeTCD = 15
    'Initialisation de la longueur totale de l'axe des ordonnées en mètres
    frmD.maLongueurAxeY = LongueurAxeY
    'Affichage d'un titre d'études par défaut. Commentaires du site
    frmD.monTitreEtude = ""
    frmD.TitreEtude.Text = frmD.monTitreEtude
    'Affectation à 50 de la durée du cycle par défaut
    frmD.maDuréeDeCycle = 50
    frmD.DuréeCycle.Text = Format(frmD.maDuréeDeCycle)
    'Affectation du nombre d'objet graphiques de TC (label NomCarf) créés
    frmD.monNbObjGraphicCarf = 0
    'Affectation du nombre d'objet graphiques de TC (label NumFeu) créés
    frmD.monNbObjGraphicFeu = 0
    'Affectation du nombre d'objet graphiques de TC (label NomArret) créés
    frmD.monNbObjGraphicTC = 0
    'Affectation avec des valeurs par défaut de l'onglet Cadrage onde verte
    frmD.monTypeOnde = OndeDouble
    frmD.OptionOndeDouble = True
    frmD.monPoidsSensM = 1
    frmD.TextPoidsM.Text = frmD.monPoidsSensM
    frmD.monPoidsSensD = 1
    frmD.TextPoidsD.Text = frmD.monPoidsSensD
    'Affichage ou masquage des colonnes de saisies des vitesses montantes
    ' et descendantes de chaque carrefour si vitesse variable ou constante
    frmD.monTypeVit = VitConst
    If frmD.monTypeVit = VitConst Then
        unIsVitConst = True
        frmD.OptionVitConst = True
    Else
        unIsVitConst = False
        frmD.OptionVitVar = True
    End If
    frmD.TabInfoCalc.Col = 3
    frmD.TabInfoCalc.ColHidden = unIsVitConst
    frmD.TabInfoCalc.Col = 4
    frmD.TabInfoCalc.ColHidden = unIsVitConst
    frmD.TextVitM.Enabled = unIsVitConst
    frmD.TextVitD.Enabled = unIsVitConst
    frmD.LabelVitSensM.Enabled = unIsVitConst
    frmD.LabelVitSensD.Enabled = unIsVitConst
    
    frmD.maVitSensM = 30
    frmD.TextVitM.Text = frmD.maVitSensM
    frmD.maVitSensD = 30
    frmD.TextVitD.Text = frmD.maVitSensD
    frmD.maTransDec = 0
    frmD.TextTransDec.Text = frmD.maTransDec
    'Création du premier carrefour qui crééra un premier feu
    CreerCarrefour frmD
    'Modification du titre de la fenetre Site
    frmD.Caption = "Site : Sans Nom " & monDocumentCount
    'Stockage de la fenetre du site courant
    Set monSite = frmD
    'Origine mise au milieu
    frmD.Origine.Top = (frmD.AxeOrdonnée.Y1 + frmD.AxeOrdonnée.Y2) / 2 - frmD.Origine.Height
    'masquage de la frame FrameTC
    frmD.FrameTC.Visible = False
    'Inhibition des boutons de TC
    frmD.RenameTC.Enabled = False
    frmD.DelTC.Enabled = False
    'Affichage de la fenetre Site
    frmD.Show
    'Initialisation des variables indiquant les modifications aprés
    'saisies et calculs
    frmD.InitIndiqModif
    'Indication de la création d'un carrefour par défaut avec un feu
    frmD.maModifDataCarf = True
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    'Indication d'une fermeture de la fenêtre mère MDI
    monFermerParMereMDI = True
    
    'Sauvegarde des options dans la base de registre à la place du fichier
    'OndeV.ini (fait à partir de la version 1.00.0002),
    'juste les fichiers récents
    SauverOptionsAffImp True
End Sub


Private Sub mnuFile_Click()
    If Forms.Count = 1 Then
        'Aucune fenêtre fille ouverte
        'La seul fenetre ouverte la MDI mère
        uneMiseEnGrisé = False
    Else
        'Des fenêtres filles ouvertes
        uneMiseEnGrisé = True
    End If
    
    'Mise à jour des items du menu Site (= mnuFile)
    mnuFileClose.Enabled = uneMiseEnGrisé
    mnuFileExport.Enabled = uneMiseEnGrisé
    mnuFileSave.Enabled = uneMiseEnGrisé
    mnuFileSaveAs.Enabled = uneMiseEnGrisé
    mnuFileSaveAll.Enabled = uneMiseEnGrisé
    mnuFilePrint.Enabled = uneMiseEnGrisé
End Sub

Private Sub mnuFileExport_Click()
    Dim unCarf As Carrefour, unFeu As Feu
    Dim unTC As TC, unArret As ArretTC
    Dim unNbFeuxM As Integer, unNbFeuxD As Integer

    'Impossible en version démo
    If maDemoVersion Then
        unMsg = Mid(mnuFileExport.Caption, 2) 'Suppression du &
        unMsg = Mid(unMsg, 1, Len(unMsg) - 3) 'Suppression des ... finaux
        MsgBox UCase(unMsg) + " n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    
    'Vérification de la validité de la protection
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Test si on exporte des données et des résultats cohérents sinon abandon
    
    'Calcul de l'état de cohérence entre les données et les résultats
    'des calculs dans l'étude en cours
    If monSite.TabFeux.Tab > 2 Then
        unEtatIncoherenceDataCalc = False
    Else
        unEtatIncoherenceDataCalc = monSite.maModifDataCarf Or monSite.maModifDataOndeTC = True Or monSite.maModifDataOnde
    End If
    
    'Test si une ou plusieurs données du calcul d'onde ont
    'changé ou si incohérence entre données et résultats
    '==> Pas d'impression des résultats et du dessin
    'd'onde verte tant qu'il y a incohérence
    If monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul Or unEtatIncoherenceDataCalc Or monSite.maCoherenceDataCalc = CalculImpossible Then
        If monSite.maCoherenceDataCalc = CalculImpossible Then
            unMsgMilieu = unMsgMilieu + "Raison : Le calcul d'onde verte est impossible avec les données de ce site."
            If monSite.monTypeOnde = 3 And monSite.monTCM = 0 And monSite.monTCD = 0 Then
                unMsgMilieu = unMsgMilieu + Chr(13) + "En effet, dans l'onglet Cadrage Onde Verte, aucun TC montant et/ou descendant n'ont été choisis." + Chr(13) + Chr(13) + "Calcul d'onde verte prenant en compte les TC impossible"
            End If
        ElseIf monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul Or unEtatIncoherenceDataCalc Then
            monSite.maCoherenceDataCalc = IncoherenceDonneeCalcul
            unMsgMilieu = unMsgMilieu + "Raison : une ou plusieurs données du calcul d'onde verte ont changé." + Chr(13)
            unMsgMilieu = unMsgMilieu + "Ces données sont incohérentes avec les résultats des calculs précédant ces changements."
        End If
        
        unMsg = "Impossible d'exporter dans un fichier texte" + Chr(13) + Chr(13)
        unMsg = unMsg + unMsgMilieu + Chr(13) + Chr(13)
        unMsg = unMsg + "Vous pouvez recalculer les ondes vertes en sélectionnant l'un des 3 onglets suivants :" + Chr(13)
        unMsg = unMsg + "     - Résultat décalages" + Chr(13)
        unMsg = unMsg + "     - Dessin onde verte" + Chr(13)
        unMsg = unMsg + "     - Fiche Résultats"
        MsgBox unMsg, vbCritical
        Exit Sub
    End If
    
    With dlgCommonDialog
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurExport
        'définir les indicateurs et attributs
        'du contrôle des dialogues communs
        .CancelError = True
        .DialogTitle = "Exporter vers"
        .Filter = "Tous les fichiers (*.txt)|*.txt"
        .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        
        'Affichage d'un nom par défaut
        'titre du site actif moins la chaine "Site : " faisant 7 caractères
        'et l'extension .tal pour les sites différents de Sans Nom ...
        If Mid(monSite.Caption, 8, 8) = "Sans Nom" Then
            uneExt = 0
        Else
            uneExt = 4
        End If
        .FileName = Mid(monSite.Caption, 8, Len(monSite.Caption) - 7 - uneExt)
                
        'Ouverture du sélectionneur de fichier
        .ShowSave
        unFich = .FileName
        
        ' Ouvre le fichier en écriture.
        unFichId = FreeFile(0)
        Open unFich For Output As #unFichId
        
        'Ecriture dans le fichier texte choisi des données
        'et résultats du site actif (fiche carrefours, fiche TC et
        'fiche Résultats)
        'Le séparateur entre données est le point virgule
        
        'Nom du fichier site et titre de l'étude
        Print #unFichId, monSite.TitreEtude
        Print #unFichId, Mid(monSite.Caption, 8)
        
        'Fiche Données Carrefour
        Print #unFichId,
        Print #unFichId, "Fiche Données des carrefours"
        Print #unFichId,
        unNtot = monSite.mesCarrefours.Count
        If unNtot = 0 Then
            Print #unFichId, "Aucun carrefour dans ce site étudié"
        Else
            Print #unFichId, "Carrefour"; ";"; "Feu"; ";"; "Sens"; ";"; "Y (m)"; ";"; "Durée de vert (s)"; ";"; "Point de référence (s)"; ";"; "Demande montante (uvpd/h/v)"; ";"; "Demande descendante (uvpd/h/v)"; ";"; "Débit de saturation montant (uvpd/h/v)"; ";"; "Débit de saturation descendant (uvpd/h/v)"
        End If
        For i = 1 To unNtot
            Set unCarf = monSite.mesCarrefours(i)
            For j = 1 To unCarf.mesFeux.Count
                Set unFeu = unCarf.mesFeux(j)
                If unFeu.monSensMontant Then
                    unSens = "Montant"
                Else
                    unSens = "Descendant"
                End If
                'Ecriture dans fichier texte
                If j = 1 Then
                    uneChaine = unCarf.monNom
                Else
                    uneChaine = " "
                End If
                Print #unFichId, uneChaine; ";"; j; ";"; unSens; ";"; unFeu.monOrdonnée; ";"; unFeu.maDuréeDeVert; ";"; -unFeu.maPositionPointRef; ";"; unCarf.maDemandeM; ";"; unCarf.maDemandeD; ";"; unCarf.monDebSatM; ";"; unCarf.monDebSatD
                '- pour le point de référence car en interne il est inversé par rapport à la saisie
            Next j
        Next i
        
        'Fiche Données TC
        Print #unFichId,
        Print #unFichId, "Fiche Données des Transports Collectifs"
        Print #unFichId,
        unNtot = monSite.mesTC.Count
        If unNtot = 0 Then
            Print #unFichId, "Aucun TC dans ce site étudié"
        Else
            For i = 1 To unNtot
                Set unTC = monSite.mesTC(i)
                Print #unFichId, "TC"; ";"; "Instant de départ (s)"; ";"; "Distance Accél. + Frein (m)"; ";"; "Durée Accél. + Frein (s)"; ";"; "Carrefour de départ"; ";"; "Carrefour d'arrivée"
                Print #unFichId, unTC.monNom; ";"; unTC.monTDep; ";"; unTC.maDistAccFrein; ";"; unTC.maDureeAccFrein; ";"; unTC.monCarfDep.monNom; ";"; unTC.monCarfArr.monNom
                For j = 1 To unTC.mesArrets.Count
                    Set unArret = unTC.mesArrets(j)
                    If j = 1 Then
                        'ligne d'entête des arrêts du TC
                        Print #unFichId, " "; ";"; "Arrêt"; ";"; "Y (m)"; ";"; "V (km/h)"; ";"; "Temps d'arrêt (s)"; ";"; "Libellé"
                    End If
                    Print #unFichId, " "; ";"; j; ";"; unArret.monOrdonnee; ";"; unArret.maVitesseMarche; ";"; unArret.monTempsArret; ";"; unArret.monLibelle
                Next j
            Next i
        End If
        
        'Fiche Résultats
        'Affectation d'un titre de fiche
        unTitreFiche = "Résultats du calcul d'onde verte "
        If EstModifierManuel Then
            'Cas d'une modification manuelle des décalages
            unTitreFiche = unTitreFiche + "avec décalages modifiés manuellement"
        Else
            If monSite.monTypeOnde = OndeDouble Then
                unTitreFiche = unTitreFiche + "à double sens"
            ElseIf monSite.monTypeOnde = OndeSensM Then
                unTitreFiche = unTitreFiche + "à sens privilégié montant"
            ElseIf monSite.monTypeOnde = OndeSensD Then
                unTitreFiche = unTitreFiche + "à sens privilégié descendant"
            ElseIf monSite.monTypeOnde = OndeTC Then
                unTitreFiche = unTitreFiche + "prenant en compte les TC"
            End If
        End If
        'Affichage du titre de la fiche
        Print #unFichId,
        Print #unFichId, unTitreFiche
        Print #unFichId,
        
        'Calcul des temps de parcours
        TrouverTempsParcoursEtCarrefours unIndCarfM, unIndCarfD, unTmpM, unTmpD
        'Calcul des vitesses maximales possibles
        CalculerVitMax monSite
        'Affichage avec deux chiffres après la virgule
        'val des vit max = 0 si VM est > VitMaxLim (150 km/h) ou < vitMinLim (20 km/h)
        If Val(monSite.maVitMaxM) <> 0 Then monSite.maVitMaxM = Format(monSite.maVitMaxM, "fixed")
        If Val(monSite.maVitMaxD) <> 0 Then monSite.maVitMaxD = Format(monSite.maVitMaxD, "fixed")
        'Recherche des TC cadrant l'onde verte
        If monSite.monTCM = 0 Then
            unNomTCM = "Aucun"
        Else
            unNomTCM = monSite.mesTC(monSite.monTCM).monNom
        End If
        If monSite.monTCD = 0 Then
            unNomTCD = "Aucun"
        Else
            unNomTCD = monSite.mesTC(monSite.monTCD).monNom
        End If
        'Recherche des poids utilisés
        If monSite.monTypeOnde = OndeDouble Then
            unPM = Format(monSite.monPoidsSensM)
            unPD = Format(monSite.monPoidsSensD)
        Else
            unPM = "Aucun"
            unPD = "Aucun"
        End If
        
        'Remplissage de la 1ère partie des résultats
        Print #unFichId, "Sens de parcours"; ";"; "Temps de parcours (s)"; ";"; "Bande passante (s)"; ";"; "Vitesse max (km/h)"; ";"; "Poids"; ";"; "TC pris en compte"
        Print #unFichId, "MONTANT"; ";"; Format(unTmpM, "fixed"); ";"; monSite.maBandeModifM; ";"; monSite.maVitMaxM; ";"; unPM; ";"; unNomTCM
        Print #unFichId, "DESCENDANT"; ";"; Format(unTmpD, "fixed"); ";"; monSite.maBandeModifD; ";"; monSite.maVitMaxD; ";"; unPD; ";"; unNomTCD
        
        uneChaineVide = " "
        'Remplissage de la 2ème partie des résultats, ceux des carrefours
        Print #unFichId, "Carrefour"; ";"; "Décalages (s)"; ";"; "R Capacité Mont (%)"; ";"; "R Capacité Desc (%)"; ";"; "Vitesse Mon (km/h)"; ";"; "Vitesse Des (km/h)"; ";"; "Décalage ouverture (s)"
        For i = 1 To monSite.mesCarrefours.Count
            Set unCarf = monSite.mesCarrefours(i)
            If unCarf.monDecCalcul = -99 Then
                'Cas des carrefours inutilisés ou non compris entre
                'Ymin et Ymax d'une onde cadrée par un TC
                Print #unFichId, unCarf.monNom; ";"; uneChaineVide; ";"; uneChaineVide; ";"; uneChaineVide; ";"; uneChaineVide; ";"; uneChaineVide; ";"; uneChaineVide
            Else
                'Affichage du décalage en arrondissant à l'entier le plus
                'proche, d'où l'utilisation de la fonction VB5 CInt
                If CIntCorrigé(unCarf.monDecModif) = monSite.maDuréeDeCycle Then
                    'Une valeur valant durée du cycle s'affiche 0
                    unDecal = "0"
                Else
                    unDecal = Format(CIntCorrigé(unCarf.monDecModif))
                End If
                'Affichage en fonction du type de carrefour
                'réduit (double sens ou sens unique)
                If TypeOf unCarf.monCarfRed Is CarfReduitSensDouble Then
                    uneRCap = unCarf.monCarfRed.maDureeVertM / monSite.maDuréeDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                    If unCarf.maDemandeM = 0 Then
                        uneRCapM = "+infini"
                    Else
                        uneRCap = uneRCap / unCarf.maDemandeM * 100 'RCap en %
                        uneRCapM = Format(CInt(uneRCap))
                    End If
                    uneRCap = unCarf.monCarfRed.maDureeVertD / monSite.maDuréeDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                    If unCarf.maDemandeD = 0 Then
                        uneRCapD = "+infini"
                    Else
                        uneRCap = uneRCap / unCarf.maDemandeD * 100 'RCap en %
                        uneRCapD = Format(CInt(uneRCap))
                    End If
                    uneVM = Format(CInt(unCarf.DonnerVitSens(True) * 3.6))
                    uneVD = Format(CInt(unCarf.DonnerVitSens(False) * -3.6))
                    'Ecriture du Décalage à l'ouverture
                    'Il est indéterminé si plusieurs lignes de feux dans le
                    'même sens (Carrefour <> Carf réduit)==> Affichage "Indéfini"
                    unCarf.DonnerNbFeuxMetD unNbFeuxM, unNbFeuxD
                    If unNbFeuxM = 1 And unNbFeuxD = 1 Then
                        unDecOuv = Format(CInt(unCarf.monCarfRed.maPosRefM - unCarf.monCarfRed.maPosRefD))
                    Else
                        unDecOuv = "Indéfini"
                    End If
                Else
                    If unCarf.monCarfRed.monSensMontant Then
                        'Cas d'un carrefour à sens unique montant
                        uneRCap = unCarf.monCarfRed.maDureeVert / monSite.maDuréeDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                       If unCarf.maDemandeM = 0 Then
                            uneRCapM = "+infini"
                        Else
                            uneRCap = uneRCap / unCarf.maDemandeM * 100 'RCap en %
                            uneRCapM = Format(CInt(uneRCap))
                        End If
                        uneRCapD = " "
                        uneVM = Format(CInt(unCarf.DonnerVitSens(True) * 3.6))
                        uneVD = " "
                    Else
                        'Cas d'un carrefour à sens unique descendant
                        uneRCapM = " "
                        uneRCap = unCarf.monCarfRed.maDureeVert / monSite.maDuréeDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                       If unCarf.maDemandeD = 0 Then
                            uneRCapD = "+infini"
                        Else
                            uneRCap = uneRCap / unCarf.maDemandeD * 100 'RCap en %
                            uneRCapD = Format(CInt(uneRCap))
                        End If
                        uneVM = " "
                        uneVD = Format(CInt(unCarf.DonnerVitSens(False) * -3.6))
                    End If
                    'Décalage à l'ouverture indéterminé ==> Affichage "Indéfini"
                    unDecOuv = "Indéfini"
                End If
                Print #unFichId, unCarf.monNom; ";"; unDecal; ";"; uneRCapM; ";"; uneRCapD; ";"; uneVM; ";"; uneVD; ";"; unDecOuv
            End If
        Next i
        
        'Remplissage de la 3ème partie des résultats, ceux des TC utilisés
        unNtot = monSite.mesTCutil.Count
        If unNtot = 0 Then
            Print #unFichId, "Transport Collectif"; ";"; "Aucun résultat"
        Else
            Print #unFichId, "Transport Collectif"; ";"; "Sens de parcours"; ";"; "Instant de départ (s)"; ";"; "Nb d'arrêts aux feux"; ";"; "Temps d'arrêt aux feux (s)"; ";"; "Temps de parcours (s)"; ";"; "Vit moyenne (km/h)"
        End If
        For i = 1 To unNtot
            Set unTC = monSite.mesTCutil(i)
            If monSite.maModifDataTC Or monSite.maModifDataOndeTC Then
                'Recalcul du tableau de marche de progression s'il y a eu une
                'modif dans les données TC, de plus cela donne le sens du TC
                unSens = unTC.CalculerTableauMarcheProg()
            Else
                'Détermination du sens du TC
                If DonnerYCarrefour(unTC.monCarfDep) >= DonnerYCarrefour(unTC.monCarfArr) Then
                    'Cas d'un TC descendant
                    unSens = -1
                Else
                    'Cas d'un TC montant
                    unSens = 1
                End If
            End If
            
            'Ecriture dans fichier texte
            If unSens = 1 Then
                unSensText = "Montant"
            ElseIf unSens = -1 Then
                unSensText = "Descendant"
            Else
                MsgBox "ERREUR de programmation dans OndeV dans mnuFileExport", vbCritical
            End If
            'Calcul du temps de parcours du TC
            unNbPhases = unTC.mesPhasesTMProg.Count
            unTmpPar = unTC.mesPhasesTMProg(unNbPhases).monTDeb + unTC.mesPhasesTMProg(unNbPhases).maDureePhase - unTC.mesPhasesTMProg(1).monTDeb
            'Calcul de la distance parcourue par le TC
            uneDistPar = unTC.mesPhasesTMProg(unNbPhases).monYDeb + unTC.mesPhasesTMProg(unNbPhases).maLongPhase - unTC.mesPhasesTMProg(1).monYDeb
            'Affichage du temps de parcours et de la vitesse moyenne du TC en km/h
            unTmpParText = Format(CInt(unTmpPar))
            uneVMoyText = Format(CInt(uneDistPar / unTmpPar * 3.6))
            Print #unFichId, unTC.monNom; ";"; unSensText; ";"; unTC.monTDep; ";"; unTC.monNbArretsFeux; ";"; unTC.monTempsArretFeux; ";"; unTmpParText; ";"; uneVMoyText
        Next i
                
        ' Fermeture du fichier.
        Close #unFichId
        
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
    End With
    
    ' Quitte pour éviter le gestionnaire d'erreur.
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurExport:
    
    Select Case Err.Number
        Case 55 'Erreur "Ce fichier est déjà ouvert".
            MsgBox "Le fichier " + unFich + " est déjà ouvert", vbCritical
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    'fermeture et Sortie du menu Ouvrir
    Close #unFichId
    Exit Sub
End Sub

Private Sub mnuFileSite1_Click()
    'Ouverture du fichier récent numéro 1
    OuvrirFichierRecent mnuFileSite1.Caption
End Sub

Private Sub mnuFileSite2_Click()
    'Ouverture du fichier récent numéro 2
    OuvrirFichierRecent mnuFileSite2.Caption
End Sub

Private Sub mnuFileSite3_Click()
    'Ouverture du fichier récent numéro 3
    OuvrirFichierRecent mnuFileSite3.Caption
End Sub

Private Sub mnuFileSite4_Click()
    'Ouverture du fichier récent numéro 4
    OuvrirFichierRecent mnuFileSite4.Caption
End Sub

Private Sub mnuGraphicOndeAnnul_Click()
    'Annulation de la dernière modification
    'dans le graphique d'onde verte
    AnnulerLastModifGraphic monSite.ZoneDessin
    
    'Remise en grisée après l'utilisation du menu Annuler
    mnuGraphicOndeAnnul.Enabled = False
End Sub

Private Sub mnuGraphicOndeFindBandes_Click()
    frmInfoVitBande.Show vbModal
End Sub

Private Sub mnuGraphicOndeInterCarfVP_Click()
    'Redessin avec affichage des bandes inter-carrefours voitures
    'Cela ne se produit que si on est en onde cadrée par un TC
    mnuGraphicOndeInterCarfVP.Checked = Not monSite.monDessinInterCarfVP
    monSite.monDessinInterCarfVP = mnuGraphicOndeInterCarfVP.Checked
    MettreAJourDessin
End Sub

Private Sub mnuGraphicOndeOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuGraphicOndePleinEcran_Click()
    'Masquage des poignées si aucune n'a été sélectionnée
    monSite.PoigneeDroite.Visible = False
    monSite.PoigneeGauche.Visible = False
    
    'Affichage de la fenêtre plein écran
    Enabled = False
    frmPleinEcran.Show
End Sub

Private Sub mnuGraphicOndeProp_Click()
    'Affichage des propriétés du dernier objet sélectionné graphiquement
    AfficherPropsObjPick
End Sub

Private Sub mnuGraphicOndeTempsTC_Click()
    frmInfoTpsTC.Show vbModal
End Sub

Private Sub mnuGraphicOndeTracerTC_Click()
    frmTracerTC.Show vbModal
    
    'Affichage de la fenêtre du site étudié pour remise au premier plan
    'et ainsi éviter qu'une autre fenêtre Windows vienne au 1er plan
    'Bug affichage windows à priori
    monSite.Show
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuLicence_Click()
    'menu de saisie de licence QLM
    frmKey.Show 1
    'Mise à jour de l'ihm
    Call InitQlm
End Sub

Private Sub mnuObjGraphicDel_Click()
    If monTypeObjetSelect = "arrêt TC" Then
       SupprimerArretTC ActiveForm
    ElseIf monTypeObjetSelect = "carrefour" Then
        SupprimerCarrefour ActiveForm
    ElseIf monTypeObjetSelect = "feu" Then
        SupprimerFeu ActiveForm
    End If

End Sub

Private Sub mnuObjGraphicNew_Click()
    If monTypeObjetSelect = "arrêt TC" Then
       CreerArretTC ActiveForm
    ElseIf monTypeObjetSelect = "carrefour" Then
        CreerCarrefour ActiveForm
    ElseIf monTypeObjetSelect = "feu" Then
        CreerFeu ActiveForm
    End If
End Sub

Private Sub mnuObjGraphicRen_Click()
    If monTypeObjetSelect = "arrêt TC" Then
       ActiveForm.NewOrRenameTC "rename"
    ElseIf monTypeObjetSelect = "carrefour" Then
        RenommerCarrefour ActiveForm
    End If
End Sub

Private Sub mnuView_Click()
    If Forms.Count = 1 Then
        'Des fenêtres filles ouvertes
        uneMiseEnGrisé = False
    Else
        'Aucune fenêtre fille ouverte
        uneMiseEnGrisé = True
    End If
    
    'Mise à jour des items du menu Affichage (= mnuView)
    mnuViewOptions.Enabled = uneMiseEnGrisé
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal
End Sub


Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub


Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbToolBar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbToolBar.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

'Ajout O.Forel 12/07/2005 : modification du menu aide (méthode décrites dans Chelp.bas)
Private Sub mnuHelpIndex_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    'Modif fait par Frank Trifiletti on utilise le contextid de la fenêtre étude en cours
    'qui est dans la globale monetude dont son helpcontextid est mis à jour dans la sub ChangerHelpId
    'qui est appellé à chaque Form_Activate et dans le TabData_Click de frmDocument.frm
    'car le contextid était toujours nulle avec showindex normal on ne le passe pas en argument.
    If monSite Is Nothing Then
        'Cas d'appel  de F1 si aucun étude ouverte sinon plantage
        Call objHelp.ShowIndex(App.HelpFile, "Main")
    Else
        Call objHelp.Show(App.HelpFile, "Main", monSite.HelpContextID)
    End If
    'Fin modif F.Trifiletti
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpSearch_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.ShowSearch(App.HelpFile, "Main")
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpContents_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.Show(App.HelpFile, "Main")
    Set objHelp = Nothing
End Sub

'fin ajout o.Forel

'Private Sub mnuHelpContents_Click()
    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub


'Private Sub mnuHelpSearch_Click()
    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
'    Else
'        If HelpContextID = 0 Then
'            'Lance le sommaire de l'aide
'            mnuHelpContents_Click
'        Else
'            'Lance l'aide du bon contexte
'            dlgCommonDialog.HelpCommand = cdlHelpContext
'            dlgCommonDialog.HelpContext = HelpContextID
'            dlgCommonDialog.ShowHelp  ' affiche la rubrique
'        End If
'    End If
'End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Public Sub mnuFileOpen_Click()
    Dim unFich As String, unFichId As Integer
    Dim uneString As String, unNumLigne As Integer
    Dim frmD As frmDocument
    Dim unNom As String, uneVitM As Integer, uneVitD As Integer
    Dim unIsUtil As Boolean, unIsSensMontant As Boolean, unDecImp As Integer
    Dim unY As Integer, uneDureeVert As Single, unePosRef As Single
    Dim unTDep As Integer, uneDistAccFrein As Integer, uneDureeAccFrein As Integer
    Dim unIndCarfArr As Integer, unIndCarfDep As Integer
    Dim uneCoul As Long, unLibelle As String
    Dim unCarfDep As Carrefour, unCarfArr As Carrefour, unTC As TC
    Dim unYlong As Integer, uneVitMarche As Integer, unTempsArret As Integer
    Dim uneDureeCycle As Integer, unYMinFeu As Integer, unYMaxFeu As Integer
    Dim uneVitesseTC As Integer, unTAccelTC As Integer, uneDureeArret As Integer
    Dim unTypeOnde As Integer, unPoidsSensM As Integer, unPoidsSensD As Integer
    Dim uneVitSensM As Integer, uneVitSensD As Integer
    Dim uneVitTCM As Single, uneVitTCD As Single
    Dim unTypeVit As Integer, uneTransDec As Integer
    Dim uneDemM As Long, unDebSatM As Long
    Dim uneDemD As Long, unDebStaD As Long
    Dim uneBM As Single, uneBD As Single
    Dim uneBMmodif As Single, uneBDmodif As Single
    Dim unDecal As Single, unDecalModif As Single
    Dim unTCM As Integer, unTCD As Integer
    Dim uneBandeTCM As Single, uneBandeTCD As Single
    Dim uneOndeDoubleTrouve As Boolean, unEtatDeCoherence As Integer
    
    If maDemoVersion Then
        MsgBox "OUVERTURE D'UN FICHIER n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    
    With dlgCommonDialog
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurLecture
        
        'Test si OndeV est lancé avec un fichier de démarrage
        'double click sur un .TAL
        If monFichierDemarrage = "" Then
            'Cas où OndeV lancé sans fichier
            'Ouverture d'une fenêtre Ouvrir fichier
            
            'définir les indicateurs et attributs
            'du contrôle des dialogues communs
            .InitDir = CurDir
            .CancelError = True
            .DialogTitle = "Ouverture"
            .Filter = "Tous les fichiers (*.tal)|*.tal"
            .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
            .FileName = ""
            .ShowOpen
            If Len(.FileName) = 0 Then
                'Cas du click sur annuler
                ' Désactive la récupération d'erreur.
                On Error GoTo 0
                'Sortie de la procédure
                Exit Sub
            End If
            'Affectation du fichier à ouvrir
            unFich = .FileName
        Else
            'Cas où OndeV lancé avec un fichier
            
            'Affectation du fichier à ouvrir
            unFich = monFichierDemarrage
            'Effacement du fichier pour utiliser les prochains Ouvrir
            monFichierDemarrage = ""
        End If
                
        ' Ouvre le fichier en lecture et lock en read write.
        unFichId = FreeFile(0)
        Open unFich For Input Lock Read Write As #unFichId
        'Lecture de la première ligne
        If Not EOF(unFichId) Then Input #unFichId, uneString
        If uneString <> "Fichier Talon 3.0" Then
            MsgBox "Ce fichier n'est pas un fichier du logiciel OndeV version 1.0", vbCritical
        Else
            'Création d'une nouvelle fenetre fille
            Set frmD = New frmDocument
            frmD.monFichId = unFichId 'Stockage du Fichier Id
            'mise à jour du titre de la fenetre
            frmD.Caption = "Site : " + unFich
            'Affectation du nombre d'objet graphiques de TC (label NomCarf) créés
            frmD.monNbObjGraphicCarf = 0
            'Affectation du nombre d'objet graphiques de TC (label NumFeu) créés
            frmD.monNbObjGraphicFeu = 0
            'Affectation du nombre d'objet graphiques de TC (label NomArret) créés
            frmD.monNbObjGraphicTC = 0
            ' Effectue la boucle jusqu'à la fin du fichier.
            unNumLigne = 2
            Do While Not EOF(unFichId)
                With frmD
                    'Lecture des données et alimentation
                    'des attributs du site
                    If unNumLigne = 2 Then
                        Input #unFichId, uneString
                        .monTitreEtude = uneString
                    ElseIf unNumLigne = 3 Then
                        Input #unFichId, uneDureeCycle, unYMinFeu, unYMaxFeu, unEtatDeCoherence
                        .maDuréeDeCycle = uneDureeCycle
                        .monYMinFeu = unYMinFeu
                        .monYMaxFeu = unYMaxFeu
                        .maCoherenceDataCalc = unEtatDeCoherence
                        'Calcul de la longueur réelle de l'axe des Y
                        .maLongueurAxeY = .monYMaxFeu - .monYMinFeu
                    ElseIf unNumLigne = 4 Then
                        Input #unFichId, unTypeOnde, unPoidsSensM, unPoidsSensD, unTCM, unTCD, uneBandeTCM, uneBandeTCD, uneOndeDoubleTrouve
                        .monTypeOnde = unTypeOnde
                        .monPoidsSensM = unPoidsSensM
                        .monPoidsSensD = unPoidsSensD
                        .monTCM = unTCM
                        .monTCD = unTCD
                        .maBandeTCM = uneBandeTCM
                        .maBandeTCD = uneBandeTCD
                        .monOndeDoubleTrouve = uneOndeDoubleTrouve
                    ElseIf unNumLigne = 5 Then
                        Input #unFichId, unTypeVit, uneVitSensM, uneVitSensD
                        .monTypeVit = unTypeVit
                        .maVitSensM = uneVitSensM
                        .maVitSensD = uneVitSensD
                    ElseIf unNumLigne = 6 Then
                        Input #unFichId, uneTransDec, uneBM, uneBD, uneBMmodif, uneBDmodif
                        .maTransDec = uneTransDec
                        .maBandeM = uneBM
                        .maBandeD = uneBD
                        .maBandeModifM = uneBMmodif
                        .maBandeModifD = uneBDmodif
                    Else
                        Input #unFichId, uneString
                        If uneString = "Carrefour" Then
                            'Lecture des carrefours
                            Input #unFichId, unNom, uneVitM, uneVitD, unIsUtil, uneDemM, unDebSatM, uneDemD, unDebStaD, unDecal, unDecalModif, uneVitTCM, uneVitTCD, unDecImp
                            'Création du nouveau carrefour avec son nom unique
                            Set .monCarrefourCourant = .mesCarrefours.Add(unNom, uneVitM, uneVitD, unIsUtil, unDecImp)
                            'Stockage des demandes et des débits de saturation
                            .monCarrefourCourant.SetDemDeb uneDemM, unDebSatM, uneDemD, unDebStaD
                            'Stockage des décalages
                            .monCarrefourCourant.monDecCalcul = unDecal
                            .monCarrefourCourant.monDecModif = unDecalModif
                            'Stockage des vitesses TC
                            .monCarrefourCourant.maVitTCSensM = uneVitTCM
                            .monCarrefourCourant.maVitTCSensD = uneVitTCD
                            'Mise à jour de la combobox listant les noms de carrefours
                            .ComboNomCarf.AddItem unNom
                            'Mise à jour des combobox des TC listant les carrefours
                            'de départ et d'arrivée possibles
                            .ComboCarfDep.AddItem unNom
                            .ComboCarfArr.AddItem unNom
                            'Création du label NomCarf du carrefour qui sera mis
                            'en dernier position dans la collection mesCarrefours
                            DessinerCarrefour frmD, .mesCarrefours.Count
                        ElseIf uneString = "Feu" Then
                            'Lecture des feux
                            Input #unFichId, unIsSensMontant, unY, uneDureeVert, unePosRef
                            'Ajout d'un nouveau feu
                            Set unFeu = .monCarrefourCourant.mesFeux.Add(unIsSensMontant, unY, uneDureeVert, -unePosRef) '-PosRef car définition inverse entre dossier programmation et doc logiciel OndeV
                            'Stockage du carrefour du feu créé
                            Set unFeu.monCarrefour = .monCarrefourCourant
                            'Création des objets graphiques du feu numéro .monCarrefourCourant.mesFeux.Count
                            DessinerFeu frmD, .monCarrefourCourant.maPosition, .monCarrefourCourant.mesFeux.Count
                        ElseIf uneString = "TC" Then
                            'Lecture des TC
                            Input #unFichId, unNom, unTDep, uneDistAccFrein, uneDureeAccFrein, unIndCarfDep, unIndCarfArr, uneCoul
                            'Création d'une instance de TC
                            Set unCarfDep = .mesCarrefours(unIndCarfDep)
                            Set unCarfArr = .mesCarrefours(unIndCarfArr)
                            Set unTC = .mesTC.Add(unNom, unTDep, uneDistAccFrein, uneDureeAccFrein, unCarfDep, unCarfArr, uneCoul)
                            'Mise à jour de la combobox listant les TC
                            .ComboTC.AddItem unNom
                            'Mise à jour des combobox TC pour l'onde verte TC
                            RemplirComboboxOndeTC frmD, unTC
                        ElseIf uneString = "Arret" Then
                            'Lecture des arrets du dernier TC lu
                            Input #unFichId, unYlong, unTempsArret, uneVitMarche, unLibelle
                            unTC.mesArrets.Add unYlong, unTempsArret, uneVitMarche, unLibelle
                            'Dessin des objets graphiques de l'arrêt TC numéro i
                            frmD.DessinerArretTC .ComboTC.ListCount, CLng(unYlong)
                        ElseIf Trim(uneString) = "" Then
                            'Cas de ligne vide à la fin ==> on passe à la ligne suivante
                            'La dernière ligne doit être vide sans blanc sinon erreur 62
                            'lecture au-delà de la fin du fichier
                        Else
                            'Cas d'un fichier OndeV mal formatté
                            MsgBox "Ce fichier a été endommagé ou mal formatté, il n'est plus utilisable dans OndeV", vbCritical
                            'Fermeture du fichier et sortie de menu Ouvrir
                            Close #unFichId
                            Unload frmD
                            Exit Sub
                        End If
                    End If
                    'Augmentation du nombre de lignes lues
                    unNumLigne = unNumLigne + 1
                End With
            Loop
            'Stockage de la fenetre du site courant
            Set monSite = frmD
            'Réduction des carrefours pour lier le carrefour
            'et son carrefour réduit
            Call ReduireCarrefourSite(frmD, frmD.mesCarrefours, frmD.monTypeOnde)
            'Calcul des temps de parcours dans chaque sens à
            'chaque carrefour. Ces temps servent dans le recalcul
            'des bandes passantes lors d'une modif d'un décalage
            CalculerTempsParcours frmD
            'Remplissage des onglets de la fenetre site
            RemplirFenetreSite frmD
            'Affichage de la fenetre Site
            frmD.Show
            'Avertissement si le site qu'on ouvre est incohérent entre
            'les données et les résultats
            If frmD.maCoherenceDataCalc = IncoherenceDonneeCalcul Then
                unMsg = "Les données de ce site ne correspondent pas à celles des résultats de calcul" + Chr(13)
                unMsg = unMsg + Chr(13) + "Aller dans l'un des 3 onglets Résultat décalages, Dessin onde verte ou Fiche Résultats pour retrouver des données et des résultats cohérents"
                MsgBox unMsg, vbInformation
            End If
        End If
        
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
    End With
    
    'Mise en tête dans la liste des derniers fichiers ouverts
    MettreEnTeteRecents unFich, False
    
    ' Quitte pour éviter le gestionnaire d'erreur.
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurLecture:
    
    Select Case Err.Number
        Case 55, 70 'Erreur "Ce fichier est déjà ouvert".
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg + Chr(13) + "Le fichier " + unFich + " est déjà ouvert", vbCritical
            If Not frmD Is Nothing Then Unload frmD
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg + Chr(13) + "Ce fichier a été endommagé ou mal formatté, il n'est plus utilisable dans OndeV", vbCritical
            If Not frmD Is Nothing Then Unload frmD
    End Select
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    'fermeture et Sortie du menu Ouvrir
    Close #unFichId
    Exit Sub
End Sub


Private Sub mnuFileClose_Click()
    'Fermeture de la fenetre active
    Unload frmMain.ActiveForm
End Sub


Private Sub mnuFileSave_Click()
    'Ecriture dans le fichier du site courant
    'Nom du fichier = titre de la fenetre site moins la chaine "Site : "
    'ou enregistrement sous un nouveau de fichier choisi par l'utilisateur
    'si le titre de la fenêtre commence par "Site : Sans Nom"
    
    'Impossible en version démo
    If maDemoVersion Then
        unePos = InStr(1, mnuFileSave.Caption, "&") 'Suppression du &
        unMsg = Mid(mnuFileSave.Caption, 1, unePos - 1)
        unMsg = unMsg + Mid(mnuFileSave.Caption, unePos + 1)
        MsgBox UCase(unMsg) + " n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    
    If Mid(monSite.Caption, 1, 15) = "Site : Sans Nom" Then
        'Appel de la fonction Enregistrer sous...
        RunSaveAs monSite
    Else
        EcrireDansFichier Mid(monSite.Caption, 8), monSite
    End If
    'Mise en tête dans la liste des derniers fichiers ouverts
    MettreEnTeteRecents Mid(monSite.Caption, 8), False
End Sub


Private Sub mnuFileSaveAs_Click()
    If maDemoVersion Then
        unePos = InStr(1, mnuFileSaveAs.Caption, "&") 'Suppression du &
        unMsg = Mid(mnuFileSaveAs.Caption, 1, unePos - 1)
        unMsg = unMsg + Mid(mnuFileSaveAs.Caption, unePos + 1, Len(mnuFileSaveAs.Caption) - 3 - unePos)
        '-3 pour la Suppression des ... finaux
        MsgBox UCase(unMsg) + " n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    'Enregistrer sous
    RunSaveAs monSite
    'Mise en tête dans la liste des derniers fichiers ouverts
    MettreEnTeteRecents Mid(monSite.Caption, 8), False
End Sub


Private Sub mnuFileSaveAll_Click()
    'Ecriture dans les fichiers des sites ouverts
    'Nom du fichier = titre de la fenetre site moins la chaine "Site : "
    'ou enregistrement sous un nouveau de fichier choisi par l'utilisateur
    'si le titre de la fenêtre commence par "Site : Sans Nom"
    If maDemoVersion Then
        unePos = InStr(1, mnuFileSaveAll.Caption, "&") 'Suppression du &
        unMsg = Mid(mnuFileSaveAll.Caption, 1, unePos - 1)
        unMsg = unMsg + Mid(mnuFileSaveAll.Caption, unePos + 1)
        MsgBox UCase(unMsg) + " n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    
    For i = 1 To Forms.Count - 1
        If Mid(Forms(i).Caption, 1, 15) = "Site : Sans Nom" Then
            'Appel de la fonction Enregistrer sous...
            RunSaveAs Forms(i)
        Else
            EcrireDansFichier Mid(Forms(i).Caption, 8), Forms(i)
        End If
        'Mise en tête dans la liste des derniers fichiers ouverts
        MettreEnTeteRecents Mid(Forms(i).Caption, 8), False
    Next i
End Sub



Private Sub mnuFilePrint_Click()
    'Vérification de la validité de la protection
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Test si l'accés à une imprimante est possible
    If Printers.Count = 0 Then
        MsgBox "Aucune imprimante n'est connectée à ce poste", vbCritical
    Else
        'Affichage de la fenêtre d'impression si au moins l'accés
        'a une imprimante est possible
        frmImprimer.Show vbModal
    End If
End Sub


Private Sub mnuFileExit_Click()
    'décharger la feuille
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub




Public Sub AfficherMenuContextuel(unTypeObjet As String)
    monTypeObjetSelect = unTypeObjet
    
    If unTypeObjet = "arrêt TC" Then
        unTypeObjet = "arrêt"
        unItemNewCaption = "&Nouvel "
        unItemDelCaption = "l'"
        unItemRenCaption = "le TC"
        frmMain.mnuObjGraphicRen.Visible = True
    ElseIf unTypeObjet = "feu" Then
        unItemNewCaption = "&Nouveau "
        unItemDelCaption = "le "
        unItemRenCaption = "le feu"
        'Pas de renommage de feu
        frmMain.mnuObjGraphicRen.Visible = False
    Else
        'Autre cas le carrefour
        unItemNewCaption = "&Nouveau "
        unItemDelCaption = "le "
        unItemRenCaption = "le carrefour"
        frmMain.mnuObjGraphicRen.Visible = True
    End If
    
    frmMain.mnuObjGraphicNew.Caption = unItemNewCaption + unTypeObjet
    frmMain.mnuObjGraphicDel.Caption = "&Supprimer " + unItemDelCaption + unTypeObjet
    frmMain.mnuObjGraphicRen.Caption = "&Renommer " + unItemRenCaption + "..."
    frmMain.PopupMenu mnuObjGraphic, vbPopupMenuRightButton
End Sub

Public Sub RemplirFenetreSite(uneForm As Form)
    Dim uneLongEcranAxeY As Long
    
    'Positionnement de l'origine de l'axe OY au bon niveau de zoom
    uneLongEcranAxeY = uneForm.AxeOrdonnée.Y2 - uneForm.AxeOrdonnée.Y1
    unePos = ConvertirReelEnEcran(uneForm.monYMaxFeu, uneForm.maLongueurAxeY, uneLongEcranAxeY)
    uneForm.Origine.Top = unePos + uneForm.AxeOrdonnée.Y1 - uneForm.Origine.Height
    'Affichage d'un titre d'études par défaut. Commentaires du site
    uneForm.TitreEtude.Text = uneForm.monTitreEtude
    'Affichage de la durée du cycle par défaut
    uneForm.DuréeCycle.Text = Format(uneForm.maDuréeDeCycle)
    'Affichage des labelNomCarf au bon endroit
    For i = 1 To uneForm.mesCarrefours.Count
        ModifYNomCarf uneForm, uneForm.mesCarrefours(i)
    Next i
    'Remplissage de la frame FrameTC avec le premier TC s'il existe
    If uneForm.mesTC.Count > 0 Then
        'Déclenchement du comboTC_Click event ==> remplissage FrameTC
        uneForm.ComboTC.ListIndex = 0
    Else
        'Masquage de la frame contenant les caractéristiques
        'de chaque TC s'il n'y en a aucun
        uneForm.FrameTC.Visible = False
    End If
    'Remplissage onglet carrefour par déclenchement du
    'ComboNomCarf_Click event. Mis aprés la frameTC pour
    'sélectionner le premier carrefour du site
    uneForm.ComboNomCarf.Tag = "Déroulé par Click souris"
    uneForm.ComboNomCarf.ListIndex = 0
    'Remplissage de l'onglet Cadrage onde verte
    If uneForm.monTypeOnde = OndeDouble Then
        uneForm.OptionOndeDouble = True
    ElseIf uneForm.monTypeOnde = OndeSensM Then
        uneForm.OptionSensM = True
    ElseIf uneForm.monTypeOnde = OndeSensD Then
        uneForm.OptionSensD = True
    ElseIf uneForm.monTypeOnde = OndeTC Then
        uneForm.OptionTC = True
    End If
    uneForm.TextPoidsM.Text = uneForm.monPoidsSensM
    uneForm.TextPoidsD.Text = uneForm.monPoidsSensD
    uneForm.TextVitM.Text = uneForm.maVitSensM
    uneForm.TextVitD.Text = uneForm.maVitSensD
    If uneForm.monTypeVit = VitConst Then
        unIsVitConst = True
        uneForm.OptionVitConst = True
    Else
        unIsVitConst = False
        uneForm.OptionVitVar = True
    End If
    'Affichage ou masquage des colonnes de saisies des vitesses montantes
    ' et descendantes de chaque carrefour si vitesse variable ou constante
    uneForm.TabInfoCalc.Col = 3
    uneForm.TabInfoCalc.ColHidden = unIsVitConst
    uneForm.TabInfoCalc.Col = 4
    uneForm.TabInfoCalc.ColHidden = unIsVitConst
    uneForm.TextVitM.Enabled = unIsVitConst
    uneForm.TextVitD.Enabled = unIsVitConst
    uneForm.LabelVitSensM.Enabled = unIsVitConst
    uneForm.LabelVitSensD.Enabled = unIsVitConst
    'Remplissage du tableau des vitesses montante
    'et descendante des carrefours
    uneForm.TabInfoCalc.MaxRows = uneForm.mesCarrefours.Count
    For i = 1 To uneForm.mesCarrefours.Count
        uneForm.TabInfoCalc.Row = i
        uneForm.TabInfoCalc.Col = 1
        uneForm.TabInfoCalc.Text = uneForm.mesCarrefours(i).monNom
        uneForm.TabInfoCalc.Row = i
        uneForm.TabInfoCalc.Col = 2
        If uneForm.mesCarrefours(i).monIsUtil Then
            uneForm.TabInfoCalc.Text = "Oui"
        Else
            uneForm.TabInfoCalc.Text = "Non"
        End If
        uneForm.TabInfoCalc.Row = i
        uneForm.TabInfoCalc.Col = 3
        uneForm.TabInfoCalc.Text = uneForm.mesCarrefours(i).maVitSensM
        uneForm.TabInfoCalc.Row = i
        uneForm.TabInfoCalc.Col = 4
        uneForm.TabInfoCalc.Text = uneForm.mesCarrefours(i).maVitSensD
    Next i
    'Remplissage onglet Tableau résultat
    With uneForm
        .TextTransDec.Text = .maTransDec
    End With
    
    'Initialisation des variables indiquant les modifications aprés
    'saisies et calculs
    uneForm.InitIndiqModif
End Sub

Public Sub ChangerCouleurPicBox(unePictureBox As PictureBox)
    'Changement de couleur de fond d'une picture box
    ' Attribue à CancelError la valeur True
    dlgCommonDialog.CancelError = True
    On Error GoTo ErrHandler
    ' Définit la propriété Flags
    dlgCommonDialog.flags = cdlCCRGBInit
    ' Affiche la boîte de dialogue Couleur
    dlgCommonDialog.ShowColor
    ' Attribue à l'arrière-plan de la feuille la
    ' couleur sélectionnée
    unePictureBox.BackColor = dlgCommonDialog.Color
    Exit Sub

ErrHandler:
    If Err.Number <> cdlCancel Then
        ' Erreur autre que celle déclenchée par un click sur Annuler
        MsgBox "Erreur " + Format(Err.Number) + " : " + Err.Description, vbCritical
    End If
    On Error GoTo 0
    Exit Sub
End Sub


Public Sub RunSaveAs(uneForm As Form)
    'Fonction faisant le Save as
    'Configurer le contrôle des dialogues communs
    'avant d'appeler ShowSave
    With dlgCommonDialog
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurSaveAs
        'définir les indicateurs et attributs
        'du contrôle des dialogues communs
        .CancelError = True
        .DialogTitle = "Enregistrer sous"
        .Filter = "Tous les fichiers (*.tal)|*.tal"
        .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        .FileName = Mid(uneForm.Caption, 8)
        .ShowSave
        'Ecriture dans le fichier choisi du contenu du site courant
        EcrireDansFichier .FileName, uneForm
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
    End With
    
    ' Quitte pour éviter le gestionnaire d'erreur.
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurSaveAs:
    
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

Public Sub MettreEnTeteRecents(unNomFich As String, unePermutation As Boolean)
    'Mise en tête du fichier unNomFich  dans la liste des derniers
    'fichiers ouverts avec permutation avec l'ancien fichier de tête
    'si unePermutation est Vrai ou suppresion dans la liste de l'ancien
    'fichier numéro 4 (On stocke maxi les 4 derniers fichiers ouverts) si
    'le nouveau fichier de tête n'était pas dans la liste
    'La liste est stockée dans la collection maColLastFich
    Dim unSite1 As String, unSite2 As String
    Dim unSite3 As String, unSite4 As String
    
    mnuFileBar3.Visible = True 'Affichage du séparateur
    
    'Afffectation des noms de sites
    unSite1 = Mid(mnuFileSite1.Caption, 4)
    unSite2 = Mid(mnuFileSite2.Caption, 4)
    unSite3 = Mid(mnuFileSite3.Caption, 4)
    unSite4 = Mid(mnuFileSite4.Caption, 4)
    
    If unePermutation = False And (unNomFich = unSite1 Or unNomFich = unSite2 Or unNomFich = unSite3 Or unNomFich = unSite4) Then
        'Si on ne permute et que le nom de fichier fait
        'partie des 4 derniers fichiers ouverts on permute
        'juste l'ancien fichier de tête avec le nouveau fichier
        unePermutation = True
    End If
    
    If unePermutation Then
        'Si unNomfich déjà en tête on ne fait rien
        If unNomFich = unSite2 Then
            mnuFileSite2.Caption = "&2 " + unSite1
            unSite2 = mnuFileSite2.Caption
       ElseIf unNomFich = unSite3 Then
            mnuFileSite3.Caption = "&3 " + unSite1
            unSite3 = mnuFileSite3.Caption
        ElseIf unNomFich = unSite4 Then
            mnuFileSite4.Caption = "&4 " + unSite1
            unSite4 = mnuFileSite4.Caption
       End If
        'Mise en tête du nouveau nom
        mnuFileSite1.Caption = "&1 " + unNomFich
        unSite1 = mnuFileSite1.Caption
    Else
        '1 devient 2, 2 devient 3, 3 devient 4 et nouveau 1
        mnuFileSite4.Caption = "&4 " + unSite3
        unSite4 = mnuFileSite4.Caption
        
        mnuFileSite3.Caption = "&3 " + unSite2
        unSite3 = mnuFileSite3.Caption
        
        mnuFileSite2.Caption = "&2 " + unSite1
        unSite2 = mnuFileSite2.Caption
        
        mnuFileSite1.Caption = "&1 " + unNomFich
        unSite1 = mnuFileSite1.Caption
    End If
    
    'Affichage des noms de fichiers si différents des
    'valeurs par défaut (SiteN ou &N SiteN) long <=8 alors
    'qu'un fichier vaut au moins &N c:\g.tal ==> 11 caractéres
    'ou &N \\g.tal ==> 10 caractères
    If Len(unSite1) > 8 Then mnuFileSite1.Visible = True
    If Len(unSite2) > 8 Then mnuFileSite2.Visible = True
    If Len(unSite3) > 8 Then mnuFileSite3.Visible = True
    If Len(unSite4) > 8 Then mnuFileSite4.Visible = True
End Sub

Private Sub OuvrirFichierRecent(unNomFich As String)
    'Ouverture du fichier récent de nom unNomFich
    monFichierDemarrage = Mid(unNomFich, 4)
    unSiteDejàOuvert = False
    For i = 0 To Forms.Count - 1
        If "Site : " + monFichierDemarrage = Forms(i).Caption Then
            Forms(i).SetFocus
            unSiteDejàOuvert = True
            monFichierDemarrage = ""
            Exit For
        End If
    Next i
    If unSiteDejàOuvert = False Then mnuFileOpen_Click
End Sub

'Code pour modifier l'ihm suite à l'implémentation de Qlm
Private Sub InitQlm()
    'Initialisation des menus modifiés par QLM
    'les variables globales sont maj par protection.bas
    'ATTENTION : vérifier les noms des menus!!!
    Me.mnuHelpBar2.Visible = GvisibiliteMnuBarre
    Me.mnuLicence.Visible = GvisibiliteMnuLicence
    'a adapter en fonction du clogiciel
    Me.Caption = "OndeV v" + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision) + GmodifTitreApplication
    'fin initialisation qlm
    'fin initialisation qlm
End Sub
