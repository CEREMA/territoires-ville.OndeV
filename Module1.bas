Attribute VB_Name = "Module1"
'Variable indiquant si on travaille sur une version protégée ou pas
Public maProtectVersion As Boolean

'Variable indiquant si on travaille sur une version démo non protégée ou pas
Public maDemoVersion As Boolean

'Constante pour la décomposition des composantes RGB d'une couleur
Public Const CarreDe256 As Long = 65536 '256 * 256

'Coef de passage entre les cm et les twips
Public Const unTwipToCm = 567      '567 twips = 1 cm

'Constantes pour le type de phase des tableaux de marche TC
Public Const VConst As Integer = 0   'Phase à vitesse constante
Public Const Accel As Integer = 1    'Phase d'accélération
Public Const Decel As Integer = 2    'Phase d'décélération
Public Const Arret As Integer = 3    'Phase d'arrêt

'Constantes pour le type d'onde verte
Public Const OndeDouble As Integer = 0 'Type double sens
Public Const OndeSensM As Integer = 1  'Type sens montant privilégié
Public Const OndeSensD As Integer = 2  'Type sens descendant privilégié
Public Const OndeTC As Integer = 3     'Type cadrage pour un TC montant
                                       'et/ou descendant

'Constantes pour le type de vitesse à chaque carrefour
Public Const VitConst As Integer = 0 'Type vitesse constant pour tous les carrefours
Public Const VitVar As Integer = 1   'Type vitesse variable à chaque carrefour

'Constantes pour la modification des objets graphiques d'un TC
Public Const ModifNomTC As Integer = 0 'Modification du nom dans les labels
Public Const ModifColTC As Integer = 1 'Modification de la couleur
Public Const SupprTC As Integer = 2    'Suppression du TC

'Constantes pour la sélection des objets graphiques d'une fenetre site
Public Const CarfSel As Integer = 0  'Sélection graphique d'un carrefour
Public Const FeuSel As Integer = 1   'Sélection graphique d'un feu
Public Const ArretSel As Integer = 2 'Sélection graphique d'un arrêt

'Constantes pour la saisie des feux de départ et d'arrivée des TC
Public Const FeuDep As Integer = 0  'Saisie du feu de départ
Public Const FeuArr As Integer = 1  'Saisie du feu d'arrivée

'Constantes pour faire le trait jusqu'à l'axe des ordonnées
Public Const StringTrait As String = "___________________________________________________________________________"
    
'Constante fixant le nombre de caratères maximuns pour le nom des TC
'5 pour en mettre au moins 4 cote à cote dans le FrameVisuCarf
Public Const NbCarMaxNomTC As Integer = 5

'Public formMain As frmMain
Public monCallOptionByPrint As Boolean 'Variable indiquant si la fenêtre frmOptions a été ouverte par la fenêtre frmImprimer
Public monPleinEcranVisible As Boolean 'Variable indiquant si la fenêtre plein écran est ouverte
Public monFermerParMereMDI As Boolean 'Variable indiquant si on a fermé la MDI mère
Public monFichierDemarrage As String 'Fichier de démarrage éventuel sur la ligne de commande
Public monSite As frmDocument 'fenetre site dans lequel on travail
Public IsEtatCreation As Boolean 'True valeur pour nouveau site, False pour ouvrir un site existant
Public monDocumentCount As Long 'Nombre de fenetres filles sans Nom ouvertes
Public monNbFenFilles As Integer 'Nombre de fenetres filles ouvertes

'Variables stockant la position de la souris lors d'un mouve down event
'dans un control de type NumFeu ou IconeFeu
Public monXSouris As Single
Public monYSouris As Single

'Constante donnant la valeur réelle en mètres de la
'longueur de l'axe des Y à l'ouverture d'une nouvelle fenêtre de site
Public Const LongueurAxeY As Integer = 200

'Constante d'index d'aide en ligne
Public Const IDhlp_OngletCarf As Integer = 207
Public Const IDhlp_OngletTC As Integer = 208
Public Const IDhlp_OngletCadrage As Integer = 209
Public Const IDhlp_WinOptiVit As Integer = 210
Public Const IDhlp_OngletResDec As Integer = 211
Public Const IDhlp_OngletDesOnde As Integer = 212
Public Const IDhlp_WinFindBande As Integer = 216
Public Const IDhlp_WinTempsTC As Integer = 219
Public Const IDhlp_OngletFicRes As Integer = 221

Public Const IDhlp_NewSite As Integer = 205
Public Const IDhlp_OpenSite As Integer = 204
Public Const IDhlp_Save As Integer = 222
Public Const IDhlp_SaveAs As Integer = 223
Public Const IDhlp_SaveAll As Integer = 224
Public Const IDhlp_PrintSite As Integer = 225
Public Const IDhlp_MenuAffichageOptions As Integer = 100
Public Const IDhlp_WinAffichageOptions As Integer = 217
Public Const IDhlp_WinTracerTC As Integer = 219


Public Type CarfRed
    maDureeVert As Single
    monEcart As Single
End Type


Sub Main()

    'Variable indiquant si on travaille sur une version
    'protégée cc3.x (TRUE) ou pas (FALSE)
    'Utile lors du développement sous VB
    
    'maProtectVersion = True
    
    'désactivation de la protection de copycontrol
    maProtectVersion = False
    
    'Variable indiquant si on est en version DEMO ou pas
    maDemoVersion = False
    'maDemoVersion = True
    If maDemoVersion Then App.Title = App.Title + " version DEMO"
    
'********************************
'test Protection
'********************************
    'Type de protection
          TYPPROTECTION = CPM
    ' Vérification de l'enregistrement
    If ProtectCheck("its00+-k") = "its00+-k" Then
      ' Affichage de la feuille principale
         frmMain.Show
    Else 'la licence n'a pas été validée on ferme
       End
    End If
'********************************
    
    'Cas où protection valide
    If maDemoVersion Then
        'Changement du titre de l'application
        frmMain.Caption = frmMain.Caption + " version DEMO"
    End If

    'Initialisation
    monFermerParMereMDI = False
    monPleinEcranVisible = False
    monCallOptionByPrint = False
    
    monFichierDemarrage = Command()
    If monFichierDemarrage <> "" Then
        'Lancement de OndeV avec ouverture sur le fichier choisi
        'Double click sur .TAL
        frmMain.mnuFileOpen_Click
    End If
End Sub


Public Sub VerifSaisieEntierPositif(KeyAscii As Integer, unControl As Control, uneValeurDefaut As Integer)
    Dim uneChaineTmp
    
    If KeyAscii = 27 Or KeyAscii = 13 Then
        'Cas de la frappe des touches Echap ou Retour Chariot
        Exit Sub
    End If
    
    uneChaineTmp = " " + unControl.Text 'car la fonction Str rajoute un blanc pour les valeurs > 0
    If unControl.Text = "" Then
        'Cas où la zone de saisie est vide, on remet la valeur par défaut
        unControl.Text = uneValeurDefaut
    ElseIf uneChaineTmp <> str(Val(unControl.Text)) Or InStr(1, uneChaineTmp, ".") <> 0 Then
        unMsg = "Saisie d'entiers positifs uniquement"
        MsgBox unMsg
        unControl.Text = uneValeurDefaut
    End If
End Sub




Public Sub VideoInverse(unControl As Control)
    Dim uneColTmp As Long
    
    uneColTmp = unControl.BackColor
    unControl.BackColor = unControl.ForeColor
    unControl.ForeColor = uneColTmp
End Sub

Public Function PosInListe(uneChaine As String, unControl As Control) As Integer
    'Retourne la position dans la liste si trouvé (entre 0 et listcount-1)
    'ou -1 si non trouvé
    'La casse Min/majuscule est ignorée
    Dim i As Integer
    
    i = 0
    unInd = -1
    
    Do
        If UCase(uneChaine) = UCase(unControl.List(i)) Then
            unInd = i
        Else
            i = i + 1
        End If
    Loop While i < unControl.ListCount And unInd = -1
    
    PosInListe = unInd
End Function

Public Function ConvertirReelEnEcran(unYreel As Long, unDyReel As Long, unDyEcran As Long) As Long
    'Règle de trois pour convertir les mètres, stockés en long, en twips
    ConvertirReelEnEcran = unDyEcran / unDyReel * unYreel
End Function

Public Function ConvertirSingleEnEcran(unYreel As Single, unDyReel As Long, unDyEcran As Long) As Single
    'Règle de trois pour convertir les mètres, stockés en single, en twips
    ConvertirSingleEnEcran = unDyEcran / unDyReel * unYreel
End Function

Public Sub MiseAJourNomArret(uneFenetreSite As Form, uneListeIndexTC As Collection, uneListeIndexArret As Collection)
    'Mise à jour de tous les noms des arrêts qui étaient confondus
    'en le re-décalant correctement après la suppression d'arrêts
    Dim unIndexTC0 As Integer
    Dim unTC As TC
    Dim uneStringDecal As String
    Dim unNbTC As Integer
    Dim unObjGraphic As Control
    
    unNbTC = uneListeIndexTC.Count
    If unNbTC <> 0 Then
        unStringDecal = ""
        unIndexTC0 = uneListeIndexTC(1)
        For i = 1 To unNbTC
            Set unTC = uneFenetreSite.mesTC(uneListeIndexTC(i))
            If unIndexTC0 <> uneListeIndexTC(i) Then
                'Cas d'un TC différent ==> Augmentation du décalage du nom
                uneStringDecal = uneStringDecal + DonnerStringDecalage
            End If
            'Récupération du label NomArret rangé dans la collection mesObjgraphics
            Set unObjGraphic = unTC.mesObjGraphics(uneListeIndexArret(i))
            'Modification du label
            unObjGraphic.Caption = uneStringDecal + unTC.monNom + StringTrait
            'Coupure de la chaine de caractéres à l'axe des ordonnées
            unObjGraphic.Width = uneFenetreSite.AxeOrdonnée.X1 - unObjGraphic.Left
            'Stockage pour le i suivant
            unIndexTC0 = uneListeIndexTC(i)
        Next i
    End If
End Sub

Public Sub ViderCollection(uneCol As Collection)
    'Procédure vidant une collection
    
    'Algo : Puisque les collections sont réindexées
    '       automatiquement, en supprimant le premier
    '       membre à chaque itération, on supprime tout.
    For i = 1 To uneCol.Count
        uneCol.Remove 1
    Next i
End Sub



Public Sub MiseAJourSelection(uneFenetreFille As Form, unObjSel As Integer, unIndSel As Integer, Optional unControl As Control, Optional unX As Single)
    'Sélection graphique de l'objet graphique Index représentant un carrefour,
    'un feu ou un arret TC et déselection de l'ancien objet sélectionné
    '==> mise en gras de l'ancienne et de la nouvelle sélection
    
    If unObjSel = CarfSel Then
        'Déselection de la sélection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        'Cas d'un carrefour à sélectionner
        MettreEnGras uneFenetreFille.NomCarf(unIndSel)
        'Affichage de l'onglet Carrefours
        uneFenetreFille.TabFeux.Tab = 0
    ElseIf unObjSel = FeuSel Then
        'Déselection de la sélection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        'Cas d'un feu à sélectionner
        MettreEnGras uneFenetreFille.NumFeu(unIndSel)
        'Affichage de l'onglet Carrefours
        uneFenetreFille.TabFeux.Tab = 0
    ElseIf unObjSel = ArretSel Then
        'Cas d'un arrêt TC à sélectionner
        'Affichage de l'onglet TC déclenché après une sélection graphique
        uneFenetreFille.TabFeux.Tab = 1
        'Déselection de la sélection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        If TypeOf unControl Is Label Then
            'Cas d'un arrêt sélectionné par son label nomArret,
            'si sélection par l'icone STOP on ne passe pas ici
            'Recherche du label NomArret placé en unX parmi les arrêts confondus
            RechercherArretEnX uneFenetreFille, unControl, unX
            unIndSel = uneFenetreFille.monIndSel
        End If
        'Mise en gras de la sélection
        MettreEnGras uneFenetreFille.NomArret(unIndSel)
        'Rafraichissement de la frame contenant les donnéees d'un TC
        'Pour éviter l'apparition d'un tableau TabYarret à moitié (Bug Spread)
        uneFenetreFille.FrameTC.Refresh
        'Récupération des positions du TC et du Y de l'arrêt sélectionné dans leurs
        ' collections respectives à partir du tag (codage pos TC-pos YArret) du NomArret cliqué
        unePos = InStr(1, uneFenetreFille.NomArret(unIndSel).Tag, "-")
        unePosTC = Val(Mid$(uneFenetreFille.NomArret(unIndSel).Tag, 1, unePos - 1))
        unePosY = Val(Mid$(uneFenetreFille.NomArret(unIndSel).Tag, unePos + 1))
        'Modification du nombre de colonnes de TabYArret pour la future cellule
        'active en fasse partie, sinon plantage
        uneFenetreFille.TabYArret.MaxCols = uneFenetreFille.mesTC(unePosTC).mesArrets.Count
        'Mise en actif de la cellule contenant le Y de l'arrêt sélectionné déclenché après une sélection graphique
        uneFenetreFille.TabYArret.Row = 1
        uneFenetreFille.TabYArret.Col = unePosY
        uneFenetreFille.TabYArret.Action = SS_ACTION_ACTIVE_CELL
        'Affichage du TC de l'arrêt sélectionné dans comboTC déclenché après une sélection graphique
        uneFenetreFille.ComboTC.ListIndex = unePosTC - 1
        'Affichage du libellé de l'arrêt sélectionné
        uneFenetreFille.TextArret.Text = uneFenetreFille.mesTC(unePosTC).mesArrets(unePosY).monLibelle
    End If
    'Stockage des valeurs de la nouvelle sélection
    uneFenetreFille.monObjSel = unObjSel
    uneFenetreFille.monIndSel = unIndSel
    'Mise à jour des contextes d'aide
    ChangerHelpID uneFenetreFille.TabFeux.Tab
End Sub

Public Sub Deselectionner(uneFenetreFille As Form)
    Dim unControl As Control
    
    If uneFenetreFille.monIndSel <> 0 Then
        'Cas d'une sélection précédente
        If uneFenetreFille.monObjSel = CarfSel Or uneFenetreFille.monObjSel = FeuSel Then
            'Cas d'un feu et d'un carrefour à désélectionner
            'Déselection du dernier feu sélectionné
            Set unControl = uneFenetreFille.NumFeu(uneFenetreFille.monIndSel)
            EnleverGras unControl
            'Récupération du carrefour contenant le dernier feu sélectionné
            'par décodage du tag de l'objet graphique NumFeu à déselectionner
            unePos = InStr(1, unControl.Tag, "-")
            unePosCarf = Val(Mid$(unControl.Tag, 1, unePos - 1))
            'Récupération de l'objet graphique du carrefour
            Set unControl = uneFenetreFille.mesCarrefours(unePosCarf).monCarfGraphic
            'Déselection du carrefour contenant le dernier feu sélectionné
            EnleverGras unControl
        ElseIf uneFenetreFille.monObjSel = ArretSel Then
            'Cas d'un arrêt TC à désélectionner
            EnleverGras uneFenetreFille.NomArret(uneFenetreFille.monIndSel)
        End If
    End If
    uneFenetreFille.monIndSel = 0
End Sub
Public Sub RechercherArretEnX(uneFenetreFille As Form, unControl As Control, unX As Single)
    'Recherche de l'arrêt se trouve sous la souris en unX, unY
    'parmi les arrêts confondus éventuels
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unNbArretsConfondus As Integer, unYArret As Long
    Dim unTC As TC
    
    'Récupération des positions du TC et du Y de l'arrêt dans les collections
    'à partir du tag (codage pos TC-pos YArret) du NomArret cliqué
    unePos = InStr(1, unControl.Tag, "-")
    unePosTC = Val(Mid$(unControl.Tag, 1, unePos - 1))
    unePosY = Val(Mid$(unControl.Tag, unePos + 1))
    'Recherche des arrêts confondus
    unYArret = uneFenetreFille.mesTC(unePosTC).mesArrets(unePosY).monOrdonnee
    unNbArretsConfondus = uneFenetreFille.RechercherArretConfondu(unYArret, uneListeIndexTC, uneListeIndexArret)
    'Recherche de la position du label NomArret se trouvant sous le click souris en unX, unY
    uneFenetreFille.LabelTrait.Caption = DonnerStringDecalage
    'Calcul du nombre de décalage ce qui donne le label dont le nom est cliqué
    unIndex = 1 + unX \ uneFenetreFille.LabelTrait.Width '\ = division entière entre 2 entiers
    If unIndex <= unNbArretsConfondus Then
        'Cas d'un click sur un des noms TC et pas en dehors
        'sur la droite dans le souligné, dans ce cas la sélection
        'sera celui au premier plan, c'est celui que donne VB par défaut
        
        'Récupération du TC touvé
        Set unTC = uneFenetreFille.mesTC(uneListeIndexTC(unIndex))
        'Récupération du control NomArret ou IconeArret trouvé
        Set unControl = unTC.mesObjGraphics(uneListeIndexArret(unIndex))
    End If
    'Affectation de l'indice sélectionné
    uneFenetreFille.monIndSel = unControl.Index
End Sub


Public Function DonnerStringDecalage()
    'Donner la chaine permettant de décaler les noms d'arrêts TC confondus
    'Elle est de la même longueur que la longueur maximun des Noms de TC
    DonnerStringDecalage = ""
    For i = 1 To NbCarMaxNomTC + 3
        '+ 3 pour tenir compte de la mise en gras lors
        'de la sélectionet des fontes proportionnelles
        DonnerStringDecalage = DonnerStringDecalage + "_"
    Next i
End Function

Public Sub MettreEnGras(unControl As Control)
    'Stockage de la largeur avant mise en gras
    uneWidth = unControl.Width
    'Mise en gras
    unControl.Font.Bold = True
    'Ajustement de la largeur à celle initiale
    unControl.Width = uneWidth
End Sub
Public Sub EnleverGras(unControl As Control)
    'Stockage de la largeur avant la suppression de la mise en gras
    uneWidth = unControl.Width
    'Suppression de la mise en gras
    unControl.Font.Bold = False
    'Ajustement de la largeur à celle initiale
    unControl.Width = uneWidth
End Sub

Public Sub MiseAJourSelectionParCellule(uneFenetreFille As Form, unObjSel As Integer, unIndPere As Long, unIndFils As Long)
    'Sélection graphique de l'objet graphique Index représentant un carrefour,
    'un feu ou un arret TC et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    'Déselection de la sélection courante qui va donc devenir l'ancienne
    Call Deselectionner(uneFenetreFille)
    If unObjSel = CarfSel Or unObjSel = FeuSel Then
        'Cas d'un carrefour à sélectionner
        '==> Sélection du carrefour donc on le met en gras
        Set unControl = uneFenetreFille.mesCarrefours(unIndPere).monCarfGraphic
        MettreEnGras unControl
        '==> Et sélection de son feu numéro unIndFeu qu'on met donc en gras
        Set unControl = uneFenetreFille.mesCarrefours(unIndPere).mesFeuxGraphics(unIndFils)
        MettreEnGras unControl
        'Stockage de l'indice de l'objet de la nouvelle sélection
        'On stocke l'objet graphique du feu car grâce son tag
        'on retourve le carrefour et son feu
        uneFenetreFille.monIndSel = unControl.Index
    ElseIf unObjSel = ArretSel Then
        'Cas d'un arrêt TC à sélectionner
        'Récupération du control NomArret correspondant à la colonne active
        Set unControl = uneFenetreFille.mesTC(unIndPere).mesObjGraphics(unIndFils)
        'Mise à jour de la sélection
        MettreEnGras unControl
        'Affichage du libellé de l'arrêt sélectionné
        uneFenetreFille.TextArret.Text = uneFenetreFille.mesTC(unIndPere).mesArrets(unIndFils).monLibelle
        'Stockage de l'indice de l'objet de la nouvelle sélection
        uneFenetreFille.monIndSel = unControl.Index
    End If
    'Stockage du type d'objet de la nouvelle sélection
    uneFenetreFille.monObjSel = unObjSel
End Sub


Public Sub CreerArretTC(uneFenetreFille As Form)
    Dim unIndTC As Long, unTC As TC
    Dim unNbArret As Integer
    Dim unYArret As Integer
    Dim unArret As ArretTC
    Dim unLibelle As String
    Dim unYMax As Integer, unYMin As Integer
    Dim unYFeuMax As Integer, unYFeuMin As Integer
    
    'Récupération du TC par sa position
    unIndTC = uneFenetreFille.ComboTC.ListIndex + 1
    Set unTC = uneFenetreFille.mesTC(unIndTC)
    'Récupération du nombre d'arrêt avant la création du nouveau
    unNbArret = uneFenetreFille.TabYArret.MaxCols
    
    'Recherche de l'ordonnée unYFeuMax qui est la plus grande
    'parmi les feux du carrefour de départ et d'arrivée
    unYFeuMax = DonnerYMaxCarf(unTC.monCarfDep)
    unYMax = DonnerYMaxCarf(unTC.monCarfArr)
    If unYMax > unYFeuMax Then unYFeuMax = unYMax
    
    'Recherche de l'ordonnée unYFeuMin qui est la plus petite
    'parmi les feux du carrefour de départ et d'arrivée
    unYFeuMin = DonnerYMinCarf(unTC.monCarfDep)
    unYMin = DonnerYMinCarf(unTC.monCarfArr)
    If unYMin < unYFeuMin Then unYFeuMin = unYMin
    
    'Recherche de l'arrêt ayant l'ordonnée la plus grande mais inférieure
    'à unYFeuMax. Certains Y d'arrêts peuvent > à unYFeuMax lors de changement
    'de carrefour de départ et/ou d'arrivée ou d'inversion du sens du TC
    unYMax = unYFeuMin 'Intialisation du Ymax avec le Ymin des feux
    For i = 1 To unNbArret
        Set unArret = uneFenetreFille.mesTC(unIndTC).mesArrets(i)
        If unArret.monOrdonnee > unYMax And unArret.monOrdonnee <= unYFeuMax Then
            unYMax = unArret.monOrdonnee
        End If
    Next i
            
    'On met le nouvel arrêt au milieu entre le feu ayant le plus grand Y parmi
    'les carrefours départ et arrivée et l'arrêt ayant le Y le plus grand du TC
    unYArret = (unYMax + unYFeuMax) / 2
    If CInt(unYArret) = unYMax Then
        unMsg = "Plus de place pour insérer un nouvel arrêt TC."
        unMsg = unMsg + Chr(13) + "Changer soit le carrefour d'arrivée pour un TC montant, soit celui de départ pour un TC descendant."
        MsgBox unMsg, vbInformation
        Exit Sub
    End If
    
    'Ajout d'un nouvel Y d'arrêt TC
    uneFenetreFille.TabYArret.MaxCols = unNbArret + 1
    uneFenetreFille.TabYArret.Col = uneFenetreFille.TabYArret.MaxCols
    uneFenetreFille.TabYArret.Row = 1
    uneFenetreFille.TabYArret.Text = Format(unYArret)
    
    'Création d'une instance d'arrêt
    unLibelle = "Arrêt " + Format(unNbArret + 1) + " de " + uneFenetreFille.mesTC(unIndTC).monNom
    Set unArret = uneFenetreFille.mesTC(unIndTC).mesArrets.Add(unYArret, 15, 30, unLibelle)
    'alimentation des lignes 2 et 3 du nouvel arrêt
    uneFenetreFille.TabYArret.Row = 2
    uneFenetreFille.TabYArret.Text = Format(unArret.maVitesseMarche)
    uneFenetreFille.TabYArret.Row = 3
    uneFenetreFille.TabYArret.Text = Format(unArret.monTempsArret)
    'On rend active la colonne nouvellement créée
    uneFenetreFille.TabYArret.Action = SS_ACTION_ACTIVE_CELL
    'Création des objets graphiques
    uneFenetreFille.DessinerArretTC unIndTC, CLng(unYArret)
    'Indication d'une modification dans les données TC
    IndiquerModifTC
End Sub

Public Sub SupprimerArretTC(uneFenetreFille As Form)
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unControl As Control
    Dim unY As Long, unePosTC As Long, uneColDel As Long
    
    unMsg = "Etes-vous sûr de vouloir supprimer l'arrêt " + str(uneFenetreFille.TabYArret.ActiveCol)
    unMsg = unMsg + " du transport collectif " + UCase(uneFenetreFille.ComboTC.Text) + " ?"
    If uneFenetreFille.TabYArret.MaxCols = 1 Then
        unMsg = "Un transport collectif sans arrêt ne sert à rien." + Chr(13) + Chr(13)
        unMsg = unMsg + "Supprimer plutôt le transport collectif."
        MsgBox unMsg, vbInformation
    ElseIf MsgBox(unMsg, vbYesNo + vbQuestion) = vbYes Then
        uneColDel = uneFenetreFille.TabYArret.ActiveCol
        'Stockage du Y de l'arrêt supprimé pour modifier
        'les décalages un peu plus bas
        uneFenetreFille.TabYArret.Row = 1
        uneFenetreFille.TabYArret.Col = uneColDel
        unY = Val(uneFenetreFille.TabYArret.Text)
        'Suppression du Y de l'arrêt du TC
        unePosTC = uneFenetreFille.ComboTC.ListIndex + 1
        uneFenetreFille.mesTC(unePosTC).mesArrets.Remove uneColDel
        'Suppression des objets graphiques (NomArret et IconeArret)
        'de l'arrêt TC en sachant que mesObjGraphics sont des NomArret
        Set unControl = uneFenetreFille.mesTC(unePosTC).mesObjGraphics(uneColDel)
        Unload uneFenetreFille.IconeArret(unControl.Index)
        Unload unControl
        uneFenetreFille.mesTC(unePosTC).mesObjGraphics.Remove uneColDel
        'Recherche des arrêts confondus en unY pour alimenter
        'les listes d'arrêts et de TC trouvés
        unNb = uneFenetreFille.RechercherArretConfondu(unY, uneListeIndexTC, uneListeIndexArret)
        'Mise à jour des décalages des labels NomArrêt
        Call MiseAJourNomArret(uneFenetreFille, uneListeIndexTC, uneListeIndexArret)
        'Décalage des colonnes du tableau représentant les arrêts du TC
        'et mise à jour des tags des objets graphiques des arrêts suivants
        For i = uneColDel To uneFenetreFille.TabYArret.MaxCols - 1
            For j = 1 To 3
                'positionnment en ligne j
                uneFenetreFille.TabYArret.Row = j
                'Récupération du contenu de la cellule i + 1
                uneFenetreFille.TabYArret.Col = i + 1
                uneStrTmp = uneFenetreFille.TabYArret.Text
                'Affectation de la cellule i
                uneFenetreFille.TabYArret.Col = i
                uneFenetreFille.TabYArret.Text = uneStrTmp
            Next j
            'Mise à jour des tags des objets graphiques des arrêts suivants
            Set unControl = uneFenetreFille.mesTC(unePosTC).mesObjGraphics(i)
            unControl.Tag = Format(unePosTC) + "-" + Format(i)
            uneFenetreFille.IconeArret(unControl.Index).Tag = unControl.Tag
        Next i
        'Suppression de la colonne de l'arrêt TC dans le spread TabYArret
        'Sélection d'une colonne
        uneFenetreFille.TabYArret.Col = uneColDel
        ' Suppression de la colonne sélectionnée
        uneFenetreFille.TabYArret.Action = SS_ACTION_DELETE_COL
        uneFenetreFille.TabYArret.MaxCols = uneFenetreFille.TabYArret.MaxCols - 1
        
        'Mise à jour de la sélection graphique
        'Le dernier sélectionné a été détruit ==> Sélection vide pour ne rien déselectionner
        uneFenetreFille.monIndSel = 0
        If uneColDel = uneFenetreFille.TabYArret.MaxCols + 1 Then
            'Cas où l'on supprime le dernier arrêt,
            'on sélectionnera le nouveau dernier
            uneColDel = uneColDel - 1
        End If
        'Sélection graphique à partir de la cellule active, c'est le nouveau
        'dernier arrêt si l'on a supprimé le dernir arrêt, ou l'arrêt précédent
        'celui qui a été supprimé dans les autres cas.
        'Sélection graphique de l'arrêt correspondant à la cellule active
        '==> colonne active celle d'indice ColDel
        MiseAJourSelectionParCellule uneFenetreFille, ArretSel, unePosTC, uneColDel
        
        'Indication d'une modification dans les données TC
        IndiquerModifTC
    End If
End Sub
Public Sub CreerFeu(uneFenetreFille As Form)
    'Création d'un feu du carrefour courant du site courant (uneFenetreFille)
    Dim unY As Integer
    Dim unFeu As Feu
    Dim unSensMontant As Boolean
    
    With uneFenetreFille
        If .monCarrefourCourant.mesFeux.Count = 0 Then
        'Cas de la création du premier feu d'un carrefour
            If .mesCarrefours.Count = 1 Then
                'Cas de la création du premier carrefour
                unY = 0
            ElseIf .mesCarrefours.Count = 2 Then
                'Cas de la création du deuxième carrefour
                'Premier feu mis à 500 m du feu max
                unY = .monYMaxFeu + 500
                'Recalcul du Y min des feux = Y min du premier carrefour
                '==> Englobant en Y min et max OK
                .monYMinFeu = DonnerYMinCarf(.mesCarrefours(1))
            Else
                'Cas d'un nouveau carrefour autre que le premier
                'le nouveau sera mis en dernier lors de sa création
                unY = .monYMaxFeu
                'Mise à 500 mètres du premier feu du nouveau carrefour par
                'rapport au carrefour dont le feu correspond au Y le plus grand
                unY = unY + 500
            End If
        Else
            'Cas de la création d'un feu autre que le premier
            'On le met à 20 mètres du Y le plus grand parmi tous les Y
            'des feux du carrefour courant auquel on ajoute ce nouveau feu
            unY = 20 + DonnerYMaxCarf(.monCarrefourCourant)
        End If
    
        'Calcul du sens du nouveau feu
        'par défaut basé sur l'indice de création :
        'si impair ==> montant, si pair descendant
        If (.monCarrefourCourant.mesFeux.Count + 1) Mod 2 = 0 Then
            unSensMontant = False
        Else
            unSensMontant = True
        End If
        
        'Ajout d'un nouveau feu
        Set unFeu = .monCarrefourCourant.mesFeux.Add(unSensMontant, unY, .maDuréeDeCycle / 2, 0)
        'Stockage du carrefour du feu créé
        Set unFeu.monCarrefour = .monCarrefourCourant
        'Ajout d'une nouvelle ligne pour le nouveau feu
        .TabPropCarf.MaxRows = .monCarrefourCourant.mesFeux.Count
        'Mise à jour des titres des rangées
        .TabPropCarf.Col = 0
        .TabPropCarf.Row = .TabPropCarf.MaxRows
        .TabPropCarf.Text = "Feu " + str(.TabPropCarf.Row)
        'Affichage des valeurs par défaut
        .TabPropCarf.Col = 1
        If unSensMontant Then
            .TabPropCarf.Text = "Montant"
        Else
            .TabPropCarf.Text = "Descendant"
        End If
        
        .TabPropCarf.Col = 2
        .TabPropCarf.Text = Format(unFeu.monOrdonnée)
        
        .TabPropCarf.Col = 3
        .TabPropCarf.Text = Format(unFeu.maDuréeDeVert)
        
        .TabPropCarf.Col = 4
        .TabPropCarf.Text = Format(unFeu.maPositionPointRef)
        
        'On rend actif dans TabPropCarf la ligne du dernier feu créé
        .TabPropCarf.Col = 1
        .TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
        'Modification de la position du label Nom de carrefour
        'au barycentre des Y de ses feux
        ModifYNomCarf uneFenetreFille, .monCarrefourCourant
    End With
    'Création des objets graphiques du feu numéro TabPropCarf.MaxRows
    DessinerFeu uneFenetreFille, uneFenetreFille.monCarrefourCourant.maPosition, uneFenetreFille.TabPropCarf.MaxRows
    'Sélection du dernier feu créé avec son carrefour
    MiseAJourSelectionParCellule uneFenetreFille, FeuSel, uneFenetreFille.monCarrefourCourant.maPosition, uneFenetreFille.TabPropCarf.MaxRows
    'Redessin avec le bon niveau de zoom, celui maximun englobant tous les feux
    RedessinerTout uneFenetreFille, CLng(unY) 'unY converti en entier long
    'Indication d'une modification dans les données carrefour
    uneFenetreFille.maModifDataCarf = True
End Sub

Public Function DonnerYCarrefour(unCarf As Carrefour) As Integer
    'Calcul de l'ordonnée du carrefour en prenant le barycentre du Y de ses feux
    Dim unYMoyen As Double
    Dim unNbFeux As Integer
    
    unYMoyen = 0
    unNbFeux = unCarf.mesFeux.Count
    
    For i = 1 To unNbFeux
        unYMoyen = unYMoyen + unCarf.mesFeux(i).monOrdonnée
    Next i
    
    DonnerYCarrefour = Fix(unYMoyen / unNbFeux)
End Function

Public Sub CreerCarrefour(uneFenetreFille As Form)
    Dim unNom As String, uneClé As String
    Dim uneValeurDefaut As String
    
    With uneFenetreFille
        'Nom générique à partir du nombre total d'objets graphiques carrefours
        'créés dans ce site.
        'Uniquement pour le premier carrefour lors de la création d'un nouveau site d'études
        'On ajoute +1 car monNbObjGraphicCarf est incrémentée plus tard dans
        'DessinerCarrefour ==> cohérence
        If .monNbObjGraphicCarf = 0 Then
            'Cas du premier carrefour créé, c'est nouveau site
            'qui le fait ==> Nom généré automatiquement
            unNom = "Carrefour " + Format(.monNbObjGraphicCarf + 1)
        Else
            'Saisie et Verification que le nom saisie n'existe pas,
            'sinon demande de modif à l'utilisateur
            Do
                unMsg = "Entrez un nom de carrefour (15 caractères maximun):"
                unTitre = "Création d'un carrefour" ' Définit le titre.
                unNom = InputBox(unMsg, unTitre, uneValeurDefaut)
                unNom = Trim(unNom) 'Suppression des blancs avant et après
                uneValeurDefaut = unNom
                If Len(unNom) > 15 Then
                    unMsg = "Le nom d'un carrefour est limité à 15 caractères"
                    MsgBox unMsg, vbCritical
                    uneSortie = False
                ElseIf Trim(unNom) = "" Then
                    'Cas du click sur le bouton annuler ou sur OK sans rentrer de nom
                    '==> Sortie sans rien faire comme un annuler
                    Exit Sub
                ElseIf PosInListe(unNom, uneFenetreFille.ComboNomCarf) <> -1 Then
                    'Cas où le nom existe déjà
                    unMsg = "Le carrefour " + UCase(unNom) + " existe déjà."
                    MsgBox unMsg, vbCritical
                    uneValeurDefaut = unNom
                    uneSortie = False
                Else
                    uneSortie = True
                End If
            Loop While uneSortie = False
        End If
        'Création du nouveau carrefour avec son nom unique
        Set .monCarrefourCourant = .mesCarrefours.Add(unNom, .maVitSensM, .maVitSensD)
        'Affectation des valeurs par défaut des demandes et des débits de saturation
        .monCarrefourCourant.SetDemDeb 0, 1800, 0, 1800
        'Création du label NomCarf du carrefour qui sera mis en dernier position
        'dans la collection mesCarrefours
        DessinerCarrefour uneFenetreFille, uneFenetreFille.mesCarrefours.Count
        'Création et Affichage du premier feu
        CreerFeu uneFenetreFille
        'Mise à jour de la combobox listant les noms de carrefours
        .ComboNomCarf.AddItem unNom
        .ComboNomCarf.ListIndex = .ComboNomCarf.ListCount - 1
        'Mise à jour des combobox des TC listant les carrefours
        'de départ et d'arrivée possibles
        .ComboCarfDep.AddItem unNom
        .ComboCarfArr.AddItem unNom
        'Mise à jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
        .TabInfoCalc.MaxRows = .mesCarrefours.Count
        RemplirLigneTabInfoCalc uneFenetreFille, .monCarrefourCourant.maPosition
        'Mise à jour du tableau TabDecal de l'onglet Tableau de résultat
        .TabDecal.MaxRows = .mesCarrefours.Count
    End With
End Sub

Public Sub DessinerFeu(uneFenetreFille As Form, unIndCarf As Integer, unIndFeu As Integer)
    Dim unePos As Long, unYreel As Long
    Dim unYMaxFeu As Integer
    
    With uneFenetreFille
        unYreel = .monYMaxFeu - .mesCarrefours(unIndCarf).mesFeux(unIndFeu).monOrdonnée
        'Conversion du Yréel en Y écran dans la FrameVisuCarf
        unePos = ConvertirReelEnEcran(unYreel, .maLongueurAxeY, .AxeOrdonnée.Y2 - .AxeOrdonnée.Y1)
        'Incrémentation du nombre d'objets graphiques Feu créés
        .monNbObjGraphicFeu = .monNbObjGraphicFeu + 1
        i = .monNbObjGraphicFeu
        'Création du label pour le numéro du feu
        Load .NumFeu(i)
        'Création de l'icone graphique FEU tricolore du feu
        Load .IconeFeu(i)
        'Positionnement du feu (Numéro et icône Feu) à droite de l'axe des Y
        'pour un feu montant et à gauche pour un feu descendant
        PlacerFeuAxeY uneFenetreFille, unIndCarf, unIndFeu, i
        'Positionnement en Y écran
        .NumFeu(i).Top = unePos + .AxeOrdonnée.Y1 - .NumFeu(i).Height
        .IconeFeu(i).Top = unePos + .AxeOrdonnée.Y1 - .IconeFeu(i).Height
        'Affichage des objets graphiques du feu
        .NumFeu(i).Visible = True
        .IconeFeu(i).Visible = True
        'Codage permettant de retrouver le carrefour et son feu à partir des objets graphiques
        'Tag = index dans la collection des carrefours plus un tiret et le numéro du feu
        .NumFeu(i).Tag = Format(unIndCarf) + "-" + Format(unIndFeu)
        .IconeFeu(i).Tag = .NumFeu(i).Tag
        'Stockage dans la liste des objets graphiques représentant le feu
        .mesCarrefours(unIndCarf).mesFeuxGraphics.Add .NumFeu(i)
    End With
End Sub
Public Sub DessinerCarrefour(uneFenetreFille As Form, unIndCarf As Integer)
    Dim unePos As Long
    
    With uneFenetreFille
        'Conversion du Yréel = 0 en Y écran dans la FrameVisuCarf
        unePos = ConvertirReelEnEcran(0, .maLongueurAxeY, .AxeOrdonnée.Y2 - .AxeOrdonnée.Y1)
        'Incrémentation du nombre d'objets graphiques Carrefour créés
        .monNbObjGraphicCarf = .monNbObjGraphicCarf + 1
        i = .monNbObjGraphicCarf
        'Création du label pour le nom du carrefour
        Load .NomCarf(i)
        .NomCarf(i).Caption = .mesCarrefours(unIndCarf).monNom
        'Positionnement en Y écran
        .NomCarf(i).Top = unePos + (.AxeOrdonnée.Y2 + .AxeOrdonnée.Y1) / 2 - .NomCarf(i).Height
        'Affichage des objets graphiques du feu
        .NomCarf(i).Visible = True
        'Codage permettant de retrouver le carrefour à partir des objets graphiques
        'Tag = index dans la collection des carrefours
        .NomCarf(i).Tag = Format(unIndCarf)
        'Stockage dans la liste des objets graphiques représentant le feu
        Set .mesCarrefours(unIndCarf).monCarfGraphic = .NomCarf(i)
    End With
End Sub


Public Sub ModifYNomCarf(uneFenetreFille As Form, unCarf As Carrefour)
    'Modification de la position du label Nom du carrefour unCarf
    'au barycentre des Y de ses feux
    Dim unYreel As Long
    
    unYreel = DonnerYCarrefour(unCarf)
    'Conversion du Yréel en Y écran dans la FrameVisuCarf
    unePos = ConvertirReelEnEcran(uneFenetreFille.monYMaxFeu - unYreel, uneFenetreFille.maLongueurAxeY, uneFenetreFille.AxeOrdonnée.Y2 - uneFenetreFille.AxeOrdonnée.Y1)
    'Positionnement en Y écran
    unCarf.monCarfGraphic.Top = unePos + uneFenetreFille.AxeOrdonnée.Y1 - unCarf.monCarfGraphic.Height
End Sub

Public Sub AfficherValeursCarrefour(uneForm As Form, unCarf As Carrefour)
    With uneForm
        'Mise à jour de la combobox listant les noms de carrefours
        .ComboNomCarf.ListIndex = unCarf.maPosition - 1
        
        .TabPropCarf.MaxRows = unCarf.mesFeux.Count
        For i = 1 To unCarf.mesFeux.Count
            .TabPropCarf.Row = i
            .TabPropCarf.Col = 0
            .TabPropCarf.Text = "Feu " + Format(i)
            .TabPropCarf.Col = 1
            If unCarf.mesFeux(i).monSensMontant Then
                .TabPropCarf.Text = "Montant"
            Else
                .TabPropCarf.Text = "Descendant"
            End If
            .TabPropCarf.Col = 2
            .TabPropCarf.Text = Format(unCarf.mesFeux(i).monOrdonnée)
            .TabPropCarf.Col = 3
            .TabPropCarf.Text = Format(unCarf.mesFeux(i).maDuréeDeVert)
            .TabPropCarf.Col = 4
            .TabPropCarf.Text = Format(-unCarf.mesFeux(i).maPositionPointRef)
            '-PosRef car définition inverse entre dossier programmation et doc logiciel OndeV
        Next i
    End With
End Sub

Public Sub MiseAJourSelectionEtOngletCarrefour(uneForm As Form, unObjSel As Integer, unePosCarf As Long, unePosFeu As Long)
    'Mise à jour sélection graphique et l'onglet Carrefour
    
    'On rend actif dans TabPropCarf la 1 ère ligne et 1ère colonne
    'pour corriger un bug dans une cellule combobox du spread
    'En fait si juste avant de cliquer sur un carrefour, on a tapé M ou D dans
    'la 1ère colonne et sur un feu de numéro > au nombre de feux du carrefour cliqué
    '==> plantage indébuggable pas à pas, c'est la seule correction trouvée
    uneForm.TabPropCarf.Row = 1
    uneForm.TabPropCarf.Col = 1
    uneForm.TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
    
    'Affichage de l'onglet Carrefour
    uneForm.TabFeux.Tab = 0
    'Mise à jour sélection grâce à la cellule courante
    Call MiseAJourSelectionParCellule(uneForm, unObjSel, unePosCarf, unePosFeu)
    'Affichage des valeurs du carrefour sélectionné qui devient le courant
    Set uneForm.monCarrefourCourant = uneForm.mesCarrefours(unePosCarf)
    AfficherValeursCarrefour uneForm, uneForm.monCarrefourCourant
    'On rend actif dans TabPropCarf la ligne du feu sélectionné
    uneForm.TabPropCarf.Row = unePosFeu
    uneForm.TabPropCarf.Col = 1
    uneForm.TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
End Sub

Public Sub DonnerPosCarfFeu(unControl As Control, unePosCarf As Long, unePosFeu As Long)
    'Récupération du carrefour et du feu par décodage du tag
    'de l'objet graphique unControl de type NumFeu ou IconeFeu
    unePos = InStr(1, unControl.Tag, "-")
    unePosCarf = Val(Mid$(unControl.Tag, 1, unePos - 1))
    unePosFeu = Val(Mid$(unControl.Tag, unePos + 1))
End Sub

Public Sub SupprimerFeu(uneForm As Form)
    Dim unCarf As Carrefour
    Dim unControl As Control
    Dim unIndFeu As Long
    Dim unIndTC As Integer
    Dim uneColPosTC As Collection
    
    Set unCarf = uneForm.monCarrefourCourant
    unIndFeu = uneForm.TabPropCarf.ActiveRow
    
    'Test préliminaire avant la destruction du feu
    unMsg = "Etes-vous sûr de vouloir supprimer le feu " + Format(unIndFeu)
    unMsg = unMsg + " du carrefour " + unCarf.monNom + " ?"
    If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
        'Cas de confirmation négative
        Exit Sub
    End If
    
    If unCarf.mesFeux.Count = 1 Then
        'Cas où le carrefour ne contient qu'un feu
        unMsg = "Le carrefour " + unCarf.monNom + " ne contient qu'un feu. Un carrefour sans feu n'ayant aucun interêt, supprimez plutôt le carrefour."
        MsgBox unMsg, vbCritical
    Else
        'Cas où l'on peut faire la suppression
        'Suppression des objets graphiques du feu
        Unload uneForm.IconeFeu(unCarf.mesFeuxGraphics(unIndFeu).Index)
        Unload unCarf.mesFeuxGraphics(unIndFeu)
        'Stockage du Y du feu qui va être supprimé pour utilisation plus loin
        unYOld = unCarf.mesFeux(unIndFeu).monOrdonnée
        'Suppression dans les feux et les objets graphiques du carrefour du feu
        unCarf.mesFeuxGraphics.Remove unIndFeu
        unCarf.mesFeux.Remove unIndFeu
        'Modification des feux restants du carrefour
        For i = unIndFeu To unCarf.mesFeux.Count
            'Modification de l'attribut maPosition des autres feux
            unCarf.mesFeux(i).maPosition = i
            'Modification du contenu du tag valant PositionCarrefour-PositionFeu
            Set unControl = unCarf.mesFeuxGraphics(i)
            unControl.Tag = Format(unCarf.maPosition) + "-" + Format(i)
            'Modification du label des autres feux
            uneStringDecalage = "___" 'Décalage par rapport à l'axe des Y
            If unCarf.mesFeux(i).monSensMontant Then
                'Cas des feux montant
                '==> Positionnement à droite de l'axe des Y avec 2 blancs à la fin
                'pour que la mise en gras ne chevauche pas l'icone Feu
                unControl.Caption = uneStringDecalage + Format(i) + "  "
                unControl.Left = uneForm.AxeOrdonnée.X1
                'Positionnement de l'icone Feu tricolore
                'à l'extrémité droite du numéro de feux
                uneForm.IconeFeu(unControl.Index).Left = unControl.Left + unControl.Width
            Else
                'Cas des feux descendant
                '==> Positionnement à gauche de l'axe des Y avec un
                'souligné plus grand pour intersecter l'axe des Y
                unControl.Caption = Format(i) + uneStringDecalage + uneStringDecalage
                'Ajustement de la chaine de caractéres à l'axe des ordonnées car la
                'propriété AutoSize est à true ==> restriction du souligné précédent
                unControl.Width = uneForm.AxeOrdonnée.X1 - uneForm.NumFeu(0).Left 'Le left de l'indice n'a pas bougé
                unControl.Left = uneForm.AxeOrdonnée.X1 - unControl.Width
                'Positionnement de l'icone Feu tricolore
                'à l'extrémité gauche du numéro de feux
                uneForm.IconeFeu(unControl.Index).Left = unControl.Left - uneForm.IconeFeu(unControl.Index).Width
            End If
        Next i
        'Mise à jour à 0 de l'indice d'objet graphique selectionné ==> Déselection
        uneForm.monIndSel = 0
        'Mise à jour de la sélection et de l'onglet carrefour en
        'sélectionnant le feu précédent celui supprimé
        If unIndFeu > unCarf.mesFeux.Count Then unIndFeu = unIndFeu - 1
        MiseAJourSelectionEtOngletCarrefour uneForm, FeuSel, unCarf.maPosition, unIndFeu
        'Modification de la position du label Nom de carrefour
        'au barycentre des Y de ses feux
        ModifYNomCarf uneForm, unCarf
        'Redessin au bon niveau de zoom si le feu détruit était
        'une des limites de l'englobant
        If unYOld = uneForm.monYMaxFeu Or unYOld = uneForm.monYMinFeu Then
            'Cas où le Y du feu détruit était le maximun ou le minimun des Y
            '==> Modification de l'englobant d'où recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        End If
        'Indication d'une modification dans les données carrefour
        uneForm.maModifDataCarf = True
    End If
End Sub

Public Sub SupprimerCarrefour(uneForm As Form)
    Dim unCarf As Carrefour
    Dim unControl As Control
    Dim unIndCarf As Long
    Dim i As Long
         
    Set unCarf = uneForm.monCarrefourCourant
    unIndCarf = unCarf.maPosition
    
    'Test préliminaire avant la destruction du carrefour
    unMsg = "Etes-vous sûr de vouloir supprimer le carrefour "
    unMsg = unMsg + unCarf.monNom + " ?"
    If uneForm.mesCarrefours.Count = 1 Then
        'Cas où l'on a détruit l'unique carrefour
        unMsg = "Il n'y a aucun intérêt à supprimer le seul et unique carrefour."
        unMsg = unMsg + Chr(13) + Chr(13)
        unMsg = unMsg + "Modifiez ou supprimez plutôt ses feux."
        MsgBox unMsg, vbCritical
        Exit Sub
    ElseIf MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
        'Cas de confirmation négative
        Exit Sub
    End If
    
    'Test de l'utilisation de ce carrefour dans les TC
    unMsg = "Impossible de supprimer le carrefour " + unCarf.monNom
    unMsg = unMsg + " car il est carrefour de départ ou d'arrivée "
    unMsg = unMsg + "des transports collectifs ci-dessous :" + Chr(13) + Chr(13)
    
    unNbTC = 0
    For i = 1 To uneForm.mesTC.Count
        If unCarf.monNom = uneForm.mesTC(i).monCarfDep.monNom Or unCarf.monNom = uneForm.mesTC(i).monCarfArr.monNom Then
            'Cas où le carrefour est un carrefour de départ ou d'arrivée d'un TC
            'car les noms de carrefour sont uniques dans un site
            unMsg = unMsg + "            " + uneForm.mesTC(i).monNom
            unMsg = unMsg + Chr(13)
            unNbTC = unNbTC + 1
        End If
    Next i
    
    If unNbTC > 0 Then
            'Cas où le carrefour est carrefour de départ ou d'arrivée d'un TC
            MsgBox unMsg, vbCritical
    Else
        'Cas où l'on peut faire la suppression du carrefour
        'Suppression de tous les feux et de leurs objets graphiques du carrefour
        uneModifEnglobant = False
        unNbFeux = unCarf.mesFeux.Count
        For i = unNbFeux To 1 Step -1
            Unload uneForm.IconeFeu(unCarf.mesFeuxGraphics(i).Index)
            Unload unCarf.mesFeuxGraphics(i)
            'Test si on supprime un feu dont le Y était le min
            'ou le max des Y des feux
            unY = unCarf.mesFeux(i).monOrdonnée
            If unY = uneForm.monYMaxFeu Or unY = uneForm.monYMinFeu Then
                uneModifEnglobant = True
            End If
            'Suppression dans les feux et les objets graphiques feu tricolore
            unCarf.mesFeuxGraphics.Remove i
            unCarf.mesFeux.Remove i
        Next i
        Set unCarf.mesFeuxGraphics = Nothing
        Set unCarf.mesFeux = Nothing
        'Suppression du carrefour et de son objet graphique
        Unload unCarf.monCarfGraphic
        uneForm.mesCarrefours.Remove unCarf.maPosition
        'Suppresion dans les combobox des TC listant les carrefours
        'de départ et d'arrivée possibles
        uneForm.ComboCarfDep.RemoveItem unCarf.maPosition - 1
        uneForm.ComboCarfArr.RemoveItem unCarf.maPosition - 1
        'Suppression dans la combobox ComboNomCarf
        uneForm.ComboNomCarf.RemoveItem unCarf.maPosition - 1
        'Modification des carrefours restants et de leurs feux, et qui suivait le
        'carrefour supprimé : mise à jour des attributs maPosition et des tag
        For i = unIndCarf To uneForm.mesCarrefours.Count
            'Modification de l'attribut maPosition des autres carrefours
            uneForm.mesCarrefours(i).maPosition = i
            'Modification du tag de l'objet graphique des autres carrefours
            uneForm.mesCarrefours(i).monCarfGraphic.Tag = Format(i)
            'Modification du contenu du tag valant PositionCarrefour-PositionFeu
            'Pour tous les feux du carrefour
            For j = 1 To uneForm.mesCarrefours(i).mesFeux.Count
                Set unControl = uneForm.mesCarrefours(i).mesFeuxGraphics(j)
                unControl.Tag = Format(i) + "-" + Format(j)
            Next j
        Next i
        'Mise à jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
        For i = unIndCarf To uneForm.TabInfoCalc.MaxRows - 1
            'On remplit la ligne i par la i+1 pour les 4 colonnes
            For j = 1 To 4
                uneForm.TabInfoCalc.Row = i + 1
                uneForm.TabInfoCalc.Col = j
                uneVarTmp = uneForm.TabInfoCalc.Text
                uneForm.TabInfoCalc.Row = i
                uneForm.TabInfoCalc.Text = uneVarTmp
            Next j
        Next i
        'Mise à jour du nombre de ligne ==> perte de la dernière ligne
        uneForm.TabInfoCalc.MaxRows = uneForm.mesCarrefours.Count
        'Mise à jour à 0 de l'indice d'objet graphique selectionné ==> Déselection
        uneForm.monIndSel = 0
        'Mise à jour de la sélection et de l'onglet carrefour en
        'sélectionnant le carrefour précédent celui supprimé et son feu 1
        If unIndCarf > uneForm.mesCarrefours.Count Then unIndCarf = unIndCarf - 1
        MiseAJourSelectionEtOngletCarrefour uneForm, CarfSel, unIndCarf, 1
        'Redessin au bon niveau de zoom si le carrefour détruit contenait un
        'feu qui était une des limites de l'englobant
        If uneModifEnglobant Then
            'Cas où le Y du feu détruit était le maximun ou le minimun des Y
            '==> Modification de l'englobant d'où recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        End If
        'Indication d'une modification dans les données carrefour
        uneForm.maModifDataCarf = True
    End If
End Sub

Public Sub ModifierYFeu(uneForm As Form, unCarf As Carrefour, unIndFeu As Integer, unYNew As Long)
    Dim unYOld As Integer
    
    'Stockage de l'ancien Y du feu modifié
    unYOld = unCarf.mesFeux(unIndFeu).monOrdonnée
    'Modification du Y du feu
    unCarf.mesFeux(unIndFeu).monOrdonnée = unYNew
    'Déplacement des objets graphiques du feu
    'Conversion du unYNew valeur réelle en Y écran dans la FrameVisuCarf
    unePos = ConvertirReelEnEcran(uneForm.monYMaxFeu - unYNew, uneForm.maLongueurAxeY, uneForm.AxeOrdonnée.Y2 - uneForm.AxeOrdonnée.Y1)
    'Positionnement en Y écran
    unCarf.mesFeuxGraphics(unIndFeu).Top = unePos + uneForm.AxeOrdonnée.Y1 - unCarf.mesFeuxGraphics(unIndFeu).Height
    unInd = unCarf.mesFeuxGraphics(unIndFeu).Index
    uneForm.IconeFeu(unInd).Top = unePos + uneForm.AxeOrdonnée.Y1 - uneForm.IconeFeu(unInd).Height
    'Modification de la position de l'objet graphique du carrefour
    ModifYNomCarf uneForm, unCarf
    'Test si la modif concerne un des limites de l'englobant
    With uneForm
        If (unYOld = .monYMaxFeu And unYNew < .monYMaxFeu) Or (unYOld = .monYMinFeu And unYNew > .monYMinFeu) Then
            'Cas où l'ancien Y était le maximun et que le nouvel Y est plus petit
            'ou l'ancien Y était le minimun et que le nouvel Y est plus grand
            '==> Modification de l'englobant d'où recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        Else
            'Tous les autres cas sont réglés par la fonction RedessinerTout
            'ci-dessous, qui redessinne avec le bon niveau de zoom, celui
            'maximun englobant tous les feux si unYnew est extérieur à l'englobant
            RedessinerTout uneForm, unYNew
        End If
    End With
End Sub

Public Sub RenommerCarrefour(uneForm As Form)
    Dim unNomCarf As String
    
    If uneForm.mesCarrefours.Count = 0 Then
        MsgBox "Le site ne contient aucun carrefour.", vbCritical
        Exit Sub
    End If
    
    ' Définit le message.
    unMsg = "Entrez le nouveau nom du carrefour (15 caractères maximun):"
    unTitre = "Changement du nom d'un carrefour" ' Définit le titre.
    uneValeurDefaut = uneForm.monCarrefourCourant.monNom
    ' Affiche le message, le titre et la valeur par défaut.
    Do
        unNomCarf = InputBox(unMsg, unTitre, uneValeurDefaut)
        unNomCarf = Trim(unNomCarf) 'Suppression des blancs avant et après
        uneValeurDefaut = unNomCarf
        If Len(unNomCarf) > 15 Then
            unMsg1 = "Le nom d'un carrefour est limité à 15 caractères"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        ElseIf Trim(unNomCarf) = "" Then
            'Cas du click sur le bouton annuler ou sur OK sans rentrer de nom
            '==> Sortie sans rien faire comme un annuler
            uneSortie = True
        ElseIf PosInListe(unNomCarf, uneForm.ComboNomCarf) <> -1 Then
            'Cas où le nom existe déjà
            unMsg1 = "Le carrefour " + UCase(unNomCarf) + " existe déjà"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        Else
            uneSortie = True
            unePos = uneForm.monCarrefourCourant.maPosition - 1
            'Renommage dans la combobox listant les carrefours de départ de TC
            RenommerCarfInCombobox uneForm.ComboCarfDep, unNomCarf, unePos
            'Renommage dans la combobox listant les carrefours d'arrivée de TC
            RenommerCarfInCombobox uneForm.ComboCarfArr, unNomCarf, unePos
            'Renommage du carrefour dans la combobox listant les carrefours
            RenommerCarfInCombobox uneForm.ComboNomCarf, unNomCarf, unePos
            'Positionnement sur ce carrefour
            uneForm.ComboNomCarf.ListIndex = unePos
            'Changement du label NomCarf
            uneForm.monCarrefourCourant.monCarfGraphic.Caption = unNomCarf
            'Changement du nom du carrefour courant
            uneForm.monCarrefourCourant.monNom = unNomCarf
            'Mise à jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
            RemplirLigneTabInfoCalc uneForm, uneForm.monCarrefourCourant.maPosition
            'Indication d'une modification dans les données du site et pas
            'carrefour car le changement de nom n'influence pas les calculs
            maModifDataSite = True
        End If
    Loop While uneSortie = False
End Sub

Public Sub ZoomTout(uneForm As Form)
    'Zoom de tous les carrefours avec leurs feux et de tous les arrêts TC
    'entre le minimun des Y des feux et le maximun des Y des feux
    Dim unYreel As Long, unePos As Long
    Dim unCarf As Carrefour
    Dim unTC As TC
    Dim uneLongEcranAxeY As Long
    Dim unNumFeu As Control
    Dim unNomArret As Control
    Dim unIndex As Integer
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    
    'Calcul de la longueur écran de l'axe des ordonnées
    uneLongEcranAxeY = uneForm.AxeOrdonnée.Y2 - uneForm.AxeOrdonnée.Y1
    'Changement de la longueur réelle de l'axe des Y
    uneForm.maLongueurAxeY = uneForm.monYMaxFeu - uneForm.monYMinFeu
    'Positionnement de l'origine au bon niveau de zoom
    'si elle est entre l'englobant en Y des feux
    If uneForm.monYMinFeu <= 0 And uneForm.monYMaxFeu >= 0 Then
        uneForm.Origine.Visible = True
        unePos = ConvertirReelEnEcran(uneForm.monYMaxFeu, uneForm.maLongueurAxeY, uneLongEcranAxeY)
        uneForm.Origine.Top = unePos + uneForm.AxeOrdonnée.Y1 - uneForm.Origine.Height
    Else
        uneForm.Origine.Visible = False
    End If
    'Redessin de tous les carrefours et de leurs feux au bon zoom
    For i = 1 To uneForm.mesCarrefours.Count
        'Redessin de tous les feux au zoom
        Set unCarf = uneForm.mesCarrefours(i)
        For j = 1 To unCarf.mesFeux.Count
            unYreel = uneForm.monYMaxFeu - unCarf.mesFeux(j).monOrdonnée
            'Conversion du Yréel en Y écran dans la FrameVisuCarf
            unePos = ConvertirReelEnEcran(unYreel, uneForm.maLongueurAxeY, uneLongEcranAxeY)
            'Positionnement en Y écran des objets graphiques du feu
            Set unNumFeu = unCarf.mesFeuxGraphics(j)
            unIndex = unNumFeu.Index
            unNumFeu.Top = unePos + uneForm.AxeOrdonnée.Y1 - unNumFeu.Height
            uneForm.IconeFeu(unIndex).Top = unePos + uneForm.AxeOrdonnée.Y1 - uneForm.IconeFeu(unIndex).Height
        Next j
        'Déplacement du label NomCarf au bon endroit par rapport au zoom
        ModifYNomCarf uneForm, unCarf
    Next i
    
    'Redessin de tous les arrêts TC au bon zoom
    For i = 1 To uneForm.mesTC.Count
        Set unTC = uneForm.mesTC(i)
        For j = 1 To unTC.mesArrets.Count
            unYreel = uneForm.monYMaxFeu - unTC.mesArrets(j).monOrdonnee
            'Conversion du Yréel en Y écran dans la FrameVisuCarf
            unePos = ConvertirReelEnEcran(unYreel, uneForm.maLongueurAxeY, uneLongEcranAxeY)
            'Positionnement en Y écran des objets graphiques de l'arrêt TC
            Set unNomArret = unTC.mesObjGraphics(j)
            unNomArret.Top = unePos + uneForm.AxeOrdonnée.Y1 - unNomArret.Height
            'Recherche des arrêts confondus en un Y valant unTC.mesArrets(j).monOrdonnee pour
            'alimenter les listes d'arrêts et de TC trouvés
            unNb = uneForm.RechercherArretConfondu(unTC.mesArrets(j).monOrdonnee, uneListeIndexTC, uneListeIndexArret, i - 1)
            'Mise à jour des décalages des labels NomArrêt confondus en ce nouveau Y
            Call MiseAJourNomArret(uneForm, uneListeIndexTC, uneListeIndexArret)
            'On vide les listes pour le j suivant
            ViderCollection uneListeIndexTC
            ViderCollection uneListeIndexArret
            'Ajustement de la chaine de caractéres à l'axe des ordonnées
            unNomArret.Width = uneForm.AxeOrdonnée.X1 - unNomArret.Left
            unInd = unNomArret.Index
            uneForm.IconeArret(unInd).Top = unePos + uneForm.AxeOrdonnée.Y1 - uneForm.IconeArret(unInd).Height
        Next j
    Next i
End Sub

Public Sub RedessinerTout(uneFenetreFille As Form, unY As Long)
    'Si l'englobant des Y change, on fait un zoom maximun englobant tous les feux
    '==> redessin de toutes les entités dans le nouveau repère (translation + zoom)
    uneModifEnglobant = False
    With uneFenetreFille
        If unY > .monYMaxFeu Then
            .monYMaxFeu = unY
            uneModifEnglobant = True
        ElseIf unY < .monYMinFeu Then
            .monYMinFeu = unY
            uneModifEnglobant = True
        End If
    End With
    If uneModifEnglobant Then Call ZoomTout(uneFenetreFille)
End Sub

Public Sub CalculerEnglobant(uneForm As Form)
    'Recheche du maximun et du mininum des ordonnées des Feux
    Dim unCarf As Carrefour
    Dim unNbCarf As Integer, unYCarf As Integer
    
    'Mise à jour des Y max et Y min des feux
    With uneForm
        unNbCarf = .mesCarrefours.Count
        'Réinitialisation de l'englobant
        If unNbCarf = 1 And .mesCarrefours(1).mesFeux.Count = 1 Then
            'Cas où on a un seul carrefour avec un seul feu
            'On prend l'englobant autour du seul feu du seul carrefour
            'à + ou - 100 mètres
            unYCarf = DonnerYCarrefour(.mesCarrefours(1))
            .monYMaxFeu = unYCarf + 100
            .monYMinFeu = unYCarf - 100
        Else
            'Tous les autres cas
            .monYMaxFeu = -9999
            monSite.monYMinFeu = 9999
        End If
            
        'recherche du min et max des Y de tous les feux de tous les carrefours
        For i = 1 To unNbCarf
            Set unCarf = .mesCarrefours(i)
            For j = 1 To unCarf.mesFeux.Count
                unY = unCarf.mesFeux(j).monOrdonnée
                If unY > .monYMaxFeu Then
                    'Modif du maximun
                    .monYMaxFeu = unY
                End If
                If unY < .monYMinFeu Then
                    'Modif du minimun
                    .monYMinFeu = unY
                End If
            Next j
        Next i
    End With
End Sub


Public Sub RemplirLigneTabInfoCalc(uneForm As Form, unIndLig As Integer)
    'Remplit la ligne numéro unIndLig du tableau TabInfoCalc
    'avec les valeurs par défaut
    Dim unCarf As Carrefour
    
    With uneForm
        Set unCarf = .mesCarrefours(unIndLig)
        .TabInfoCalc.Row = unIndLig
        .TabInfoCalc.Col = 1
        .TabInfoCalc.Text = unCarf.monNom
        .TabInfoCalc.Col = 2
        If unCarf.monIsUtil Then
            .TabInfoCalc.Text = "Oui"
        Else
            .TabInfoCalc.Text = "Non"
        End If
        .TabInfoCalc.Col = 3
        .TabInfoCalc.Text = Format(unCarf.maVitSensM)
        .TabInfoCalc.Col = 4
        .TabInfoCalc.Text = Format(unCarf.maVitSensD)
    End With
End Sub

Public Sub RemplirOngletTabDecalage(uneForm As Form)
    'Remplit les tableaux de l'onglet Tableau Décalages
    
    'Remplissage des bandes passantes
    With uneForm
        'Affichage des largeurs de bandes
        .TabBande.Row = 1
        .TabBande.Col = 1
        .TabBande.Text = Format(.maBandeM)
        .TabBande.Row = 2
        .TabBande.Col = 1
        .TabBande.Text = Format(.maBandeD)
        .TabBande.Row = 1
        .TabBande.Col = 2
        .TabBande.Text = Format(.maBandeModifM)
        .TabBande.Row = 2
        .TabBande.Col = 2
        .TabBande.Text = Format(.maBandeModifD)
    End With
    
    'Remplissage des décalages
    With uneForm
        .TabDecal.MaxRows = .mesCarrefours.Count
        For i = 1 To .mesCarrefours.Count
            .TabDecal.Row = i
            .TabDecal.Col = 1
            .TabDecal.Text = .mesCarrefours(i).monNom
            'Affichage dans l'onglet Tableau de résultat en arrondissant
            'à l'entier le plus proche grâce à la fonction VB5 CInt
            'si le carrefour est pris en compte dans le calcul
            'sinon affichage vide pour les décalages
            If .mesCarrefours(i).monDecCalcul <> -99 Then
                .TabDecal.Col = 2
                If CIntCorrigé(.mesCarrefours(i).monDecCalcul) = .maDuréeDeCycle Then
                    'Une valeur valant durée du cycle s'affiche 0
                    .TabDecal.Text = "0"
                Else
                    .TabDecal.Text = CIntCorrigé(.mesCarrefours(i).monDecCalcul)
                End If
                .TabDecal.Col = 3
                .TabDecal.Lock = False
                If CIntCorrigé(.mesCarrefours(i).monDecModif) = .maDuréeDeCycle Then
                    'Une valeur valant durée du cycle s'affiche 0
                    .TabDecal.Text = "0"
                Else
                    .TabDecal.Text = CIntCorrigé(.mesCarrefours(i).monDecModif)
                End If
                .TabDecal.Col = 4
                .TabDecal.Lock = False
                'Mise de la BackColor de la colonne 4 à celle de
                'l'image des checkbox
                '.TabDecal.BackColor = uneForm.BackColor
                If .mesCarrefours(i).monDecImp = 1 Then
                    'Cas d'un carrefour à décalage imposé
                    .TabDecal.Text = "Oui"
                Else
                    'Cas d'un carrefour sans décalage imposé
                    .TabDecal.Text = "Non"
                End If
            Else
                'Si le carrefour n'est pas pris en compte dans le calcul
                'affichage vide pour les décalages
                .TabDecal.Col = 2
                .TabDecal.Text = ""
                
                .TabDecal.Col = 3
                .TabDecal.Lock = True
                .TabDecal.Text = ""
                
                'Ajout de la chaine vide dans la liste (oui,Non)
                .TabDecal.TypeComboBoxIndex = 2
                .TabDecal.TypeComboBoxString = ""
                .TabDecal.Col = 4
                .TabDecal.Lock = True
                .TabDecal.Text = ""
                'Effacement de la chaine vide dans la liste (Oui, Non)
                .TabDecal.TypeComboBoxIndex = 2
                .TabDecal.Action = 27 'SS_ACTION_COMBO_REMOVE
            End If
        Next i
    End With
End Sub

Public Function SaisieEntierPositifEntreMinMax(KeyCode As Integer, unControl As Control, uneValeurDefaut As Integer, unIntMin As Integer, unIntMax As Integer, uneString) As Boolean
    SaisieEntierPositifEntreMinMax = False
    Call VerifSaisieEntierPositif(KeyCode, unControl, uneValeurDefaut)
    If Val(unControl.Text) < unIntMin Or Val(unControl.Text) > unIntMax Then
        unMsg = uneString + " doit être >= à " + Format(unIntMin)
        unMsg = unMsg + " et <= à " + Format(unIntMax)
        MsgBox unMsg, vbCritical
        unControl.Text = uneValeurDefaut
   Else
        SaisieEntierPositifEntreMinMax = True
    End If
End Function


Public Sub VerifSaisieEntier(KeyAscii As Integer, unControl As Control)
    'Vérification de la saisie d'entier grâce à la touche tapée
    'dans le KeyPress event du control unControl.
    'On utilise apres le Keyup event de ce même control
    Dim unEntier As Integer
    Dim uneChaineTmp As String
    
    If KeyAscii > 47 And KeyAscii < 58 Then
        'Cas de saisie d'un chiffre
        '==> on ne fait rien car saisie OK
    ElseIf KeyAscii = 8 Then
        'Cas de saisie d'un retour arrière
        '==> on ne fait rien car saisie OK
    ElseIf KeyAscii = 45 Then
        'Cas de saisie d'un moins
        unePos = InStr(1, unControl.Text, "-")
        If unControl.Text = "0" Then
            'On réaffiche 0
            unControl.Text = "0"
            KeyAscii = 0
        ElseIf unePos > 0 Then
            'Cas d'un moins existant ==> suppression du moins existant en tête
            unControl.Text = Mid$(unControl.Text, 2)
            'On n'affichage pas le moins dernièrement saisi
            KeyAscii = 0
        Else
            'Cas d'abscence de moins ==> rajout d'un moins en tête
            unControl.Text = "-" + unControl.Text
            'On n'affichage pas le moins dernièrement saisi
            KeyAscii = 0
        End If
    Else
        'Cas des autres touches ==> on n'affiche pas le caractère erroné
        KeyAscii = 0
        Beep
    End If
End Sub

Public Sub RemplirFrameTC(uneForm As Form, unInd As Long)
    With uneForm
        'Affectation avec les nouvelles valeurs
        .TextTDep.Text = Format(.mesTC(unInd).monTDep)
        .TextDistAF_TC.Text = Format(.mesTC(unInd).maDistAccFrein)
        .TextDuréeAF_TC.Text = Format(.mesTC(unInd).maDureeAccFrein)
        .ColorTC.BackColor = .mesTC(unInd).maCouleur
       'Mise à vide des combobox pour éviter le test de différence
        'des carrefours de départ et d'arrivée
        .ComboCarfDep.ListIndex = -1
        .ComboCarfArr.ListIndex = -1
        'Mise à jour des carrefours de départ et d'arrivée
        .ComboCarfDep.ListIndex = .mesTC(unInd).monCarfDep.maPosition - 1
        .ComboCarfArr.ListIndex = .mesTC(unInd).monCarfArr.maPosition - 1
        'Remplissage des arrêts
        unNbArret = .mesTC(unInd).mesArrets.Count
        .TabYArret.MaxCols = unNbArret
        For i = 1 To unNbArret
            .TabYArret.Col = i
            .TabYArret.Row = 1
            .TabYArret.Text = Format(.mesTC(unInd).mesArrets(i).monOrdonnee)
            .TabYArret.Row = 2
            .TabYArret.Text = Format(.mesTC(unInd).mesArrets(i).maVitesseMarche)
            .TabYArret.Row = 3
            .TabYArret.Text = Format(.mesTC(unInd).mesArrets(i).monTempsArret)
        Next i
    End With
End Sub

Public Function DonnerYMaxCarf(unCarf As Carrefour) As Integer
    'Recherche du plus grand Y parmi les Y des feux d'un carrefour
    Dim unFeu As Feu
    
    DonnerYMaxCarf = -30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonnée > DonnerYMaxCarf Then
            DonnerYMaxCarf = unFeu.monOrdonnée
        End If
    Next i
End Function

Public Function DonnerYMinCarf(unCarf As Carrefour) As Integer
    'Recherche du plus petit Y parmi les Y des feux d'un carrefour
    Dim unFeu As Feu
    
    DonnerYMinCarf = 30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonnée < DonnerYMinCarf Then
            DonnerYMinCarf = unFeu.monOrdonnée
        End If
    Next i
End Function

Public Function DonnerYMinCarfSens(unCarf As Carrefour, unSensMontant As Boolean, unIndFeu As Integer) As Integer
    'Recherche du plus petit Y parmi les Y
    'des feux de même sens d'un carrefour
    'unIndFeu renvoie le feu du carrefour réalisant ce minimun
    Dim unFeu As Feu
    
    DonnerYMinCarfSens = 30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonnée < DonnerYMinCarfSens And unFeu.monSensMontant = unSensMontant Then
            DonnerYMinCarfSens = unFeu.monOrdonnée
            unIndFeu = i
        End If
    Next i
    
    'Cas où aucun feu dans le sens cherché, on prend le feu
    'd'Y min de l'autre sens
    If DonnerYMinCarfSens = 30000 Then
        DonnerYMinCarfSens = DonnerYMinCarfSens(unCarf, Not unSensMontant, unIndFeu)
    End If
End Function

Public Function DonnerYMaxCarfSens(unCarf As Carrefour, unSensMontant As Boolean, unIndFeu As Integer) As Integer
    'Recherche du plus grand Y parmi les Y
    'des feux de même sens d'un carrefour
    'unIndFeu renvoie le feu du carrefour réalisant ce maximun
    Dim unFeu As Feu
    
    DonnerYMaxCarfSens = -30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonnée > DonnerYMaxCarfSens And unFeu.monSensMontant = unSensMontant Then
            DonnerYMaxCarfSens = unFeu.monOrdonnée
            unIndFeu = i
        End If
    Next i

    'Cas où aucun feu dans le sens cherché, on prend le feu
    'd'Y max de l'autre sens
    If DonnerYMaxCarfSens = -30000 Then
        DonnerYMaxCarfSens = DonnerYMaxCarfSens(unCarf, Not unSensMontant, unIndFeu)
    End If
End Function

Public Function VerifierExistenceArret(unY As Long, unTabArret As vaSpread, uneListeArret As ColArretTC) As Boolean
    'Test de l'existence d'un arrêt pour le TC courant en  unY
    VerifierExistenceArret = True
    unNb = uneListeArret.Count
    i = 1
    Do While unNb > 1 And i <= unNb
        'On boucle sur toutes les ordonnées des arrêts du TC
        If unY = uneListeArret(i).monOrdonnee And i <> unTabArret.Col Then
            'Cas où les Y, qui sont des entiers, sont égaux avec un arrêt
            'différent de celui en cours de modification
            unMsg = "Ce transport collectif a déjà un arrêt d'ordonnée " + Format(unY) + Chr(13)
            unMsg = unMsg + Chr(13) + "Saisissez une nouvelle valeur entre -9999 et 9999 mètres :"
            uneNewVal = InputBox(unMsg, "Message d'erreur de OndeV", unY)
            If Trim(uneNewVal) = "" Then
                'Cas d'un click sur annuler ou d'une saisie vide
                'Sortie sans rien modifier en remettant la valeur précédente
                unTabArret.Text = Format(uneListeArret(unTabArret.Col).monOrdonnee)
                VerifierExistenceArret = False
                Exit Function
            Else
                'Pour reboucler sur tous les Y
                'et vérifier l'unicité du Y saisi
                unY = Val(uneNewVal)
                i = 0
                'Test du domaine de validité
                Do While unY < -9999 Or unY > 9999
                    unMsg = "L'ordonnée doit être comprise entre -9999 et 9999 mètres"
                    unMsg = unMsg + Chr(13) + Chr(13) + "Saisissez une nouvelle valeur entre -9999 et 9999 mètres :"
                    uneNewVal = InputBox(unMsg, "Message d'erreur de OndeV", unY)
                    If Trim(uneNewVal) = "" Then
                        'Cas d'un click sur annuler ou d'une saisie vide
                        'Sortie sans rien modifier en remettant la valeur précédente
                        unY = Format(uneListeArret(unTabArret.Col).monOrdonnee)
                        VerifierExistenceArret = False
                    Else
                        unY = Val(uneNewVal)
                    End If
                Loop
                'Mise à jour de la colonne avec un valeur valide
                unTabArret.Text = Format(unY)
            End If
        End If
        i = i + 1
    Loop
End Function

Public Function TrouverTCParNom(unSite As Form, unNomTC As String) As Integer
    'Recherche d'un TC par son nom et
    'retour de sa position dans la liste des TC du site
    i = 0
    If unNomTC <> "Aucun" Then
        Do
            i = i + 1
        Loop Until unSite.mesTC(i).monNom = unNomTC
    End If
    TrouverTCParNom = i
End Function

Public Sub RemplirComboboxOndeTC(unSite As Form, unTC As TC)
    'Mise à jour des combobox des TC pour l'onde verte TC
    If DonnerYCarrefour(unTC.monCarfDep) < DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC montant
        unSite.ComboTCM.AddItem unTC.monNom
    Else
        'Cas d'un TC descendant
        unSite.ComboTCD.AddItem unTC.monNom
    End If
End Sub


Public Function ChangerParamOndeTC(unSite As Form, unIndTC, unNewCarfDep As Carrefour, unNewCarfArr As Carrefour) As Boolean
    'Mise à jour des controls de la frame FrameOndeTC de
    'l'onglet Cadrage Onde verte, lors d'un changement de
    'carrefours départ et/ou arrivée, ce qui peut changer le sens du TC
    'Retourne true si le TC d'indice unIndTC n'est pas utilisé dans les ondes TC
    'ou si les TC cadrant les ondes TC ne changent pas de sens.
    'Retourne faux dans les autres cas
    
    unDY = DonnerYCarrefour(unSite.mesTC(unIndTC).monCarfArr) - DonnerYCarrefour(unSite.mesTC(unIndTC).monCarfDep)
    unDYnew = DonnerYCarrefour(unNewCarfArr) - DonnerYCarrefour(unNewCarfDep)
    If unDY * unDYnew < 0 Then
        'Cas d'un changement de sens du TC
        If unSite.monTCM = unIndTC Or unSite.monTCD = unIndTC Then
            'Cas d'un TC servant à cadrer les ondes TC
            unMsg = "Impossible de changer les carrefours départ ou arrivée du TC " + unSite.mesTC(unIndTC).monNom
            unMsg = unMsg + " car son sens de parcours est changé or il est utilisé dans le calcul d'onde verte prenant en compte des TC"
            MsgBox unMsg, vbCritical
            ChangerParamOndeTC = False
        Else
            'Cas d'un TC non utilisée dans les ondes vertes TC
            If unDY > 0 Then
                'Cas du TC montant devenant descendant
                'Suppression dans la liste des TC montant
                i = -1
                Do
                    i = i + 1
                Loop Until unSite.mesTC(unIndTC).monNom = unSite.ComboTCM.List(i)
                unSite.ComboTCM.RemoveItem i
                'Ajout dans la liste des TC descendant
                unSite.ComboTCD.AddItem unSite.mesTC(unIndTC).monNom
            Else
                'Cas du TC descendant devenant montant
                'Suppression dans la liste des TC descendant
                i = -1
                Do
                    i = i + 1
                Loop Until unSite.mesTC(unIndTC).monNom = unSite.ComboTCD.List(i)
                unSite.ComboTCD.RemoveItem i
                'Ajout dans la liste des TC montant
                unSite.ComboTCM.AddItem unSite.mesTC(unIndTC).monNom
            End If
            ChangerParamOndeTC = True
        End If
    Else
        'Cas où le sens du TC reste le même
        ChangerParamOndeTC = True
    End If
End Function

Public Sub SauverOptionsAffImp(unSaveRecentsOnly As Boolean)
    'Sauvegarde des options d'affichage et d'impression dans la base de
    'registre à la place du fichier OndeV.ini (fait à partir de la version 1.00.0002)
    Dim unSite1 As String, unSite2 As String
    Dim unSite3 As String, unSite4 As String
    Dim unSite As frmDocument, unFileName As String
    
    If unSaveRecentsOnly Then
        'Cas où l'on ne sauvegarde que les fichiers récents
        'Appel par le unload de la MDI
        'Récup des options d'affichage et d'impression
        Set unSite = New frmDocument
        Set unSite.mesOptionsAffImp = New OptionsAffImp
        ChargerOptionsAffImp unSite
    Else
        'Cas où l'on sauve tout
        'Appel par le click dans Conserver par défaut des options
        Set unSite = monSite
    End If
    
    With unSite.mesOptionsAffImp
        SaveSetting App.Title, "OptionsAffImp", "monEpaisseurLigne", .monEpaisseurLigne
        SaveSetting App.Title, "OptionsAffImp", "monNbSecondesRappel", .monNbSecondesRappel
    
        SaveSetting App.Title, "OptionsAffImp", "maCoulBandComD", .maCoulBandComD
        SaveSetting App.Title, "OptionsAffImp", "maCoulBandComM", .maCoulBandComM
        SaveSetting App.Title, "OptionsAffImp", "maCoulBandInterCarfD", .maCoulBandInterCarfD
        SaveSetting App.Title, "OptionsAffImp", "maCoulBandInterCarfM", .maCoulBandInterCarfM
        SaveSetting App.Title, "OptionsAffImp", "maCoulLigne", .maCoulLigne
        SaveSetting App.Title, "OptionsAffImp", "maCoulNomArret", .maCoulNomArret
        SaveSetting App.Title, "OptionsAffImp", "maCoulNomCarf", .maCoulNomCarf
        SaveSetting App.Title, "OptionsAffImp", "maCoulPtRef", .maCoulPtRef
        SaveSetting App.Title, "OptionsAffImp", "maCoulTitreEch", .maCoulTitreEch
    
        SaveSetting App.Title, "OptionsAffImp", "maVisuBandComD", CInt(.maVisuBandComD)
        SaveSetting App.Title, "OptionsAffImp", "maVisuBandComM", CInt(.maVisuBandComM)
        SaveSetting App.Title, "OptionsAffImp", "maVisuBandInterCarfD", CInt(.maVisuBandInterCarfD)
        SaveSetting App.Title, "OptionsAffImp", "maVisuBandInterCarfM", CInt(.maVisuBandInterCarfM)
        SaveSetting App.Title, "OptionsAffImp", "maVisuLigne", CInt(.maVisuLigne)
    
        'Remplissage des 4 derniers ouverts éventuels
        If frmMain.mnuFileSite1.Visible Then
            unFileName = Mid(frmMain.mnuFileSite1.Caption, 4)
        Else
            unFileName = ""
        End If
        SaveSetting App.Title, "Recent Files", "File1", unFileName
        
        If frmMain.mnuFileSite2.Visible Then
            unFileName = Mid(frmMain.mnuFileSite2.Caption, 4)
        Else
            unFileName = ""
        End If
        SaveSetting App.Title, "Recent Files", "File2", unFileName
        
        If frmMain.mnuFileSite3.Visible Then
            unFileName = Mid(frmMain.mnuFileSite3.Caption, 4)
        Else
            unFileName = ""
        End If
        SaveSetting App.Title, "Recent Files", "File3", unFileName
        
        If frmMain.mnuFileSite4.Visible Then
            unFileName = Mid(frmMain.mnuFileSite4.Caption, 4)
        Else
            unFileName = ""
        End If
        SaveSetting App.Title, "Recent Files", "File4", unFileName
    End With
End Sub

Public Sub ChargerOptionsAffImp(unSite As Form)
    'Alimentation de l'instance d'options d'affichage et d'impression
    'à partir des infos de la base de registre à la place du fichier
    'OndeV.ini (fait à partir de la version 1.00.0002)
    Dim uneString As String, unLong As Long
    Dim unBool As Boolean
    
    unSite.mesOptionsAffImp.monEpaisseurLigne = GetSetting(App.Title, "OptionsAffImp", "monEpaisseurLigne", 1)
    unSite.mesOptionsAffImp.monNbSecondesRappel = GetSetting(App.Title, "OptionsAffImp", "monNbSecondesRappel", 10)
    
    unSite.mesOptionsAffImp.maCoulBandComD = GetSetting(App.Title, "OptionsAffImp", "maCoulBandComD", 255)
    unSite.mesOptionsAffImp.maCoulBandComM = GetSetting(App.Title, "OptionsAffImp", "maCoulBandComM", 16711680)
    unSite.mesOptionsAffImp.maCoulBandInterCarfD = GetSetting(App.Title, "OptionsAffImp", "maCoulBandInterCarfD", 16711935)
    unSite.mesOptionsAffImp.maCoulBandInterCarfM = GetSetting(App.Title, "OptionsAffImp", "maCoulBandInterCarfM", 16776960)
    unSite.mesOptionsAffImp.maCoulLigne = GetSetting(App.Title, "OptionsAffImp", "maCoulLigne", 0)
    unSite.mesOptionsAffImp.maCoulNomArret = GetSetting(App.Title, "OptionsAffImp", "maCoulNomArret", 16711935)
    unSite.mesOptionsAffImp.maCoulNomCarf = GetSetting(App.Title, "OptionsAffImp", "maCoulNomCarf", 0)
    unSite.mesOptionsAffImp.maCoulPtRef = GetSetting(App.Title, "OptionsAffImp", "maCoulPtRef", 49152)
    unSite.mesOptionsAffImp.maCoulTitreEch = GetSetting(App.Title, "OptionsAffImp", "maCoulTitreEch", 49152)
    
    unSite.mesOptionsAffImp.maVisuBandComD = GetSetting(App.Title, "OptionsAffImp", "maVisuBandComD", -1) 'True par défaut
    unSite.mesOptionsAffImp.maVisuBandComM = GetSetting(App.Title, "OptionsAffImp", "maVisuBandComM", -1) 'True par défaut
    unSite.mesOptionsAffImp.maVisuBandInterCarfD = GetSetting(App.Title, "OptionsAffImp", "maVisuBandInterCarfD", 0) 'False par défaut
    unSite.mesOptionsAffImp.maVisuBandInterCarfM = GetSetting(App.Title, "OptionsAffImp", "maVisuBandInterCarfM", 0) 'False par défaut
    unSite.mesOptionsAffImp.maVisuLigne = GetSetting(App.Title, "OptionsAffImp", "maVisuLigne", -1) 'True par défaut
End Sub

Public Sub ChargerOptionsAffImpParDefaut(unSite As Form)
    'Alimentation de l'instance d'options d'affichage et d'impression
    'avec les valeurs par défaut
    unSite.mesOptionsAffImp.monEpaisseurLigne = 1
    unSite.mesOptionsAffImp.monNbSecondesRappel = 10
    
    unSite.mesOptionsAffImp.maCoulBandComD = 255
    unSite.mesOptionsAffImp.maCoulBandComM = 16711680
    unSite.mesOptionsAffImp.maCoulBandInterCarfD = 16711935
    unSite.mesOptionsAffImp.maCoulBandInterCarfM = 16776960
    unSite.mesOptionsAffImp.maCoulLigne = 0
    unSite.mesOptionsAffImp.maCoulNomArret = 16711935
    unSite.mesOptionsAffImp.maCoulNomCarf = 0
    unSite.mesOptionsAffImp.maCoulPtRef = 49152
    unSite.mesOptionsAffImp.maCoulTitreEch = 49152
    
    unSite.mesOptionsAffImp.maVisuBandComD = True
    unSite.mesOptionsAffImp.maVisuBandComM = True
    unSite.mesOptionsAffImp.maVisuBandInterCarfD = False
    unSite.mesOptionsAffImp.maVisuBandInterCarfM = False
    unSite.mesOptionsAffImp.maVisuLigne = True
End Sub


Public Sub PlacerFeuAxeY(uneFenetreFille As Form, unIndCarf As Integer, unIndFeu As Integer, unIndObjGraphicFeu)
    'Positionnement du feu (Numéro et icône Feu) à droite de l'axe des Y
    'pour un feu montant et à gauche pour un feu descendant
    Dim unSensMontant As Boolean
    
    With uneFenetreFille
        i = unIndObjGraphicFeu
        'Récupération du sens du feu
        unSensMontant = .mesCarrefours(unIndCarf).mesFeux(unIndFeu).monSensMontant
    
        'Utilisation du label LabelTrait pour calculer le
        'décalage à droite ou à gauche par rapport à l'axe des Y
        .LabelTrait.Caption = "___"
        
        'Stockage de la valeur Gras de la font du label NumFeu
        unSaveBold = .NumFeu(i).Font.Bold
        'Mise en non gras de la font du label NumFeu car les calculs de Width
        'des label NumFeu sont calibrés avec une fonte non grasse
        '(la propriété AutoSize d'un label tient compte du type de fonte)
        .NumFeu(i).Font.Bold = False
        
        If unSensMontant Then
            'Cas des feux montant
            '==> Positionnement à droite de l'axe des Y avec 2 blancs à la fin
            'pour que la mise en gras ne chevauche pas l'icone Feu
            .NumFeu(i).Caption = .LabelTrait.Caption + Format(unIndFeu) + "  "
            .NumFeu(i).Left = .AxeOrdonnée.X1
            .IconeFeu(i).Left = .NumFeu(i).Left + .NumFeu(i).Width
        Else
            'Cas des feux descendant
            '==> Positionnement à gauche de l'axe des Y un numéro + un souligné
            .NumFeu(i).Caption = Format(unIndFeu) + .LabelTrait.Caption
            'Ajustement de la chaine de caractéres à l'axe des ordonnées car la
            'propriété AutoSize est à true ==> restriction du souligné précédent
            .NumFeu(i).Left = .AxeOrdonnée.X1 - .NumFeu(i).Width
            .IconeFeu(i).Left = .NumFeu(i).Left - .IconeFeu(i).Width
        End If
        
        'Restauration de la valeur Gras initiale de la font du label NumFeu
        .NumFeu(i).Font.Bold = unSaveBold
    End With
End Sub

Public Sub RenommerCarfInCombobox(uneComboBox As ComboBox, unNomCarf As String, unePos)
    'Renommage d'un nom de carrefour situé dans une
    'liste de noms d'une combobox
    
    'Suppression dans la combobox listant
    'de l'item correspondant à l'ancien nom
    uneComboBox.RemoveItem unePos
    'Création en ajoutant le nouveau nom dans la liste à la même
    'position que l'ancien nom
    uneComboBox.AddItem unNomCarf, unePos
End Sub

Public Sub ConfigurerSpreadToPrint(unSpread As vaSpread, unHeader As String, unNomFiche As String, unTitreFiche As String)
    'Affectation des options d'impression du spread donné par
    'la variable unSpread passé en paramètre
    
    unSpread.PrintAbortMsg = "Impression sur " + Printer.DeviceName + " en cours..."
    unSpread.PrintJobName = "Impression de la fiche " + unNomFiche + " de OndeV"
    
    If unHeader <> "/fb1" Then unHeader = unHeader + "/n"
    uneFontSize = "/fz""12"""
    unHeader = unHeader + uneFontSize + "/n/cFiche " + unTitreFiche
    unSpread.PrintHeader = unHeader + "/n"
    If maDemoVersion Then
        unSpread.PrintFooter = "/fb1OndeV version 1.0 DEMO"
    Else
        unSpread.PrintFooter = "/fb1OndeV version 1.0"
    End If
    
    unSpread.PrintColor = True
    unSpread.PrintBorder = True
    unSpread.PrintRowHeaders = False
    unSpread.PrintColHeaders = False
    unSpread.PrintGrid = True
    
    'Marge droite et gauche
    unSpread.PrintMarginRight = unTwipToCm
    unSpread.PrintMarginLeft = unSpread.PrintMarginRight
    'Marge en haut et en bas
    unSpread.PrintMarginBottom = unTwipToCm ' = 1 cm
    unSpread.PrintMarginTop = unSpread.PrintMarginBottom
    
    unSpread.PrintType = 0 ' = SS_PRINT_ALL
    unSpread.PrintUseDataMax = True
    unSpread.PrintShadows = True
    
    'Mise en noir des lignes de séparation du spread
    unSpread.GridColor = RGB(0, 0, 0)
End Sub

Public Function TrouverCarfParNom(unNom As String) As Integer
    'Fonction retournant l'indice du carrefour de nom unNom parmi
    'les carrefours du site sinon elle retourne 0 si non trouvé
    TrouverCarfParNom = 0
    For i = 1 To monSite.mesCarrefours.Count
        If monSite.mesCarrefours(i).monNom = unNom Then
            TrouverCarfParNom = i
            Exit For
        End If
    Next i
End Function

Public Sub ChangerHelpID(unNumOnglet As Integer)
    'Changer des contextes Id de l'aide
    'en fonction de l'onglet courant du site actif
    'de la MDI, de la fille active et de ses textbox
    'TitreEtude et DuréeCycle
    Select Case unNumOnglet
        Case 0 ' Onglet Carrefour
            unHelpID = IDhlp_OngletCarf
        Case 1 ' Onglet TC
            unHelpID = IDhlp_OngletTC
        Case 2 ' Onglet Cadrage
            unHelpID = IDhlp_OngletCadrage
        Case 3 ' Onglet Résultat décalages
            unHelpID = IDhlp_OngletResDec
        Case 4 ' Onglet Dessin onde verte
            unHelpID = IDhlp_OngletDesOnde
        Case 5 ' Onglet Fiche résultats
            unHelpID = IDhlp_OngletFicRes
    End Select
    
    'Affectation du nouveau contexte d'aide
    frmMain.HelpContextID = unHelpID
    monSite.HelpContextID = unHelpID
    monSite.TitreEtude.HelpContextID = unHelpID
    monSite.FrameVisuCarf.HelpContextID = unHelpID
    monSite.DuréeCycle.HelpContextID = unHelpID
    monSite.TabFeux.HelpContextID = unHelpID
    'Remplir tous les éléments d'un onglet TabFeux actif
    'For i = 0 To monSite.TabFeux(unNumOnglet).Controls.Count - 1
    '    monSite.TabFeux.Controls(unNumOnglet).HelpContextID = unHelpID
    'Next i
End Sub

Public Function CIntCorrigé(unSingle As Single) As Integer
    'Fonction corrigeant la fonction VB CInt qui est buggée
    'pour les flottants (=Single) positifs
    
    'En effet, CInt(20.5) = 20 alors que CInt(21.5)=22
    'CIntCorrigé doit rendre l'entier supérieur toujours
    'CintCorrigé(xx.yyy) = xx si 0.yyy < à 0.5
    'et (xx + 1) si 0.yyy >= 0.5
    
    'A utiliser pour les modifs et affichage de décalages
    'Ces décalages toujours entre 0 et cycle ==> >=0
    Dim unReste As Single
    
    If unSingle < 0 Then
        unReste = unSingle - Int(unSingle)
        'car Int(-xx.yy)= -xx-1 ==> unReste = -xx.yy + xx + 1
        '==> unReste =  1- 0.yy > 0
    Else
        unReste = unSingle - Fix(unSingle)
        'car Fix(xx.yy)= xx ==> unReste = xx.yy - xx
        '==> unReste =  0.yy > 0
    End If
    
    'Calcul de l'arrondi grâce à Int
    'car Int(xx.yy) = xx et Int(-xx.yy) = -xx - 1
    If unReste < 0.5 Then
        CIntCorrigé = Int(unSingle)
    Else
        CIntCorrigé = Int(unSingle) + 1
    End If
End Function

Public Sub TestCIntCorrigé(unIntMin As Integer, unIntMax As Integer, unPas As Single)
    'Fonction de test de CIntCorrigé entre unIntMin et unIntMax
    'avec unPas
    Dim unSingle As Single

    Do
    unSingle = CSng(InputBox("Entrez un réel :", "TestCIntCorrigé"))
    uneRep = MsgBox(Format(unSingle) + " : CInt = " + Format(CInt(unSingle)) + " et CIntNew = " + Format(CIntCorrigé(unSingle)), vbRetryCancel)
    Loop Until uneRep = vbCancel
    
    unSingle = unIntMin
    unSingle = unIntMax + 1
    Do While unSingle <= unIntMax
        Debug.Print unSingle; " : CInt = "; CInt(unSingle); " et CIntNew = "; CIntCorrigé(unSingle)
        unSingle = unSingle + unPas
    Loop
End Sub

Public Function GetAppPath() As String
    'Fonction retournant le répertoire de l'application
    'avec un \ à la fin, toujours.
    'Car si RepInstall = c:\Test, App.Path rend "c:\Test"
    'mais si repInstall = c:\, alors App.Path rend "c:\"
    'donc on rajoute un \ pour faire homogène
    If Mid(App.Path, Len(App.Path)) = "\" Then
        GetAppPath = App.Path
    Else
        GetAppPath = App.Path + "\"
    End If
End Function

