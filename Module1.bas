Attribute VB_Name = "Module1"
'Variable indiquant si on travaille sur une version prot�g�e ou pas
Public maProtectVersion As Boolean

'Variable indiquant si on travaille sur une version d�mo non prot�g�e ou pas
Public maDemoVersion As Boolean

'Constante pour la d�composition des composantes RGB d'une couleur
Public Const CarreDe256 As Long = 65536 '256 * 256

'Coef de passage entre les cm et les twips
Public Const unTwipToCm = 567      '567 twips = 1 cm

'Constantes pour le type de phase des tableaux de marche TC
Public Const VConst As Integer = 0   'Phase � vitesse constante
Public Const Accel As Integer = 1    'Phase d'acc�l�ration
Public Const Decel As Integer = 2    'Phase d'd�c�l�ration
Public Const Arret As Integer = 3    'Phase d'arr�t

'Constantes pour le type d'onde verte
Public Const OndeDouble As Integer = 0 'Type double sens
Public Const OndeSensM As Integer = 1  'Type sens montant privil�gi�
Public Const OndeSensD As Integer = 2  'Type sens descendant privil�gi�
Public Const OndeTC As Integer = 3     'Type cadrage pour un TC montant
                                       'et/ou descendant

'Constantes pour le type de vitesse � chaque carrefour
Public Const VitConst As Integer = 0 'Type vitesse constant pour tous les carrefours
Public Const VitVar As Integer = 1   'Type vitesse variable � chaque carrefour

'Constantes pour la modification des objets graphiques d'un TC
Public Const ModifNomTC As Integer = 0 'Modification du nom dans les labels
Public Const ModifColTC As Integer = 1 'Modification de la couleur
Public Const SupprTC As Integer = 2    'Suppression du TC

'Constantes pour la s�lection des objets graphiques d'une fenetre site
Public Const CarfSel As Integer = 0  'S�lection graphique d'un carrefour
Public Const FeuSel As Integer = 1   'S�lection graphique d'un feu
Public Const ArretSel As Integer = 2 'S�lection graphique d'un arr�t

'Constantes pour la saisie des feux de d�part et d'arriv�e des TC
Public Const FeuDep As Integer = 0  'Saisie du feu de d�part
Public Const FeuArr As Integer = 1  'Saisie du feu d'arriv�e

'Constantes pour faire le trait jusqu'� l'axe des ordonn�es
Public Const StringTrait As String = "___________________________________________________________________________"
    
'Constante fixant le nombre de carat�res maximuns pour le nom des TC
'5 pour en mettre au moins 4 cote � cote dans le FrameVisuCarf
Public Const NbCarMaxNomTC As Integer = 5

'Public formMain As frmMain
Public monCallOptionByPrint As Boolean 'Variable indiquant si la fen�tre frmOptions a �t� ouverte par la fen�tre frmImprimer
Public monPleinEcranVisible As Boolean 'Variable indiquant si la fen�tre plein �cran est ouverte
Public monFermerParMereMDI As Boolean 'Variable indiquant si on a ferm� la MDI m�re
Public monFichierDemarrage As String 'Fichier de d�marrage �ventuel sur la ligne de commande
Public monSite As frmDocument 'fenetre site dans lequel on travail
Public IsEtatCreation As Boolean 'True valeur pour nouveau site, False pour ouvrir un site existant
Public monDocumentCount As Long 'Nombre de fenetres filles sans Nom ouvertes
Public monNbFenFilles As Integer 'Nombre de fenetres filles ouvertes

'Variables stockant la position de la souris lors d'un mouve down event
'dans un control de type NumFeu ou IconeFeu
Public monXSouris As Single
Public monYSouris As Single

'Constante donnant la valeur r�elle en m�tres de la
'longueur de l'axe des Y � l'ouverture d'une nouvelle fen�tre de site
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
    'prot�g�e cc3.x (TRUE) ou pas (FALSE)
    'Utile lors du d�veloppement sous VB
    
    'maProtectVersion = True
    
    'd�sactivation de la protection de copycontrol
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
    ' V�rification de l'enregistrement
    If ProtectCheck("its00+-k") = "its00+-k" Then
      ' Affichage de la feuille principale
         frmMain.Show
    Else 'la licence n'a pas �t� valid�e on ferme
       End
    End If
'********************************
    
    'Cas o� protection valide
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
        'Cas o� la zone de saisie est vide, on remet la valeur par d�faut
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
    'Retourne la position dans la liste si trouv� (entre 0 et listcount-1)
    'ou -1 si non trouv�
    'La casse Min/majuscule est ignor�e
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
    'R�gle de trois pour convertir les m�tres, stock�s en long, en twips
    ConvertirReelEnEcran = unDyEcran / unDyReel * unYreel
End Function

Public Function ConvertirSingleEnEcran(unYreel As Single, unDyReel As Long, unDyEcran As Long) As Single
    'R�gle de trois pour convertir les m�tres, stock�s en single, en twips
    ConvertirSingleEnEcran = unDyEcran / unDyReel * unYreel
End Function

Public Sub MiseAJourNomArret(uneFenetreSite As Form, uneListeIndexTC As Collection, uneListeIndexArret As Collection)
    'Mise � jour de tous les noms des arr�ts qui �taient confondus
    'en le re-d�calant correctement apr�s la suppression d'arr�ts
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
                'Cas d'un TC diff�rent ==> Augmentation du d�calage du nom
                uneStringDecal = uneStringDecal + DonnerStringDecalage
            End If
            'R�cup�ration du label NomArret rang� dans la collection mesObjgraphics
            Set unObjGraphic = unTC.mesObjGraphics(uneListeIndexArret(i))
            'Modification du label
            unObjGraphic.Caption = uneStringDecal + unTC.monNom + StringTrait
            'Coupure de la chaine de caract�res � l'axe des ordonn�es
            unObjGraphic.Width = uneFenetreSite.AxeOrdonn�e.X1 - unObjGraphic.Left
            'Stockage pour le i suivant
            unIndexTC0 = uneListeIndexTC(i)
        Next i
    End If
End Sub

Public Sub ViderCollection(uneCol As Collection)
    'Proc�dure vidant une collection
    
    'Algo : Puisque les collections sont r�index�es
    '       automatiquement, en supprimant le premier
    '       membre � chaque it�ration, on supprime tout.
    For i = 1 To uneCol.Count
        uneCol.Remove 1
    Next i
End Sub



Public Sub MiseAJourSelection(uneFenetreFille As Form, unObjSel As Integer, unIndSel As Integer, Optional unControl As Control, Optional unX As Single)
    'S�lection graphique de l'objet graphique Index repr�sentant un carrefour,
    'un feu ou un arret TC et d�selection de l'ancien objet s�lectionn�
    '==> mise en gras de l'ancienne et de la nouvelle s�lection
    
    If unObjSel = CarfSel Then
        'D�selection de la s�lection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        'Cas d'un carrefour � s�lectionner
        MettreEnGras uneFenetreFille.NomCarf(unIndSel)
        'Affichage de l'onglet Carrefours
        uneFenetreFille.TabFeux.Tab = 0
    ElseIf unObjSel = FeuSel Then
        'D�selection de la s�lection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        'Cas d'un feu � s�lectionner
        MettreEnGras uneFenetreFille.NumFeu(unIndSel)
        'Affichage de l'onglet Carrefours
        uneFenetreFille.TabFeux.Tab = 0
    ElseIf unObjSel = ArretSel Then
        'Cas d'un arr�t TC � s�lectionner
        'Affichage de l'onglet TC d�clench� apr�s une s�lection graphique
        uneFenetreFille.TabFeux.Tab = 1
        'D�selection de la s�lection courante qui va donc devenir l'ancienne
        Call Deselectionner(uneFenetreFille)
        If TypeOf unControl Is Label Then
            'Cas d'un arr�t s�lectionn� par son label nomArret,
            'si s�lection par l'icone STOP on ne passe pas ici
            'Recherche du label NomArret plac� en unX parmi les arr�ts confondus
            RechercherArretEnX uneFenetreFille, unControl, unX
            unIndSel = uneFenetreFille.monIndSel
        End If
        'Mise en gras de la s�lection
        MettreEnGras uneFenetreFille.NomArret(unIndSel)
        'Rafraichissement de la frame contenant les donn�ees d'un TC
        'Pour �viter l'apparition d'un tableau TabYarret � moiti� (Bug Spread)
        uneFenetreFille.FrameTC.Refresh
        'R�cup�ration des positions du TC et du Y de l'arr�t s�lectionn� dans leurs
        ' collections respectives � partir du tag (codage pos TC-pos YArret) du NomArret cliqu�
        unePos = InStr(1, uneFenetreFille.NomArret(unIndSel).Tag, "-")
        unePosTC = Val(Mid$(uneFenetreFille.NomArret(unIndSel).Tag, 1, unePos - 1))
        unePosY = Val(Mid$(uneFenetreFille.NomArret(unIndSel).Tag, unePos + 1))
        'Modification du nombre de colonnes de TabYArret pour la future cellule
        'active en fasse partie, sinon plantage
        uneFenetreFille.TabYArret.MaxCols = uneFenetreFille.mesTC(unePosTC).mesArrets.Count
        'Mise en actif de la cellule contenant le Y de l'arr�t s�lectionn� d�clench� apr�s une s�lection graphique
        uneFenetreFille.TabYArret.Row = 1
        uneFenetreFille.TabYArret.Col = unePosY
        uneFenetreFille.TabYArret.Action = SS_ACTION_ACTIVE_CELL
        'Affichage du TC de l'arr�t s�lectionn� dans comboTC d�clench� apr�s une s�lection graphique
        uneFenetreFille.ComboTC.ListIndex = unePosTC - 1
        'Affichage du libell� de l'arr�t s�lectionn�
        uneFenetreFille.TextArret.Text = uneFenetreFille.mesTC(unePosTC).mesArrets(unePosY).monLibelle
    End If
    'Stockage des valeurs de la nouvelle s�lection
    uneFenetreFille.monObjSel = unObjSel
    uneFenetreFille.monIndSel = unIndSel
    'Mise � jour des contextes d'aide
    ChangerHelpID uneFenetreFille.TabFeux.Tab
End Sub

Public Sub Deselectionner(uneFenetreFille As Form)
    Dim unControl As Control
    
    If uneFenetreFille.monIndSel <> 0 Then
        'Cas d'une s�lection pr�c�dente
        If uneFenetreFille.monObjSel = CarfSel Or uneFenetreFille.monObjSel = FeuSel Then
            'Cas d'un feu et d'un carrefour � d�s�lectionner
            'D�selection du dernier feu s�lectionn�
            Set unControl = uneFenetreFille.NumFeu(uneFenetreFille.monIndSel)
            EnleverGras unControl
            'R�cup�ration du carrefour contenant le dernier feu s�lectionn�
            'par d�codage du tag de l'objet graphique NumFeu � d�selectionner
            unePos = InStr(1, unControl.Tag, "-")
            unePosCarf = Val(Mid$(unControl.Tag, 1, unePos - 1))
            'R�cup�ration de l'objet graphique du carrefour
            Set unControl = uneFenetreFille.mesCarrefours(unePosCarf).monCarfGraphic
            'D�selection du carrefour contenant le dernier feu s�lectionn�
            EnleverGras unControl
        ElseIf uneFenetreFille.monObjSel = ArretSel Then
            'Cas d'un arr�t TC � d�s�lectionner
            EnleverGras uneFenetreFille.NomArret(uneFenetreFille.monIndSel)
        End If
    End If
    uneFenetreFille.monIndSel = 0
End Sub
Public Sub RechercherArretEnX(uneFenetreFille As Form, unControl As Control, unX As Single)
    'Recherche de l'arr�t se trouve sous la souris en unX, unY
    'parmi les arr�ts confondus �ventuels
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unNbArretsConfondus As Integer, unYArret As Long
    Dim unTC As TC
    
    'R�cup�ration des positions du TC et du Y de l'arr�t dans les collections
    '� partir du tag (codage pos TC-pos YArret) du NomArret cliqu�
    unePos = InStr(1, unControl.Tag, "-")
    unePosTC = Val(Mid$(unControl.Tag, 1, unePos - 1))
    unePosY = Val(Mid$(unControl.Tag, unePos + 1))
    'Recherche des arr�ts confondus
    unYArret = uneFenetreFille.mesTC(unePosTC).mesArrets(unePosY).monOrdonnee
    unNbArretsConfondus = uneFenetreFille.RechercherArretConfondu(unYArret, uneListeIndexTC, uneListeIndexArret)
    'Recherche de la position du label NomArret se trouvant sous le click souris en unX, unY
    uneFenetreFille.LabelTrait.Caption = DonnerStringDecalage
    'Calcul du nombre de d�calage ce qui donne le label dont le nom est cliqu�
    unIndex = 1 + unX \ uneFenetreFille.LabelTrait.Width '\ = division enti�re entre 2 entiers
    If unIndex <= unNbArretsConfondus Then
        'Cas d'un click sur un des noms TC et pas en dehors
        'sur la droite dans le soulign�, dans ce cas la s�lection
        'sera celui au premier plan, c'est celui que donne VB par d�faut
        
        'R�cup�ration du TC touv�
        Set unTC = uneFenetreFille.mesTC(uneListeIndexTC(unIndex))
        'R�cup�ration du control NomArret ou IconeArret trouv�
        Set unControl = unTC.mesObjGraphics(uneListeIndexArret(unIndex))
    End If
    'Affectation de l'indice s�lectionn�
    uneFenetreFille.monIndSel = unControl.Index
End Sub


Public Function DonnerStringDecalage()
    'Donner la chaine permettant de d�caler les noms d'arr�ts TC confondus
    'Elle est de la m�me longueur que la longueur maximun des Noms de TC
    DonnerStringDecalage = ""
    For i = 1 To NbCarMaxNomTC + 3
        '+ 3 pour tenir compte de la mise en gras lors
        'de la s�lectionet des fontes proportionnelles
        DonnerStringDecalage = DonnerStringDecalage + "_"
    Next i
End Function

Public Sub MettreEnGras(unControl As Control)
    'Stockage de la largeur avant mise en gras
    uneWidth = unControl.Width
    'Mise en gras
    unControl.Font.Bold = True
    'Ajustement de la largeur � celle initiale
    unControl.Width = uneWidth
End Sub
Public Sub EnleverGras(unControl As Control)
    'Stockage de la largeur avant la suppression de la mise en gras
    uneWidth = unControl.Width
    'Suppression de la mise en gras
    unControl.Font.Bold = False
    'Ajustement de la largeur � celle initiale
    unControl.Width = uneWidth
End Sub

Public Sub MiseAJourSelectionParCellule(uneFenetreFille As Form, unObjSel As Integer, unIndPere As Long, unIndFils As Long)
    'S�lection graphique de l'objet graphique Index repr�sentant un carrefour,
    'un feu ou un arret TC et d�selection de l'ancien objet s�lectionn�
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle s�lection
    Dim unControl As Control
    
    'D�selection de la s�lection courante qui va donc devenir l'ancienne
    Call Deselectionner(uneFenetreFille)
    If unObjSel = CarfSel Or unObjSel = FeuSel Then
        'Cas d'un carrefour � s�lectionner
        '==> S�lection du carrefour donc on le met en gras
        Set unControl = uneFenetreFille.mesCarrefours(unIndPere).monCarfGraphic
        MettreEnGras unControl
        '==> Et s�lection de son feu num�ro unIndFeu qu'on met donc en gras
        Set unControl = uneFenetreFille.mesCarrefours(unIndPere).mesFeuxGraphics(unIndFils)
        MettreEnGras unControl
        'Stockage de l'indice de l'objet de la nouvelle s�lection
        'On stocke l'objet graphique du feu car gr�ce son tag
        'on retourve le carrefour et son feu
        uneFenetreFille.monIndSel = unControl.Index
    ElseIf unObjSel = ArretSel Then
        'Cas d'un arr�t TC � s�lectionner
        'R�cup�ration du control NomArret correspondant � la colonne active
        Set unControl = uneFenetreFille.mesTC(unIndPere).mesObjGraphics(unIndFils)
        'Mise � jour de la s�lection
        MettreEnGras unControl
        'Affichage du libell� de l'arr�t s�lectionn�
        uneFenetreFille.TextArret.Text = uneFenetreFille.mesTC(unIndPere).mesArrets(unIndFils).monLibelle
        'Stockage de l'indice de l'objet de la nouvelle s�lection
        uneFenetreFille.monIndSel = unControl.Index
    End If
    'Stockage du type d'objet de la nouvelle s�lection
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
    
    'R�cup�ration du TC par sa position
    unIndTC = uneFenetreFille.ComboTC.ListIndex + 1
    Set unTC = uneFenetreFille.mesTC(unIndTC)
    'R�cup�ration du nombre d'arr�t avant la cr�ation du nouveau
    unNbArret = uneFenetreFille.TabYArret.MaxCols
    
    'Recherche de l'ordonn�e unYFeuMax qui est la plus grande
    'parmi les feux du carrefour de d�part et d'arriv�e
    unYFeuMax = DonnerYMaxCarf(unTC.monCarfDep)
    unYMax = DonnerYMaxCarf(unTC.monCarfArr)
    If unYMax > unYFeuMax Then unYFeuMax = unYMax
    
    'Recherche de l'ordonn�e unYFeuMin qui est la plus petite
    'parmi les feux du carrefour de d�part et d'arriv�e
    unYFeuMin = DonnerYMinCarf(unTC.monCarfDep)
    unYMin = DonnerYMinCarf(unTC.monCarfArr)
    If unYMin < unYFeuMin Then unYFeuMin = unYMin
    
    'Recherche de l'arr�t ayant l'ordonn�e la plus grande mais inf�rieure
    '� unYFeuMax. Certains Y d'arr�ts peuvent > � unYFeuMax lors de changement
    'de carrefour de d�part et/ou d'arriv�e ou d'inversion du sens du TC
    unYMax = unYFeuMin 'Intialisation du Ymax avec le Ymin des feux
    For i = 1 To unNbArret
        Set unArret = uneFenetreFille.mesTC(unIndTC).mesArrets(i)
        If unArret.monOrdonnee > unYMax And unArret.monOrdonnee <= unYFeuMax Then
            unYMax = unArret.monOrdonnee
        End If
    Next i
            
    'On met le nouvel arr�t au milieu entre le feu ayant le plus grand Y parmi
    'les carrefours d�part et arriv�e et l'arr�t ayant le Y le plus grand du TC
    unYArret = (unYMax + unYFeuMax) / 2
    If CInt(unYArret) = unYMax Then
        unMsg = "Plus de place pour ins�rer un nouvel arr�t TC."
        unMsg = unMsg + Chr(13) + "Changer soit le carrefour d'arriv�e pour un TC montant, soit celui de d�part pour un TC descendant."
        MsgBox unMsg, vbInformation
        Exit Sub
    End If
    
    'Ajout d'un nouvel Y d'arr�t TC
    uneFenetreFille.TabYArret.MaxCols = unNbArret + 1
    uneFenetreFille.TabYArret.Col = uneFenetreFille.TabYArret.MaxCols
    uneFenetreFille.TabYArret.Row = 1
    uneFenetreFille.TabYArret.Text = Format(unYArret)
    
    'Cr�ation d'une instance d'arr�t
    unLibelle = "Arr�t " + Format(unNbArret + 1) + " de " + uneFenetreFille.mesTC(unIndTC).monNom
    Set unArret = uneFenetreFille.mesTC(unIndTC).mesArrets.Add(unYArret, 15, 30, unLibelle)
    'alimentation des lignes 2 et 3 du nouvel arr�t
    uneFenetreFille.TabYArret.Row = 2
    uneFenetreFille.TabYArret.Text = Format(unArret.maVitesseMarche)
    uneFenetreFille.TabYArret.Row = 3
    uneFenetreFille.TabYArret.Text = Format(unArret.monTempsArret)
    'On rend active la colonne nouvellement cr��e
    uneFenetreFille.TabYArret.Action = SS_ACTION_ACTIVE_CELL
    'Cr�ation des objets graphiques
    uneFenetreFille.DessinerArretTC unIndTC, CLng(unYArret)
    'Indication d'une modification dans les donn�es TC
    IndiquerModifTC
End Sub

Public Sub SupprimerArretTC(uneFenetreFille As Form)
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unControl As Control
    Dim unY As Long, unePosTC As Long, uneColDel As Long
    
    unMsg = "Etes-vous s�r de vouloir supprimer l'arr�t " + str(uneFenetreFille.TabYArret.ActiveCol)
    unMsg = unMsg + " du transport collectif " + UCase(uneFenetreFille.ComboTC.Text) + " ?"
    If uneFenetreFille.TabYArret.MaxCols = 1 Then
        unMsg = "Un transport collectif sans arr�t ne sert � rien." + Chr(13) + Chr(13)
        unMsg = unMsg + "Supprimer plut�t le transport collectif."
        MsgBox unMsg, vbInformation
    ElseIf MsgBox(unMsg, vbYesNo + vbQuestion) = vbYes Then
        uneColDel = uneFenetreFille.TabYArret.ActiveCol
        'Stockage du Y de l'arr�t supprim� pour modifier
        'les d�calages un peu plus bas
        uneFenetreFille.TabYArret.Row = 1
        uneFenetreFille.TabYArret.Col = uneColDel
        unY = Val(uneFenetreFille.TabYArret.Text)
        'Suppression du Y de l'arr�t du TC
        unePosTC = uneFenetreFille.ComboTC.ListIndex + 1
        uneFenetreFille.mesTC(unePosTC).mesArrets.Remove uneColDel
        'Suppression des objets graphiques (NomArret et IconeArret)
        'de l'arr�t TC en sachant que mesObjGraphics sont des NomArret
        Set unControl = uneFenetreFille.mesTC(unePosTC).mesObjGraphics(uneColDel)
        Unload uneFenetreFille.IconeArret(unControl.Index)
        Unload unControl
        uneFenetreFille.mesTC(unePosTC).mesObjGraphics.Remove uneColDel
        'Recherche des arr�ts confondus en unY pour alimenter
        'les listes d'arr�ts et de TC trouv�s
        unNb = uneFenetreFille.RechercherArretConfondu(unY, uneListeIndexTC, uneListeIndexArret)
        'Mise � jour des d�calages des labels NomArr�t
        Call MiseAJourNomArret(uneFenetreFille, uneListeIndexTC, uneListeIndexArret)
        'D�calage des colonnes du tableau repr�sentant les arr�ts du TC
        'et mise � jour des tags des objets graphiques des arr�ts suivants
        For i = uneColDel To uneFenetreFille.TabYArret.MaxCols - 1
            For j = 1 To 3
                'positionnment en ligne j
                uneFenetreFille.TabYArret.Row = j
                'R�cup�ration du contenu de la cellule i + 1
                uneFenetreFille.TabYArret.Col = i + 1
                uneStrTmp = uneFenetreFille.TabYArret.Text
                'Affectation de la cellule i
                uneFenetreFille.TabYArret.Col = i
                uneFenetreFille.TabYArret.Text = uneStrTmp
            Next j
            'Mise � jour des tags des objets graphiques des arr�ts suivants
            Set unControl = uneFenetreFille.mesTC(unePosTC).mesObjGraphics(i)
            unControl.Tag = Format(unePosTC) + "-" + Format(i)
            uneFenetreFille.IconeArret(unControl.Index).Tag = unControl.Tag
        Next i
        'Suppression de la colonne de l'arr�t TC dans le spread TabYArret
        'S�lection d'une colonne
        uneFenetreFille.TabYArret.Col = uneColDel
        ' Suppression de la colonne s�lectionn�e
        uneFenetreFille.TabYArret.Action = SS_ACTION_DELETE_COL
        uneFenetreFille.TabYArret.MaxCols = uneFenetreFille.TabYArret.MaxCols - 1
        
        'Mise � jour de la s�lection graphique
        'Le dernier s�lectionn� a �t� d�truit ==> S�lection vide pour ne rien d�selectionner
        uneFenetreFille.monIndSel = 0
        If uneColDel = uneFenetreFille.TabYArret.MaxCols + 1 Then
            'Cas o� l'on supprime le dernier arr�t,
            'on s�lectionnera le nouveau dernier
            uneColDel = uneColDel - 1
        End If
        'S�lection graphique � partir de la cellule active, c'est le nouveau
        'dernier arr�t si l'on a supprim� le dernir arr�t, ou l'arr�t pr�c�dent
        'celui qui a �t� supprim� dans les autres cas.
        'S�lection graphique de l'arr�t correspondant � la cellule active
        '==> colonne active celle d'indice ColDel
        MiseAJourSelectionParCellule uneFenetreFille, ArretSel, unePosTC, uneColDel
        
        'Indication d'une modification dans les donn�es TC
        IndiquerModifTC
    End If
End Sub
Public Sub CreerFeu(uneFenetreFille As Form)
    'Cr�ation d'un feu du carrefour courant du site courant (uneFenetreFille)
    Dim unY As Integer
    Dim unFeu As Feu
    Dim unSensMontant As Boolean
    
    With uneFenetreFille
        If .monCarrefourCourant.mesFeux.Count = 0 Then
        'Cas de la cr�ation du premier feu d'un carrefour
            If .mesCarrefours.Count = 1 Then
                'Cas de la cr�ation du premier carrefour
                unY = 0
            ElseIf .mesCarrefours.Count = 2 Then
                'Cas de la cr�ation du deuxi�me carrefour
                'Premier feu mis � 500 m du feu max
                unY = .monYMaxFeu + 500
                'Recalcul du Y min des feux = Y min du premier carrefour
                '==> Englobant en Y min et max OK
                .monYMinFeu = DonnerYMinCarf(.mesCarrefours(1))
            Else
                'Cas d'un nouveau carrefour autre que le premier
                'le nouveau sera mis en dernier lors de sa cr�ation
                unY = .monYMaxFeu
                'Mise � 500 m�tres du premier feu du nouveau carrefour par
                'rapport au carrefour dont le feu correspond au Y le plus grand
                unY = unY + 500
            End If
        Else
            'Cas de la cr�ation d'un feu autre que le premier
            'On le met � 20 m�tres du Y le plus grand parmi tous les Y
            'des feux du carrefour courant auquel on ajoute ce nouveau feu
            unY = 20 + DonnerYMaxCarf(.monCarrefourCourant)
        End If
    
        'Calcul du sens du nouveau feu
        'par d�faut bas� sur l'indice de cr�ation :
        'si impair ==> montant, si pair descendant
        If (.monCarrefourCourant.mesFeux.Count + 1) Mod 2 = 0 Then
            unSensMontant = False
        Else
            unSensMontant = True
        End If
        
        'Ajout d'un nouveau feu
        Set unFeu = .monCarrefourCourant.mesFeux.Add(unSensMontant, unY, .maDur�eDeCycle / 2, 0)
        'Stockage du carrefour du feu cr��
        Set unFeu.monCarrefour = .monCarrefourCourant
        'Ajout d'une nouvelle ligne pour le nouveau feu
        .TabPropCarf.MaxRows = .monCarrefourCourant.mesFeux.Count
        'Mise � jour des titres des rang�es
        .TabPropCarf.Col = 0
        .TabPropCarf.Row = .TabPropCarf.MaxRows
        .TabPropCarf.Text = "Feu " + str(.TabPropCarf.Row)
        'Affichage des valeurs par d�faut
        .TabPropCarf.Col = 1
        If unSensMontant Then
            .TabPropCarf.Text = "Montant"
        Else
            .TabPropCarf.Text = "Descendant"
        End If
        
        .TabPropCarf.Col = 2
        .TabPropCarf.Text = Format(unFeu.monOrdonn�e)
        
        .TabPropCarf.Col = 3
        .TabPropCarf.Text = Format(unFeu.maDur�eDeVert)
        
        .TabPropCarf.Col = 4
        .TabPropCarf.Text = Format(unFeu.maPositionPointRef)
        
        'On rend actif dans TabPropCarf la ligne du dernier feu cr��
        .TabPropCarf.Col = 1
        .TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
        'Modification de la position du label Nom de carrefour
        'au barycentre des Y de ses feux
        ModifYNomCarf uneFenetreFille, .monCarrefourCourant
    End With
    'Cr�ation des objets graphiques du feu num�ro TabPropCarf.MaxRows
    DessinerFeu uneFenetreFille, uneFenetreFille.monCarrefourCourant.maPosition, uneFenetreFille.TabPropCarf.MaxRows
    'S�lection du dernier feu cr�� avec son carrefour
    MiseAJourSelectionParCellule uneFenetreFille, FeuSel, uneFenetreFille.monCarrefourCourant.maPosition, uneFenetreFille.TabPropCarf.MaxRows
    'Redessin avec le bon niveau de zoom, celui maximun englobant tous les feux
    RedessinerTout uneFenetreFille, CLng(unY) 'unY converti en entier long
    'Indication d'une modification dans les donn�es carrefour
    uneFenetreFille.maModifDataCarf = True
End Sub

Public Function DonnerYCarrefour(unCarf As Carrefour) As Integer
    'Calcul de l'ordonn�e du carrefour en prenant le barycentre du Y de ses feux
    Dim unYMoyen As Double
    Dim unNbFeux As Integer
    
    unYMoyen = 0
    unNbFeux = unCarf.mesFeux.Count
    
    For i = 1 To unNbFeux
        unYMoyen = unYMoyen + unCarf.mesFeux(i).monOrdonn�e
    Next i
    
    DonnerYCarrefour = Fix(unYMoyen / unNbFeux)
End Function

Public Sub CreerCarrefour(uneFenetreFille As Form)
    Dim unNom As String, uneCl� As String
    Dim uneValeurDefaut As String
    
    With uneFenetreFille
        'Nom g�n�rique � partir du nombre total d'objets graphiques carrefours
        'cr��s dans ce site.
        'Uniquement pour le premier carrefour lors de la cr�ation d'un nouveau site d'�tudes
        'On ajoute +1 car monNbObjGraphicCarf est incr�ment�e plus tard dans
        'DessinerCarrefour ==> coh�rence
        If .monNbObjGraphicCarf = 0 Then
            'Cas du premier carrefour cr��, c'est nouveau site
            'qui le fait ==> Nom g�n�r� automatiquement
            unNom = "Carrefour " + Format(.monNbObjGraphicCarf + 1)
        Else
            'Saisie et Verification que le nom saisie n'existe pas,
            'sinon demande de modif � l'utilisateur
            Do
                unMsg = "Entrez un nom de carrefour (15 caract�res maximun):"
                unTitre = "Cr�ation d'un carrefour" ' D�finit le titre.
                unNom = InputBox(unMsg, unTitre, uneValeurDefaut)
                unNom = Trim(unNom) 'Suppression des blancs avant et apr�s
                uneValeurDefaut = unNom
                If Len(unNom) > 15 Then
                    unMsg = "Le nom d'un carrefour est limit� � 15 caract�res"
                    MsgBox unMsg, vbCritical
                    uneSortie = False
                ElseIf Trim(unNom) = "" Then
                    'Cas du click sur le bouton annuler ou sur OK sans rentrer de nom
                    '==> Sortie sans rien faire comme un annuler
                    Exit Sub
                ElseIf PosInListe(unNom, uneFenetreFille.ComboNomCarf) <> -1 Then
                    'Cas o� le nom existe d�j�
                    unMsg = "Le carrefour " + UCase(unNom) + " existe d�j�."
                    MsgBox unMsg, vbCritical
                    uneValeurDefaut = unNom
                    uneSortie = False
                Else
                    uneSortie = True
                End If
            Loop While uneSortie = False
        End If
        'Cr�ation du nouveau carrefour avec son nom unique
        Set .monCarrefourCourant = .mesCarrefours.Add(unNom, .maVitSensM, .maVitSensD)
        'Affectation des valeurs par d�faut des demandes et des d�bits de saturation
        .monCarrefourCourant.SetDemDeb 0, 1800, 0, 1800
        'Cr�ation du label NomCarf du carrefour qui sera mis en dernier position
        'dans la collection mesCarrefours
        DessinerCarrefour uneFenetreFille, uneFenetreFille.mesCarrefours.Count
        'Cr�ation et Affichage du premier feu
        CreerFeu uneFenetreFille
        'Mise � jour de la combobox listant les noms de carrefours
        .ComboNomCarf.AddItem unNom
        .ComboNomCarf.ListIndex = .ComboNomCarf.ListCount - 1
        'Mise � jour des combobox des TC listant les carrefours
        'de d�part et d'arriv�e possibles
        .ComboCarfDep.AddItem unNom
        .ComboCarfArr.AddItem unNom
        'Mise � jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
        .TabInfoCalc.MaxRows = .mesCarrefours.Count
        RemplirLigneTabInfoCalc uneFenetreFille, .monCarrefourCourant.maPosition
        'Mise � jour du tableau TabDecal de l'onglet Tableau de r�sultat
        .TabDecal.MaxRows = .mesCarrefours.Count
    End With
End Sub

Public Sub DessinerFeu(uneFenetreFille As Form, unIndCarf As Integer, unIndFeu As Integer)
    Dim unePos As Long, unYreel As Long
    Dim unYMaxFeu As Integer
    
    With uneFenetreFille
        unYreel = .monYMaxFeu - .mesCarrefours(unIndCarf).mesFeux(unIndFeu).monOrdonn�e
        'Conversion du Yr�el en Y �cran dans la FrameVisuCarf
        unePos = ConvertirReelEnEcran(unYreel, .maLongueurAxeY, .AxeOrdonn�e.Y2 - .AxeOrdonn�e.Y1)
        'Incr�mentation du nombre d'objets graphiques Feu cr��s
        .monNbObjGraphicFeu = .monNbObjGraphicFeu + 1
        i = .monNbObjGraphicFeu
        'Cr�ation du label pour le num�ro du feu
        Load .NumFeu(i)
        'Cr�ation de l'icone graphique FEU tricolore du feu
        Load .IconeFeu(i)
        'Positionnement du feu (Num�ro et ic�ne Feu) � droite de l'axe des Y
        'pour un feu montant et � gauche pour un feu descendant
        PlacerFeuAxeY uneFenetreFille, unIndCarf, unIndFeu, i
        'Positionnement en Y �cran
        .NumFeu(i).Top = unePos + .AxeOrdonn�e.Y1 - .NumFeu(i).Height
        .IconeFeu(i).Top = unePos + .AxeOrdonn�e.Y1 - .IconeFeu(i).Height
        'Affichage des objets graphiques du feu
        .NumFeu(i).Visible = True
        .IconeFeu(i).Visible = True
        'Codage permettant de retrouver le carrefour et son feu � partir des objets graphiques
        'Tag = index dans la collection des carrefours plus un tiret et le num�ro du feu
        .NumFeu(i).Tag = Format(unIndCarf) + "-" + Format(unIndFeu)
        .IconeFeu(i).Tag = .NumFeu(i).Tag
        'Stockage dans la liste des objets graphiques repr�sentant le feu
        .mesCarrefours(unIndCarf).mesFeuxGraphics.Add .NumFeu(i)
    End With
End Sub
Public Sub DessinerCarrefour(uneFenetreFille As Form, unIndCarf As Integer)
    Dim unePos As Long
    
    With uneFenetreFille
        'Conversion du Yr�el = 0 en Y �cran dans la FrameVisuCarf
        unePos = ConvertirReelEnEcran(0, .maLongueurAxeY, .AxeOrdonn�e.Y2 - .AxeOrdonn�e.Y1)
        'Incr�mentation du nombre d'objets graphiques Carrefour cr��s
        .monNbObjGraphicCarf = .monNbObjGraphicCarf + 1
        i = .monNbObjGraphicCarf
        'Cr�ation du label pour le nom du carrefour
        Load .NomCarf(i)
        .NomCarf(i).Caption = .mesCarrefours(unIndCarf).monNom
        'Positionnement en Y �cran
        .NomCarf(i).Top = unePos + (.AxeOrdonn�e.Y2 + .AxeOrdonn�e.Y1) / 2 - .NomCarf(i).Height
        'Affichage des objets graphiques du feu
        .NomCarf(i).Visible = True
        'Codage permettant de retrouver le carrefour � partir des objets graphiques
        'Tag = index dans la collection des carrefours
        .NomCarf(i).Tag = Format(unIndCarf)
        'Stockage dans la liste des objets graphiques repr�sentant le feu
        Set .mesCarrefours(unIndCarf).monCarfGraphic = .NomCarf(i)
    End With
End Sub


Public Sub ModifYNomCarf(uneFenetreFille As Form, unCarf As Carrefour)
    'Modification de la position du label Nom du carrefour unCarf
    'au barycentre des Y de ses feux
    Dim unYreel As Long
    
    unYreel = DonnerYCarrefour(unCarf)
    'Conversion du Yr�el en Y �cran dans la FrameVisuCarf
    unePos = ConvertirReelEnEcran(uneFenetreFille.monYMaxFeu - unYreel, uneFenetreFille.maLongueurAxeY, uneFenetreFille.AxeOrdonn�e.Y2 - uneFenetreFille.AxeOrdonn�e.Y1)
    'Positionnement en Y �cran
    unCarf.monCarfGraphic.Top = unePos + uneFenetreFille.AxeOrdonn�e.Y1 - unCarf.monCarfGraphic.Height
End Sub

Public Sub AfficherValeursCarrefour(uneForm As Form, unCarf As Carrefour)
    With uneForm
        'Mise � jour de la combobox listant les noms de carrefours
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
            .TabPropCarf.Text = Format(unCarf.mesFeux(i).monOrdonn�e)
            .TabPropCarf.Col = 3
            .TabPropCarf.Text = Format(unCarf.mesFeux(i).maDur�eDeVert)
            .TabPropCarf.Col = 4
            .TabPropCarf.Text = Format(-unCarf.mesFeux(i).maPositionPointRef)
            '-PosRef car d�finition inverse entre dossier programmation et doc logiciel OndeV
        Next i
    End With
End Sub

Public Sub MiseAJourSelectionEtOngletCarrefour(uneForm As Form, unObjSel As Integer, unePosCarf As Long, unePosFeu As Long)
    'Mise � jour s�lection graphique et l'onglet Carrefour
    
    'On rend actif dans TabPropCarf la 1 �re ligne et 1�re colonne
    'pour corriger un bug dans une cellule combobox du spread
    'En fait si juste avant de cliquer sur un carrefour, on a tap� M ou D dans
    'la 1�re colonne et sur un feu de num�ro > au nombre de feux du carrefour cliqu�
    '==> plantage ind�buggable pas � pas, c'est la seule correction trouv�e
    uneForm.TabPropCarf.Row = 1
    uneForm.TabPropCarf.Col = 1
    uneForm.TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
    
    'Affichage de l'onglet Carrefour
    uneForm.TabFeux.Tab = 0
    'Mise � jour s�lection gr�ce � la cellule courante
    Call MiseAJourSelectionParCellule(uneForm, unObjSel, unePosCarf, unePosFeu)
    'Affichage des valeurs du carrefour s�lectionn� qui devient le courant
    Set uneForm.monCarrefourCourant = uneForm.mesCarrefours(unePosCarf)
    AfficherValeursCarrefour uneForm, uneForm.monCarrefourCourant
    'On rend actif dans TabPropCarf la ligne du feu s�lectionn�
    uneForm.TabPropCarf.Row = unePosFeu
    uneForm.TabPropCarf.Col = 1
    uneForm.TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
End Sub

Public Sub DonnerPosCarfFeu(unControl As Control, unePosCarf As Long, unePosFeu As Long)
    'R�cup�ration du carrefour et du feu par d�codage du tag
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
    
    'Test pr�liminaire avant la destruction du feu
    unMsg = "Etes-vous s�r de vouloir supprimer le feu " + Format(unIndFeu)
    unMsg = unMsg + " du carrefour " + unCarf.monNom + " ?"
    If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
        'Cas de confirmation n�gative
        Exit Sub
    End If
    
    If unCarf.mesFeux.Count = 1 Then
        'Cas o� le carrefour ne contient qu'un feu
        unMsg = "Le carrefour " + unCarf.monNom + " ne contient qu'un feu. Un carrefour sans feu n'ayant aucun inter�t, supprimez plut�t le carrefour."
        MsgBox unMsg, vbCritical
    Else
        'Cas o� l'on peut faire la suppression
        'Suppression des objets graphiques du feu
        Unload uneForm.IconeFeu(unCarf.mesFeuxGraphics(unIndFeu).Index)
        Unload unCarf.mesFeuxGraphics(unIndFeu)
        'Stockage du Y du feu qui va �tre supprim� pour utilisation plus loin
        unYOld = unCarf.mesFeux(unIndFeu).monOrdonn�e
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
            uneStringDecalage = "___" 'D�calage par rapport � l'axe des Y
            If unCarf.mesFeux(i).monSensMontant Then
                'Cas des feux montant
                '==> Positionnement � droite de l'axe des Y avec 2 blancs � la fin
                'pour que la mise en gras ne chevauche pas l'icone Feu
                unControl.Caption = uneStringDecalage + Format(i) + "  "
                unControl.Left = uneForm.AxeOrdonn�e.X1
                'Positionnement de l'icone Feu tricolore
                '� l'extr�mit� droite du num�ro de feux
                uneForm.IconeFeu(unControl.Index).Left = unControl.Left + unControl.Width
            Else
                'Cas des feux descendant
                '==> Positionnement � gauche de l'axe des Y avec un
                'soulign� plus grand pour intersecter l'axe des Y
                unControl.Caption = Format(i) + uneStringDecalage + uneStringDecalage
                'Ajustement de la chaine de caract�res � l'axe des ordonn�es car la
                'propri�t� AutoSize est � true ==> restriction du soulign� pr�c�dent
                unControl.Width = uneForm.AxeOrdonn�e.X1 - uneForm.NumFeu(0).Left 'Le left de l'indice n'a pas boug�
                unControl.Left = uneForm.AxeOrdonn�e.X1 - unControl.Width
                'Positionnement de l'icone Feu tricolore
                '� l'extr�mit� gauche du num�ro de feux
                uneForm.IconeFeu(unControl.Index).Left = unControl.Left - uneForm.IconeFeu(unControl.Index).Width
            End If
        Next i
        'Mise � jour � 0 de l'indice d'objet graphique selectionn� ==> D�selection
        uneForm.monIndSel = 0
        'Mise � jour de la s�lection et de l'onglet carrefour en
        's�lectionnant le feu pr�c�dent celui supprim�
        If unIndFeu > unCarf.mesFeux.Count Then unIndFeu = unIndFeu - 1
        MiseAJourSelectionEtOngletCarrefour uneForm, FeuSel, unCarf.maPosition, unIndFeu
        'Modification de la position du label Nom de carrefour
        'au barycentre des Y de ses feux
        ModifYNomCarf uneForm, unCarf
        'Redessin au bon niveau de zoom si le feu d�truit �tait
        'une des limites de l'englobant
        If unYOld = uneForm.monYMaxFeu Or unYOld = uneForm.monYMinFeu Then
            'Cas o� le Y du feu d�truit �tait le maximun ou le minimun des Y
            '==> Modification de l'englobant d'o� recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        End If
        'Indication d'une modification dans les donn�es carrefour
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
    
    'Test pr�liminaire avant la destruction du carrefour
    unMsg = "Etes-vous s�r de vouloir supprimer le carrefour "
    unMsg = unMsg + unCarf.monNom + " ?"
    If uneForm.mesCarrefours.Count = 1 Then
        'Cas o� l'on a d�truit l'unique carrefour
        unMsg = "Il n'y a aucun int�r�t � supprimer le seul et unique carrefour."
        unMsg = unMsg + Chr(13) + Chr(13)
        unMsg = unMsg + "Modifiez ou supprimez plut�t ses feux."
        MsgBox unMsg, vbCritical
        Exit Sub
    ElseIf MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then
        'Cas de confirmation n�gative
        Exit Sub
    End If
    
    'Test de l'utilisation de ce carrefour dans les TC
    unMsg = "Impossible de supprimer le carrefour " + unCarf.monNom
    unMsg = unMsg + " car il est carrefour de d�part ou d'arriv�e "
    unMsg = unMsg + "des transports collectifs ci-dessous :" + Chr(13) + Chr(13)
    
    unNbTC = 0
    For i = 1 To uneForm.mesTC.Count
        If unCarf.monNom = uneForm.mesTC(i).monCarfDep.monNom Or unCarf.monNom = uneForm.mesTC(i).monCarfArr.monNom Then
            'Cas o� le carrefour est un carrefour de d�part ou d'arriv�e d'un TC
            'car les noms de carrefour sont uniques dans un site
            unMsg = unMsg + "            " + uneForm.mesTC(i).monNom
            unMsg = unMsg + Chr(13)
            unNbTC = unNbTC + 1
        End If
    Next i
    
    If unNbTC > 0 Then
            'Cas o� le carrefour est carrefour de d�part ou d'arriv�e d'un TC
            MsgBox unMsg, vbCritical
    Else
        'Cas o� l'on peut faire la suppression du carrefour
        'Suppression de tous les feux et de leurs objets graphiques du carrefour
        uneModifEnglobant = False
        unNbFeux = unCarf.mesFeux.Count
        For i = unNbFeux To 1 Step -1
            Unload uneForm.IconeFeu(unCarf.mesFeuxGraphics(i).Index)
            Unload unCarf.mesFeuxGraphics(i)
            'Test si on supprime un feu dont le Y �tait le min
            'ou le max des Y des feux
            unY = unCarf.mesFeux(i).monOrdonn�e
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
        'de d�part et d'arriv�e possibles
        uneForm.ComboCarfDep.RemoveItem unCarf.maPosition - 1
        uneForm.ComboCarfArr.RemoveItem unCarf.maPosition - 1
        'Suppression dans la combobox ComboNomCarf
        uneForm.ComboNomCarf.RemoveItem unCarf.maPosition - 1
        'Modification des carrefours restants et de leurs feux, et qui suivait le
        'carrefour supprim� : mise � jour des attributs maPosition et des tag
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
        'Mise � jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
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
        'Mise � jour du nombre de ligne ==> perte de la derni�re ligne
        uneForm.TabInfoCalc.MaxRows = uneForm.mesCarrefours.Count
        'Mise � jour � 0 de l'indice d'objet graphique selectionn� ==> D�selection
        uneForm.monIndSel = 0
        'Mise � jour de la s�lection et de l'onglet carrefour en
        's�lectionnant le carrefour pr�c�dent celui supprim� et son feu 1
        If unIndCarf > uneForm.mesCarrefours.Count Then unIndCarf = unIndCarf - 1
        MiseAJourSelectionEtOngletCarrefour uneForm, CarfSel, unIndCarf, 1
        'Redessin au bon niveau de zoom si le carrefour d�truit contenait un
        'feu qui �tait une des limites de l'englobant
        If uneModifEnglobant Then
            'Cas o� le Y du feu d�truit �tait le maximun ou le minimun des Y
            '==> Modification de l'englobant d'o� recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        End If
        'Indication d'une modification dans les donn�es carrefour
        uneForm.maModifDataCarf = True
    End If
End Sub

Public Sub ModifierYFeu(uneForm As Form, unCarf As Carrefour, unIndFeu As Integer, unYNew As Long)
    Dim unYOld As Integer
    
    'Stockage de l'ancien Y du feu modifi�
    unYOld = unCarf.mesFeux(unIndFeu).monOrdonn�e
    'Modification du Y du feu
    unCarf.mesFeux(unIndFeu).monOrdonn�e = unYNew
    'D�placement des objets graphiques du feu
    'Conversion du unYNew valeur r�elle en Y �cran dans la FrameVisuCarf
    unePos = ConvertirReelEnEcran(uneForm.monYMaxFeu - unYNew, uneForm.maLongueurAxeY, uneForm.AxeOrdonn�e.Y2 - uneForm.AxeOrdonn�e.Y1)
    'Positionnement en Y �cran
    unCarf.mesFeuxGraphics(unIndFeu).Top = unePos + uneForm.AxeOrdonn�e.Y1 - unCarf.mesFeuxGraphics(unIndFeu).Height
    unInd = unCarf.mesFeuxGraphics(unIndFeu).Index
    uneForm.IconeFeu(unInd).Top = unePos + uneForm.AxeOrdonn�e.Y1 - uneForm.IconeFeu(unInd).Height
    'Modification de la position de l'objet graphique du carrefour
    ModifYNomCarf uneForm, unCarf
    'Test si la modif concerne un des limites de l'englobant
    With uneForm
        If (unYOld = .monYMaxFeu And unYNew < .monYMaxFeu) Or (unYOld = .monYMinFeu And unYNew > .monYMinFeu) Then
            'Cas o� l'ancien Y �tait le maximun et que le nouvel Y est plus petit
            'ou l'ancien Y �tait le minimun et que le nouvel Y est plus grand
            '==> Modification de l'englobant d'o� recalcul de ce dernier
            CalculerEnglobant uneForm
            'Redessin avec le bon niveau de zoom
            ZoomTout uneForm
        Else
            'Tous les autres cas sont r�gl�s par la fonction RedessinerTout
            'ci-dessous, qui redessinne avec le bon niveau de zoom, celui
            'maximun englobant tous les feux si unYnew est ext�rieur � l'englobant
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
    
    ' D�finit le message.
    unMsg = "Entrez le nouveau nom du carrefour (15 caract�res maximun):"
    unTitre = "Changement du nom d'un carrefour" ' D�finit le titre.
    uneValeurDefaut = uneForm.monCarrefourCourant.monNom
    ' Affiche le message, le titre et la valeur par d�faut.
    Do
        unNomCarf = InputBox(unMsg, unTitre, uneValeurDefaut)
        unNomCarf = Trim(unNomCarf) 'Suppression des blancs avant et apr�s
        uneValeurDefaut = unNomCarf
        If Len(unNomCarf) > 15 Then
            unMsg1 = "Le nom d'un carrefour est limit� � 15 caract�res"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        ElseIf Trim(unNomCarf) = "" Then
            'Cas du click sur le bouton annuler ou sur OK sans rentrer de nom
            '==> Sortie sans rien faire comme un annuler
            uneSortie = True
        ElseIf PosInListe(unNomCarf, uneForm.ComboNomCarf) <> -1 Then
            'Cas o� le nom existe d�j�
            unMsg1 = "Le carrefour " + UCase(unNomCarf) + " existe d�j�"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        Else
            uneSortie = True
            unePos = uneForm.monCarrefourCourant.maPosition - 1
            'Renommage dans la combobox listant les carrefours de d�part de TC
            RenommerCarfInCombobox uneForm.ComboCarfDep, unNomCarf, unePos
            'Renommage dans la combobox listant les carrefours d'arriv�e de TC
            RenommerCarfInCombobox uneForm.ComboCarfArr, unNomCarf, unePos
            'Renommage du carrefour dans la combobox listant les carrefours
            RenommerCarfInCombobox uneForm.ComboNomCarf, unNomCarf, unePos
            'Positionnement sur ce carrefour
            uneForm.ComboNomCarf.ListIndex = unePos
            'Changement du label NomCarf
            uneForm.monCarrefourCourant.monCarfGraphic.Caption = unNomCarf
            'Changement du nom du carrefour courant
            uneForm.monCarrefourCourant.monNom = unNomCarf
            'Mise � jour du tableau TabInfoCalc de l'onglet Cadrage d'onde verte
            RemplirLigneTabInfoCalc uneForm, uneForm.monCarrefourCourant.maPosition
            'Indication d'une modification dans les donn�es du site et pas
            'carrefour car le changement de nom n'influence pas les calculs
            maModifDataSite = True
        End If
    Loop While uneSortie = False
End Sub

Public Sub ZoomTout(uneForm As Form)
    'Zoom de tous les carrefours avec leurs feux et de tous les arr�ts TC
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
    
    'Calcul de la longueur �cran de l'axe des ordonn�es
    uneLongEcranAxeY = uneForm.AxeOrdonn�e.Y2 - uneForm.AxeOrdonn�e.Y1
    'Changement de la longueur r�elle de l'axe des Y
    uneForm.maLongueurAxeY = uneForm.monYMaxFeu - uneForm.monYMinFeu
    'Positionnement de l'origine au bon niveau de zoom
    'si elle est entre l'englobant en Y des feux
    If uneForm.monYMinFeu <= 0 And uneForm.monYMaxFeu >= 0 Then
        uneForm.Origine.Visible = True
        unePos = ConvertirReelEnEcran(uneForm.monYMaxFeu, uneForm.maLongueurAxeY, uneLongEcranAxeY)
        uneForm.Origine.Top = unePos + uneForm.AxeOrdonn�e.Y1 - uneForm.Origine.Height
    Else
        uneForm.Origine.Visible = False
    End If
    'Redessin de tous les carrefours et de leurs feux au bon zoom
    For i = 1 To uneForm.mesCarrefours.Count
        'Redessin de tous les feux au zoom
        Set unCarf = uneForm.mesCarrefours(i)
        For j = 1 To unCarf.mesFeux.Count
            unYreel = uneForm.monYMaxFeu - unCarf.mesFeux(j).monOrdonn�e
            'Conversion du Yr�el en Y �cran dans la FrameVisuCarf
            unePos = ConvertirReelEnEcran(unYreel, uneForm.maLongueurAxeY, uneLongEcranAxeY)
            'Positionnement en Y �cran des objets graphiques du feu
            Set unNumFeu = unCarf.mesFeuxGraphics(j)
            unIndex = unNumFeu.Index
            unNumFeu.Top = unePos + uneForm.AxeOrdonn�e.Y1 - unNumFeu.Height
            uneForm.IconeFeu(unIndex).Top = unePos + uneForm.AxeOrdonn�e.Y1 - uneForm.IconeFeu(unIndex).Height
        Next j
        'D�placement du label NomCarf au bon endroit par rapport au zoom
        ModifYNomCarf uneForm, unCarf
    Next i
    
    'Redessin de tous les arr�ts TC au bon zoom
    For i = 1 To uneForm.mesTC.Count
        Set unTC = uneForm.mesTC(i)
        For j = 1 To unTC.mesArrets.Count
            unYreel = uneForm.monYMaxFeu - unTC.mesArrets(j).monOrdonnee
            'Conversion du Yr�el en Y �cran dans la FrameVisuCarf
            unePos = ConvertirReelEnEcran(unYreel, uneForm.maLongueurAxeY, uneLongEcranAxeY)
            'Positionnement en Y �cran des objets graphiques de l'arr�t TC
            Set unNomArret = unTC.mesObjGraphics(j)
            unNomArret.Top = unePos + uneForm.AxeOrdonn�e.Y1 - unNomArret.Height
            'Recherche des arr�ts confondus en un Y valant unTC.mesArrets(j).monOrdonnee pour
            'alimenter les listes d'arr�ts et de TC trouv�s
            unNb = uneForm.RechercherArretConfondu(unTC.mesArrets(j).monOrdonnee, uneListeIndexTC, uneListeIndexArret, i - 1)
            'Mise � jour des d�calages des labels NomArr�t confondus en ce nouveau Y
            Call MiseAJourNomArret(uneForm, uneListeIndexTC, uneListeIndexArret)
            'On vide les listes pour le j suivant
            ViderCollection uneListeIndexTC
            ViderCollection uneListeIndexArret
            'Ajustement de la chaine de caract�res � l'axe des ordonn�es
            unNomArret.Width = uneForm.AxeOrdonn�e.X1 - unNomArret.Left
            unInd = unNomArret.Index
            uneForm.IconeArret(unInd).Top = unePos + uneForm.AxeOrdonn�e.Y1 - uneForm.IconeArret(unInd).Height
        Next j
    Next i
End Sub

Public Sub RedessinerTout(uneFenetreFille As Form, unY As Long)
    'Si l'englobant des Y change, on fait un zoom maximun englobant tous les feux
    '==> redessin de toutes les entit�s dans le nouveau rep�re (translation + zoom)
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
    'Recheche du maximun et du mininum des ordonn�es des Feux
    Dim unCarf As Carrefour
    Dim unNbCarf As Integer, unYCarf As Integer
    
    'Mise � jour des Y max et Y min des feux
    With uneForm
        unNbCarf = .mesCarrefours.Count
        'R�initialisation de l'englobant
        If unNbCarf = 1 And .mesCarrefours(1).mesFeux.Count = 1 Then
            'Cas o� on a un seul carrefour avec un seul feu
            'On prend l'englobant autour du seul feu du seul carrefour
            '� + ou - 100 m�tres
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
                unY = unCarf.mesFeux(j).monOrdonn�e
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
    'Remplit la ligne num�ro unIndLig du tableau TabInfoCalc
    'avec les valeurs par d�faut
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
    'Remplit les tableaux de l'onglet Tableau D�calages
    
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
    
    'Remplissage des d�calages
    With uneForm
        .TabDecal.MaxRows = .mesCarrefours.Count
        For i = 1 To .mesCarrefours.Count
            .TabDecal.Row = i
            .TabDecal.Col = 1
            .TabDecal.Text = .mesCarrefours(i).monNom
            'Affichage dans l'onglet Tableau de r�sultat en arrondissant
            '� l'entier le plus proche gr�ce � la fonction VB5 CInt
            'si le carrefour est pris en compte dans le calcul
            'sinon affichage vide pour les d�calages
            If .mesCarrefours(i).monDecCalcul <> -99 Then
                .TabDecal.Col = 2
                If CIntCorrig�(.mesCarrefours(i).monDecCalcul) = .maDur�eDeCycle Then
                    'Une valeur valant dur�e du cycle s'affiche 0
                    .TabDecal.Text = "0"
                Else
                    .TabDecal.Text = CIntCorrig�(.mesCarrefours(i).monDecCalcul)
                End If
                .TabDecal.Col = 3
                .TabDecal.Lock = False
                If CIntCorrig�(.mesCarrefours(i).monDecModif) = .maDur�eDeCycle Then
                    'Une valeur valant dur�e du cycle s'affiche 0
                    .TabDecal.Text = "0"
                Else
                    .TabDecal.Text = CIntCorrig�(.mesCarrefours(i).monDecModif)
                End If
                .TabDecal.Col = 4
                .TabDecal.Lock = False
                'Mise de la BackColor de la colonne 4 � celle de
                'l'image des checkbox
                '.TabDecal.BackColor = uneForm.BackColor
                If .mesCarrefours(i).monDecImp = 1 Then
                    'Cas d'un carrefour � d�calage impos�
                    .TabDecal.Text = "Oui"
                Else
                    'Cas d'un carrefour sans d�calage impos�
                    .TabDecal.Text = "Non"
                End If
            Else
                'Si le carrefour n'est pas pris en compte dans le calcul
                'affichage vide pour les d�calages
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
        unMsg = uneString + " doit �tre >= � " + Format(unIntMin)
        unMsg = unMsg + " et <= � " + Format(unIntMax)
        MsgBox unMsg, vbCritical
        unControl.Text = uneValeurDefaut
   Else
        SaisieEntierPositifEntreMinMax = True
    End If
End Function


Public Sub VerifSaisieEntier(KeyAscii As Integer, unControl As Control)
    'V�rification de la saisie d'entier gr�ce � la touche tap�e
    'dans le KeyPress event du control unControl.
    'On utilise apres le Keyup event de ce m�me control
    Dim unEntier As Integer
    Dim uneChaineTmp As String
    
    If KeyAscii > 47 And KeyAscii < 58 Then
        'Cas de saisie d'un chiffre
        '==> on ne fait rien car saisie OK
    ElseIf KeyAscii = 8 Then
        'Cas de saisie d'un retour arri�re
        '==> on ne fait rien car saisie OK
    ElseIf KeyAscii = 45 Then
        'Cas de saisie d'un moins
        unePos = InStr(1, unControl.Text, "-")
        If unControl.Text = "0" Then
            'On r�affiche 0
            unControl.Text = "0"
            KeyAscii = 0
        ElseIf unePos > 0 Then
            'Cas d'un moins existant ==> suppression du moins existant en t�te
            unControl.Text = Mid$(unControl.Text, 2)
            'On n'affichage pas le moins derni�rement saisi
            KeyAscii = 0
        Else
            'Cas d'abscence de moins ==> rajout d'un moins en t�te
            unControl.Text = "-" + unControl.Text
            'On n'affichage pas le moins derni�rement saisi
            KeyAscii = 0
        End If
    Else
        'Cas des autres touches ==> on n'affiche pas le caract�re erron�
        KeyAscii = 0
        Beep
    End If
End Sub

Public Sub RemplirFrameTC(uneForm As Form, unInd As Long)
    With uneForm
        'Affectation avec les nouvelles valeurs
        .TextTDep.Text = Format(.mesTC(unInd).monTDep)
        .TextDistAF_TC.Text = Format(.mesTC(unInd).maDistAccFrein)
        .TextDur�eAF_TC.Text = Format(.mesTC(unInd).maDureeAccFrein)
        .ColorTC.BackColor = .mesTC(unInd).maCouleur
       'Mise � vide des combobox pour �viter le test de diff�rence
        'des carrefours de d�part et d'arriv�e
        .ComboCarfDep.ListIndex = -1
        .ComboCarfArr.ListIndex = -1
        'Mise � jour des carrefours de d�part et d'arriv�e
        .ComboCarfDep.ListIndex = .mesTC(unInd).monCarfDep.maPosition - 1
        .ComboCarfArr.ListIndex = .mesTC(unInd).monCarfArr.maPosition - 1
        'Remplissage des arr�ts
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
        If unFeu.monOrdonn�e > DonnerYMaxCarf Then
            DonnerYMaxCarf = unFeu.monOrdonn�e
        End If
    Next i
End Function

Public Function DonnerYMinCarf(unCarf As Carrefour) As Integer
    'Recherche du plus petit Y parmi les Y des feux d'un carrefour
    Dim unFeu As Feu
    
    DonnerYMinCarf = 30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonn�e < DonnerYMinCarf Then
            DonnerYMinCarf = unFeu.monOrdonn�e
        End If
    Next i
End Function

Public Function DonnerYMinCarfSens(unCarf As Carrefour, unSensMontant As Boolean, unIndFeu As Integer) As Integer
    'Recherche du plus petit Y parmi les Y
    'des feux de m�me sens d'un carrefour
    'unIndFeu renvoie le feu du carrefour r�alisant ce minimun
    Dim unFeu As Feu
    
    DonnerYMinCarfSens = 30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonn�e < DonnerYMinCarfSens And unFeu.monSensMontant = unSensMontant Then
            DonnerYMinCarfSens = unFeu.monOrdonn�e
            unIndFeu = i
        End If
    Next i
    
    'Cas o� aucun feu dans le sens cherch�, on prend le feu
    'd'Y min de l'autre sens
    If DonnerYMinCarfSens = 30000 Then
        DonnerYMinCarfSens = DonnerYMinCarfSens(unCarf, Not unSensMontant, unIndFeu)
    End If
End Function

Public Function DonnerYMaxCarfSens(unCarf As Carrefour, unSensMontant As Boolean, unIndFeu As Integer) As Integer
    'Recherche du plus grand Y parmi les Y
    'des feux de m�me sens d'un carrefour
    'unIndFeu renvoie le feu du carrefour r�alisant ce maximun
    Dim unFeu As Feu
    
    DonnerYMaxCarfSens = -30000
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monOrdonn�e > DonnerYMaxCarfSens And unFeu.monSensMontant = unSensMontant Then
            DonnerYMaxCarfSens = unFeu.monOrdonn�e
            unIndFeu = i
        End If
    Next i

    'Cas o� aucun feu dans le sens cherch�, on prend le feu
    'd'Y max de l'autre sens
    If DonnerYMaxCarfSens = -30000 Then
        DonnerYMaxCarfSens = DonnerYMaxCarfSens(unCarf, Not unSensMontant, unIndFeu)
    End If
End Function

Public Function VerifierExistenceArret(unY As Long, unTabArret As vaSpread, uneListeArret As ColArretTC) As Boolean
    'Test de l'existence d'un arr�t pour le TC courant en  unY
    VerifierExistenceArret = True
    unNb = uneListeArret.Count
    i = 1
    Do While unNb > 1 And i <= unNb
        'On boucle sur toutes les ordonn�es des arr�ts du TC
        If unY = uneListeArret(i).monOrdonnee And i <> unTabArret.Col Then
            'Cas o� les Y, qui sont des entiers, sont �gaux avec un arr�t
            'diff�rent de celui en cours de modification
            unMsg = "Ce transport collectif a d�j� un arr�t d'ordonn�e " + Format(unY) + Chr(13)
            unMsg = unMsg + Chr(13) + "Saisissez une nouvelle valeur entre -9999 et 9999 m�tres :"
            uneNewVal = InputBox(unMsg, "Message d'erreur de OndeV", unY)
            If Trim(uneNewVal) = "" Then
                'Cas d'un click sur annuler ou d'une saisie vide
                'Sortie sans rien modifier en remettant la valeur pr�c�dente
                unTabArret.Text = Format(uneListeArret(unTabArret.Col).monOrdonnee)
                VerifierExistenceArret = False
                Exit Function
            Else
                'Pour reboucler sur tous les Y
                'et v�rifier l'unicit� du Y saisi
                unY = Val(uneNewVal)
                i = 0
                'Test du domaine de validit�
                Do While unY < -9999 Or unY > 9999
                    unMsg = "L'ordonn�e doit �tre comprise entre -9999 et 9999 m�tres"
                    unMsg = unMsg + Chr(13) + Chr(13) + "Saisissez une nouvelle valeur entre -9999 et 9999 m�tres :"
                    uneNewVal = InputBox(unMsg, "Message d'erreur de OndeV", unY)
                    If Trim(uneNewVal) = "" Then
                        'Cas d'un click sur annuler ou d'une saisie vide
                        'Sortie sans rien modifier en remettant la valeur pr�c�dente
                        unY = Format(uneListeArret(unTabArret.Col).monOrdonnee)
                        VerifierExistenceArret = False
                    Else
                        unY = Val(uneNewVal)
                    End If
                Loop
                'Mise � jour de la colonne avec un valeur valide
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
    'Mise � jour des combobox des TC pour l'onde verte TC
    If DonnerYCarrefour(unTC.monCarfDep) < DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC montant
        unSite.ComboTCM.AddItem unTC.monNom
    Else
        'Cas d'un TC descendant
        unSite.ComboTCD.AddItem unTC.monNom
    End If
End Sub


Public Function ChangerParamOndeTC(unSite As Form, unIndTC, unNewCarfDep As Carrefour, unNewCarfArr As Carrefour) As Boolean
    'Mise � jour des controls de la frame FrameOndeTC de
    'l'onglet Cadrage Onde verte, lors d'un changement de
    'carrefours d�part et/ou arriv�e, ce qui peut changer le sens du TC
    'Retourne true si le TC d'indice unIndTC n'est pas utilis� dans les ondes TC
    'ou si les TC cadrant les ondes TC ne changent pas de sens.
    'Retourne faux dans les autres cas
    
    unDY = DonnerYCarrefour(unSite.mesTC(unIndTC).monCarfArr) - DonnerYCarrefour(unSite.mesTC(unIndTC).monCarfDep)
    unDYnew = DonnerYCarrefour(unNewCarfArr) - DonnerYCarrefour(unNewCarfDep)
    If unDY * unDYnew < 0 Then
        'Cas d'un changement de sens du TC
        If unSite.monTCM = unIndTC Or unSite.monTCD = unIndTC Then
            'Cas d'un TC servant � cadrer les ondes TC
            unMsg = "Impossible de changer les carrefours d�part ou arriv�e du TC " + unSite.mesTC(unIndTC).monNom
            unMsg = unMsg + " car son sens de parcours est chang� or il est utilis� dans le calcul d'onde verte prenant en compte des TC"
            MsgBox unMsg, vbCritical
            ChangerParamOndeTC = False
        Else
            'Cas d'un TC non utilis�e dans les ondes vertes TC
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
        'Cas o� le sens du TC reste le m�me
        ChangerParamOndeTC = True
    End If
End Function

Public Sub SauverOptionsAffImp(unSaveRecentsOnly As Boolean)
    'Sauvegarde des options d'affichage et d'impression dans la base de
    'registre � la place du fichier OndeV.ini (fait � partir de la version 1.00.0002)
    Dim unSite1 As String, unSite2 As String
    Dim unSite3 As String, unSite4 As String
    Dim unSite As frmDocument, unFileName As String
    
    If unSaveRecentsOnly Then
        'Cas o� l'on ne sauvegarde que les fichiers r�cents
        'Appel par le unload de la MDI
        'R�cup des options d'affichage et d'impression
        Set unSite = New frmDocument
        Set unSite.mesOptionsAffImp = New OptionsAffImp
        ChargerOptionsAffImp unSite
    Else
        'Cas o� l'on sauve tout
        'Appel par le click dans Conserver par d�faut des options
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
    
        'Remplissage des 4 derniers ouverts �ventuels
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
    '� partir des infos de la base de registre � la place du fichier
    'OndeV.ini (fait � partir de la version 1.00.0002)
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
    
    unSite.mesOptionsAffImp.maVisuBandComD = GetSetting(App.Title, "OptionsAffImp", "maVisuBandComD", -1) 'True par d�faut
    unSite.mesOptionsAffImp.maVisuBandComM = GetSetting(App.Title, "OptionsAffImp", "maVisuBandComM", -1) 'True par d�faut
    unSite.mesOptionsAffImp.maVisuBandInterCarfD = GetSetting(App.Title, "OptionsAffImp", "maVisuBandInterCarfD", 0) 'False par d�faut
    unSite.mesOptionsAffImp.maVisuBandInterCarfM = GetSetting(App.Title, "OptionsAffImp", "maVisuBandInterCarfM", 0) 'False par d�faut
    unSite.mesOptionsAffImp.maVisuLigne = GetSetting(App.Title, "OptionsAffImp", "maVisuLigne", -1) 'True par d�faut
End Sub

Public Sub ChargerOptionsAffImpParDefaut(unSite As Form)
    'Alimentation de l'instance d'options d'affichage et d'impression
    'avec les valeurs par d�faut
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
    'Positionnement du feu (Num�ro et ic�ne Feu) � droite de l'axe des Y
    'pour un feu montant et � gauche pour un feu descendant
    Dim unSensMontant As Boolean
    
    With uneFenetreFille
        i = unIndObjGraphicFeu
        'R�cup�ration du sens du feu
        unSensMontant = .mesCarrefours(unIndCarf).mesFeux(unIndFeu).monSensMontant
    
        'Utilisation du label LabelTrait pour calculer le
        'd�calage � droite ou � gauche par rapport � l'axe des Y
        .LabelTrait.Caption = "___"
        
        'Stockage de la valeur Gras de la font du label NumFeu
        unSaveBold = .NumFeu(i).Font.Bold
        'Mise en non gras de la font du label NumFeu car les calculs de Width
        'des label NumFeu sont calibr�s avec une fonte non grasse
        '(la propri�t� AutoSize d'un label tient compte du type de fonte)
        .NumFeu(i).Font.Bold = False
        
        If unSensMontant Then
            'Cas des feux montant
            '==> Positionnement � droite de l'axe des Y avec 2 blancs � la fin
            'pour que la mise en gras ne chevauche pas l'icone Feu
            .NumFeu(i).Caption = .LabelTrait.Caption + Format(unIndFeu) + "  "
            .NumFeu(i).Left = .AxeOrdonn�e.X1
            .IconeFeu(i).Left = .NumFeu(i).Left + .NumFeu(i).Width
        Else
            'Cas des feux descendant
            '==> Positionnement � gauche de l'axe des Y un num�ro + un soulign�
            .NumFeu(i).Caption = Format(unIndFeu) + .LabelTrait.Caption
            'Ajustement de la chaine de caract�res � l'axe des ordonn�es car la
            'propri�t� AutoSize est � true ==> restriction du soulign� pr�c�dent
            .NumFeu(i).Left = .AxeOrdonn�e.X1 - .NumFeu(i).Width
            .IconeFeu(i).Left = .NumFeu(i).Left - .IconeFeu(i).Width
        End If
        
        'Restauration de la valeur Gras initiale de la font du label NumFeu
        .NumFeu(i).Font.Bold = unSaveBold
    End With
End Sub

Public Sub RenommerCarfInCombobox(uneComboBox As ComboBox, unNomCarf As String, unePos)
    'Renommage d'un nom de carrefour situ� dans une
    'liste de noms d'une combobox
    
    'Suppression dans la combobox listant
    'de l'item correspondant � l'ancien nom
    uneComboBox.RemoveItem unePos
    'Cr�ation en ajoutant le nouveau nom dans la liste � la m�me
    'position que l'ancien nom
    uneComboBox.AddItem unNomCarf, unePos
End Sub

Public Sub ConfigurerSpreadToPrint(unSpread As vaSpread, unHeader As String, unNomFiche As String, unTitreFiche As String)
    'Affectation des options d'impression du spread donn� par
    'la variable unSpread pass� en param�tre
    
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
    
    'Mise en noir des lignes de s�paration du spread
    unSpread.GridColor = RGB(0, 0, 0)
End Sub

Public Function TrouverCarfParNom(unNom As String) As Integer
    'Fonction retournant l'indice du carrefour de nom unNom parmi
    'les carrefours du site sinon elle retourne 0 si non trouv�
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
    'TitreEtude et Dur�eCycle
    Select Case unNumOnglet
        Case 0 ' Onglet Carrefour
            unHelpID = IDhlp_OngletCarf
        Case 1 ' Onglet TC
            unHelpID = IDhlp_OngletTC
        Case 2 ' Onglet Cadrage
            unHelpID = IDhlp_OngletCadrage
        Case 3 ' Onglet R�sultat d�calages
            unHelpID = IDhlp_OngletResDec
        Case 4 ' Onglet Dessin onde verte
            unHelpID = IDhlp_OngletDesOnde
        Case 5 ' Onglet Fiche r�sultats
            unHelpID = IDhlp_OngletFicRes
    End Select
    
    'Affectation du nouveau contexte d'aide
    frmMain.HelpContextID = unHelpID
    monSite.HelpContextID = unHelpID
    monSite.TitreEtude.HelpContextID = unHelpID
    monSite.FrameVisuCarf.HelpContextID = unHelpID
    monSite.Dur�eCycle.HelpContextID = unHelpID
    monSite.TabFeux.HelpContextID = unHelpID
    'Remplir tous les �l�ments d'un onglet TabFeux actif
    'For i = 0 To monSite.TabFeux(unNumOnglet).Controls.Count - 1
    '    monSite.TabFeux.Controls(unNumOnglet).HelpContextID = unHelpID
    'Next i
End Sub

Public Function CIntCorrig�(unSingle As Single) As Integer
    'Fonction corrigeant la fonction VB CInt qui est bugg�e
    'pour les flottants (=Single) positifs
    
    'En effet, CInt(20.5) = 20 alors que CInt(21.5)=22
    'CIntCorrig� doit rendre l'entier sup�rieur toujours
    'CintCorrig�(xx.yyy) = xx si 0.yyy < � 0.5
    'et (xx + 1) si 0.yyy >= 0.5
    
    'A utiliser pour les modifs et affichage de d�calages
    'Ces d�calages toujours entre 0 et cycle ==> >=0
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
    
    'Calcul de l'arrondi gr�ce � Int
    'car Int(xx.yy) = xx et Int(-xx.yy) = -xx - 1
    If unReste < 0.5 Then
        CIntCorrig� = Int(unSingle)
    Else
        CIntCorrig� = Int(unSingle) + 1
    End If
End Function

Public Sub TestCIntCorrig�(unIntMin As Integer, unIntMax As Integer, unPas As Single)
    'Fonction de test de CIntCorrig� entre unIntMin et unIntMax
    'avec unPas
    Dim unSingle As Single

    Do
    unSingle = CSng(InputBox("Entrez un r�el :", "TestCIntCorrig�"))
    uneRep = MsgBox(Format(unSingle) + " : CInt = " + Format(CInt(unSingle)) + " et CIntNew = " + Format(CIntCorrig�(unSingle)), vbRetryCancel)
    Loop Until uneRep = vbCancel
    
    unSingle = unIntMin
    unSingle = unIntMax + 1
    Do While unSingle <= unIntMax
        Debug.Print unSingle; " : CInt = "; CInt(unSingle); " et CIntNew = "; CIntCorrig�(unSingle)
        unSingle = unSingle + unPas
    Loop
End Sub

Public Function GetAppPath() As String
    'Fonction retournant le r�pertoire de l'application
    'avec un \ � la fin, toujours.
    'Car si RepInstall = c:\Test, App.Path rend "c:\Test"
    'mais si repInstall = c:\, alors App.Path rend "c:\"
    'donc on rajoute un \ pour faire homog�ne
    If Mid(App.Path, Len(App.Path)) = "\" Then
        GetAppPath = App.Path
    Else
        GetAppPath = App.Path + "\"
    End If
End Function

