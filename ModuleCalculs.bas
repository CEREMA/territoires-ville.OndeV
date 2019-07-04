Attribute VB_Name = "ModuleCalculs"
'Variable globale et public à ce module permettant de pointer sur les
'carrefours réduit du site courant avec un Y correspondant à la moyenne des
'ordonnées des feux équivalents du carrefour réduit, ce tableau sera
'réorganiser en classant les carrefours réduits par ordonnée croissante
Public monTabCarfY() As CarfY

'Variable privée stockant le début de vert lors du calcul du feu équivalent
'Elle n'est utilisée que pour le dessin de l'onde verte
'(Procédure DessinerOndeVerte de ce module)
Private monDebutVert As Single
Private monTStartVert As Single

'Constantes donnant les résultats possibles des bandes passantes
Public Const AucuneSolution As Integer = 0
Public Const DoubleSensPossible As Integer = 1
Public Const DoubleSensImpossible As Integer = 2

'Constantes donnant le type de dessin à réaliser dans la fonction
'TracerProgressionTC
Public Const DessinProgTC As Integer = 0  'Dessin de la progression du TC
Public Const DessinOndeTCM As Integer = 1 'Dessin de l'onde montante cadrée TC
Public Const DessinOndeTCD As Integer = 2 'Dessin de l'onde descendante cadrée TC

'Constante pour la variable indiquant la cohérence entre les données
'et les résultats du calcul d'onde
Public Const OK As Integer = 0
Public Const CalculImpossible As Integer = 1
Public Const IncoherenceDonneeCalcul As Integer = 2

'Constantes pour la sélection graphique dans
'l'onglet Graphique Onde Verte et dans la fenêtre frmPleinEcran
Public Const NoSel As Integer = 0  'Aucune Sélection graphique trouvée
Public Const PgGSel As Integer = 1 'Sélection graphique de la poignée gauche
Public Const PgDSel As Integer = 2 'Sélection graphique de la poignée droite
Public Const PlaSel As Integer = 3 'Sélection graphique d'une plage de vert
Public Const RefSel As Integer = 4 'Sélection graphique d'un point de référence

'Constante pour la précision du pick écran en Twips
Public Const PrecPick As Integer = 60

'Variables stockant les X écran de début et de fin d'une modification
'graphique interactive, Initialisée dans la fonction SelectionGraphique.
'Ainsi la différence entre ce début et cette fin permet d'avoir la valeur
'de la modification.
Public monXEcranDebModif As Single
Public monXEcranFinModif As Single

'Collection stockant les valeurs avant une modif graphique du dernier
'objet graphique sélectionné, pour les restaurer si besoin est
Private maColValPred As New Collection

'Variable privée stockant l'objet pické et son type dans
'Onglet Graphique ou frmPleinEcran
Private monTypeObjPick As Integer
Private monObjPick As Object

'Variable privée stockant le Temps total qui est l'englobant
'en temps donc suivant les X en coordonnées réelles
'Lors d'une annulation modif graphique il faut réutiliser ce nombre
'pour retrouver la même échelle en X car la modif a pu la changer
'car elle lance un redessin d'onde
Private monTmpTotalAvantModif As Long

Public Function ReduireCarrefourSite(uneForm As Form, uneColCarf As ColCarrefour, unTypeOnde As Integer) As Boolean
    'Procedure réduisant tous les carrefours du site
    'en créant les carrefours réduits à sens unique ou double
    Dim uneDureeVert As Single, unePosRef As Single
    Dim unCarf As Carrefour, uneOrdonnee As Integer
    Dim uneDureeVertD As Single, unePosRefD As Single
    Dim uneOrdonneeD As Integer
    Dim i As Integer, unNbCarfUtil As Integer
    Dim unCarfRedSens2 As CarfReduitSensDouble
    Dim unCarfRedSensU As CarfReduitSensUnique
    Dim unYFeuMinM As Integer, unYFeuMaxM As Integer
    Dim unTCM As TC, unTCD As TC
    Dim unYFeuMinD As Integer, unYFeuMaxD As Integer
    Dim unCarfEntreYminYmax As Boolean
    Dim unIndFeu As Integer
    
    'Initialisation des TC montant et descendant
    Set unTCM = Nothing
    Set unTCD = Nothing
        
    'Mise à vide des listes de Carrefour réduit
    uneForm.mesCarfReduitsSensM.Vider
    uneForm.mesCarfReduitsSensD.Vider
    uneForm.mesCarfReduitsSens2.Vider
    unNbCarf = uneColCarf.Count
    unNbCarfUtil = 0
    
    'Recherche des Y des feux d'Y min et d'Y max pour une onde TC
    If unTypeOnde = OndeTC Then
        If uneForm.monTCM > 0 Then
            'Pour une onde TC montante on ne prend que les feux dont
            'le Y est compris entre le Y min des feux du carrefour de départ
            'et le Y max des feux du carrefour d'arrivée.
            'Dans ce cas, le départ a un Y < à celui de l'arrivée
            Set unTCM = uneForm.mesTC(uneForm.monTCM)
            unYFeuMinM = DonnerYMinCarfSens(unTCM.monCarfDep, True, unIndFeu)
            unYFeuMaxM = DonnerYMaxCarfSens(unTCM.monCarfArr, True, unIndFeu)
            'Calcul du tableau de marche prenant en compte les arrêts
            'mais pas les feux
            unTCM.CalculerTableauMarcheOnde
        End If
        
        If uneForm.monTCD > 0 Then
            'Pour une onde TC descendante on ne prend que les feux dont
            'le Y est compris entre le Y min des feux du carrefour d'arrivée
            'et le Y max des feux du carrefour de départ
            'Dans ce cas, le départ a un Y > à celui de l'arrivée
            Set unTCD = uneForm.mesTC(uneForm.monTCD)
            unYFeuMinD = DonnerYMinCarfSens(unTCD.monCarfArr, False, unIndFeu)
            unYFeuMaxD = DonnerYMaxCarfSens(unTCD.monCarfDep, False, unIndFeu)
            'Calcul du tableau de marche prenant en compte les arrêts
            'mais pas les feux
            unTCD.CalculerTableauMarcheOnde
        End If
    End If
    
    'Initialisation de la valeur de retour de ReduireCarrefourSite
    ReduireCarrefourSite = True
    
    'Parcours de tous les carrefours passés en paramètre
    For i = 1 To unNbCarf
        Set unCarf = uneColCarf(i)
        'On ne travaille que sur les carrefours choisis par l'utilisateur
        If unCarf.monIsUtil Then
            'Parcours de tous les feux du carrefour pour voir s'ils sont
            'tous dans le même sens ou dans deux sens différents
            j = 2
            unIsSensDouble = False
            'Test des sens de deux feux consécutifs
            'Sortie si les sens sont différents ==> Carrefour à double sens
            'Sinon Carrefour à sens unique celui du feu 1 par exemple
            'Si un seul feu dans le carrefour, on ne rentre pas dans la boucle
            '==> Carrefour à sens unique, celui du seul feu
            Do While j <= unCarf.mesFeux.Count And unIsSensDouble = False
                If unCarf.mesFeux(j - 1).monSensMontant <> unCarf.mesFeux(j).monSensMontant Then
                    unIsSensDouble = True
                End If
                j = j + 1
            Loop
            
            'Alimentation des listes de Carrefour réduit
            If Not unIsSensDouble And unCarf.mesFeux(1).monSensMontant Then
                'Cas d'un carrefour ayant tous ses feux dans le sens montant
                'Calcul du feu équivalent dans le sens unique montant
                unIsFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinM, unYFeuMaxM, unTCM)
                If unIsFeuEquivSensMExist Then
                    'Retaillage dynamique du tableau des carrefours réduits avec ordonnée
                    'On ne stocke que les carrefours utilisés
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout à la liste des Carrefours réduits à sens unique montant
                    Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
                    'Alimentation du tableau des carrefours réduits avec ordonnée
                    AjouterCarfY unNbCarfUtil, unCarfRedSensU
                End If
                '(unIsFeuEquivSensMExist Or monNbFeuxMpris = 1) est VRAI si on
                'a trouvé un feu équivalent montant ou s'il n'y a pas de feu
                'équivalent mais que tous les feux du carrefour sont en dehors
                'de unYFeuMin et unYFeuMax donc ce n'est pas une erreur
                '==> pas de changement
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensMExist Or uneForm.monNbFeuxMpris = 1)
            ElseIf Not unIsSensDouble And Not unCarf.mesFeux(1).monSensMontant Then
                'Cas d'un carrefour ayant tous ses feux dans le sens descendant
                'Calcul du feu équivalent dans le sens unique descendant
                unIsFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinD, unYFeuMaxD, unTCD)
                If unIsFeuEquivSensDExist Then
                    'Retaillage dynamique du tableau des carrefours réduits avec ordonnée
                    'On ne stocke que les carrefours utilisés
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout à la liste des Carrefours réduits à sens unique descendant
                    Set unCarfRedSensU = uneForm.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVert, unePosRef, uneOrdonnee)
                    'Alimentation du tableau des carrefours réduits avec ordonnée
                    AjouterCarfY unNbCarfUtil, unCarfRedSensU
                End If
                '(unIsFeuEquivSensDExist Or monNbFeuxDpris = 1) est VRAI si on
                'a trouvé un feu équivalent descendant ou s'il n'y a pas de feu
                'équivalent mais que tous les feux du carrefour sont en dehors
                'de unYFeuMin et unYFeuMax donc ce n'est pas une erreur
                '==> pas de changement
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensDExist Or uneForm.monNbFeuxDpris = 1)
            Else
                'Cas d'un carrefour ayant des feux dans les deux sens
                'Calcul du feu équivalent dans le sens montant
                unIsFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinM, unYFeuMaxM, unTCM)
                'Calcul du feu équivalent dans le sens descendant
                unIsFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unePosRefD, uneOrdonneeD, , , unYFeuMinD, unYFeuMaxD, unTCD)
                If unIsFeuEquivSensMExist And unIsFeuEquivSensDExist Then
                    'Cas d'un carrefour réduit à double sens
                    'Pour toutes les ondes sauf celle TC on doit passer là
                    'sinon calcul impossible car double feu équivalent imposible
                    
                    'Retaillage dynamique du tableau des carrefours réduits avec ordonnée
                    'On ne stocke que les carrefours utilisés
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout à la liste des Carrefours réduits à double sens
                    Set unCarfRedSens2 = uneForm.mesCarfReduitsSens2.Add(unCarf)
                    'Mise à jour des propriétés dans le sens montant du carrefour réduit
                    unCarfRedSens2.SetPropsSensM uneDureeVert, unePosRef, uneOrdonnee
                    'Mise à jour des propriétés dans le sens descendant du carrefour réduit
                    unCarfRedSens2.SetPropsSensD uneDureeVertD, unePosRefD, uneOrdonneeD
                    'Alimentation du tableau des carrefours réduits avec ordonnée
                    AjouterCarfY unNbCarfUtil, unCarfRedSens2
                Else
                    'Ce cas n'est pas une erreur uniquement si onde TC
                    'On aura un carrefour réduit à sens unique
                    If unIsFeuEquivSensMExist Then
                        'Cas du sens unique montant
                        'Retaillage dynamique du tableau des carrefours réduits avec ordonnée
                        'On ne stocke que les carrefours utilisés
                        unNbCarfUtil = unNbCarfUtil + 1
                        ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                        'Ajout à la liste des Carrefours réduits à sens unique montant
                        Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
                        'Alimentation du tableau des carrefours réduits avec ordonnée
                        AjouterCarfY unNbCarfUtil, unCarfRedSensU
                    End If
                    If unIsFeuEquivSensDExist Then
                        'Cas du sens unique descendant
                        'Retaillage dynamique du tableau des carrefours réduits avec ordonnée
                        'On ne stocke que les carrefours utilisés
                        unNbCarfUtil = unNbCarfUtil + 1
                        ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                        'Ajout à la liste des Carrefours réduits à sens unique descendant
                        Set unCarfRedSensU = uneForm.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVertD, unePosRefD, uneOrdonneeD)
                        'Alimentation du tableau des carrefours réduits avec ordonnée
                        AjouterCarfY unNbCarfUtil, unCarfRedSensU
                    End If
                End If
                
                'Affectation de la valeur de retour de cette fonction
                'Retour Vrai si on a trouvé un feu équivalent
                'dans les deux sens ou que le carrfour n'a aucun feux
                'entre Ymin et Ymax, faux sinon
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensMExist Or uneForm.monNbFeuxMpris = 1) And (unIsFeuEquivSensDExist Or uneForm.monNbFeuxDpris = 1)
            End If
        End If
    Next i
    
    'Vérification de la cohérence entre le type d'onde verte choisi
    'et les carrefours réduits trouvés
    If uneForm.monTypeOnde = OndeSensM And uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
        ReduireCarrefourSite = False
        MsgBox "Impossible de privilégier le sens montant. Aucun Carrefour n'est dans ce sens", vbCritical
    ElseIf uneForm.monTypeOnde = OndeSensD And uneForm.mesCarfReduitsSensD.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
        ReduireCarrefourSite = False
        MsgBox "Impossible de privilégier le sens descendant. Aucun Carrefour n'est dans ce sens", vbCritical
    End If

    If uneForm.monTypeOnde = OndeTC And uneForm.monTCM > 0 Then
        'Cas d'une onde cadrée par un TC montant
        If uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
            ReduireCarrefourSite = False
            MsgBox "Impossible de cadrer le sens montant par le TC : " + unTCM.monNom + Chr(13) + "Aucun Carrefour n'a de feu dans ce sens entre le départ et l'arrivée de ce TC", vbCritical
        End If
    End If
    
    If uneForm.monTypeOnde = OndeTC And uneForm.monTCD > 0 Then
        'Cas d'une onde cadrée par un TC descendant
        If uneForm.mesCarfReduitsSensD.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
            ReduireCarrefourSite = False
            MsgBox "Impossible de cadrer le sens descendant par le TC : " + unTCD.monNom + Chr(13) + "Aucun Carrefour n'a de feu dans ce sens entre le départ et l'arrivée de ce TC", vbCritical
        End If
    End If
    
    If unNbCarfUtil = 0 Then
        'Cas où aucun carrefour réduit créé, on n'en crée un fictif
        'pour que ubound(monTabCarfY,1) ne plante pas (cf CalculerTempsparcours)
        'Ce carrefour réduit ne sert à rien d'autre
        ReDim Preserve monTabCarfY(1 To 1)
        'Ajout à la liste des Carrefours réduits à sens unique montant
        Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
        'Alimentation du tableau des carrefours réduits avec ordonnée
        AjouterCarfY 1, unCarfRedSensU
    End If
End Function

Public Function CalculerFeuEquivalent(unCarf As Carrefour, unSensMontant As Boolean, uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer, Optional unDecalModif As Boolean = False, Optional unSansMsgErreur As Boolean = False, Optional unYFeuMin As Integer = 0, Optional unYFeuMax As Integer = 0, Optional unTC As TC = Nothing) As Boolean
    'Calcul, lors de la réduction d'un carrefour, le feu équivalent
    'd'un carrefour qui sera utilisé dans le carrefour réduit
    'Si unSensMontant est vrai, calcul du feu équivalent dans le sens montant
    'Sinon calcul du feu équivalent dans le sens descendant
    'Cette fonction modifie les valeurs de ses trois derniers
    'paramètres d'entrée
    
    Dim uneVitesse As Single, unDecalCauseVitesse As Single
    Dim unYExtremun As Single, unCoefSens As Integer
    Dim uneColFeuSens1 As New Collection
    Dim unFeuSens1 As FeuSens1
    Dim unFeu As Feu, unFeuExt As Feu
    Dim unDebutVert As Single, unDecalage As Single
    Dim unTabBorne() As Single, unFeuPris As Boolean
    Dim uneColPeriodeVert As New Collection
    Dim unePeriodeVert As PeriodeVert
    
    'Initialisation du nombre de feux du sens choisi (FeuSens1) créé
    unNbFeu = 1
    
    'Affectation de la vitesse en m/s du carrefour dans le sens étudié
    'Elle dépend du type d'onde verte, du type de vitesse choisi (constante
    'ou variable) et des TC cadrant l'onde verte en sens montant et/ou
    'descendant
    uneVitesse = unCarf.DonnerVitSens(unSensMontant)
    If unSensMontant Then
        'Cas de recherche du feu équivalent dans le sens montant
        unSens = "montant"
        unCoefSens = 1
        'Initialisation de l'extremun des Y en sens montant
        '==> c'est un maximun, donc valeur initiale petite
        unYExtremun = -10000
    Else
        'Cas de recherche du feu équivalent dans le sens descendant
        unSens = "descendant"
        unCoefSens = -1
        'Initialisation de l'extremun des Y en sens descendant
        '==> c'est un minimun, donc valeur initiale grande
        unYExtremun = 10000
    End If
        
    'Calcul des plages de période de vert des feux
    'du carrefour entre t=0 et t<durée du cycle dans
    'le sens montant si unSensmontant = True, descendant sinon
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monSensMontant = unSensMontant Then
            'Cas d'un feu ayant le sens cherché pour le calcul
            
            If Not (unTC Is Nothing) And unDecalModif = False Then
                'Cas d'une onde verte cadrée par un TC dans le sens donné
                'par unSensMontant celui du feu équivalent recherché
                If unFeu.monOrdonnée >= unYFeuMin And unFeu.monOrdonnée <= unYFeuMax Then
                    'Cas où le feu est entre le départ et l'arrivée de ce TC
                    '==> Feu pris en compte pour le calcul du feu équivalent,
                    'sinon non
                    unFeuPris = True
                Else
                    unFeuPris = False
                End If
            Else
                'Cas d'une onde verte qui n'est pas cadrée par un TC
                'dans le sens du feu équivalent recherché
                'Tous les feux dans ce sens sont pris
                unFeuPris = True
            End If
            
            If unFeuPris Then
                'Cas d'un feu intervenant dans le calcul du feu équivalent
                
                'Calcul du maximun des Y en sens montant ou
                'du minimun des Y en sens descendant
                If (unSensMontant And unFeu.monOrdonnée > unYExtremun) Or (Not unSensMontant And unFeu.monOrdonnée < unYExtremun) Then
                    unYExtremun = unFeu.monOrdonnée
                    Set unFeuExt = unFeu 'Stockage du feu d'Y extremun
                End If
                
                'Calcul de son début de vert
                If unDecalModif Then
                    'Cas où l'on recalcule les bandes passantes
                    'après une modification des décalages
                    'unCarf est un carrefour réduit global en prenant
                    'tous les feux equivalents des carrefours réduits
                    '==> la vitesse n'est plus local au carrefour
                    'si vitesse non constante
                    '==> Utilisation du décalage induit par les vitesses
                    'constantes ou variables éventuelles
                    
                    'Le feu du carrefour global est le feu équivalent montant
                    'ou descendant d'un carrefour réduit, mais son champ
                    'carrefour pointe sur le carrefour dont il est issu
                    'après réduction
                    'Le décalage du aux vitesses doit être multiplié par le sens
                    '(1 = montant et -1 = descendant)
                    If unSensMontant Then
                        unDecalCauseVitesse = unFeu.monCarrefour.monDecVitSensM
                    Else
                        unDecalCauseVitesse = -unFeu.monCarrefour.monDecVitSensD
                    End If
                    'Récupération du décalage modifié du carrefour dont
                    'le feu unFeu est le feu équivalent
                    unDecalage = unFeu.monCarrefour.monDecModif
                Else
                    'Cas où l'on calcule les bandes passantes,
                    'les décalages seront calculés plus tard
                    '==> Mis à zéro pour avoir une formule de
                    'calcul du début vert valable pour tous les cas
                    unDecalage = 0
                    'Dans ce cas, même en vitesse variable, le décalage du
                    'aux vitesses est du uniquement à la vitesse locale du
                    'carrefour
                    If unTC Is Nothing Then
                        'Cas d'une onde non cadrée par unTC dans le sens du feu équivalent
                        unDecalCauseVitesse = unFeu.monOrdonnée / uneVitesse
                    Else
                        'Cas d'une onde cadrée par un TC dans le sens du feu équivalent
                        'Le coefsens permet d'inverser le signe des Y pour le cas descendant
                        'et de rendre le décalage négatif dans le cas descendant pour respecter
                        'l'analogie avec les vitesses variables ci-dessus,
                        'uneVitesse > 0 si montant, < 0 si descendant
                        unDecalCauseVitesse = unCoefSens * unTC.CalculerDecalCauseProgTC(unTC.mesPhasesTMOnde, unFeu.monOrdonnée, unCoefSens)
                    End If
                End If 'Fin du cas de modif des décalages
                
                unDebutVert = unDecalage + unFeu.maPositionPointRef - unDecalCauseVitesse
                'On ramène modulo entre [0, duréee du cycle[
                unDebutVert = ModuloZeroCycle(unDebutVert, monSite.maDuréeDeCycle)
                
                'Création d'un feu avec ses plages de vert
                Set unFeuSens1 = New FeuSens1
                Set unFeuSens1.monFeu = unFeu
                If unDebutVert + unFeu.maDuréeDeVert > monSite.maDuréeDeCycle Then
                    'Cas de l'existence de deux plages de vert
                    unFeuSens1.monNbPlageVert = 2
                    unFeuSens1.maBorneVert1 = unDebutVert + unFeu.maDuréeDeVert - monSite.maDuréeDeCycle
                    unFeuSens1.maBorneVert2 = unDebutVert
                Else
                    'Cas de l'existence d'une seule plage de vert
                    unFeuSens1.monNbPlageVert = 1
                    unFeuSens1.maBorneVert1 = unDebutVert
                    unFeuSens1.maBorneVert2 = unDebutVert + unFeu.maDuréeDeVert
                End If
                'Augmentation du tableau dynamique pour stocker
                'les 2 bornes de vert en preservant les précédents
                ReDim Preserve unTabBorne(unNbFeu + 1)
                'Ajout au tableau des bornes de vert
                unTabBorne(unNbFeu) = unFeuSens1.maBorneVert1
                unTabBorne(unNbFeu + 1) = unFeuSens1.maBorneVert2
                unNbFeu = unNbFeu + 2
                'Création d'un feu avec ses plages de vert
                uneColFeuSens1.Add unFeuSens1
            End If 'Fin de if unFeuPris
        End If 'Fin de if Feu.sensmontant = sensmontant
    Next i
        
    'Calcul des caractéristiques feu équivalent
    '(ordonnée, durée de vert, position du point de réference)
    If uneColFeuSens1.Count = 0 Then
        'Cas d'une onde cadrée par un TC dans le sens du feu équivalent
        'cherché mais n'ayant aucun feu situé entre le départ et l'arrivée
        'de ce TC
        CalculerFeuEquivalent = False
    ElseIf uneColFeuSens1.Count = 1 Then
        'Cas d'un carrefour n'ayant qu'un seul feu dans le sens étudié
        'Le feu équivalent trouvé sera cet unique feu
        CalculerFeuEquivalent = True
        Set unFeu = uneColFeuSens1(1).monFeu
        uneDureeVert = unFeu.maDuréeDeVert
        unePosRef = unFeu.maPositionPointRef
        uneOrdonnee = unFeu.monOrdonnée
    Else
        'Cas d'un carrefour ayant plusieurs feux dans le sens étudié
        'Tri dans l'ordre croissant des différentes bornes de vert trouvées
        TrierOrdreCroissant unTabBorne
        'Rajout de la dernière borne de vert valant Durée du cycle
        ReDim Preserve unTabBorne(unNbFeu)
        unTabBorne(unNbFeu) = monSite.maDuréeDeCycle
        'Test de l'état des feux entre les bornes ordonnées
        'avec la borne d'indice 0 valant 0
        For i = 0 To UBound(unTabBorne, 1) - 1
            If unTabBorne(i) < unTabBorne(i + 1) Then
                'Cas où deux bornes successives sont différentes
                'Car tri précédent par ordre croissant sans virer les doublons
                j = 1
                IsTousFeuxVert = True
                Do
                    'On regarde si tous les feux du sens étudiés sont verts
                    'entre deux bornes successives différentes en regardant
                    'à la borne inf de la période, unTabBorne(i)
                    If Not uneColFeuSens1(j).IsVert(unTabBorne(i)) Then
                        'Cas où le feu rouge à cet instant
                        IsTousFeuxVert = False
                    End If
                    j = j + 1
                'Fin de boucle sur les feux du sens étudié
                Loop While j <= uneColFeuSens1.Count And IsTousFeuxVert = True
                
                'Ajout à la collection des périodes de vert trouvée
                'si tous les feux sont verts dans cette période
                If IsTousFeuxVert Then
                    Set unePeriodeVert = New PeriodeVert
                    unePeriodeVert.monDebutVert = unTabBorne(i)
                    unePeriodeVert.maDuree = unTabBorne(i + 1) - unTabBorne(i)
                    uneColPeriodeVert.Add unePeriodeVert
                End If
            End If 'Fin du if entre 2 bornes consécutives
        Next i 'Fin de boucle sur les bornes de vert
        
        If uneColPeriodeVert.Count = 0 Then
            'Cas où aucun période de vert n'a été trouvé
            '==> Aucun feu du carrefour vert en même temps
            'dans le sens étudié, donc pas de feu équivalent trouvé
            CalculerFeuEquivalent = False
            
            If Not unDecalModif And unSansMsgErreur = False Then
                'Affichage d'un message d'erreur ciblée dans le cas où
                'l'on ne recalcule pas les bandes aprés une modification
                'd'un décalage
                unMsg = "Impossible de trouver une plage de vert commune pour "
                unMsg = unMsg + "les feux du carrefour " + unCarf.monNom
                unMsg = unMsg + " dans le sens " + unSens + "." + Chr(13) + Chr(13)
                unMsg = unMsg + "Modifier un ou plusieurs des paramètres de ce carrefour "
                unMsg = unMsg + "dans le sens " + unSens + " :" + Chr(13)
                unMsg = unMsg + "    sa durée de vert, sa position du point de référence, sa vitesse"
                MsgBox unMsg, vbCritical
            End If
        Else
            'Cas où le feu équivalent existe, car une période de vert trouvé
            CalculerFeuEquivalent = True
            
            'Si la première et dernière période a tous ses feux verts et que la
            'première période commence à 0 et que la dernière finit à la durée du cycle
            'la durée de la dernière devient la somme des durées de ces 2 périodes
            'et la date de début de vert reste celle de la dernière période et
            'on supprime la première période
            unNbPeriodeVert = uneColPeriodeVert.Count
            If unNbPeriodeVert <> 1 And uneColPeriodeVert(1).monDebutVert = 0 And (uneColPeriodeVert(unNbPeriodeVert).monDebutVert + uneColPeriodeVert(unNbPeriodeVert).maDuree = monSite.maDuréeDeCycle) Then
                'Calcul des caractéristiques feu équivalent
                '(ordonnée, durée de vert, position du point de réference)
                uneColPeriodeVert(unNbPeriodeVert).maDuree = uneColPeriodeVert(unNbPeriodeVert).maDuree + uneColPeriodeVert(1).maDuree
                'Suppression de la première période de vert
                uneColPeriodeVert.Remove 1
                unNbPeriodeVert = uneColPeriodeVert.Count
            End If
            'Recherche de la période de vert la plus grande
            uneDureeMax = 0
            unIndPeriodeMax = 0
            For i = 1 To unNbPeriodeVert
                If uneColPeriodeVert(i).maDuree > uneDureeMax Then
                    uneDureeMax = uneColPeriodeVert(i).maDuree
                    unIndPeriodeMax = i
                End If
            Next i
            'Calcul des caractéristiques feu équivalent
            '(ordonnée, durée de vert, position du point de réference)
            uneDureeVert = uneDureeMax
            If unTC Is Nothing Then
                'Cas d'une onde non cadrée par unTC dans le sens du feu équivalent
                If unDecalModif Then
                    'Cas de la réduction pour le calcul des bandes passantes lors
                    'de la modif manuelle d'un décalage calculé ou lors de la
                    'réduction des carrefours à date imposée ==> Projection sur le
                    'carrefour le plus bas en Y pour le cas montant et sur le
                    'carrefour le plus haut en Y pour le cas descendant
                    If unSensMontant Then
                        unDecalCauseVitesse = unFeuExt.monCarrefour.monDecVitSensM
                    Else
                        unDecalCauseVitesse = -unFeuExt.monCarrefour.monDecVitSensD
                    End If
                    unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unDecalCauseVitesse
                Else
                    'Cas de la réduction d'un carrefour pour calculer les
                    'décalages ==> Projection sur Y = 0
                    unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unYExtremun / uneVitesse
                End If
            Else
                'Cas d'une onde cadrée par un TC dans le sens du feu équivalent
                'Le coefsens permet d'inverser le signe des Y pour le cas descendant
                'et de rendre le décalage négatif dans le cas descendant pour respecter
                'l'analogie avec les vitesses variables ci-dessus,
                'uneVitesse > 0 si montant, < 0 si descendant
                unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unCoefSens * unTC.CalculerDecalCauseProgTC(unTC.mesPhasesTMOnde, unYExtremun, unCoefSens)
            End If
            uneOrdonnee = unYExtremun
            'Stockage du début de vert dans une variable privée
            'à ce module ModuleCalculs (cf Déclarations de ce module)
            monDebutVert = uneColPeriodeVert(unIndPeriodeMax).monDebutVert
        End If 'Fin du if période vert trouvé
    End If 'Fin du calcul des caractéristiques du feu équivalent
    
    'Stockage du nombre de feux pris pour le calcul du feu équivalent
    'dans le sens choisi
    If unSensMontant Then
        monSite.monNbFeuxMpris = unNbFeu
    Else
        monSite.monNbFeuxDpris = unNbFeu
    End If
    
    'Libération de la mémoire
    Set uneColFeuSens1 = Nothing
    Set uneColPeriodeVert = Nothing
End Function

Public Function ModuloZeroCycle(unReel As Single, uneDureeCycle As Integer) As Single
    'Fonction ramenant un nombre réel dans l'intervalle [0, durée du cycle[
    If unReel >= 0 Then
        ModuloZeroCycle = unReel - uneDureeCycle * Fix(unReel / uneDureeCycle)
    Else
        ModuloZeroCycle = unReel + uneDureeCycle * (1 - Fix(unReel / uneDureeCycle))
    End If
    
    If ModuloZeroCycle = uneDureeCycle Then ModuloZeroCycle = 0
End Function

Public Sub TrierOrdreCroissant(unTabBorne() As Single)
    'Réorganisation par ordre croissant d'un tableau de réels
    'indexé entre 1 et n avec l'indice zéro nul mais qui ne sert pas.
    'Algo choisi : Le tri insertion (récupérer sur Internet)
    'Il consiste à comparer successivement un élément
    'à tous les précédents et à décaler les éléments intermédiaires

    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unTmp As Single
    
    'Mise à zéro du contenu d'indice 0
    unTabBorne(0) = 0
    
    'Tri
    unNbTotal = UBound(unTabBorne, 1)
    For j = 2 To unNbTotal
            unTmp = unTabBorne(j)
            i = j - 1
            Do While i > 0 And unTabBorne(i) > unTmp 'Indice zéro nul, ça évite le plantage du And pour i = 0
                unTabBorne(i + 1) = unTabBorne(i)
                i = i - 1
            Loop
            unTabBorne(i + 1) = unTmp
    Next j
End Sub

Public Function CalculerOndeVerte(uneForm As Form, Optional uneModifDec As Boolean = False) As Boolean
    'Procédure essayant de calculer l'onde verte
    'Si uneModifDec = true c'est que CalculerOndeVerte a été appelée
    'après une modif manuelle dans l'onglet Tableau Décalage
    'ou après une modif graphique à la souris d'un décalage d'un carrefour
    'à décalage imposé ==> on ne change que le champ monDecModif de ce carrefour
    
    'Variables locales caractéristiques des bandes passantes cherchées
    'Sens 1 = sens montant pour onde verte double sens ou le sens
    '         privilégié pour une onde verte à sens privilégié
    'Sens 2 = sens descendant pour onde verte double sens ou l'autre sens
    '         que celui privilégié pour une onde verte à sens privilégié
    Dim unB1 As Single 'valeur de la bande passante de vert du sens 1
    Dim unB2 As Single 'valeur de la bande passante de vert du sens 2
    Dim unH As Single  'Temps écoulé entre les événements
                       '"Passage au vert montant" et "Fin de vert descendant"
    Dim unCarfRed As Object
    Dim unCarf As Carrefour
    Dim unDecImp As Single
    
    'Affectation à vrai de la réalisation du calcul de l'onde
    'On mettra faux chaque fois que le calcul de l'onde est impossible
    CalculerOndeVerte = True
        
    With uneForm
        'Calcul à double sens des ondes vertes TC
        'si aucun TC montant et descendant
        If .monTypeOnde = 3 And .ComboTCM.Text = "Aucun" And .ComboTCD.Text = "Aucun" Then
                MsgBox "Dans l'onglet Cadrage Onde Verte, aucun TC montant et/ou descendant n'ont été choisis." + Chr(13) + Chr(13) + "Calcul d'onde verte prenant en compte les TC impossible", vbCritical
                CalculerOndeVerte = False
                monSite.maCoherenceDataCalc = CalculImpossible
                Exit Function
        End If
                
        'Le calcul n'est pas effectué s'il n'y a pas eu une modif dans les
        'données carrefours, TC qui cadre l'onde verte, calculs d'onde et de
        'modifications graphiques
        If Not .maModifDataCarf And Not .maModifDataOndeTC And Not .maModifDataOnde And Not .maModifDataDes Then
            If .maCoherenceDataCalc = CalculImpossible Then
                MsgBox "Le calcul d'onde verte est impossible avec les données de ce site", vbCritical
                'Mise à zéro des bandes et des décalages
                RendreNulleBandesEtDecalages uneForm
            End If
            If .maCoherenceDataCalc <> IncoherenceDonneeCalcul Then Exit Function
            'Si les données sont incohérentes avec les
            'résultats (modif de données sans recalcul) ont refait le calcul
        End If
                
        'Mise en grisée du menu Annuler dernière modif graphique si on
        'a fait une modif par saisie et pas par interaction graphique
        If .maModifDataDes = False Then
            frmMain.mnuGraphicOndeAnnul.Enabled = False
        End If
        
        'Affectation à vrai de la réalisation d'une onde à sens privilégié
        'mais possible à cadrer à double sens
        'On mettra faux si le cadrage à double sens de l'onde à sens privilégié
        'est impossible
        .monOndeDoubleTrouve = True
        
        'Stockage d'une modification de valeurs dans les décalages
        'Ceci permettra aussi de demander une sauvegarde à la fermeture
        .maModifDataDec = True
        
        'Remise à FALSE des autres indicateurs pour pouvoir relancer un
        'calcul d'onde verte s'il repasse à TRUE
        .maModifDataCarf = False
        .maModifDataOndeTC = False
        .maModifDataOnde = False
        .maModifDataDes = False
    End With
    
    'Remise à zéro d'une translation précédente globale à tous les
    'carrefours
    uneForm.maTransDec = 0
    uneForm.TextTransDec.Text = Format(uneForm.maTransDec)
    
    'Réduction des carrefours
    If ReduireCarrefourSite(uneForm, uneForm.mesCarrefours, uneForm.monTypeOnde) Then
        'Cas où tous les carrefours ont pu être réduits
        '==> tous les feux équivalents ont pu être calculés
        
        'Initialisation des décalages à -99 des carrefours à décalages non imposés
        'Cas des carrefours non pris en compte dans le calcul (monIsUtil = False
        'ou Y carrefour pas entre Ymin et Ymax pour les ondes TC)
        'Seule possibilité d'avoir cette valeur qui reste à -99
        'car sinon les décalages sont entre 0 et la durée du cycle
        For i = 1 To uneForm.mesCarrefours.Count
            'Sauvegarde des décalages calculés et modifiés
            'avant un calcul avec décalage imposé
            uneForm.mesCarrefours(i).monDecCalculSave = uneForm.mesCarrefours(i).monDecCalcul
            uneForm.mesCarrefours(i).monDecModifSave = uneForm.mesCarrefours(i).monDecModif
        
        
            ' LCHAMMBON Correction
            If uneForm.mesCarrefours(i).monDecImp = 1 And Not uneForm.mesCarrefours(i).monIsUtil Then
                uneForm.mesCarrefours(i).monDecImp = 0
            End If
            
            If uneForm.mesCarrefours(i).monDecImp = 0 Then
                'Cas d'un carrefour à décalage non imposé
                'Initialisation à -99
                uneForm.mesCarrefours(i).monDecCalcul = -99
                uneForm.mesCarrefours(i).monDecModif = -99
            End If
        Next i
                                
        'Lancement du cas où des décalages sont imposés
        'unCarfRed est différent de nothing si le calcul
        'a eu lieu avec des dates imposées
        Set unCarfRed = UtiliserDecalagesImposes(unIndUniqCarfImp, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD)
        
        'Sauvegarde des bandes passantes avant leur modif
        uneBMsave = uneForm.maBandeM
        uneBDsave = uneForm.maBandeD
        uneBmodifMsave = uneForm.maBandeModifM
        uneBmodifDsave = uneForm.maBandeModifD
    
        'Calcul des bandes passantes maximales
        unResCalculBandes = CalculerBandesPassantesMaxi(uneForm, unB1, unB2, unH, Not (unCarfRed Is Nothing))
        Select Case unResCalculBandes
            Case AucuneSolution
                'Cas où aucune solution de bande passante n'a été trouvé
                MsgBox "Aucune solution d'onde verte à double sens n'a pu être trouvée", vbCritical
                If unCarfRed Is Nothing Then
                    'Cas d'un calcul sans aucun décalage imposé
                    'Le calcul n'a rien trouvé donc on signale l'impossibilité
                    'on ne pourra pas voir les dessins des plages de vert car
                    'les décalages sont inconnus
                    CalculerOndeVerte = False
                Else
                    'Cas d'un calcul avec des décalages imposés
                    'Le calcul n'a rien trouvé donc on signale l'impossibilité
                    'mais on trouve des bandes nulles et on garde les décalages
                    'saisis donnant cette impossibilité
                    'mais on peut voir le dessin des plages de vert car les
                    'décalages sont connus
                    CalculerOndeVerte = True
                End If
            Case DoubleSensPossible
                'Cas où les bandes passantes existent
                'avec une solution à double sens
                
                If Not (unCarfRed Is Nothing) Then
                    'Correction de la solution à date imposée
                    '(cf commentaires dans la fonction CorrectionDateImposée)
                    CorrectionDateImposée uneForm, unCarfRed, unB1, unB2, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD
                End If
                
                'Stockage et affichage des bandes passantes calculées
                StockerEtAfficherBandes uneForm, unB1, unB2
                'Calcul des décalages pour chaque carrefour
                CalculerDecalageDoubleSens uneForm, unB2, unH
            Case DoubleSensImpossible
                'Cas où les bandes passantes existent
                'mais pas de solution à double sens
                unTousCarfSensUniqueM = uneForm.mesCarfReduitsSens2.Count = 0 And uneForm.mesCarfReduitsSensM.Count > 0 And uneForm.mesCarfReduitsSensD.Count = 0
                unTousCarfSensUniqueD = uneForm.mesCarfReduitsSens2.Count = 0 And uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSensD.Count > 0
                If uneForm.monTypeOnde <> OndeTC And unTousCarfSensUniqueM = False And unTousCarfSensUniqueD = False Then
                    'Message affiché si on n'est pas en onde cadrée par TC et s'il
                    'n'y a pas que des carrefours à sens unique dans le même sens
                    MsgBox "Une solution pour le sens privilégié a été calculée, mais sans arriver à trouver une solution pour l'autre sens", vbInformation
                End If
                
                If Not (unCarfRed Is Nothing) Then
                    'Correction de la solution à date imposée
                    '(cf commentaires dans la fonction CorrectionDateImposée)
                    CorrectionDateImposée uneForm, unCarfRed, unB1, unB2, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD
                End If
                
                'Stockage et affichage des bandes passantes calculées
                StockerEtAfficherBandes uneForm, unB1, unB2
                'Calcul des décalages pour chaque carrefour
                CalculerDecalageSansDoubleSens uneForm
                'Cadrage à double sens de l'onde à sens privilégié impossible
                '==> On ne dessinera que l'onde dans le sens privilégié
                uneForm.monOndeDoubleTrouve = False
            Case Else
                CalculerOndeVerte = False
                MsgBox "Erreur de programmation dans OndeV dans CalculerOndeVerte", vbCritical
        End Select
    
        If Not (unCarfRed Is Nothing) And CalculerOndeVerte Then
            'Cas où le calcul a eu lieu avec des dates imposées
            'et il s'est bien passé
            
            'Récupération du décalage calculé du carrefour réduisant
            'tous les carrefours à date imposée
            If unIndUniqCarfImp = 0 Then
                'Cas avec plusieurs carrefours à décalage imposé
                unDecImp = unCarfRed.monCarrefour.monDecCalcul
            Else
                'Cas particulier d'un seul carrefour avec décalage imposé
                'Nouveau décalage moins l'ancien
                unDecImp = unCarfRed.monCarrefour.monDecCalcul - uneForm.mesCarrefours(unIndUniqCarfImp).monDecModif
            End If
            
            'Libération mémoire du carrefour de unCarfRed et de lui-même
            Set unCarfRed.monCarrefour = Nothing
            Set unCarfRed = Nothing
            
            If unResCalculBandes Then
                'Cas où des solutions avec des décalages imposés existent
            
                'Soustraction de ce décalage à tous les décalages des carrefours
                'sans date imposée et mise de ces décalages entre 0 et durée du cycle
                For i = 1 To uneForm.mesCarrefours.Count
                    Set unCarf = uneForm.mesCarrefours(i)
                    If unCarf.monDecImp = 0 And unCarf.monDecCalcul <> -99 Then
                        'Cas d'un carrefour à décalage non imposé
                        'et utilisé dans le calcul
                        unCarf.monDecCalcul = ModuloZeroCycle(unCarf.monDecCalcul - unDecImp, uneForm.maDuréeDeCycle)
                        unCarf.monDecModif = unCarf.monDecCalcul
                        If uneModifDec Then unCarf.monDecCalcul = unCarf.monDecCalculSave
                    ElseIf unCarf.monDecImp = 1 And unCarf.monDecCalcul <> -99 Then
                        'Cas d'un carrefour à décalage imposé
                        'et utilisé dans le calcul
                        If uneModifDec = False Then unCarf.monDecCalcul = unCarf.monDecModif
                    End If
                    'Si uneModifDec = true c'est que CalculerOndeVerte a été appelée
                    'après une modif manuelle dans l'onglet Tableau Décalage (cf procédure TabDecal_EditMode de frmDocument et RecalculerAvecDateImp dans ModuleCalculs)
                    'ou après une modif graphique (cf procédure MettreAJourSelection dans ModuleCalculs) à la souris d'un décalage d'un carrefour
                    'à décalage imposé ==> on remet la valeur du champ monDecCalcul de ce carrefour avant le calcul à date imposée
                Next i
                    
                'Restauration des bandes passantes avant leur modif si on a fait
                'une modif manuelle ou graphique d'un décalage
                If uneModifDec Then
                    uneForm.maBandeM = uneBMsave
                    uneForm.maBandeD = uneBDsave
                End If
            Else
                'Cas où aucune solution trouvée avec des décalages imposés
                
                'Restauration des décalages précédent le calcul
                For i = 1 To uneForm.mesCarrefours.Count
                    uneForm.mesCarrefours(i).monDecCalcul = uneForm.mesCarrefours(i).monDecCalculSave
                    uneForm.mesCarrefours(i).monDecModif = uneForm.mesCarrefours(i).monDecModifSave
                Next i
                'Restauration des bandes passantes précédent le calcul
                uneForm.maBandeM = uneBMsave
                uneForm.maBandeD = uneBDsave
                uneForm.maBandeModifM = unB1
                uneForm.maBandeModifD = unB2
            End If
            
            'Affichage dans l'onglet Tableau de résultat
            RemplirOngletTabDecalage uneForm
            
            'Remise à jour de la réduction des carrefours du site
            'et des temps de parcours pour le dessin des ondes
            ReduireCarrefourSite uneForm, uneForm.mesCarrefours, uneForm.monTypeOnde
            CalculerTempsParcours uneForm
        End If
    Else
        'Réduction de tous les carrefours impossible
        CalculerOndeVerte = False
    End If

    'Indication du niveau de cohérence entre les données
    'et le résultat du calcul d'onde verte
    If CalculerOndeVerte = False Then
        monSite.maCoherenceDataCalc = CalculImpossible
    Else
        monSite.maCoherenceDataCalc = OK
    End If
End Function

Public Function CalculerBandesPassantesMaxi(uneForm As Form, unB1 As Single, unB2 As Single, unH As Single, unCasDateImp As Boolean) As Integer
    'Fonction cherchant les bandes passantes maximales
    'dans les deux sens ou en privilégiant un sens
    
    Dim unCarfRedSens2 As CarfReduitSensDouble
    'Variables locales donnant les périodes de verte minimales
    'dans les sens montant et descendant
    Dim unMinVertSensM As Single
    Dim unMinVertSensD As Single
    
    Dim unS As Single 'Correspond au S des spécifs
    Dim unMin As Single, unMinLoc As Single
    Dim unPMsurPD As Single
    Dim A1 As Single, A2 As Single, K As Single
    Dim B1 As Single, B2 As Single
    Dim unB1M As Single, unB1D As Single
    Dim unB2M As Single, unB2D As Single

    'Initialisation
    unMinVertSensM = uneForm.maDuréeDeCycle
    unMinVertSensD = uneForm.maDuréeDeCycle
    unNbCarfRedSens2 = uneForm.mesCarfReduitsSens2.Count
    
    'Calcul des temps de parcours dans chaque sens à chaque carrefour
    'en prenant comme origine le premier carrefour dans chaque sens
    'considéré et en ayant trier par ordonnée croissante les carrefours
    'réduits.
    'De plus dans CalculerTempsParcours on calcul les écarts
    'des carrefours réduits à double sens
    'Calcul temps de parcours si on n'est pas dans le cas date imposée
    If unCasDateImp = False Then CalculerTempsParcours uneForm
        
    '1ère condition sur l'onde verte
    'La bande passante <= à la plus petite période de
    'verte rencontrée dans le sens considéré
    '==> Calcul des périodes vertes minimun unMinVertSensM et unMinVertSensD
    
    'Recherche sur tous les carrefours réduits à sens unique montant
    'pour la durée de vert minimale dans ce sens
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        If uneForm.mesCarfReduitsSensM(i).maDureeVert < unMinVertSensM Then
            unMinVertSensM = uneForm.mesCarfReduitsSensM(i).maDureeVert
        End If
    Next i
    
    'Recherche sur tous les carrefours réduits à sens unique descendant
    'pour la durée de vert minimale dans ce sens
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        If uneForm.mesCarfReduitsSensD(i).maDureeVert < unMinVertSensD Then
            unMinVertSensD = uneForm.mesCarfReduitsSensD(i).maDureeVert
        End If
    Next i
    
    'Recherche sur tous les carrefours réduits à double sens
    'pour les durées de vert minimales dans chaque sens
    
    'Initialisation du unS qui est un maximun, pour le trouver dans
    'le code plus bas explicant la 2ème condition
    unS = -3 * uneForm.maDuréeDeCycle 'Car toutes les valeurs sont dans [0,durée du cycle[
    For i = 1 To unNbCarfRedSens2
        'Récupération du carrefour réduit i
        Set unCarfRedSens2 = uneForm.mesCarfReduitsSens2(i)
        '1ère condition sur l'onde verte
        If unCarfRedSens2.maDureeVertM < unMinVertSensM Then
            unMinVertSensM = unCarfRedSens2.maDureeVertM
        End If
        If unCarfRedSens2.maDureeVertD < unMinVertSensD Then
            unMinVertSensD = unCarfRedSens2.maDureeVertD
        End If
                
        '2ème condition sur l'onde verte s'il y a
        'des carrefours réduits à double sens
                
        'Calcul sur tous les carrefours réduits double sens de la fonction :
        'Min(Z) = Minimun(DureeVertSensM_Carf_i + DureeVertSensD_Carf_i - Ecart_Carf_i(Z) + Z)
        'pour tout i variant de 1 à nombre de carrefours réduits double sens
        'et Z variant de monEcart du premier carrefour réduit double sens à
        'monEcart du dernier carrefour réduit double sens
        'En même temps on cherche unS = Max de ces min et on stocke dans unH
        'le Z correspondant au Max
        unMin = 3 * uneForm.maDuréeDeCycle 'Car toutes les valeurs sont dans [0,durée du cycle[
        For j = 1 To unNbCarfRedSens2
            unMinLoc = uneForm.mesCarfReduitsSens2(j).maDureeVertM + uneForm.mesCarfReduitsSens2(j).maDureeVertD - Ecart(uneForm.mesCarfReduitsSens2(j).monEcart, unCarfRedSens2.monEcart, uneForm.maDuréeDeCycle) + unCarfRedSens2.monEcart
            If unMinLoc < unMin Then
                'Stockage du minimun
                unMin = unMinLoc
            End If
        Next j
        
        'Stockage du maximun des minimuns sur tous les
        'carrefours réduits et l'écart réalisant ce max
        If unMin > unS Then
            unS = unMin
            unH = unCarfRedSens2.monEcart
        End If
    Next i
    
    'Détermination des bandes passantes maximales
    If uneForm.mesCarfReduitsSens2.Count = 0 Then
        'Cas où tous les carrefours sont à sens unique
        CalculerBandesPassantesMaxi = DoubleSensImpossible
        If uneForm.mesCarfReduitsSensM.Count = 0 Then
            'Cas où tous les carrefours à sens unique descendant
            unB1 = 0
            unB2 = unMinVertSensD
        ElseIf uneForm.mesCarfReduitsSensD.Count = 0 Then
            'Cas où tous les carrefours à sens unique montant
            unB1 = unMinVertSensM
            unB2 = 0
        Else
            unB1 = unMinVertSensM
            unB2 = unMinVertSensD
            unH = 0 'Tous les H dans [0, Durée du cycle[ sont possibles
        End If
    Else
        'Cas où il y a des carrefours réduits à double sens
        
        If uneForm.monTypeOnde = OndeTC And (uneForm.monTCM > 0 Or uneForm.monTCD > 0) Then
            'Cas d'une onde cadrée par un TC
                
            'Coordonnées des points A et B segment sur lequel se trouve la
            ' solution (cf Dossier de programmation / Solution temps imposé)
            If unMinVertSensD < unS Then
                A1 = unS - unMinVertSensD
                A2 = unMinVertSensD
            Else
                A1 = 0
                A2 = unS
            End If
            If unMinVertSensM < unS Then
                B1 = unMinVertSensM
                B2 = unS - unMinVertSensM
            Else
                B1 = unS
                B2 = 0
            End If
            
            'Calcul des bandes passantes par temps imposés
            If unS <= 0 Then
                If uneForm.monTCM > 0 And uneForm.monTCD > 0 Then
                    CalculerBandesPassantesMaxi = AucuneSolution
                Else
                    If uneForm.monTCM > 0 Then
                        unB1 = unMinVertSensM
                        unB2 = 0
                    End If
                    If uneForm.monTCD > 0 Then
                        unB1 = 0
                        unB2 = unMinVertSensD
                    End If
                    CalculerBandesPassantesMaxi = DoubleSensImpossible
                End If
            ElseIf unS >= (unMinVertSensM + unMinVertSensD) Then
                unB1 = unMinVertSensM
                unB2 = unMinVertSensD
                CalculerBandesPassantesMaxi = DoubleSensPossible
            ElseIf uneForm.monTCM > 0 And uneForm.monTCD > 0 Then
                'Cas où le temps est imposé par un TC dans les deux sens
                '(cf Dossier programmation)
                unTimp = (unS - uneForm.maBandeTCD + uneForm.maBandeTCM) / 2
                K = (unTimp - A1) / (B1 - A1)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            ElseIf uneForm.monTCM > 0 Then
                'Cas où le temps est imposé par un TC dans le sens montant
                K = (uneForm.maBandeTCM - A1) / (B1 - A1)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            ElseIf uneForm.monTCD > 0 Then
                'Cas où le temps est imposé par un TC dans le sens desendant
                'Solution sens descendant
                K = (uneForm.maBandeTCD - A2) / (B2 - A2)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            Else
                MsgBox "ERREUR de programmation dans OndeV dans CalculerBandesPassantesMaxi", vbCritical
            End If
            
            'Sortie pour éviter de faire le code qui suit
            Exit Function
        End If
        
        'Cas d'une onde non cadrée par un TC
        If uneForm.monTypeOnde = OndeDouble Then
            'Cas de l'onde double
            CalculerBandesPassantesMaxi = DoubleSensPossible
            'La valeur ci-dessus sera modifiée uniquement
            'si aucune solution trouvée
            unPMsurPD = uneForm.monPoidsSensM / uneForm.monPoidsSensD
            If unS <= 0 Then
                'Cas sans solution
                CalculerBandesPassantesMaxi = AucuneSolution
            ElseIf unS >= unMinVertSensM + unMinVertSensD Then
                unB1 = unMinVertSensM
                unB2 = unMinVertSensD
            ElseIf unS / (1 + 1 / unPMsurPD) <= unMinVertSensM And unS / (1 + unPMsurPD) <= unMinVertSensD Then
                unB1 = unS / (1 + 1 / unPMsurPD)
                unB2 = unS / (1 + unPMsurPD)
            ElseIf unS / (1 + 1 / unPMsurPD) > unMinVertSensM Then
                unB1 = unMinVertSensM
                unB2 = unS - unMinVertSensM
            ElseIf unS / (1 + unPMsurPD) > unMinVertSensD Then
                unB1 = unS - unMinVertSensD
                unB2 = unMinVertSensD
            End If
        ElseIf uneForm.monTypeOnde = OndeSensM Then
            'Cas de l'onde à sens privilégié montant
            CalculerBandesPassantesMaxi = CalculerBandeSensPrivi(unS, unB1, unB2, unMinVertSensM, unMinVertSensD)
        ElseIf uneForm.monTypeOnde = OndeSensD Then
            'Cas de l'onde à sens privilégié descendant
            CalculerBandesPassantesMaxi = CalculerBandeSensPrivi(unS, unB2, unB1, unMinVertSensD, unMinVertSensM)
        Else
            'Cas d'une erreur de programmation
            MsgBox "Erreur dans le calcul des bandes passantes maximales", vbCritical
        End If
    End If
End Function

Public Function Ecart(unEi As Single, unH As Single, uneDureeCycle As Integer) As Single
    'Fonction correspondant à la fonction ECART_i(h) = Ei+Ai*Cycle des spécifs
    If unH > unEi Then
        Ecart = unEi + uneDureeCycle
    Else
        Ecart = unEi
    End If
End Function

Public Sub CalculerTempsParcours(uneForm As Form)
    'Calcul des temps de parcours cumulés de chaque carrefour dans le
    'sens montant (respectivement descendant) à partir du premier carrefour
    'dans ce sens montant (respectivement descendant), les carrefours ayant
    'été au préalable classés par ordre croissant grâce à la moyenne des
    'ordonnées des feux équivalents du carrefour réduit.
    
    'De plus, on en profite pour faire le Calcul des écarts de chaque
    'carrefour réduit à double sens.
    'L'écart est le temps s'écoulant entre les événements "passage au vert
    'dans le sens montant" et "fin du vert dns le sens descendant" après
    'projection sur une référence commune à l'ensemble des carrefours
    '(cf Dossier de programmation et spécifs)
    
    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unCarfTmp As CarfY
    Dim unCarfRedSensU As CarfReduitSensUnique
    Dim unCarfRedSens2 As CarfReduitSensDouble
    Dim unDecVitSensM As Single, unDecVitSensD As Single
    Dim unIndCarfM As Integer, unIndCarfD As Integer
    Dim uneOndeTCM As Boolean, uneOndeTCD As Boolean
    Dim unIndPhase As Integer, unTCM As TC, unTCD As TC
    
    'Détermination du type d'onde verte à calculer
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde verte cadrée par un TC dans le sens montant
        uneOndeTCM = True
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde verte non cadrée par un TC dans le sens montant
        uneOndeTCM = False
        Set unTCM = Nothing
    End If
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde verte cadrée par un TC dans le sens descendant
        uneOndeTCD = True
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde verte non cadrée par un TC dans le sens descendant
        uneOndeTCD = False
        Set unTCD = Nothing
    End If
    
    'Réorganisation du tableau de carrefours réduits avec leur ordonnée
    'par classement suivant les ordonnées croissantes
    'Algo choisi : Le tri insertion (récupérer sur Internet)
    'Il consiste à comparer successivement un élément
    'à tous les précédents et à décaler les éléments intermédiaires
    unNbTotal = UBound(monTabCarfY, 1)
    For j = 2 To unNbTotal
            uneFinBoucle = False
            Set unCarfTmp = monTabCarfY(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                If monTabCarfY(i).monY > unCarfTmp.monY Then
                    Set monTabCarfY(i + 1) = monTabCarfY(i)
                    i = i - 1
                Else
                    'Fin de boucle
                    uneFinBoucle = True
                End If
            Loop
            Set monTabCarfY(i + 1) = unCarfTmp
    Next j

    'Calcul des temps de parcours cumulés de chaque carrefour dans le
    'sens montant (respectivement descendant) à partir du premier carrefour
    'dans ce sens montant (respectivement descendant), les carrefours ayant
    'été réorganisé ci-dessus
    unIndCarfM = 0
    unIndCarfD = 0
    For i = 1 To UBound(monTabCarfY, 1)
        'Initialisation à 0 des décalages dues aux vitesses du carrefour
        monTabCarfY(i).monCarfReduit.monCarrefour.monDecVitSensM = 0
        monTabCarfY(i).monCarfReduit.monCarrefour.monDecVitSensD = 0
        
        If TypeOf monTabCarfY(i).monCarfReduit Is CarfReduitSensUnique Then
            'Cas où le carrefour réduit est à sens unique
            Set unCarfRedSensU = monTabCarfY(i).monCarfReduit
            If unCarfRedSensU.monSensMontant Then
                'Cas d'un carrefour à sens unique montant
                If uneOndeTCM Then
                    unDecVitSensM = unTCM.CalculerDecalCauseProgTC(unTCM.mesPhasesTMOnde, unCarfRedSensU.monOrdonnee, 1)
                Else
                    CalculerDecVitesse unCarfRedSensU, True, i, unIndCarfM, unDecVitSensM
                End If
                'Stockage du décalage cumulé au carrefour
                unCarfRedSensU.monCarrefour.monDecVitSensM = unDecVitSensM
            Else
                'Cas d'un carrefour à sens unique descendant
                If uneOndeTCD Then
                    '-1 pour inversion du signe du Y pour le cas descendant
                    unDecVitSensD = unTCD.CalculerDecalCauseProgTC(unTCD.mesPhasesTMOnde, unCarfRedSensU.monOrdonnee, -1)
                Else
                    CalculerDecVitesse unCarfRedSensU, False, i, unIndCarfD, unDecVitSensD
                End If
                'Stockage du décalage cumulé au carrefour
                unCarfRedSensU.monCarrefour.monDecVitSensD = unDecVitSensD
           End If
        Else
            'Cas où le carrefour réduit est à double sens
            Set unCarfRedSens2 = monTabCarfY(i).monCarfReduit
            
            'Calcul du décalage cumulé au carrefour dans le sens montant
            If uneOndeTCM Then
                unDecVitSensM = unTCM.CalculerDecalCauseProgTC(unTCM.mesPhasesTMOnde, unCarfRedSens2.monOrdonneeM, 1)
            Else
                CalculerDecVitesse unCarfRedSens2, True, i, unIndCarfM, unDecVitSensM
            End If
            unCarfRedSens2.monCarrefour.monDecVitSensM = unDecVitSensM
            
            'Calcul du décalage cumulé au carrefour dans le sens descendant
            If uneOndeTCD Then
                '-1 pour Inversion du signe des Y pour le cas descendant
                unDecVitSensD = unTCD.CalculerDecalCauseProgTC(unTCD.mesPhasesTMOnde, unCarfRedSens2.monOrdonneeD, -1)
            Else
                CalculerDecVitesse unCarfRedSens2, False, i, unIndCarfD, unDecVitSensD
            End If
            unCarfRedSens2.monCarrefour.monDecVitSensD = unDecVitSensD
            
            'Calcul des écarts de chaque carrefour réduit à double sens
            'l'écart est le temps s'écoulant entre les événements "passage au vert
            'dans le sens montant" et "fin du vert dns le sens descendant" après
            'projection sur une référence commune à l'ensemble des carrefours
            '(cf Dossier de programmation et spécifs)
            'On utilise des décalages dus aux vitesses variables ou
            'constantes de chaque carrefour
            unCarfRedSens2.monEcart = unCarfRedSens2.maPosRefD + unCarfRedSens2.monCarrefour.monDecVitSensD + unCarfRedSens2.maDureeVertD
            unCarfRedSens2.monEcart = unCarfRedSens2.monEcart - (unCarfRedSens2.maPosRefM - unCarfRedSens2.monCarrefour.monDecVitSensM)
            'On ramène l'écart modulo entre [0, duréee du cycle[
            unCarfRedSens2.monEcart = ModuloZeroCycle(unCarfRedSens2.monEcart, uneForm.maDuréeDeCycle)
        End If
    Next i
End Sub

Public Sub AjouterCarfY(unIndex As Integer, unCarfReduit As Object)
    'Alimentation du tableau des carrefours réduits avec ordonnée
    'Cette ordonnée est calculée dans cette procédure et elle correspond
    'à la moyenne des ordonnées des feux équivalents du carrefour réduit
    Dim unCarfY As New CarfY
    
    Set unCarfY.monCarfReduit = unCarfReduit
    'Calcul du Y
    If TypeOf unCarfReduit Is CarfReduitSensUnique Then
            'Cas où le carrefour réduit est à sens unique
            unCarfY.monY = unCarfReduit.monOrdonnee
    Else
            'Cas où le carrefour réduit est à double sens
            unCarfY.monY = (unCarfReduit.monOrdonneeM + unCarfReduit.monOrdonneeD) / 2
    End If
    'Ajout dans le tableau des carrefours réduits avec ordonnée
    Set monTabCarfY(unIndex) = unCarfY
    'Stockage dans le carrefour non réduit de son carrefour réduit
    Set unCarfReduit.monCarrefour.monCarfRed = unCarfReduit
End Sub

Public Sub CalculerDecVitesse(unCarfReduit As Object, unSensMontant As Boolean, unIndCarf As Integer, unIndCarfPred As Integer, unDecVitesse As Single)
    'Procedure appelé par CalculerTempsParcours
    'Elle calcule dans un sens considéré, le décalage due aux vitesses
    'variable entre deux carrefours de même sens
    'Le décalage unDecVitesse et l'indice du dernier carrefour dans le même
    'sens sont modifiés pour être utilisés au prochain appel de cette procédure
       
    If unIndCarfPred = 0 Then
        'Cas du premier carrefour dans le sens considéré
        '==> son décalage est nul car il sert d'origine aux autres
        unDecVitesse = 0
    Else
        'Cumul du décalage en ajoutant le décalage due à la
        'vitesse variable entre les deux derniers carrefours
        'du sens considéré avec décalage > 0 en sens montant, < 0 sinon
        
        'Calcul de la distance entre les deux derniers carrefours dans
        'le sens considéré, celui donné par la valeur de unSensMontant
        uneDistance = unCarfReduit.DonnerYSens(unSensMontant) - monTabCarfY(unIndCarfPred).monCarfReduit.DonnerYSens(unSensMontant)
        
        'Calcul de la vitesse entre les deux derniers carrefours
        'On prend la vitesse variable du carrefour d'arrivée dans le sens considéré
        'Si sens montant, c'est la vitesse du carrefour donné par unCarfReduit
        'Si sens descendant, c'est la vitesse du carrefour précédent dans ce sens
        'c'est à dire le carrefour réduit d'indice unIndCarfPred
        If unSensMontant Then
            uneVitesse = unCarfReduit.DonnerVitSens(unSensMontant)
        Else
            uneVitesse = monTabCarfY(unIndCarfPred).monCarfReduit.DonnerVitSens(unSensMontant)
        End If
        'Explication du code ci-dessus : Par polymorphisme, les méthodes
        'DonnerVitSens et DonnerYSens sont appelées sur les bonnes
        'instances des classes CarfReduitSensUnique et CarfReduitSensDouble
        
        'Cumul du décalage entre les deux derniers carrefours
        'dans le sens considéré, celui donné par unSensMontant
        'Le décalage en temps est toujours > 0, on le multiplie ailleurs
        'dans le code par -1 pour le sens descendant et par 1 sinon
        unDecVitesse = unDecVitesse + Abs(uneDistance / uneVitesse)
    End If
    
    'Stockage du dernier carrefour rencontré dans le sens considéré
    unIndCarfPred = unIndCarf
    
End Sub

Public Function CalculerBandeSensPrivi(unS As Single, unB1 As Single, unB2 As Single, unMinVert1 As Single, unMinVert2 As Single) As Integer
    'Calcul de la bande passante dans le sens privilégié qui est le sens 1
    If unS <= 0 Then
        'Cas sans solution à double sens
        CalculerBandeSensPrivi = DoubleSensImpossible
        unB1 = unMinVert1
    ElseIf unS >= unMinVert1 + unMinVert2 Then
        'Cas avec solution à double sens
        CalculerBandeSensPrivi = DoubleSensPossible
        unB1 = unMinVert1
        unB2 = unMinVert2
    ElseIf unS <= unMinVert1 Then
        'Cas sans solution à double sens
        CalculerBandeSensPrivi = DoubleSensImpossible
        unB1 = unMinVert1
    ElseIf unS > unMinVert1 Then
        'Cas avec solution à double sens
        CalculerBandeSensPrivi = DoubleSensPossible
        unB1 = unMinVert1
        unB2 = unS - unMinVert1
    End If
End Function

Public Sub StockerEtAfficherBandes(uneForm As Form, unB1 As Single, unB2 As Single, Optional unDecalModif As Boolean = False)
    'Stockage et affichage des bandes passantes calculées
    'si undecalModif est vrai on ne stocke et n'affiche que
    'les bandes modifiables
    
    'Arrondi au deuxième chiffre après la virgule
    unB1 = Format(unB1, "Fixed")
    unB2 = Format(unB2, "Fixed")
    
    With uneForm
        'Stockage des largeurs de bandes
        If Not unDecalModif Then
            .maBandeM = unB1
            .maBandeD = unB2
        End If
        .maBandeModifM = unB1
        .maBandeModifD = unB2
        'Affichage des largeurs de bandes
        If Not unDecalModif Then
            .TabBande.Row = 1
            .TabBande.Col = 1
            .TabBande.Text = Format(.maBandeM)
            .TabBande.Row = 2
            .TabBande.Col = 1
            .TabBande.Text = Format(.maBandeD)
        End If
        .TabBande.Row = 1
        .TabBande.Col = 2
        .TabBande.Text = Format(.maBandeModifM)
        .TabBande.Row = 2
        .TabBande.Col = 2
        .TabBande.Text = Format(.maBandeModifD)
    End With
End Sub

Public Sub CalculerDecalageDoubleSens(uneForm As Form, unB2 As Single, unH As Single)
    'Calcul des décalages des carrefours si une solution à double sens
    'pour les bandes passantes a été trouvée
    Dim unCarfRed As Object
    Dim unK1 As Single, unK2 As Single
    
    'Parcours des carrefours  réduits à double sens
    For i = 1 To uneForm.mesCarfReduitsSens2.Count
        'On prend pour chaque carrefour le unK1 suivant (cf Dossier programmation)
        'Min (0, Durée de vert sens Desc - unB2 + unH - Ecart du carrefour - Ai * durée du cycle)
        'avec Ai =1 si unH > Ecart du carrefour, 0 sinon
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        unK1 = unCarfRed.maDureeVertD - unB2 + unH - unCarfRed.monEcart
        If unH > unCarfRed.monEcart Then
            unK1 = unK1 - uneForm.maDuréeDeCycle
        End If
        If unK1 > 0 Then unK1 = 0
        'Calcul et affichage des décalages calculés et modifiables
        'ramenés modulo durée du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRefM, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours réduits à sens montant
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Valeurs possibles de unK1 = tout l'intervalle
        '[Largeur bande sens montante - Durée de vert sens Montant, 0]
        'On prend unK1 = 0
        unK1 = 0
        'Calcul et affichage des décalages calculés et modifiables
        'ramenés modulo durée du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRef, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours réduits à sens descendant
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Valeurs possibles de unK2 = tout l'intervalle
        '[unH - Durée de vert sens descendant, unH - Largeur bande sens descendante]
        'On prend unK2 = unH - Largeur bande sens descendante
        unK2 = unH - unB2
        'Calcul et affichage des décalages calculés et modifiables
        'ramenés modulo durée du cycle
        CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRef, -unCarfRed.monCarrefour.monDecVitSensD
    Next i
End Sub

Public Sub CalculerDecalageSansDoubleSens(uneForm As Form)
    'Calcul des décalages des carrefours si aucune solution à double sens
    'pour les bandes passantes n'a été trouvée
    'Ceci ne se produit que pour une onde verte à sens privilégié
    Dim unCarfRed As Object
    Dim unK1 As Single, unK2 As Single
    
    'Parcours des carrefours  réduits à double sens
    'Ils ont une ligne de feux dans le sens privilégié
    For i = 1 To uneForm.mesCarfReduitsSens2.Count
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        If monSite.monTypeOnde = OndeSensM Or (monSite.monTypeOnde = OndeTC And monSite.monTCM > 0) Then
            'Cas d'une onde à sens M privi mais incadrable dans le sens D
            'Les valeurs possibles pour k1 = tout l'intervalle
            '[Durée de vert sens Montant - unB1, 0] (cf Dossier programmation)
            'On prend pour chaque carrefour unK1 = 0
            unK1 = 0
            'Calcul et affichage des décalages calculés et modifiables
            'ramenés modulo durée du cycle
            CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRefM, unCarfRed.monCarrefour.monDecVitSensM
        ElseIf monSite.monTypeOnde = OndeSensD Or (monSite.monTypeOnde = OndeTC And monSite.monTCD > 0) Then
            'Cas d'une onde à sens D privi mais incadrable dans le sens M
            'Les valeurs possibles pour k2 = tout l'intervalle
            '[Durée de vert sens Descendant - unB2, 0] (cf Dossier programmation)
            'On prend pour chaque carrefour unK2 = 0
            unK2 = 0
            'Calcul et affichage des décalages calculés et modifiables
            'ramenés modulo durée du cycle
            CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRefD, -unCarfRed.monCarrefour.monDecVitSensD
        Else
            MsgBox "Erreur de programmation dans OndeV dans CalculerDecalageSansDoubleSens", vbCritical
        End If
    Next i
    
    'Parcours des carrefours réduits à sens montant
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Cas où le sens privilégié est le montant
        '==> Valeurs possibles de unK1 = tout l'intervalle
        '[Largeur bande sens montante - Durée de vert sens Montant, 0]
        'Si le sens privilégié est le descendant k1 = 0 marche aussi
        'don on prend unK1 = 0 (cf Dossier programmation)
        unK1 = 0
        'Calcul et affichage des décalages calculés et modifiables
        'ramenés modulo durée du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRef, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours réduits à sens descendant
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Cas où le sens privilégié est le descendant
        '==> Valeurs possibles de unK2 = tout l'intervalle
        '[Largeur bande sens descendante - Durée de vert sens descendant, 0]
        'Si le sens privilégié est le montant k2 = 0 marche aussi
        'donc on prend unK2 = 0 (cf Dossier programmation)
        unK2 = 0
        'Calcul et affichage des décalages calculés et modifiables
        'ramenés modulo durée du cycle
        CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRef, -unCarfRed.monCarrefour.monDecVitSensD
    Next i
End Sub


Public Sub CalculerDecalage(uneForm As Form, unCarfRed As Object, unK As Single, unePosRef As Single, unDecVit As Single)
    'Calcul et Affichage des décalages calculés et modifiables du
    'carrefour obetenu grâce à son carrefour réduit
    
    'Calcul des décalages calculés et modifiables
    unCarfRed.monCarrefour.monDecCalcul = unK - unePosRef + unDecVit
    unCarfRed.monCarrefour.monDecCalcul = ModuloZeroCycle(unCarfRed.monCarrefour.monDecCalcul, uneForm.maDuréeDeCycle)
    unCarfRed.monCarrefour.monDecModif = unCarfRed.monCarrefour.monDecCalcul
    
    'Si le carrefour réduit est celui créé pour les dates imposées où n'
    'affiche pas sa valeur
    If unCarfRed.monCarrefour.maPosition <= 0 Then Exit Sub
    
    'Affichage dans l'onglet Tableau de résultat en arrondissant à l'entier
    'le plus proche, d'où l'utilisation de la fonction VB5 CInt
    uneForm.TabDecal.Row = unCarfRed.monCarrefour.maPosition
    uneForm.TabDecal.Col = 2
    If CIntCorrigé(unCarfRed.monCarrefour.monDecCalcul) = uneForm.maDuréeDeCycle Then
        'Une valeur valant durée du cycle s'affiche 0
        uneForm.TabDecal.Text = "0"
    Else
        uneForm.TabDecal.Text = CIntCorrigé(unCarfRed.monCarrefour.monDecCalcul)
    End If
    uneForm.TabDecal.Col = 3
    If CIntCorrigé(unCarfRed.monCarrefour.monDecModif) = uneForm.maDuréeDeCycle Then
        'Une valeur valant durée du cycle s'affiche 0
        uneForm.TabDecal.Text = "0"
    Else
        uneForm.TabDecal.Text = CIntCorrigé(unCarfRed.monCarrefour.monDecModif)
    End If
End Sub

Public Function RecalculerBandesPassantes(uneForm As Form) As Boolean
    'Recalcul des bandes passantes du site donné par uneForm
    'aprés la modification d'un décalage dans l'onglet Tableau
    'des résultats.
    'Retour : VRAI si recalcul a été possible, FAUX sinon
    
    'Création d'un nouveau carrefour qui contiendra tous les feux
    'équivalents montant et descendant des carrefours réduits dont on
    'cherchera le feu équivalent montant et descendant, les nouvelles bandes
    'passantes correspondront aux plages de vert maximales trouvées
    Dim unCarf As Carrefour
    Dim uneColCarf As New ColCarrefour
    Dim unB1 As Single, unB2 As Single
    Dim unDebVertM As Single, unDebVertD As Single
       
    'Création d'un nouveau carrefour qui contiendra tous les feux
    'équivalents montant et descendant des carrefours réduits
    'avec des vitesses non nulles.
    Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
    
    'Réduction des carrefours réduits
    unResultat = ReduireCarfReduits(uneForm, unCarf, unB1, unB2, unDebVertM, unDebVertD)
    If unResultat >= 0 Then
        '> 0 pour les cas où il y a réduction réussi et 0 sinon
        'avec >=0 on permet l'affichage et le dessin des bandes communes
        'nulles et de voir les bandes inter-carrefours (demande sites pilotes)
        RecalculerBandesPassantes = True
    Else
        RecalculerBandesPassantes = False
    End If
    
    If RecalculerBandesPassantes Then
        'Cas de réussite de la réduction des carrefours réduits
        'Stocker et afficher les nouvelles bandes passantes
        StockerEtAfficherBandes uneForm, unB1, unB2, True
    End If
    
    'Suppression de la collection ne contenant que le carrefour créé au début
    'pour libérer la mémoire.
    '(les events Terminate sont déclenchés sur les classes ColCarrefour,
    'Carrefour et ColFeu)
    uneColCarf.Remove 1
    Set unCarf = Nothing
    Set uneColCarf = Nothing
End Function
    
Public Function ReduireCarfReduits(uneForm As Form, unCarf As Carrefour, uneDureeVertM As Single, uneDureeVertD As Single, unDebVertM As Single, unDebVertD As Single) As Integer
    'Réduction du carrefour qui contient tous les feux
    'équivalents montant et descendant des carrefours réduits
    
    'Elle retourne un entier valant :
    '   - 0 si aucun feu équivalent trouvé
    '   - 1 si un feu équivalent trouvé (montant ou descendant)
    '   - 2 si deux feux équivalents trouvés (montant et descendant)
    
    Dim unFeu As Feu, unTC As TC
    Dim unCarfRed As Object
    Dim uneOrdonnee As Integer, unPosRef As Single
    Dim unNbFeuxSensM As Integer, unNbFeuxSensD As Integer
    Dim unNbFeuxSens2 As Integer
    
    'Ajout à ce carrefour global des feux équivalents
    'des carrefours réduits double sens
    unNbFeuxSens2 = uneForm.mesCarfReduitsSens2.Count
    For i = 1 To unNbFeuxSens2
        'Récup du carrefour réduit
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        'Ajout d'un nouveau feu montant
        Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonneeM, unCarfRed.maDureeVertM, unCarfRed.maPosRefM)
        'Stockage du carrefour du feu créé
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
        'Ajout d'un nouveau feu descendant
        Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonneeD, unCarfRed.maDureeVertD, unCarfRed.maPosRefD)
        'Stockage dans le feu créé du carrefour correspondant à celui réduit
        'car c'est son décalage en temps du à sa vitesse qui est utilisé dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
    
    'Ajout à ce carrefour global des feux équivalents
    'des carrefours réduits à sens unique montant
    unNbFeuxSensM = uneForm.mesCarfReduitsSensM.Count
    For i = 1 To unNbFeuxSensM
        'Récup du carrefour réduit
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Ajout d'un nouveau feu montant
        Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
        'Stockage dans le feu créé du carrefour correspondant à celui réduit
        'car c'est son décalage en temps du à sa vitesse qui est utilisé dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
    
    'Ajout à ce carrefour global des feux équivalents
    'des carrefours réduits à sens unique descendant
    unNbFeuxSensD = uneForm.mesCarfReduitsSensD.Count
    For i = 1 To unNbFeuxSensD
        'Récup du carrefour réduit
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Ajout d'un nouveau feu descendant
        Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
        'Stockage dans le feu créé du carrefour correspondant à celui réduit
        'car c'est son décalage en temps du à sa vitesse qui est utilisé dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
      
    'Initialisation de la valeur de retour de cette fonction
    ReduireCarfReduits = 0
    
    'Initialisation des résultats des feux équivalents montant et descendant
     unFeuEquivSensMExist = True
     unFeuEquivSensDExist = True
    
    'Calcul du feu équivalent montant éventuel
    If unNbFeuxSens2 > 0 Or unNbFeuxSensM > 0 Then
        If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
            'Cas d'une onde cadrée par unTC dans le sens montant
            Set unTC = monSite.mesTC(monSite.monTCM)
        Else
            'Cas d'une onde non cadrée par unTC dans le sens montant
            Set unTC = Nothing
        End If
        
        unFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVertM, unPosRef, uneOrdonnee, True, , , , unTC)
        If unFeuEquivSensMExist Then
            'Stockage du début de la plus grande plage de
            'vert trouvé lors du calcul du feu équivalent
            unDebVertM = monDebutVert
            'Calcul de la valeur de retour de cette fonction
            ReduireCarfReduits = ReduireCarfReduits + 1
        Else
            'Cas où aucune solution de bande passante montante n'a été trouvé
            unMsg = "Aucune solution de bande passante montante n'a pu être trouvée"
        End If
    End If
                      
    'Calcul du feu équivalent descendant éventuel
    If unNbFeuxSens2 > 0 Or unNbFeuxSensD > 0 Then
        If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
            'Cas d'une onde cadrée par unTC dans le sens descendant
            Set unTC = monSite.mesTC(monSite.monTCD)
        Else
            'Cas d'une onde non cadrée par unTC dans le sens descendant
            Set unTC = Nothing
        End If
        
        unFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unPosRef, uneOrdonnee, True, , , , unTC)
        If unFeuEquivSensDExist Then
            'Stockage du début de la plus grande plage de
            'vert trouvé lors du calcul du feu équivalent
            unDebVertD = monDebutVert
            'Calcul de la valeur de retour de cette fonction
            ReduireCarfReduits = ReduireCarfReduits + 1
        Else
            'Cas où aucune solution de bande passante descendante n'a été trouvé
            unMsg = unMsg + Chr(13) + "Aucune solution de bande passante descendante n'a pu être trouvée"
        End If
    End If
    
    If ReduireCarfReduits = 0 Then
        'Cas d'échec de la réduction des carrefours réduits
        'MsgBox unMsg, vbCritical
        'On n'affiche plus de message d'erreur ainsi on affichera et
        'stockera les bandes M et D même si elles sont nulles
    End If
    
    'Libération de la mémoire en supprimant tous les feux créés
    For j = 1 To unCarf.mesFeux.Count
    'Suppression du 1er dans une collection
    '==> Suppression de tout car le 2ème devient 1er, etc...
        unCarf.mesFeux.Remove 1
    Next j
End Function

Public Sub CalculerVitMax(unSite As Form)
    'Calcul et affichage dans l'onglet Fiche pour comparatif
    'des vitesses maximales possibles dans les deux sens
    
    Dim unTabFeuM() As Feu   'Tableau des feux montants
    Dim unTabFeuD() As Feu   'Tableau des feux descndants
    Dim unTabVM() As Single  'Tableau des vitesses montantes max possibles
    Dim unTabVD() As Single  'Tableau des vitesses descendantes max possibles
    Dim unTabIndFeuM() As Integer 'Tableau des indices des feux montants bas
    Dim unTabIndFeuD() As Integer 'Tableau des indices des feux descendants bas
    Dim unTabDYM() As Single 'Tableau des DY entre feux montants
    Dim unTabDYD() As Single 'Tableau des DY entre feux descendants
    Dim unTabDTM() As Single 'Tableau des DT entre feux montants
    Dim unTabDTD() As Single 'Tableau des DT entre feux descendants
    Dim unNbFeux As Integer, unNbFeuxM As Integer, unNbFeuxD As Integer
    Dim unNbCarf As Integer, unCarf As Carrefour
    Dim uneVitMaxLim As Single, uneVitMinLim As Single
    Dim uneVMax As Single
    
    'Initialisation
    unNbFeuxM = 0
    unNbFeuxD = 0
    uneVitMaxLim = 150 'Vitesse Maxi limite 150 km/h
    uneVitMinLim = 20  'Vitesse Mini limite 20 km/h
    
    'Valeurs par défaut avant réalisation de l'algo final
    unSite.maVitMaxM = "> 150"
    unSite.maVitMaxD = "< 20"

    'Remplissage des tableaux de feux montants et descendants
    unNbCarf = unSite.mesCarrefours.Count
    For i = 1 To unNbCarf
        Set unCarf = unSite.mesCarrefours(i)
        If unCarf.monDecCalcul <> -99 Then
            'Cas d'un carrefour pris en compte
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
            
    'Conversion des vitesses limites de km/h en m/s
    uneVitMaxLim = 150 / 3.6
    uneVitMinLim = 20 / 3.6
    
    'Cas de présence de feux montants
    If unNbFeuxM <> 0 Then
        'Réorganisation par ordonnée croissant des feux montants
        TrierFeuYCroissant unTabFeuM, 1
    
        'Vérification du passage à tous les verts montants
        'avec la vitesse maxi limite dans le sens montant
        uneVerif = VerifierVitessePasseToutVert(unSite, CInt(uneVitMaxLim * 3.6), unTabFeuM, True)
                              
        If uneVerif <> 0 Then
            'Cas où la vitesse maxi montante est > à la vitesse maxi limite
            unSite.maVitMaxM = "> " + Format(CInt(uneVitMaxLim * 3.6))
        Else
            'Cas où une vitesse max est < à la vitesse maxi limite
            
            'Calcul de la vitesse max montante possible en km/h
            '< à la vitesse maxi limite
            uneVMax = CalculerVMaxInfVMaxLim(unSite, uneVitMinLim, uneVitMaxLim, unTabFeuM, unTabVM, unTabIndFeuM, unTabDTM, unTabDYM, 1)
            
            'Stockage de la vitesse montante maxi trouvée
            If uneVMax < uneVitMinLim * 3.6 + 0.001 Then
                unSite.maVitMaxM = "< " + Format(CInt(uneVitMinLim * 3.6))
            Else
                unSite.maVitMaxM = Format(uneVMax)
            End If
        End If
    Else
        unSite.maVitMaxM = ""
    End If
    
    'Cas de présence de feux descendants
    If unNbFeuxD <> 0 Then
        'Réorganisation par ordonnée croissant des feux descendants
        TrierFeuYCroissant unTabFeuD, -1
    
        'Vérification du passage à tous les verts descendants
        'avec la vitesse maxi limite dans le sens descendant
        uneVerif = VerifierVitessePasseToutVert(unSite, CInt(uneVitMaxLim * 3.6), unTabFeuD, False)
                              
        If uneVerif <> 0 Then
            'Cas où la vitesse maxi descendante est > à la vitesse maxi limite
            unSite.maVitMaxD = "> " + Format(CInt(uneVitMaxLim * 3.6))
        Else
            'Cas où une vitesse max est < à la vitesse maxi limite
            
            'Calcul de la vitesse max descendante possible en km/h
            '< à la vitesse maxi limite remise en valeur positive
            uneVMax = CalculerVMaxInfVMaxLim(unSite, uneVitMinLim, uneVitMaxLim, unTabFeuD, unTabVD, unTabIndFeuD, unTabDTD, unTabDYD, -1)
            
            'Stockage de la vitesse descendante maxi trouvée
            If uneVMax < uneVitMinLim * 3.6 + 0.001 Then
                unSite.maVitMaxD = "< " + Format(CInt(uneVitMinLim * 3.6))
            Else
                unSite.maVitMaxD = Format(uneVMax)
            End If
        End If
    Else
        unSite.maVitMaxD = ""
    End If
End Sub

Public Sub RemplirOngletFicheResult(unSite As Form)
    'Remplissage de l'onglet Fiche pour comparatif
    Dim unTC As TC, unSens As Integer
    Dim unCarf As Carrefour
    Dim unNbFeuxM As Integer, unNbFeuxD As Integer
    
    With unSite
        'Affectation d'une couleur pour les cellules lockées
        'Cette onglet n'est qu'une édition ==> pas de saisie
        .TabFicheOnde.LockBackColor = .LabelTrait.BackColor
        .TabFicheCarf.LockBackColor = .LabelTrait.BackColor
        .TabFicheTC.LockBackColor = .LabelTrait.BackColor
        'Remplissage de TabFicheOnde
        'Calcul des temps de parcours
        TrouverTempsParcoursEtCarrefours unIndCarfM, unIndCarfD, unTmpM, unTmpD
        'Affichage des temps de parcours
        .TabFicheOnde.Col = 1
        .TabFicheOnde.Row = 1
        .TabFicheOnde.Text = unTmpM
        .TabFicheOnde.Row = 2
        .TabFicheOnde.Text = unTmpD
        'Affichage des bandes passantes
        .TabFicheOnde.Col = 2
        .TabFicheOnde.Row = 1
        .TabFicheOnde.Text = Format(.maBandeModifM)
        .TabFicheOnde.Row = 2
        .TabFicheOnde.Text = Format(.maBandeModifD)
        'Affichage des vitesses maximales possibles
        .TabFicheOnde.Col = 3
        .TabFicheOnde.Row = 1
        If .maVitMaxM = "" Or .maVitMaxD = "" Then
            CalculerVitMax monSite
        End If
        .TabFicheOnde.Text = .maVitMaxM
        .TabFicheOnde.Row = 2
        .TabFicheOnde.Text = .maVitMaxD
        'Affichage des poids
        .TabFicheOnde.Col = 4
        .TabFicheOnde.Row = 1
        If .monTypeOnde = OndeDouble Then
            .TabFicheOnde.Text = Format(.monPoidsSensM)
        Else
            .TabFicheOnde.Text = "Aucun"
        End If
        .TabFicheOnde.Row = 2
        If .monTypeOnde = OndeDouble Then
            .TabFicheOnde.Text = Format(.monPoidsSensD)
        Else
            .TabFicheOnde.Text = "Aucun"
        End If
        'Affichage des TC pris en compte dans l'onde TC
        .TabFicheOnde.Col = 5
        If .monTypeOnde = OndeTC Then
            .TabFicheOnde.Row = 1
            If .monTCM = 0 Then
                .TabFicheOnde.Text = "Aucun"
            Else
                .TabFicheOnde.Text = .mesTC(.monTCM).monNom
            End If
            .TabFicheOnde.Row = 2
            If .monTCD = 0 Then
                .TabFicheOnde.Text = "Aucun"
            Else
                .TabFicheOnde.Text = .mesTC(.monTCD).monNom
            End If
        Else
            .TabFicheOnde.Row = 1
            .TabFicheOnde.Text = "Aucun"
            .TabFicheOnde.Row = 2
            .TabFicheOnde.Text = "Aucun"
        End If
        
        'Remplissage des résultats carrefours
        .TabFicheCarf.MaxRows = .mesCarrefours.Count
        For i = 1 To .mesCarrefours.Count
            Set unCarf = .mesCarrefours(i)
            .TabFicheCarf.Row = i
            .TabFicheCarf.Col = 1
            .TabFicheCarf.Text = unCarf.monNom
            If unCarf.monDecCalcul = -99 Then
                'Cas des carrefours inutilisés ou non compris entre
                'Ymin et Ymax d'une onde cadrée par un TC
                For j = 2 To 7
                    .TabFicheCarf.Col = j
                    .TabFicheCarf.Text = ""
                Next j
            Else
                'Affichage du décalage en arrondissant à l'entier le plus
                'proche, d'où l'utilisation de la fonction VB5 CInt
                .TabFicheCarf.Col = 2
                If CIntCorrigé(unCarf.monDecModif) = .maDuréeDeCycle Then
                    'Une valeur valant durée du cycle s'affiche 0
                    .TabFicheCarf.Text = "0"
                Else
                    .TabFicheCarf.Text = CIntCorrigé(unCarf.monDecModif)
                End If
                'Affichage en fonction du type de carrefour
                'réduit (double sens ou sens unique)
                If TypeOf unCarf.monCarfRed Is CarfReduitSensDouble Then
                    .TabFicheCarf.Col = 3
                    uneRCap = unCarf.monCarfRed.maDureeVertM / .maDuréeDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                    'Mise en rouge des réserves de capacité négatives
                    If Val(uneRCap) < 0 Then
                        .TabFicheCarf.ForeColor = RGB(255, 0, 0)
                    Else
                        .TabFicheCarf.ForeColor = .ForeColor
                    End If
                    If unCarf.maDemandeM = 0 Then
                        .TabFicheCarf.Text = "+infini"
                    Else
                        uneRCap = uneRCap / unCarf.maDemandeM * 100 'RCap en %
                        .TabFicheCarf.Text = Format(CInt(uneRCap))
                    End If
                    .TabFicheCarf.Col = 4
                    uneRCap = unCarf.monCarfRed.maDureeVertD / .maDuréeDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                    'Mise en rouge des réserves de capacité négatives
                    If Val(uneRCap) < 0 Then
                        .TabFicheCarf.ForeColor = RGB(255, 0, 0)
                    Else
                        .TabFicheCarf.ForeColor = .ForeColor
                    End If
                    If unCarf.maDemandeD = 0 Then
                        .TabFicheCarf.Text = "+infini"
                    Else
                        uneRCap = uneRCap / unCarf.maDemandeD * 100 'RCap en %
                        .TabFicheCarf.Text = Format(CInt(uneRCap))
                    End If
                    .TabFicheCarf.Col = 5
                    .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par défaut
                    .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(True) * 3.6)
                    .TabFicheCarf.Col = 6
                    .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(False) * -3.6)
                    'Affichage du Décalage à l'ouverture
                    'Il est indéterminé si plusieurs lignes de feux dans le
                    'même sens (Carrefour <> Carf réduit)==> Affichage "Indéfini"
                    unCarf.DonnerNbFeuxMetD unNbFeuxM, unNbFeuxD
                    .TabFicheCarf.Col = 7
                    If unNbFeuxM = 1 And unNbFeuxD = 1 Then
                        .TabFicheCarf.Text = CInt(unCarf.monCarfRed.maPosRefM - unCarf.monCarfRed.maPosRefD)
                    Else
                        .TabFicheCarf.Text = "Indéfini"
                    End If
                Else
                    If unCarf.monCarfRed.monSensMontant Then
                        'Cas d'un carrefour à sens unique montant
                        .TabFicheCarf.Col = 3
                        uneRCap = unCarf.monCarfRed.maDureeVert / .maDuréeDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                        'Mise en rouge des réserves de capacité négatives
                        If Val(uneRCap) < 0 Then
                            .TabFicheCarf.ForeColor = RGB(255, 0, 0)
                        Else
                            .TabFicheCarf.ForeColor = .ForeColor
                        End If
                        If unCarf.maDemandeM = 0 Then
                            .TabFicheCarf.Text = "+infini"
                        Else
                            uneRCap = uneRCap / unCarf.maDemandeM * 100 'RCap en %
                            .TabFicheCarf.Text = Format(CInt(uneRCap))
                        End If
                        .TabFicheCarf.Col = 4
                        .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par défaut
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 5
                        .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(True) * 3.6)
                        .TabFicheCarf.Col = 6
                        .TabFicheCarf.Text = ""
                    Else
                        'Cas d'un carrefour à sens unique descendant
                        .TabFicheCarf.Col = 3
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 4
                        uneRCap = unCarf.monCarfRed.maDureeVert / .maDuréeDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                        'Mise en rouge des réserves de capacité négatives
                        If Val(uneRCap) < 0 Then
                            .TabFicheCarf.ForeColor = RGB(255, 0, 0)
                        Else
                            .TabFicheCarf.ForeColor = .ForeColor
                        End If
                        If unCarf.maDemandeD = 0 Then
                            .TabFicheCarf.Text = "+infini"
                        Else
                            uneRCap = uneRCap / unCarf.maDemandeD * 100 'RCap en %
                            .TabFicheCarf.Text = Format(CInt(uneRCap))
                        End If
                        .TabFicheCarf.Col = 5
                        .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par défaut
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 6
                        .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(False) * -3.6)
                    End If
                    'Décalage à l'ouverture indéterminé ==> Affichage "Indéfini"
                    .TabFicheCarf.Col = 7
                    .TabFicheCarf.Text = "Indéfini"
                End If
            End If
        Next i
        
        'Remplissage des résultats des TC utilisés
        .TabFicheTC.MaxRows = .mesTCutil.Count
        For i = 1 To .mesTCutil.Count
            Set unTC = .mesTCutil(i)
            If .maModifDataTC Or .maModifDataOndeTC Then
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
            
            'Remplissage du spread TabFicheTC
            .TabFicheTC.Row = i
            .TabFicheTC.Col = 1
            .TabFicheTC.Text = unTC.monNom
            .TabFicheTC.Col = 2
            If unSens = 1 Then
                .TabFicheTC.Text = "Montant"
            ElseIf unSens = -1 Then
                .TabFicheTC.Text = "Descendant"
            Else
                MsgBox "ERREUR de programmation dans OndeV dans RemplirOngletFicheResult", vbCritical
            End If
            .TabFicheTC.Col = 3
            .TabFicheTC.Text = Format(unTC.monTDep)
            .TabFicheTC.Col = 4
            .TabFicheTC.Text = Format(unTC.monNbArretsFeux)
            .TabFicheTC.Col = 5
            .TabFicheTC.Text = Format(CInt(unTC.monTempsArretFeux))
            .TabFicheTC.Col = 6
            'Calcul du temps de parcours du TC
            unNbPhases = unTC.mesPhasesTMProg.Count
            unTmpPar = unTC.mesPhasesTMProg(unNbPhases).monTDeb + unTC.mesPhasesTMProg(unNbPhases).maDureePhase - unTC.mesPhasesTMProg(1).monTDeb
            'Calcul de la distance parcourue par le TC
            uneDistPar = unTC.mesPhasesTMProg(unNbPhases).monYDeb + unTC.mesPhasesTMProg(unNbPhases).maLongPhase - unTC.mesPhasesTMProg(1).monYDeb
            'Affichage du temps de parcours et de la vitesse moyenne du TC en km/h
            .TabFicheTC.Text = Format(CInt(unTmpPar))
            .TabFicheTC.Col = 7
            .TabFicheTC.Text = Format(CInt(uneDistPar / unTmpPar * 3.6))
        Next i
    End With
End Sub

Public Sub DessinerOndeVerte(uneZoneDessin As Object, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, Optional unDessinDansOnglet As Boolean = True)
    'Dessin de l'onde verte, des plages de vert et les points de référence
    'des feux des carrefours et les parcours des TC choisis
    'unX0 et unY0 sont les coordonnées du point bas gauche
    
    'Elle réalise aussi le dessin des progressions des TC si
    'l'englobant des progressions des TC modifie celui des plages
    'de vert des feux et des ondes vertes
    
    Dim unTC As TC, unTCM As TC, unTCD As TC
    Dim unCarfRed As Object
    Dim unCarf As Carrefour
    Dim unFeu As Feu, unYFeu As Integer
    Dim unT As Long, unDY As Long
    Dim unDebVert As Single, unFinVert As Single
    Dim unDebVertMod As Single, uneSortieImprimante As Boolean
    Dim unTmpDebVert As Single, unTmpFinVert As Single
    Dim unYMax As Long, unYMin As Long, unY As Long
    Dim uneColCarf As New ColCarrefour
    Dim uneDureeVertM As Single, uneDureeVertD As Single
    Dim unDebVertM As Single, unDebVertD As Single
    Dim unMaxDebVert As Single, unMinFinVert As Single
    Dim unMaxDebVertTmp As Single, unMinFinVertTmp As Single
    Dim unDebVertMPred As Single, unDebVertDPred As Single
    Dim unNoDessinOndeM As Boolean, unNoDessinOndeD As Boolean
    Dim unTM1 As Single, unTM2 As Single
    Dim unTD1 As Single, unTD2 As Single
    Dim unYM1 As Long, unYM2 As Long, unI As Long
    Dim unLastDebVertM As Single, unLastDebVertD As Single
    Dim unXMpred As Single, unXDpred As Single
    Dim i As Integer, j As Integer
    Dim unTPtRef As Long, unX As Single, unXf As Single
    Dim unePlageGraphic As PlageGraphic
    Dim unRefGraphic As RefGraphic
    Dim unTDep As Long, unTFin As Long
    Dim unePrecM As Single, unePrecD As Single
    Dim uneDateYext As Single, uneDateYdep As Single
    Dim unIndFeu As Integer, unIndPhase As Integer
    
    'Stockage pour savoir si on dessine à l'écran ou sur imprimante
    'uneSortieImprimante = vrai si dessin sur imprimante faux si dessin écran
    uneSortieImprimante = (TypeOf uneZoneDessin Is Printer)
    
    'On vide les collections contenant les éléments graphics des ondes
    If uneSortieImprimante = False Then
        ViderCollection monSite.maColPlageGraphicD
        ViderCollection monSite.maColPlageGraphicM
        ViderCollection monSite.maColRefGraphicD
        ViderCollection monSite.maColRefGraphicM
    End If
    
    'Détermination du type d'onde verte à dessiner
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde verte cadrée par un TC dans le sens montant
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde verte non cadrée par un TC dans le sens montant
        Set unTCM = Nothing
    End If
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde verte cadrée par un TC dans le sens descendant
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde verte non cadrée par un TC dans le sens descendant
        Set unTCD = Nothing
    End If
    
    With monSite
        'Détermination de la hauteur englobante du dessin = Temps englobant
        TrouverTempsParcoursEtCarrefours unIndCarfM, unIndCarfD, unTmpM, unTmpD
                
        'Dessin des ondes vertes montantes et descendantes
        
        'Création d'un nouveau carrefour qui contiendra tous les feux
        'équivalents montant et descendant des carrefours réduits
        'avec des vitesses non nulles.
        'Ainsi on obtient le début de vert montant et descendant de la plus
        'plage de vert dans ces sens qui sert pour le dessin
        '(cf Dossier programmation, Répresentation graphique)
        Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
        
        'Réduction des carrefours réduits pour avoir unDebVertM et unDebVertD
        unResultat = ReduireCarfReduits(monSite, unCarf, uneDureeVertM, uneDureeVertD, unDebVertM, unDebVertD)
        
        'Suppression pour libérer la mémoire.
        uneColCarf.Remove 1
        Set unCarf = Nothing
        Set uneColCarf = Nothing

        If unResultat >= 0 Then
            '> 0 ==> Cas de réussite de la réduction des carrefours réduits
            'donc Dessin des ondes communes possibles
            '=0 dessin des ondes inter-carrefours même si l'onde commune
            'est impossible, donc >=0 tous les cas de figures
            '> 0 remplacé par >=0 après demande sites pilotes pour le cas = 0
            unNbCarf = UBound(monTabCarfY, 1)
            
            'Recherche du carrefour le plus bas ayant un feu montant et
            'du plus bas ayant un feu descendant
            unIndCarfBasM = 0
            unIndCarfBasD = 0
            i = 0
            'On part du carrefour le plus bas pour minimiser la boucle
            Do
                i = i + 1
                Set unCarfRed = monTabCarfY(i).monCarfReduit
                If TypeOf unCarfRed Is CarfReduitSensDouble Then
                    If unIndCarfBasM = 0 Then
                        unIndCarfBasM = i
                    End If
                    If unIndCarfBasD = 0 Then
                        unIndCarfBasD = i
                    End If
                ElseIf TypeOf unCarfRed Is CarfReduitSensUnique Then
                    If unCarfRed.monSensMontant Then
                        'Cas d'un carrefour à sens unique montant
                        If unIndCarfBasM = 0 Then
                            unIndCarfBasM = i
                        End If
                    Else
                         'Cas d'un carrefour à sens unique descendant
                        If unIndCarfBasD = 0 Then
                            unIndCarfBasD = i
                        End If
                   End If
                End If
            Loop While (unIndCarfBasM = 0 Or unIndCarfBasD = 0) And i < unNbCarf
        
            'Détermination du premier point de l'onde verte montante
            '(cf Dossier programmation : Représentation graphique)
            'Ce point correspond au carrefour le plus bas ayant un feu
            'montant, indice=unIndCarfBasM calculé juste avant
            If unIndCarfBasM > 0 Then
                'Cas où un carrefour ayant un feu montant a été trouvé
                Set unCarfRedM1 = monTabCarfY(unIndCarfBasM).monCarfReduit
                Set unCarf = unCarfRedM1.monCarrefour
                'Calcul de la référence et du Y du carrefour réduit
                'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                unePosRef = unCarfRedM1.DonnerPosRefSens(True)
                unYM1 = unCarfRedM1.DonnerYSens(True)
                'Calcul du début de vert du carrefour réduit
                unDebVertM0 = unCarf.monDecModif + unePosRef + unCarf.monDecVitSensM
                If .maBandeModifM = 0 Then
                    'Cas d'impossibilité d'avoir une bande montante commune
                    unK = 0
                Else
                    'Cas d'existence d'une bande passante montante commune
                    unK = ModuloZeroCycle(unDebVertM - unDebVertM0, .maDuréeDeCycle)
                End If
                'Abscisse du premier point, donc un temps ramené modulo cycle
                unTM1 = ModuloZeroCycle(unCarf.monDecModif + unePosRef + unK, .maDuréeDeCycle)
            
                'Correction du TM1 s'il n'y a qu'un carrefour ayant des feux
                'montants, sinon le dessin d'onde verte est erroné la bande
                'passante montante ne relie pas les plages des verts montants
                If monSite.mesCarfReduitsSens2.Count + monSite.mesCarfReduitsSensM.Count = 1 Then
                    unTM1 = 0
                    unTmpM = monSite.maDuréeDeCycle 'pour éviter une valeur nulle
                End If
            End If
            
            'Détermination du premier point de l'onde verte descendante
            '(cf Dossier programmation : Représentation graphique)
            'Ce point correspond au carrefour le plus haut ayant un feu
            'descendant, indice=unIndCarfD calculé en début de cette fonction
            If unIndCarfD > 0 Then
                'Cas où un carrefour ayant un feu descendant a été trouvé
                Set unCarfRedD1 = monTabCarfY(unIndCarfD).monCarfReduit
                Set unCarf = unCarfRedD1.monCarrefour
                'Calcul de la référence et du Y du carrefour réduit
                'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                unePosRef = unCarfRedD1.DonnerPosRefSens(False)
                unYD1 = unCarfRedD1.DonnerYSens(False)
                'Calcul du début de vert du carrefour réduit
                unDebVertD0 = unCarf.monDecModif + unePosRef + unCarf.monDecVitSensD
                If .maBandeModifD = 0 Then
                    'Cas d'impossibilité d'avoir une bande descendante commune
                    unK = 0
                Else
                    'Cas d'existence d'une bande passante descendante commune
                    unK = ModuloZeroCycle(unDebVertD - unDebVertD0, .maDuréeDeCycle)
                End If
                'Abscisse du premier point, donc un temps ramené modulo cycle
                unTD1 = ModuloZeroCycle(unCarf.monDecModif + unePosRef + unK, .maDuréeDeCycle)
                
                'Correction du TD1 s'il n'y a qu'un carrefour ayant des feux
                'descendants, sinon le dessin d'onde verte est erroné la bande
                'passante descendante ne relie pas les plages des verts descendants
                If monSite.mesCarfReduitsSens2.Count + monSite.mesCarfReduitsSensD.Count = 1 Then
                    unTD1 = unDebVertD0
                    unTmpD = monSite.maDuréeDeCycle 'pour éviter une valeur nulle
                End If
            End If
            
            
            If unIndCarfM > 0 Then
                'Calcul du nombre entier de cycle parcouru par l'onde montante
                unNbCycleM = Int((unTmpM + unTM1) / .maDuréeDeCycle)
                'Calcul du fin de vert maximun des feux du carrefour montant le + haut
                Set unCarf = monTabCarfY(unIndCarfM).monCarfReduit.monCarrefour
                unMaxFinVertHaut = unNbCycleM * .maDuréeDeCycle + TrouverMaxFinVert(unCarf) Mod .maDuréeDeCycle
                If unTmpM + unTM1 > unMaxFinVertHaut + 0.001 Then unMaxFinVertHaut = unMaxFinVertHaut + .maDuréeDeCycle
            Else
                unMaxFinVertHaut = -.maDuréeDeCycle '==> Max le + petit
            End If
            
            'Calcul du début de vert minimun des feux du
            'carrefour descendant le + haut
            If unIndCarfD > 0 Then
                Set unCarf = monTabCarfY(unIndCarfD).monCarfReduit.monCarrefour
                unMinDebVertHaut = TrouverMinDebVert(unCarf)
                'Cadrage dans le cycle
                If unTD1 < unMinDebVertHaut - 0.001 Then unMinDebVertHaut = unMinDebVertHaut - .maDuréeDeCycle
                If unTD1 > TrouverMaxFinVert(unCarf) + 0.001 Then unMinDebVertHaut = unMinDebVertHaut + .maDuréeDeCycle
            Else
                unMinDebVertHaut = .maDuréeDeCycle '==> Min le + grand
            End If
            
            'Calcul du nombre entier de cycle parcouru par l'onde descendante
            unNbCycleD = Int((unTmpD + unTD1) / .maDuréeDeCycle)
            'Calcul du fin de vert maximun des feux du carrefour descendant le + bas
            If unIndCarfBasD > 0 Then
                Set unCarf = monTabCarfY(unIndCarfBasD).monCarfReduit.monCarrefour
                unMaxFinVerBas = unNbCycleD * .maDuréeDeCycle + TrouverMaxFinVert(unCarf) Mod .maDuréeDeCycle
                If unTmpD + unTD1 > unMaxFinVerBas + 0.001 Then unMaxFinVerBas = unMaxFinVerBas + .maDuréeDeCycle
            Else
                unMaxFinVerBas = -.maDuréeDeCycle  '==> Max le + petit
            End If
            
            'Calcul du début de vert minimun des feux du
            'carrefour montant le + bas
            If unIndCarfBasM > 0 Then
                Set unCarf = monTabCarfY(unIndCarfBasM).monCarfReduit.monCarrefour
                unMinDebVertBas = TrouverMinDebVert(unCarf)
                'Cadrage dans le cycle
                If unTM1 < unMinDebVertBas - 0.001 Then
                    unMinDebVertBas = unMinDebVertBas - .maDuréeDeCycle
                End If
                If unTM1 > TrouverMaxFinVert(unCarf) + 0.001 Then
                    unMinDebVertBas = unMinDebVertBas + .maDuréeDeCycle
                End If
            Else
                unMinDebVertBas = .maDuréeDeCycle '==> Min le + grand
            End If
            
            'Calcul du max et du min en temps englobant les ondes des 2 sens
            If unMinDebVertHaut < unMinDebVertBas Then
                unMinT = unMinDebVertHaut
            Else
                unMinT = unMinDebVertBas
            End If
            If unMaxFinVertHaut < unMaxFinVerBas Then
                unMaxT = unMaxFinVerBas
            Else
                unMaxT = unMaxFinVertHaut
            End If
            
            'Modification du max et du min en temps pour avoir des lignes
            'de rappels toutes les 10 secondes englobant ce max et ce min
            ModifierMaxTempsPourVisu unMaxT
            ModifierMinTempsPourVisu unMinT
            
            'Détermination de la hauteur englobante du dessin = Temps englobant
            unT = unMaxT - unMinT
            'Stockage pour utilisation ailleurs
            monSite.monTmpTotal = unT
            monSite.monTMin = unMinT
            
            'Détermination de l'écart en Y englobant le dessin
            If unDessinDansOnglet Then
                'Cas où l'on dessine l'onde verte dans
                'l'onglet Graphique onde verte
                unYMin = .monYMinFeu
                unYMax = .monYMaxFeu
            Else
                'Cas où l'on dessine l'onde verte dans
                'la fenêtre pleine écran ou sur imprimante
                'On ne prend en compte que les carrefours utilisés
                '==> niveau de zoom différent
                TrouverMinYMaxY unYMin, unYMax
                If unYMax = unYMin Then
                    'Cas d'un seul carrefour avec un seul
                    'Pour éviter unDY = unYMax - unYMin = 0
                    'On cadre 100 au dessus et au dessous
                    unYMax = unYMax + 100
                    unYMin = unYMin - 100
                End If
                .monYMaxFeuUtil = unYMax
                .monYMinFeuUtil = unYMin
            End If
            unDY = unYMax - unYMin
            'Stockage pour utilisation ailleurs
            monSite.monDYTotal = unDY
            monSite.monYMin = unYMin
                    
            'Test si l'englobant des progressions n'est pas compris
            'dans l'englobant de l'onde ==> changement d'englobant en temps
            'plus dessin des progressions des TC
            If monSite.mesTCutil.Count > 0 Then
                'Calcul de l'origine pour les temps
                unNewX0 = unX0 - ConvertirReelEnEcran(CLng(unMinT), unT, uneLg)
                'Stockage dans une variable privée de ce module
                'pour utilisation ailleurs
                monSite.monOrigX = unNewX0
                
                'Calcul de l'englobant en temps des progressions de TC
                'sélectionnés pour connaitre l'englobant en coordonnées écran
                unNbTCUtil = monSite.mesTCutil.Count
                monSite.monTDepTCMin = 1000000
                monSite.monTFinTCMax = -1000000
                For i = 1 To unNbTCUtil
                    Set unTC = monSite.mesTCutil(i)
                    DonnerEnglobantTC unTC, unX0, unY0, uneLg, uneHt, unTDep, unTFin
                    'Calcul de l'englobant en temps pour les progressions des TC
                    If unTDep < monSite.monTDepTCMin Then monSite.monTDepTCMin = unTDep
                    If unTFin > monSite.monTFinTCMax Then monSite.monTFinTCMax = unTFin
                Next i
                
                'Englobant en Temps total en valeur écran
                unTmpTotalEcran = uneLg
                If monSite.monTDepTCMin < unX0 Then
                    'Cas où l'englobant des progressions TC commence
                    'avant l'englobant des ondes
                    '==> Changement du début de l'englobant
                    'en convertissant valeur écran en réelle
                    unMinT = (monSite.monTDepTCMin - monSite.monOrigX) / unTmpTotalEcran * monSite.monTmpTotal
                    'Modification du min en temps pour avoir des lignes
                    'de rappels toutes les 10 secondes englobant ce min
                    ModifierMinTempsPourVisu unMinT
                    monSite.monTMin = unMinT
                    'Détermination de la nouvelle largeur englobante
                    'du dessin = Temps englobant
                    monSite.monTmpTotal = unMaxT - unMinT
                End If
                
                If unTmpTotalEcran + unX0 < monSite.monTFinTCMax Then
                    'Cas où l'englobant en temps des progressions TC est plus
                    'grand que l'englobant en temps des ondes
                    '==> Changement de l'englobant en temps
                    'en convertissant valeur écran en réelle
                    unMaxT = (monSite.monTFinTCMax - unX0) / unTmpTotalEcran * monSite.monTmpTotal + monSite.monTMin
                    'Modification du min en temps pour avoir des lignes
                    'de rappels toutes les 10 secondes englobant ce min
                    ModifierMaxTempsPourVisu unMaxT
                    'Détermination de la nouvelle largeur englobante
                    'du dessin = Temps englobant
                    monSite.monTmpTotal = unMaxT - unMinT
                End If
                
                'Stockage pour utilisation ailleurs
                unT = monSite.monTmpTotal
                
                'Calcul de l'origine pour les temps
                monSite.monOrigX = unX0 - ConvertirReelEnEcran(CLng(unMinT), unT, uneLg)
                
                'Dessin des progressions des TC pour quelles soient
                'affichées avant les plages de vert des feux, qu'ainsi
                'elles ne masqueront pas
                unNbTCUtil = monSite.mesTCutil.Count
                For i = 1 To unNbTCUtil
                    Set unTC = monSite.mesTCutil(i)
                    TracerProgressionTC uneZoneDessin, unTC, unX0, unY0, uneLg, uneHt
                Next i
            End If
            
            'Conversion de la durée du cycle en valeur écran (twips)
            'Cette variable locale est utilisée plus bas
            uneLongCycle = ConvertirReelEnEcran(.maDuréeDeCycle, unT, uneLg)
        
            'Calcul de l'origine pour les temps
            unNewX0 = unX0 - ConvertirReelEnEcran(CLng(unMinT), unT, uneLg)
            'Stockage dans une variable privée de ce module
            'pour utilisation ailleurs
            monSite.monOrigX = unNewX0
            
            'Conversion en coordonnées écran du point (unTD1, unYD1)
            'Sert de premier point à l'onde verte descendante
            unXDpred = ConvertirSingleEnEcran(unTD1, unT, uneLg)
            unXDpred = unXDpred + unNewX0
            unYDpred = ConvertirReelEnEcran(unYD1 - unYMin, unDY, uneHt)
            unYDpred = unY0 - unYDpred
            
            'Conversion en coordonnées écran du point (unTM1, unYM1)
            'Sert de premier point à l'onde verte montante
            unXMpred = ConvertirSingleEnEcran(unTM1, unT, uneLg)
            unXMpred = unXMpred + unNewX0
            unYMpred = ConvertirReelEnEcran(unYM1 - unYMin, unDY, uneHt)
            unYMpred = unY0 - unYMpred
            
            'Dessin des ondes vertes
            unMsgDessinOnde = ""
            'unNoDessinOndeM = (.monOndeDoubleTrouve = False And .monTypeOnde = OndeSensD) Or .maBandeModifM = 0
            unNoDessinOndeM = (.maBandeModifM = 0)
            If unNoDessinOndeM And monSite.mesCarfReduitsSensM.Count > 0 And uneSortieImprimante = False Then
                'Cas d'une onde à sens privilégié descendant mais ayant des feux
                'montants mais ne pouvant pas être cadrer dans le sens montant
                '==> Pas d'onde verte montante, d'où pas de dessin.
                unMsgDessinOnde = "Pas de dessin de l'onde verte MONTANTE car elle n'existe pas."
            End If
            
            'unNoDessinOndeD = (.monOndeDoubleTrouve = False And .monTypeOnde = OndeSensM) Or .maBandeModifD = 0
            unNoDessinOndeD = (.maBandeModifD = 0)
            If unNoDessinOndeD And monSite.mesCarfReduitsSensD.Count > 0 And uneSortieImprimante = False Then
                'Cas d'une onde à sens privilégié montant mais ayant des feux
                'descendants ne pouvant pas être cadrer dans le sens descendant
                '==> Pas d'onde verte descendant, d'où pas de dessin.
                If unMsgDessinOnde = "" Then
                    unMsgDessinOnde = "Pas de dessin de l'onde verte DESCENDANTE car elle n'existe pas."
                Else
                    unMsgDessinOnde = unMsgDessinOnde + Chr(13) + Chr(13) + "Pas de dessin de l'onde verte DESCENDANTE car elle n'existe pas."
                End If
            End If
            
            'Affichage du non dessin d'onde éventuelle
            If unMsgDessinOnde <> "" Then MsgBox unMsgDessinOnde, vbInformation
                        
            'Initialisation de variables donnant un numéro de phases
            'elles commencent à 1
            unIndPM% = 1
            unIndPD% = 1

            'Dessin des bandes inter-carrefours voitures si on est en onde TC
            'et que l'utilisateur a choisi l'item Montrer les bandes inter-carf
            'dans le menu contextuel de zone graphique
            If monSite.monDessinInterCarfVP And monSite.monTypeOnde = OndeTC Then
                DessinerBandesInterCarfVP uneZoneDessin, unTM1, unTD1, unY0, uneHt, unDY, unT, unCarfRedM1, unCarfRedD1, unYMin, unIndCarfBasM, unIndCarfD, uneLg, unNewX0, unMaxT
            End If
            
            For i = 1 To unNbCarf
                'Parcours des carrefours dans le sens des Y croissants
                'pour l'onde verte montante car on dessine à partir du
                'carrefour le plus bas ayant un feu montant
                Set unCarfRed = monTabCarfY(i).monCarfReduit
                Set unCarf = unCarfRed.monCarrefour
                'Conversion en valeur écran de la largeur de bande montante
                uneLBM = ConvertirSingleEnEcran(.maBandeModifM, unT, uneLg)
                If unIndCarfM > 0 Then
                    'Cas d'une onde verte montante possible ==> Dessin
                    'unIndCarfM > 0 dit qu'on a trouvé des carrefours montants
                    If unCarfRed.HasFeuMontant = True Then
                        'Cas d'un carrefour contraignant de l'onde verte montante
                        'donc ayant un feu de sens montant
                        
                        'Abscisse du point suivant de l'onde montante vaut
                        'l'abscisse du point précédent plus le décalage en
                        'temps entre le carrefour courant et le premier montant
                        unTM2 = unTM1 + unCarf.monDecVitSensM - unCarfRedM1.monCarrefour.monDecVitSensM
                        'Stockage dans début onde pour trouver les plages
                        'sélectionnables graphiquement
                        unCarfRed.AffecterDebOndeSens unTM2, True
                            
                        'Ordonnée égale à l'ordonnée du carrefour réduit courant
                        'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                        unYM2 = unCarfRed.DonnerYSens(True)
                        
                        'Conversion en coordonnées écran de unYM2
                        unY = ConvertirReelEnEcran(unYM2 - unYMin, unDY, uneHt)
                        unY = unY0 - unY
                        
                        'Dessin de l'onde verte montante inter-carrefours donc
                        'entre ce carrefour réduit et son précédent en Y
                        'si choix coché dans les options d'affichage et d'impression
                        'et si ce n'est pas le 1er carrefour montant le + bas
                        
                        'Fait avant la bande commune pour ne voir que la bande
                        'commune si superposition avec la bande inter-carrefours
                        If .mesOptionsAffImp.maVisuBandInterCarfM And i <> unIndCarfBasM Then
                            'Cas d'une onde montante pas cadrée par un TC montant
                            'Calcul du début de vert de ce carrefour réduit
                            unDebVertM = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(True)
                            unDebVertM = ModuloZeroCycle(unDebVertM, .maDuréeDeCycle)
                            'Calcul du nombre de cycle séparant le début de
                            'vert du début de l'onde verte montante
                            unNbCycle = Fix((0.001 + unTM2 - unDebVertM) / .maDuréeDeCycle)
                            If unNbCycle < 0 And .maBandeModifM = 0 Then
                                'Si pas de bande commune, on ne peut pas être en retard
                                'unTM2 ne doit pas être corrigé si il est < unDebVertM
                                unNbCycle = 0
                            End If
                            If unTM2 < unDebVertM - 0.001 Then
                                'Début de vert > T de onde montante
                                '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                                unDebVertM = unDebVertM + unNbCycle * .maDuréeDeCycle
                            ElseIf unTM2 > unDebVertM + unCarfRed.DonnerDureeVertSens(True) + 0.001 Then
                                'Fin de vert < T de départ onde montante
                                '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                                unDebVertM = unDebVertM + unNbCycle * .maDuréeDeCycle
                            End If
                                                      
                           'Calcul du temps de parcours inter-carrefours
                            unTmpInterCarf = unCarf.monDecVitSensM - monTabCarfY(unIndCarfMPred).monCarfReduit.monCarrefour.monDecVitSensM
                            'Calcul de la fin de vert de ce carrefour réduit
                            unFinVertM = unDebVertM + unCarfRed.DonnerDureeVertSens(True)
                            unFinVertMPred = unDebVertMPred + monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerDureeVertSens(True)

                            unI = 0
                            unJ = 0
                            unDebOndeDéjàStocké = False
                            unLastDebVertDéjàStocké = False
                            'projection du début et fin de vert du carrefour
                            'précédent sur la droite Y = Y du carrefour courant
                            unDebVertMPred = unDebVertMPred + unTmpInterCarf
                            unFinVertMPred = unFinVertMPred + unTmpInterCarf
                            'Initialisation du LastDebVert au cas où aucun
                            'bande inter-carf trouvée
                            unLastDebVertM = unDebVertM
                            Do
                                'La boucle sert pour afficher toutes les bandes
                                'inter-carrefour pour cela on regarde dans
                                'le cycle en cours et le suivant et on prend la
                                'bande inter-carf maximale
                                
                                'On prend le minimun des fins de vert projeté
                                'sur la droite Y = Y du carrefour courant
                                If (unFinVertMPred + unJ * .maDuréeDeCycle) < (unFinVertM + unI * .maDuréeDeCycle) Then
                                    unMinFinVert = unFinVertMPred + unJ * .maDuréeDeCycle
                                Else
                                    unMinFinVert = unFinVertM + unI * .maDuréeDeCycle
                                End If
                                
                                'On prend le maximun des débuts de vert projeté
                                'sur la droite Y = Y du carrefour courant
                                If (unDebVertMPred + unJ * .maDuréeDeCycle) > (unDebVertM + unI * .maDuréeDeCycle) Then
                                    unMaxDebVert = unDebVertMPred + unJ * .maDuréeDeCycle
                                Else
                                    unMaxDebVert = unDebVertM + unI * .maDuréeDeCycle
                                End If
                                
                                'Test de l'existence d'une bande inter-carrefour
                                'supérieure à 1 seconde
                                uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                                
                                If uneBandeInterCarfExist Then
                                    If unTM2 - 0.01 < unMinFinVert And unLastDebVertDéjàStocké = False Then
                                        'Stockage du dernier debvert ayant une bande
                                        'inter-carrefour, ce stockage est fait une fois et
                                        'une seule entre deux carrefours
                                        unLastDebVertDéjàStocké = True
                                        unLastDebVertM = unDebVertM + unI * .maDuréeDeCycle
                                    End If
                                    'Stockage dans début onde pour trouver les
                                    'plages sélectionnables graphiquement,
                                    'la 1ère fois seulement
                                    If unDebOndeDéjàStocké = False Then
                                        unDebOndeDéjàStocké = True
                                        unCarfRed.AffecterDebOndeSens (unMaxDebVert + unMinFinVert) / 2, True
                                    End If
                                    'Remise dans l'englobant total si
                                    'MinFinVert en sort
                                    If unMinFinVert > unMaxT + 0.01 Then
                                        unMinFinVert = unMinFinVert - .maDuréeDeCycle
                                        unMaxDebVert = unMaxDebVert - .maDuréeDeCycle
                                    End If
                                End If
                                
                                unI = unI + 1
                                If unI = 2 Then
                                    'On se place pour essayer les début et fin de vert
                                    'du carrefour courant dans le cycle courant avec les
                                    'début et fin de vert du carrefour précédent dans le
                                    'cycle suivant
                                    unI = 0
                                    unJ = 1
                                End If

                                'Dessin de bande inter-carrefour
                                If unTCM Is Nothing And uneBandeInterCarfExist Then
                                    'Cas d'une onde montante non cadrée par un TC ou du
                                    'dessin des bandes inter-carrefours voitures d'une onde TC
                                    
                                    'Conversion en coordonnées écran
                                    unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 1ère partie bande montante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                    
                                    'Conversion en coordonnées écran
                                    unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 2ème partie bande montante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                ElseIf uneBandeInterCarfExist Then
                                    'Cas d'une onde montante cadrée par un TC
                                    'Recup du Y du carf réduit précédent
                                    unYTmp! = monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True)
                                    'Calcul de la date en Y = unYTmp%
                                    unDecT0! = unTCM.CalculerDateDansTabMarche(unTCM.mesPhasesTMOnde, unYTmp!, unIndPM%, unIndPM%)
                                    'Calcul du décalage en date pour avoir la prog partielle du TC
                                    'qui commence à la date absolue unMaxDebVert - unTmpInterCarf
                                    unDecT! = unMaxDebVert - unTmpInterCarf - unDecT0!
                                    'Dessin 1ère ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True), unCarf.monCarfRed.DonnerYSens(True), unDecT!, unIndPM%
                                    'Le dernier paramètre unDecT! sert à se cadrer au départ de l'onde sinon le TC débute à son T départ
                                    
                                    'Calcul du décalage en date pour avoir la prog partielle du TC
                                    'qui commence à la date absolue unMinFinVert - unTmpInterCarf
                                    unDecT! = unMinFinVert - unTmpInterCarf - unDecT0!
                                    'Dessin 2ème ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True), unCarf.monCarfRed.DonnerYSens(True), unDecT!, unIndPM%
                                    'Le dernier paramètre unDecT! sert à se cadrer au départ de l'onde sinon le TC débute à son T départ
                                End If
                                'Boucle fait trois fois pour trouver toutes
                                'les bande inter-carrefour
                            Loop Until unI = 1 And unJ = 1
                        End If 'Fin du dessin de bande verte montante inter-carrefours
                        
                        'Dessin de la bande verte montante commune aux carrefours
                        'si choix coché dans les options d'affichage et d'impression
                        'et si l'onde montante n'est pas cadrée par un TC
                        '(dessin fait avant ce code dans ce cas, cf for i= 1 to unNbCarf)
                        If Not unNoDessinOndeM And .mesOptionsAffImp.maVisuBandComM And unTCM Is Nothing Then
                            'Stockage dans début onde pour trouver les plages
                            'sélectionnables graphiquement
                            unCarfRed.AffecterDebOndeSens unTM2, True
                            
                            'Conversion en coordonnées écran de unTM2
                            unX = ConvertirSingleEnEcran(unTM2, unT, uneLg)
                            unX = unX + unNewX0
                            
                            'Dessin en coordonnées écran de la ligne entre les
                            'points (unTM1, unYM1) et (unTM2, unYM2) et d'une
                            'ligne // à une largeur de bande montante
                            uneZoneDessin.Line (unXMpred, unYMpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComM
                            uneZoneDessin.Line (unXMpred + uneLBM, unYMpred)-(unX + uneLBM, unY), .mesOptionsAffImp.maCoulBandComM
                        
                            'Dessin de l'onde verte dans les feux du premier
                            'carrefour dans le sens montant, sinon le dessin
                            'ne commence qu'au feu de Y maximun (cf réduction carrefour)
                            'Ce dessin va de unYMin jusqu'à Max Y 1er carrefour montant
                            If i = unIndCarfBasM Then
                                'Conversion en coordonnées écran du 1er point
                                unYFeu = DonnerYMinCarfSens(unCarf, True, unIndFeu)
                                unXMpred = unTM1 - (unYM1 - unYFeu) / unCarfRed.DonnerVitSens(True)
                                unXMpred = ConvertirSingleEnEcran(unXMpred, unT, uneLg)
                                unXMpred = unXMpred + unNewX0
                                unYMpred = ConvertirReelEnEcran(CLng(unYFeu) - unYMin, unDY, uneHt)
                                unYMpred = unY0 - unYMpred
                                uneZoneDessin.Line (unXMpred, unYMpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComM
                                uneZoneDessin.Line (unXMpred + uneLBM, unYMpred)-(unX + uneLBM, unY), .mesOptionsAffImp.maCoulBandComM
                            End If
                            
                            'Stockage du X écran  du point précédent pour le coup suivant
                            unXMpred = unX
                        End If 'Fin du dessin de bande verte montante commune
                                 
                       'Stockage du début de vert précédent
                        If i = unIndCarfBasM Then
                            'Calcul spécial pour le carrefour le + bas montant
                            unDebVertMPred = unCarfRedM1.monCarrefour.monDecModif + unCarfRedM1.DonnerPosRefSens(True)
                            unDebVertMPred = ModuloZeroCycle(unDebVertMPred, .maDuréeDeCycle)
                            If unTM1 < unDebVertMPred - 0.001 Then
                                'Début de vert > T de départ onde montante
                                '==> Recul d'un cycle
                                unDebVertMPred = unDebVertMPred - .maDuréeDeCycle
                            ElseIf unTM1 > unDebVertMPred + unCarfRedM1.DonnerDureeVertSens(True) + 0.001 Then
                                'Fin de vert < T de départ onde montante
                                '==> Avancé d'un cycle
                                unDebVertMPred = unDebVertMPred + .maDuréeDeCycle
                            End If
                        Else
                            unDebVertMPred = unLastDebVertM
                        End If
                        
                        'Stockage de l'indice de ce carrefour
                        unIndCarfMPred = i
                        'Stockage du Y écran  du point précédent pour le coup suivant
                        unYMpred = unY
                    End If
                End If
                
                'Parcours des carrefours dans le sens des Y décroissants
                'pour l'onde verte descendante car on dessine à partir du
                'carrefour le plus haut ayant un feu descendant
                Set unCarfRed = monTabCarfY(unNbCarf + 1 - i).monCarfReduit
                Set unCarf = unCarfRed.monCarrefour
                'Conversion en valeur écran de la largeur de bande descendante
                uneLBD = ConvertirSingleEnEcran(.maBandeModifD, unT, uneLg)
                
                If unIndCarfD > 0 Then
                    'Cas d'une onde verte descendante possible ==> Dessin
                    'unIndCarfD > 0 dit qu'on a trouvé des carrefours descendants
                    If unCarfRed.HasFeuDescendant = True Then
                        'Cas d'un carrefour contraignant de l'onde verte descendante
                        'donc ayant un feu de sens descendant
                        
                        'Abscisse du point suivant de l'onde descendante vaut
                        'l'abscisse du point précédent plus le décalage en
                        'temps entre le carrefour courant et le premier descendant
                        unTD2 = unTD1 - unCarf.monDecVitSensD + unCarfRedD1.monCarrefour.monDecVitSensD
                        'Signe - et + inverse par rapport au sens montant car les
                        'décalages en temps sont > 0 même en sens descendant
                        
                        'Stockage dans début onde pour trouver les plages
                        'sélectionnables graphiquement
                        unCarfRed.AffecterDebOndeSens unTD2, False
                        
                        'Ordonnée égale à l'ordonnée du carrefour réduit courant
                        'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                        unYD2 = unCarfRed.DonnerYSens(False)
                        
                        'Conversion en coordonnées écran de unYD2
                        unY = ConvertirReelEnEcran(unYD2 - unYMin, unDY, uneHt)
                        unY = unY0 - unY
                        
                        'Dessin de l'onde verte descendante inter-carrefours donc
                        'entre ce carrefour réduit et son précédent en Y
                        'si choix coché dans les options d'affichage et d'impression
                        'et si ce n'est pas le 1er carrefour descendante le + haut
                        
                        'Fait avant la bande commune pour ne voir que la bande
                        'commune si superposition avaec la bande inter-carrefours
                        If .mesOptionsAffImp.maVisuBandInterCarfD And (unNbCarf + 1 - i) <> unIndCarfD Then
                            'Calcul du début de vert de ce carrefour réduit
                            unDebVertD = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(False)
                            unDebVertD = ModuloZeroCycle(unDebVertD, .maDuréeDeCycle)
                            'Calcul du nombre de cycle séparant le début de
                            'vert du début de l'onde verte descendante
                            unNbCycle = Fix((0.001 + unTD2 - unDebVertD) / .maDuréeDeCycle)
                            If unNbCycle < 0 And .maBandeModifD = 0 Then
                                'Si pas de bande commune, on ne peut pas être en retard
                                'unTD2 ne doit pas être corrigé si il est < unDebVertD
                                unNbCycle = 0
                            End If
                            If unTD2 < unDebVertD - 0.001 Then
                                'Début de vert > T de onde descendante
                                '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                                unDebVertD = unDebVertD + unNbCycle * .maDuréeDeCycle
                            ElseIf unTD2 > unDebVertD + unCarfRed.DonnerDureeVertSens(False) + 0.001 Then
                                'Fin de vert < T de départ onde descendante
                                '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                                unDebVertD = unDebVertD + unNbCycle * .maDuréeDeCycle
                            End If
                                                      
                           'Calcul du temps de parcours inter-carrefours
                           'permutation des arguments de la soustraction par rapport
                           'au sens montant car les décalages comptés > 0 par rapport
                           'au carrefour descendant le plus bas
                            unTmpInterCarf = monTabCarfY(unIndCarfDPred).monCarfReduit.monCarrefour.monDecVitSensD - unCarf.monDecVitSensD
                            'Calcul de la fin de vert de ce carrefour réduit
                            unFinVertD = unDebVertD + unCarfRed.DonnerDureeVertSens(False)
                            unFinVertDPred = unDebVertDPred + monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerDureeVertSens(False)

                            unI = 0
                            unJ = 0
                            unDebOndeDéjàStocké = False
                            unLastDebVertDéjàStocké = False
                            'projection du début et fin de vert du carrefour
                            'précédent sur la droite Y = Y du carrefour courant
                            unDebVertDPred = unDebVertDPred + unTmpInterCarf
                            unFinVertDPred = unFinVertDPred + unTmpInterCarf
                            'Initialisation du LastDebVert au cas où aucun
                            'bande inter-carf trouvée
                            unLastDebVertD = unDebVertD
                            Do
                                'La boucle sert pour afficher toutes les bandes
                                'inter-carrefour pour cela on regarde dans
                                'le cycle en cours et le suivant et on prend la
                                'bande inter-carf maximale
                                
                                'On prend le minimun des fins de vert projeté
                                'sur la droite Y = Y du carrefour courant
                                If (unFinVertDPred + unJ * .maDuréeDeCycle) < (unFinVertD + unI * .maDuréeDeCycle) Then
                                    unMinFinVert = unFinVertDPred + unJ * .maDuréeDeCycle
                                Else
                                    unMinFinVert = unFinVertD + unI * .maDuréeDeCycle
                                End If
                                
                                'On prend le maximun des débuts de vert projeté
                                'sur la droite Y = Y du carrefour courant
                                If (unDebVertDPred + unJ * .maDuréeDeCycle) > (unDebVertD + unI * .maDuréeDeCycle) Then
                                    unMaxDebVert = unDebVertDPred + unJ * .maDuréeDeCycle
                                Else
                                    unMaxDebVert = unDebVertD + unI * .maDuréeDeCycle
                                End If
                                
                                'Test de l'existence d'une bande inter-carrefour
                                'supérieure à 1 seconde
                                uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                                
                                If uneBandeInterCarfExist Then
                                    If unTD2 - 0.01 < unMinFinVert And unLastDebVertDéjàStocké = False Then
                                        'Stockage du dernier debvert ayant une bande
                                        'inter-carrefour, ce stochage est fait une fois
                                        'et une seule entre deux carrefours
                                        unLastDebVertDéjàStocké = True
                                        unLastDebVertD = unDebVertD + unI * .maDuréeDeCycle
                                    End If
                                    'Stockage dans début onde pour trouver les
                                    'plages sélectionnables graphiquement,
                                    'la 1ère fois seulement
                                    If unDebOndeDéjàStocké = False Then
                                        unDebOndeDéjàStocké = True
                                        unCarfRed.AffecterDebOndeSens (unMaxDebVert + unMinFinVert) / 2, False
                                    End If
                                    'Remise dans l'englobant total si
                                    'MinFinVert en sort
                                    If unMinFinVert > unMaxT + 0.01 Then
                                        unMinFinVert = unMinFinVert - .maDuréeDeCycle
                                        unMaxDebVert = unMaxDebVert - .maDuréeDeCycle
                                    End If
                                End If
                                
                                unI = unI + 1
                                If unI = 2 Then
                                    'On se place pour essayer les début et fin de vert
                                    'du carrefour courant dans le cycle courant avec les
                                    'début et fin de vert du carrefour précédent dans le
                                    'cycle suivant
                                    unI = 0
                                    unJ = 1
                                End If
                                                                                    
                                'Dessin de bande inter-carrefour
                                If unTCD Is Nothing And uneBandeInterCarfExist Then
                                    'Cas d'une onde descendante non cadrée par un TC
                                    
                                    'Conversion en coordonnées écran
                                    unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 1ère partie bande descendante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                    
                                    'Conversion en coordonnées écran
                                    unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 2ème partie bande descendante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                ElseIf uneBandeInterCarfExist Then
                                    'Cas d'une onde descendante cadrée par un TC
                                    'Recup du Y du carf réduit précédent
                                    unYTmp! = monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False)
                                    'Calcul de la date en Y = -unYTmp% car inversion des Y en sens descendant
                                    unDecT0! = unTCD.CalculerDateDansTabMarche(unTCD.mesPhasesTMOnde, -unYTmp!, unIndPD%, unIndPD%)
                                    'Calcul du décalage en date pour avoir la prog partielle du TC
                                    'qui commence à la date absolue unMaxDebVert - unTmpInterCarf
                                    unDecT! = unMaxDebVert - unTmpInterCarf - unDecT0!
                                    'Dessin 1ère ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False), unCarf.monCarfRed.DonnerYSens(False), unDecT!, unIndPD%
                                    'Le dernier paramètre unDecT! sert à se cadrer au départ de l'onde sinon le TC débute à son T départ
                                    
                                    'Calcul du décalage en date pour avoir la prog partielle du TC
                                    'qui commence à la date absolue unMinFinVert - unTmpInterCarf
                                    unDecT! = unMinFinVert - unTmpInterCarf - unDecT0!
                                    'Dessin 2ème ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False), unCarf.monCarfRed.DonnerYSens(False), unDecT!, unIndPD%
                                    'Le dernier paramètre unDecT! sert à se cadrer au départ de l'onde sinon le TC débute à son T départ
                                End If
                                'Boucle fait trois fois pour trouver toutes
                                'les bande inter-carrefour
                            Loop Until unI = 1 And unJ = 1
                                
                        End If 'Fin du dessin de bande verte descendante inter-carrefours
                        
                        'Dessin de la bande verte descendante commune aux carrefours
                        'si choix coché dans les options d'affichage et d'impression
                        If Not unNoDessinOndeD And .mesOptionsAffImp.maVisuBandComD And unTCD Is Nothing Then
                            'Stockage dans début onde pour trouver les plages
                            'sélectionnables graphiquement
                            unCarfRed.AffecterDebOndeSens unTD2, False
                        
                            'Conversion en coordonnées écran de unTD2
                            unX = ConvertirSingleEnEcran(unTD2, unT, uneLg)
                            unX = unX + unNewX0
                            
                            'Dessin en coordonnées écran de la ligne entre les
                            'points (unTD1, unYD1) et (unTD2, unYD2) et d'une
                            'ligne // à une largeur de bande descendante
                            uneZoneDessin.Line (unXDpred, unYDpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComD
                            uneZoneDessin.Line (unXDpred + uneLBD, unYDpred)-(unX + uneLBD, unY), .mesOptionsAffImp.maCoulBandComD
                        
                            'Dessin de l'onde verte dans les feux du premier
                            'carrefour dans le sens descendant, sinon le dessin
                            'ne commence qu'au feu de Y minimun (cf réduction carrefour)
                            'Ce dessin va de unYMax jusqu'à Min Y 1er carrefour descendant
                            If unIndCarfD = (unNbCarf + 1 - i) Then
                                'Conversion en coordonnées écran du 1er point
                                unYFeu = DonnerYMaxCarfSens(unCarf, False, unIndFeu)
                                unXDpred = unTD1 - (unYD1 - unYFeu) / unCarfRed.DonnerVitSens(False)
                                unXDpred = ConvertirSingleEnEcran(unXDpred, unT, uneLg)
                                unXDpred = unXDpred + unNewX0
                                unYDpred = ConvertirReelEnEcran(CLng(unYFeu) - unYMin, unDY, uneHt)
                                unYDpred = unY0 - unYDpred
                                uneZoneDessin.Line (unXDpred, unYDpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComD
                                uneZoneDessin.Line (unXDpred + uneLBD, unYDpred)-(unX + uneLBD, unY), .mesOptionsAffImp.maCoulBandComD
                            End If
                            'Stockage du X écran du point précédent pour le coup suivant
                            unXDpred = unX
                        End If 'Fin du dessin de la bande verte descendante commune
                                                                                 
                        'Stockage du début de vert précédent
                        If unNbCarf + 1 - i = unIndCarfD Then
                            'Calcul spécial pour le carrefour le + haut descendant
                            unDebVertDPred = unCarfRedD1.monCarrefour.monDecModif + unCarfRedD1.DonnerPosRefSens(False)
                            unDebVertDPred = ModuloZeroCycle(unDebVertDPred, .maDuréeDeCycle)
                            If unTD1 < unDebVertDPred - 0.001 Then
                                'Début de vert > T de départ onde descendante
                                '==> Recul d'un cycle
                                unDebVertDPred = unDebVertDPred - .maDuréeDeCycle
                            ElseIf unTD1 > unDebVertDPred + unCarfRedD1.DonnerDureeVertSens(False) + 0.001 Then
                                'Fin de vert < T de départ onde descendante
                                '==> Avancé d'un cycle
                                unDebVertDPred = unDebVertDPred + .maDuréeDeCycle
                            End If
                        Else
                            unDebVertDPred = unLastDebVertD
                        End If
                        
                        'Stockage de l'indice de ce carrefour
                        unIndCarfDPred = unNbCarf + 1 - i
                        'Stockage du Y écran  du point précédent pour le coup suivant
                        unYDpred = unY
                    End If
                End If
            Next i
            
            'Dessin des ondes vertes cadrées par des TC, bandes communes
            'on le fait après les bandes inter-carrefours pour les bandes
            'communes ne soient pas écrasées par les bandes inter-carrefour
            
            'Dessin de l'onde verte commune montante dans le cas
            'd'un cadrage par TC si le dessin est possible et si
            'choisi dans les options d'affichage
            If Not (unTCM Is Nothing) And Not unNoDessinOndeM And unIndCarfM > 0 And monSite.mesOptionsAffImp.maVisuBandComM Then
                'Dessin de la 1ère ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, DessinOndeTCM, unTM1
                'Dessin de la 2ème ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, DessinOndeTCM, unTM1 + .maBandeModifM
            End If
            
            'Dessin de l'onde verte commune descendante dans le cas
            'd'un cadrage par TC si le dessin est possible et si
            'choisi dans les options d'affichage
            If Not (unTCD Is Nothing) And Not unNoDessinOndeD And unIndCarfD > 0 And monSite.mesOptionsAffImp.maVisuBandComD Then
                'Dessin de la 1ère ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, DessinOndeTCD, unTD1
                'Dessin de la 2ème ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, DessinOndeTCD, unTD1 + .maBandeModifD
            End If
        Else
            'Réduction globale impossible
            If Not TypeOf uneZoneDessin Is Printer Then uneZoneDessin.Cls
            'Initialisation pour éviter division par 0
            unT = 300
            monSite.monTmpTotal = unT
            unDY = 500
            monSite.monDYTotal = unDY
            'Idem pour éviter plantage dans PleinEcran
            monSite.monYMaxFeuUtil = 1
            monSite.monYMinFeuUtil = -1
            Exit Sub
        End If
                
        'Dessin des traits de rappels et des cycles en épaisseur 1
        'dans le cas d'une impression uniquement.
        'Raison : bug VB5 en impression les pointillés d'épaisseur > 1
        'font des traits continus sur certaines imprimantes jet d'encre
        'couleurs.
        'Restauration de l'épaisseur de trait pour imprimer plus bas
        If TypeOf uneZoneDessin Is Printer Then
            Printer.DrawWidth = 1
            
            'Si demandé dessin des lignes de rappels toutes les n secondes pour
            'une impression, n choisi dans la fenêtre d'impression (n entre 1 et
            '10) d'où :
            'unN = monSite.mesOptionsAffImp.monNbSecondesRappel en impression
            'Dessin de ses sous-divisions en trait pointillé si unN ne vaut pas
            '10 car les traits des dizaines sont faits aprés
            uneZoneDessin.DrawStyle = vbDash
            unN = monSite.mesOptionsAffImp.monNbSecondesRappel
            If .mesOptionsAffImp.maVisuLigne And unN <> 10 Then
                For unI = unMinT To unMaxT Step unN
                    If unI Mod .maDuréeDeCycle <> 0 Then
                        unX = ConvertirReelEnEcran(unI, unT, uneLg)
                        uneZoneDessin.Line (unNewX0 + unX, unY0)-(unNewX0 + unX, unY0 - uneHt), .mesOptionsAffImp.maCoulLigne
                    End If
                Next unI
            End If
        End If
        
        'Dessin des lignes de rappels toutes les 10 secondes en trait
        'plein si demandé
        uneZoneDessin.DrawStyle = vbSolid
        If .mesOptionsAffImp.maVisuLigne Then
            For unI = unMinT To unMaxT Step 10
                If unI Mod .maDuréeDeCycle <> 0 Then
                    unX = ConvertirReelEnEcran(unI, unT, uneLg)
                    uneZoneDessin.Line (unNewX0 + unX, unY0)-(unNewX0 + unX, unY0 - uneHt), .mesOptionsAffImp.maCoulLigne
                End If
            Next unI
        End If
                
        'Dessin des Traits de séparation de cycle en tiret-point (trait mixte)
        uneZoneDessin.DrawStyle = vbDashDot
        unNbCycle = Int(unT / .maDuréeDeCycle)
        If unMinT > 0 Then
            i0 = 1
        Else
            i0 = 0
        End If
        If unMaxT < unT Then
            j0 = 0
        Else
            j0 = 1
        End If
        
        'Suppression, donc effacement de l'affichage de
        'la durée du cycle dans la fenêtre plein écran
        If monPleinEcranVisible Then
            For i = frmPleinEcran.DureeCycle.Count - 1 To 1 Step -1
                Unload frmPleinEcran.DureeCycle(i)
            Next i
            frmPleinEcran.monNbCycle = 0
        End If
        
        unDT = uneLongCycle
        For i = i0 To unNbCycle + j0
            uneZoneDessin.Line (unNewX0 + i * unDT, unY0)-(unNewX0 + i * unDT, unY0 - uneHt), .mesOptionsAffImp.maCoulLigne
            If TypeOf uneZoneDessin Is Printer Then
                'Affichage de la durée du cycle sur imprimante
                ImprimerDureeCycle unX0, unY0, unNewX0 + i * unDT
            ElseIf Not unDessinDansOnglet Then
                'Affichage de la durée du cycle en fenêtre Pleine écran
                uneZoneDessin.AfficherDureeCycle (unNewX0 + i * unDT)
            End If
        Next i
        'Remise du dessin de ligne pleine
        uneZoneDessin.DrawStyle = vbSolid
        
        'Restauration de l'épaisseur de trait si on était en impression
        If TypeOf uneZoneDessin Is Printer Then
            Printer.DrawWidth = monSite.mesOptionsAffImp.monEpaisseurLigne
        End If
        
        'Calcul des limites de la zone de dessin sur écran ou sur imprimante
        If TypeOf uneZoneDessin Is Printer Then
            unDebZone = unX0
            unFinZone = unX0 + uneLg
        Else
            unDebZone = 0
            unFinZone = uneZoneDessin.Width
        End If
        
        'Calcul d'une précision montante et descendante pour les tests < et >
        '0n prend la durée du cycle convertie en largeur écran comme
        'précision
        unePrecM = ConvertirSingleEnEcran(monSite.maDuréeDeCycle / 2, unT, uneLg)
        unePrecD = unePrecM
        
        'Dessin des plages de vert montant et descendantes, et des points
        'de référence de tous feux de tous les carrefours
        For i = 1 To .mesCarrefours.Count
            Set unCarf = .mesCarrefours(i)
            If unCarf.monDecCalcul <> -99 Then
                unMin = 10000 ' Les ordonnées sont <= 9999 dans OndeV
                'récupération du point milieu de l'onde Mont ou Desc
                'converti en valeur écran
                If unCarf.monCarfRed.HasFeuMontant Then
                    unTM2 = unCarf.monCarfRed.DonnerDebOndeSens(True)
                    unTM2 = ConvertirSingleEnEcran(unTM2, unT, uneLg) + unNewX0
                End If
                If unCarf.monCarfRed.HasFeuDescendant Then
                    unTD2 = unCarf.monCarfRed.DonnerDebOndeSens(False)
                    unTD2 = ConvertirSingleEnEcran(unTD2, unT, uneLg) + unNewX0
                End If
                
                For j = 1 To unCarf.mesFeux.Count
                    Set unFeu = unCarf.mesFeux(j)
                    'Calcul du début de vert ramené entre 0 et +cycle
                    'pour dessiner les plages de vert si FinVert (= DebVert
                    'modulo cycle + la durée de vert) - le décalage est >= au
                    'cycle (donc le point de référence et cette plage de vert
                    'sont dans le même cycle) ou ramené entre -cycle et 0 sinon
                    'même si les position de référence sont trés grandes
                    '+ PosRef et pas - car les position point référence
                    'sont entrées avec un moins en interne dans OndeV
                    unDebVert = unCarf.monDecModif + unFeu.maPositionPointRef
                    unDebVertMod = ModuloZeroCycle(unDebVert, .maDuréeDeCycle)
                    If unDebVertMod + unFeu.maDuréeDeVert - unCarf.monDecModif > .maDuréeDeCycle - 0.01 Then
                        unDebVert = unDebVertMod - .maDuréeDeCycle
                    Else
                        unDebVert = unDebVertMod
                    End If
                    'Si la borne inf de l'englobant du graphic est < 0
                    'on enlève un cycle pour dessiner la partie T négative
                    If unMinT < 0 Then unDebVert = unDebVert - .maDuréeDeCycle
                    'Conversion de la plage de vert en coordonnées écran
                    unTmpDebVert = ConvertirSingleEnEcran(unDebVert, unT, uneLg)
                    unTmpFinVert = unTmpDebVert + ConvertirSingleEnEcran(CSng(unFeu.maDuréeDeVert), unT, uneLg)
                    'Conversion de l'ordonnée en coordonnées écran
                    unY = ConvertirReelEnEcran(unFeu.monOrdonnée - unYMin, unDY, uneHt)
                    
                    'Choix de la couleur du trait montant ou descendant
                    If unFeu.monSensMontant Then
                        uneCouleur = .mesOptionsAffImp.maCoulBandComM
                    Else
                        uneCouleur = .mesOptionsAffImp.maCoulBandComD
                    End If
                    
                    'Dessin de la plage de vert pour tous les cycles
                    For K = 0 To unNbCycle + 1
                        'Calcul des X écrans
                        unDebVert = unTmpDebVert + K * uneLongCycle
                        unFinVert = unTmpFinVert + K * uneLongCycle
                        unX = unDebVert + unNewX0
                        unXf = unFinVert + unNewX0
                        
                        'Stockage des lignes symbolisant la plage de
                        'vert contenant l'onde verte montante et descendante,
                        'cette ligne sera sélectionnable interactivement
                        'Précision = unePrecM ou unePrecD twips suivant le sens
                        If unFeu.monSensMontant Then
                            If unX - unePrecM < unTM2 And unTM2 < unXf + unePrecM And uneSortieImprimante = False Then
                                Set unePlageGraphic = New PlageGraphic
                                unePlageGraphic.AffecterAttributs CLng(unX), CLng(unXf), i, j
                                monSite.maColPlageGraphicM.Add unePlageGraphic
                            End If
                        Else
                            If unX - unePrecD < unTD2 And unTD2 < unXf + unePrecD And uneSortieImprimante = False Then
                                Set unePlageGraphic = New PlageGraphic
                                unePlageGraphic.AffecterAttributs CLng(unX), CLng(unXf), i, j
                                monSite.maColPlageGraphicD.Add unePlageGraphic
                            End If
                        End If
                        
                        'On ne dessine que dans la zone de dessin
                        If unX < unDebZone Then unX = unDebZone
                        If unXf < unDebZone Then unXf = unDebZone
                        If unX > unFinZone Then unX = unFinZone
                        If unXf > unFinZone Then unXf = unFinZone
                        'Dessin
                        uneZoneDessin.Line (unX, unY0 - unY)-(unXf, unY0 - unY), uneCouleur
                    Next K
                    
                    'Recherche du feu ayant l'ordonnée minimale du carrefour
                    If unFeu.monOrdonnée < unMin Then
                        unMin = unFeu.monOrdonnée
                    End If
                                    
                    If j = unCarf.mesFeux.Count Then
                        'Dessin du point de référence unique du carrefour au Y
                        'du feu le plus bas en ordonnée
                        'Conversion de l'ordonnée en coordonnées écran
                        unY = ConvertirReelEnEcran(unMin - unYMin, unDY, uneHt)
                        'Conversion du point de référence en coordonnées écran
                        unTmpPtRef = ConvertirSingleEnEcran(unCarf.monDecModif, unT, uneLg)
                        'Si la borne inf de l'englobant du graphic est < 0 on enlève
                        'un cycle en valeur écran pour dessiner la partie T < 0
                        If unMinT < 0 Then unTmpPtRef = unTmpPtRef - uneLongCycle
                        'Dessin du point de référence (triangle de hauteur écran
                        'valant 120 twips) pour tous les cycles
                        uneH = 120
                        'Indication pour l'impression du décalage
                        unDejaImprimer = False
                        For K = 0 To unNbCycle + 1
                            unTPtRef = unTmpPtRef + K * uneLongCycle + unNewX0
                            'On ne dessine dans les limites de la zone de
                            'dessin
                            If unDebZone < unTPtRef And unTPtRef < unFinZone Then
                                uneZoneDessin.Line (unTPtRef, unY0 - unY)-(unTPtRef - uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                uneZoneDessin.Line (unTPtRef, unY0 - unY)-(unTPtRef + uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                uneZoneDessin.Line (unTPtRef - uneH, unY0 - unY + uneH)-(unTPtRef + uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                'Affichage d'un cercle à l'intérieur du triangle
                                'pour les carrefours à décalages imposés
                                If unCarf.monDecImp = 1 Then
                                    uneZoneDessin.FillColor = .mesOptionsAffImp.maCoulPtRef
                                    uneZoneDessin.FillStyle = vbFSSolid
                                    uneZoneDessin.Circle (unTPtRef, unY0 - unY + uneH * 2 / 3), uneH / 3, .mesOptionsAffImp.maCoulPtRef
                                End If
                                
                                'Affichage du décalage une seule fois en impression
                                If TypeOf uneZoneDessin Is Printer And unDejaImprimer = False Then
                                    unDejaImprimer = True 'impression une seule fois
                                    Printer.ForeColor = .mesOptionsAffImp.maCoulPtRef
                                    Printer.CurrentX = unTPtRef + 1.1 * uneH
                                    Printer.CurrentY = unY0 - unY + 0.1 * uneH
                                    Printer.Print Format(CIntCorrigé(unCarf.monDecModif))
                                End If
                            End If
                            
                            'Stockage du point symbolisant la valeur du décalage
                            'de l'onde verte montante et descendante, ce point
                            'sera sélectionnable interactivement
                            'Précision = unePrecM ou unePrecD twips suivant le sens
                            unX = unTmpDebVert + K * uneLongCycle + unNewX0
                            unXf = unTmpFinVert + K * uneLongCycle + unNewX0
                            If unCarf.monCarfRed.HasFeuMontant Then
                                'If unX - unePrecM < unTM2 And unTM2 < unXf + unePrecM And uneSortieImprimante = False Then
                                If Abs(unTPtRef - unTM2) < uneLongCycle And uneSortieImprimante = False Then
                                    Set unRefGraphic = New RefGraphic
                                    unRefGraphic.AffecterAttributs unTPtRef, i
                                    monSite.maColRefGraphicM.Add unRefGraphic
                                End If
                            End If
                            If unCarf.monCarfRed.HasFeuDescendant Then
                                'If unX - unePrecD < unTD2 And unTD2 < unXf + unePrecD And uneSortieImprimante = False Then
                                If Abs(unTPtRef - unTD2) < uneLongCycle And uneSortieImprimante = False Then
                                    Set unRefGraphic = New RefGraphic
                                    unRefGraphic.AffecterAttributs unTPtRef, i
                                    monSite.maColRefGraphicD.Add unRefGraphic
                                End If
                            End If
                        Next K
                    End If
                Next j
            End If
        Next i
                
    End With
        
End Sub

Public Sub ModifierMaxTempsPourVisu(unMaxT)
    'Modification du paramètre unMaxT pour l'arrondi à la dizaine supérieure
    unReste = unMaxT - 10 * Fix(unMaxT / 10)
    If unMaxT > 0 Then
        unMaxT = unMaxT + 10 - unReste 'Ici unReste > 0
    Else
        unMaxT = unMaxT - unReste 'Ici unReste < 0
    End If
End Sub

Public Sub ModifierMinTempsPourVisu(unMinT)
    'Modification du paramètre unMinT pour l'arrondi à la dizaine inférieure
    unReste = unMinT - 10 * Fix(unMinT / 10)
    If unMinT > 0 Then
        unMinT = unMinT - unReste 'Ici unReste > 0
    Else
        unMinT = unMinT - 10 - unReste 'Ici unReste < 0
    End If
End Sub


Public Function TrouverMaxFinVert(unCarf As Carrefour) As Single
    'Retourne la date de fin de vert maximun parmi tous
    'les feux du carrefour unCarf
    'On ne le prend pas modulo cycle car une plage de vert contenant l'onde
    'verte peut se finir dans le cycle suivant
    Dim unFinVert As Single
    
    TrouverMaxFinVert = -1000
    For i = 1 To unCarf.mesFeux.Count
        unFinVert = unCarf.monDecModif + unCarf.mesFeux(i).maPositionPointRef + unCarf.mesFeux(i).maDuréeDeVert
        If unFinVert > TrouverMaxFinVert Then TrouverMaxFinVert = unFinVert
    Next i
End Function

Public Function TrouverMinDebVert(unCarf As Carrefour) As Single
    'Retourne la date de début de vert minimun parmi tous
    'les feux du carrefour unCarf
    'On ne la ramène pas modulo cycle car le début de vert minimun
    'd'où commence l'onde verte peut être négatif
    TrouverMinDebVert = 1000
    For i = 1 To unCarf.mesFeux.Count
        unDebVert = unCarf.monDecModif + unCarf.mesFeux(i).maPositionPointRef
        If unDebVert < TrouverMinDebVert Then TrouverMinDebVert = unDebVert
    Next i
End Function


Public Sub TrouverMinYMaxY(unYMin As Long, unYMax As Long)
    'Recherche du max et du min en Y des feux des carrefours
    'utilisés dans le calcul de l'onde verte
    'Les variables unYMin et unYMax sont modifiés par cette
    'procédure
    
    Dim unCarf As Carrefour
    Dim unFeu As Feu
    
    unYMin = 10000
    unYMax = -10000 'Les Y dans OndeV sont entre -9999 et 9999
    
    unNbCarf = monSite.mesCarrefours.Count
    For i = 1 To unNbCarf
        Set unCarf = monSite.mesCarrefours(i)
        If unCarf.monDecCalcul <> -99 Then
            'Cas d'un carrefour utilisé dans le calcul de l'onde
            For j = 1 To unCarf.mesFeux.Count
                Set unFeu = unCarf.mesFeux(j)
                If unFeu.monOrdonnée < unYMin Then
                    unYMin = unFeu.monOrdonnée
                End If
                If unFeu.monOrdonnée > unYMax Then
                    unYMax = unFeu.monOrdonnée
                End If
            Next j
        End If
    Next i
End Sub

Public Sub TracerProgressionTC(uneZoneDessin As Object, unTC As TC, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, Optional unTypeDessin As Integer = DessinProgTC, Optional unDecIniT As Single = 0)
    'Dessin de la progression du TC unTC de la liste des TC du site
    
    Dim unePhase As PhaseTabMarche
    Dim unSens As Integer, unIndFeu As Integer
    Dim unY As Single, unDT As Single
    Dim unT As Single, unDecT As Single
    Dim uneDateYext As Single, uneDateYdep As Single
    Dim uneCouleur As Long, unIndPhase As Integer
    Dim uneColPhases As ColPhaseTM
    Dim unNbFeuxMdep As Integer, unNbFeuxDdep As Integer
    Dim unNbFeuxMarr As Integer, unNbFeuxDarr As Integer
    
    'Information à l'utilisateur si les carrefours de départ et d'arrivée
    'n'ont pas de feu dans le sens du TC.
    'C'est juste une information les calculs continueront mais l'affichage
    'des graphiques feront apparaitre ces incohérences.
    unMsg = ""
    unTC.monCarfDep.DonnerNbFeuxMetD unNbFeuxMdep, unNbFeuxDdep
    unTC.monCarfArr.DonnerNbFeuxMetD unNbFeuxMarr, unNbFeuxDarr
    If DonnerYCarrefour(unTC.monCarfDep) <= DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC montant
        If unNbFeuxMdep = 0 Then
            unMsg = unMsg + "Le TC : " + unTC.monNom + " est de sens montant mais son carrefour de départ " + unTC.monCarfDep.monNom + " n'a aucun feu dans ce sens."
        End If
        'If unTC.monCarfArr.monCarfRed.HasFeuMontant = False Then
        If unNbFeuxMarr = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens montant mais son carrefour d'arrivée " + unTC.monCarfArr.monNom + " n'a aucun feu dans ce sens."
        End If
    Else
        'Cas d'un TC descendant
        If unNbFeuxDdep = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens descendant mais son carrefour de départ " + unTC.monCarfDep.monNom + " n'a aucun feu dans ce sens."
        End If
        'If unTC.monCarfArr.monCarfRed.HasFeuDescendant = False Then
        If unNbFeuxDarr = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens descendant mais son carrefour d'arrivée " + unTC.monCarfArr.monNom + " n'a aucun feu dans ce sens."
        End If
    End If
    If unMsg <> "" Then MsgBox unMsg, vbInformation, "Information OndeV pour correction"
    
    'Détermination du type de dessin TC à réaliser
    'et donc de sa couleur d'affichage et de la liste des phases,
    'donc du tableau de marche à dessiner (progression ou onde, cf classe TC)
    If unTypeDessin = DessinProgTC Then
        'Dessin du tableau de marche de progression du TC
        uneCouleur = unTC.maCouleur
        Set uneColPhases = unTC.mesPhasesTMProg
        'Décalage en temps à rajouter
        unDecT = 0
    ElseIf unTypeDessin = DessinOndeTCM Then
        'Dessin du tableau de marche du TC cadrant l'onde montante
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandComM
        Set uneColPhases = unTC.mesPhasesTMOnde
        'Décalage en temps entre le Y (= Y du feu montant le plus haut)
        'du feu équivalent du carrefour réduit et
        'du feu montant le plus bas du carrefour de départ
        uneDateYext = unTC.CalculerDateDansTabMarche(uneColPhases, unTC.monCarfDep.monCarfRed.DonnerYSens(True), unIndPhase, 1)
        uneDateYdep = unTC.CalculerDateDansTabMarche(uneColPhases, DonnerYMinCarfSens(unTC.monCarfDep, True, unIndFeu), unIndPhase, 1)
        unDecT = unDecIniT + uneDateYext - uneDateYdep
        'Décalage en temps à rajouter pour commencer au début de l'onde montante
        unDecT = unDecT - unTC.mesPhasesTMOnde(1).monTDeb
    ElseIf unTypeDessin = DessinOndeTCD Then
        'Dessin du tableau de marche du TC cadrant l'onde descendante
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandComD
        Set uneColPhases = unTC.mesPhasesTMOnde
        'Décalage en temps entre le Y (= Y du feu descendant le plus bas)
        'du feu équivalent du carrefour réduit et
        'du feu descendant le plus haut du carrefour de départ
        'en inversant le signe des Y pour le sens descendant
        uneDateYext = unTC.CalculerDateDansTabMarche(uneColPhases, -unTC.monCarfDep.monCarfRed.DonnerYSens(False), unIndPhase, 1)
        uneDateYdep = unTC.CalculerDateDansTabMarche(uneColPhases, -DonnerYMaxCarfSens(unTC.monCarfDep, False, unIndFeu), unIndPhase, 1)
        unDecT = unDecIniT - uneDateYext + uneDateYdep
        'Décalage en temps à rajouter pour commencer au début de l'onde descendante
        unDecT = unDecT - unTC.mesPhasesTMOnde(1).monTDeb
    Else
        MsgBox "ERREUR de programmation dans OndeV dans TracerProgressionTC", vbCritical
    End If
    
    'Détermination du sens du TC car les Y des phases sont inversés
    'pour le cas descendant
    If DonnerYCarrefour(unTC.monCarfDep) >= DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC descendant
        unSens = -1
    Else
        'Cas d'un TC montant
        unSens = 1
    End If
        
    'Parcours de toutes les phases du tableau de marche de progression du TC
    'pour trouver son englobant en temps et pour les dessiner si unAvecDessin
    'est VRAI
    unNbPhases = uneColPhases.Count
    unNbPoints = 5 'Nombre de points pour dessiner les paraboles
    For i = 1 To unNbPhases
        Set unePhase = uneColPhases(i)
        'Dessin de la phase
        
        'Conversion écran du point début de phase en inversant
        'les signes des Y début de phase si TC descendant (unSens = -1)
        unTE1 = ConvertirSingleEnEcran(unePhase.monTDeb + unDecT, monSite.monTmpTotal, uneLg)
        unTE1 = unTE1 + monSite.monOrigX
        
        unYE1 = ConvertirSingleEnEcran(unePhase.monYDeb * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
        unYE1 = unY0 - unYE1
        
        If unePhase.monType = Accel Or unePhase.monType = Decel Then
            'Cas d'une phase d'accélération ou de décélération, on dessine
            'la parabole grâce unNbPoints + 2 segments
            
            'Calcul du décalage en Temps entre chaque point de la parabole
            unDT = unePhase.maDureePhase / unNbPoints
            
            'Dessin des autres points de la phase d'acc ou de décélération
            For j = 1 To unNbPoints - 1
                'Calcul d'un point courant de la parabole
                unT = unePhase.monTDeb + j * unDT
                unY = CalculerYDansPhaseParabole(unePhase, unT)
                
                'Conversion écran du point courant de la parabole
                unTE2 = ConvertirSingleEnEcran(unT + unDecT, monSite.monTmpTotal, uneLg)
                unTE2 = unTE2 + monSite.monOrigX
                
                unYE2 = ConvertirSingleEnEcran(unY * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
                unYE2 = unY0 - unYE2
                
                'Dessin de la ligne entre le point courant et le précédent
                uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
                
                'Stockage du point écran courant pour l'incrémentation suivante
                unTE1 = unTE2
                unYE1 = unYE2
            Next j
        End If
        
        'Conversion écran du point fin de phase en inversant
        'les signes des Y début de phase si TC descendant (unSens = -1)
        unTE2 = ConvertirSingleEnEcran(unePhase.monTDeb + unDecT + unePhase.maDureePhase, monSite.monTmpTotal, uneLg)
        unTE2 = unTE2 + monSite.monOrigX
        
        unYE2 = ConvertirSingleEnEcran((unePhase.monYDeb + unePhase.maLongPhase) * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
        unYE2 = unY0 - unYE2
        
        'Dessin de la ligne entre le point précédent et le point fin de phase
        uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
    Next i
End Sub

Public Sub TracerProgPartielleTC(uneZoneDessin As Object, unTC As TC, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, unYDeb As Integer, unYFin As Integer, unDecT As Single, unIndPhaseDep As Integer)
    'Dessin de la progression partielle du TC unTC de la liste des TC du site
    'entre les ordonnées unYDeb et unYFin
    'Utiliser pour dessiner les bandes inter-carrefours d'ondes cadrées par
    'un TC montant et/ou un TC descendant
    
    Dim unePhase As PhaseTabMarche
    Dim unSens As Integer, i As Integer
    Dim unY As Single, unDT As Single
    Dim unY1 As Single, unY2 As Single
    Dim unTDeb As Single, unTFin As Single
    Dim unT As Single, unIndPhase As Integer
    Dim uneCouleur As Long
    Dim uneColPhases As ColPhaseTM
        
    'Détermination du sens du TC car les Y des phases sont inversés
    'pour le cas descendant
    If DonnerYCarrefour(unTC.monCarfDep) >= DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC descendant
        unSens = -1
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandInterCarfD
    Else
        'Cas d'un TC montant
        unSens = 1
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandInterCarfM
    End If
        
    'Parcours de toutes les phases du tableau de marche de progression du TC
    'pour trouver son englobant en temps et pour les dessiner si unAvecDessin
    'est VRAI
    Set uneColPhases = unTC.mesPhasesTMOnde
    unNbPhases = uneColPhases.Count
    unNbPoints = 5 'Nombre de points pour dessiner les paraboles
    uneFirstPhase = True 'Signale si on se trouve dans la 1ère phase
                         'de la progression partielle
    uneSortieBoucle = False
    i = unIndPhaseDep - 1
    Do
        i = i + 1
        Set unePhase = uneColPhases(i)
        'Dessin de la phase
        If unePhase.monYDeb + unePhase.maLongPhase > unYDeb * unSens - 0.001 Then
            'Cas d'une phase contenant Ydeb de la prog partielle
            If uneFirstPhase Then
                'Cas de la 1ère phase de la prog partielle
                unY1 = unYDeb * unSens
                uneFirstPhase = False
            Else
                unY1 = unePhase.monYDeb
            End If
            
            'Calcul de la date en unY1
            If unePhase.monType = Arret Then
                unTDeb = unePhase.monTDeb
            Else
                unTDeb = CalculerDateDansPhase(unePhase, unY1)
            End If
            
            If unePhase.monYDeb + unePhase.maLongPhase > unYFin * unSens - 0.001 Then
                'Cas où la phase courante sort de la prog partielle
                unY2 = unYFin * unSens
                uneSortieBoucle = True
            Else
                unY2 = unePhase.monYDeb + unePhase.maLongPhase
            End If
            
            'Calcul de la date en unY2
            If unePhase.monType = Arret Then
                unTFin = unePhase.monTDeb + unePhase.maDureePhase
            Else
                unTFin = CalculerDateDansPhase(unePhase, unY2)
            End If
            
            'Conversion écran du point début de phase en inversant
            'les signes des Y début de phase si TC descendant (unSens = -1)
            unTE1 = ConvertirSingleEnEcran(unTDeb + unDecT, monSite.monTmpTotal, uneLg)
            unTE1 = unTE1 + monSite.monOrigX
            
            unYE1 = ConvertirSingleEnEcran(unY1 * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
            unYE1 = unY0 - unYE1
            
            If unePhase.monType = Accel Or unePhase.monType = Decel Then
                'Cas d'une phase d'accélération ou de décélération, on dessine
                'la parabole grâce unNbPoints + 2 segments
                
                'Calcul du décalage en Temps entre chaque point de la parabole
                unDT = (unTFin - unTDeb) / unNbPoints
                
                'Dessin des autres points de la phase d'acc ou de décélération
                For j = 1 To unNbPoints - 1
                    'Calcul d'un point courant de la parabole
                    unT = unTDeb + j * unDT
                    unY = CalculerYDansPhaseParabole(unePhase, unT)
                    
                    If unY > unYFin * unSens + 0.001 Then Exit For
                    
                    'Conversion écran du point courant de la parabole
                    unTE2 = ConvertirSingleEnEcran(unT + unDecT, monSite.monTmpTotal, uneLg)
                    unTE2 = unTE2 + monSite.monOrigX
                    
                    unYE2 = ConvertirSingleEnEcran(unY * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
                    unYE2 = unY0 - unYE2
                    
                    'Dessin de la ligne entre le point courant et le précédent
                    uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
                    
                    'Stockage du point écran courant pour l'incrémentation suivante
                    unTE1 = unTE2
                    unYE1 = unYE2
                Next j
            End If
            
            'Conversion écran du point fin de phase en inversant
            'les signes des Y si TC descendant (unSens = -1)
            unTE2 = ConvertirSingleEnEcran(unTFin + unDecT, monSite.monTmpTotal, uneLg)
            unTE2 = unTE2 + monSite.monOrigX
            
            unYE2 = ConvertirSingleEnEcran(unY2 * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
            unYE2 = unY0 - unYE2
            
            'Dessin de la ligne entre le point précédent et le point fin de phase
            uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
        End If
    Loop Until uneSortieBoucle
End Sub

Public Sub DonnerEnglobantTC(unTC As TC, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, unTDep As Long, unTFin As Long)
    'alimentation des dates de début et de fin du parcours unTDep et unTFin pour
    'calculer l'englobant écran en temps , donc le niveau de zoom du graphique onde
    
    Dim unePhase As PhaseTabMarche
    Dim unSens As Integer
    
    'Calcul du tableau de marche de progression du TC
    'avec récupération de son sens (1 montant, -1 descendant)
    unSens = unTC.CalculerTableauMarcheProg()
        
    'Calcul du début de l'englobant écran en temps par conversion en
    'coordonnées écran du point de départ du TC, donc du T début de la
    'première phase
    Set unePhase = unTC.mesPhasesTMProg(1)
    unTDep = ConvertirSingleEnEcran(unePhase.monTDeb, monSite.monTmpTotal, uneLg)
    unTDep = unTDep + monSite.monOrigX
    
    'Calcul de la fin de l'englobant écran en temps par conversion en
    'coordonnées écran du point de fin du TC, donc du T début de la
    'dernière phase plus sa durée
    Set unePhase = unTC.mesPhasesTMProg(unTC.mesPhasesTMProg.Count)
    unTFin = ConvertirSingleEnEcran(unePhase.monTDeb + unePhase.maDureePhase, monSite.monTmpTotal, uneLg)
    unTFin = unTFin + monSite.monOrigX
End Sub

Public Sub DessinerTout(uneZoneDessin As Object, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, Optional unDessinDansOnglet As Boolean = True)
    'Dessin des ondes vertes montantes et descendantes,
    'et des progressions de TC
    Dim unTC As TC
    Dim unTDep As Long, unTFin As Long
    
    'Effacement de la zone de dessin si on m'imprime pas
    If Not (TypeOf uneZoneDessin Is Printer) Then uneZoneDessin.Cls

    'Dessin des ondes vertes montantes et descendantes
    DessinerOndeVerte uneZoneDessin, unX0, unY0, uneLg, uneHt, unDessinDansOnglet
                                               
    'Positionnement de la string "T en secondes" en bout de
    'l'axe des temps si on est dans l'onglet Dessin onde verte
    'Ainsi un pick souris n'est pas changé par un texte printé
    'contrairement à un control label qui trappe les clicks souris
    'Mais ce texte printé doit être rafraichit comme un dessin de ligne
    'd'où ce code en fin de dessin total
    If unDessinDansOnglet Then
        uneZoneDessin.CurrentX = monSite.LabelFleche.Left - uneZoneDessin.TextWidth("t en secondes")
        uneZoneDessin.CurrentY = monSite.AxeTemps.Y1 - monSite.LabelFleche.Height / 2 - uneZoneDessin.TextHeight("t en secondes")
        uneForeColor = uneZoneDessin.ForeColor 'Stockage pour restaurer après
        uneZoneDessin.ForeColor = 0 'Mise en noir
        uneZoneDessin.Print "t en secondes"
        uneZoneDessin.ForeColor = uneForeColor 'Restauration couleur initiale
    End If
End Sub

Public Function EstTCUtil(unIndTC) As Boolean
    'Fonction retournant :
    '   - vrai si le TC d'index unIndTC fait partie des TC utilisés, dont
    '     on veut tracer la progression
    '   - faux sinon
    Dim unTC As TC
    
    Set unTC = monSite.mesTC(unIndTC)
    unNbTCUtil = monSite.mesTCutil.Count
    j = 1
    EstTCUtil = False
    Do While EstTCUtil = False And j <= unNbTCUtil
        If unTC.monNom = monSite.mesTCutil(j).monNom Then
            EstTCUtil = True
        Else
            j = j + 1
        End If
    Loop
End Function

Public Sub SelectionGraphique(uneZoneDessin As Object, unXpick As Single, unYpick As Single)
    'Procédure de sélection graphique d'une plage de vert, d'une poignée
    'ou d'un point de référence
    'Lancé par l'event MouseDown sur le bouton gauche de la souris en
    'un point (unXpick, unYpick) en coordonnées écran
    'd'une Zone de dessin (PictureBox d'une fenetre site = frmDocument)
    'ou par la Form frmPleinEcran
    Dim unYreel As Long, unY As Long, unDTEcran As Long
    Dim uneLongEcranAxeY As Long
    Dim unePl As PlageGraphic, unRef As RefGraphic
    Dim unObjTrouv As Boolean
    Dim unY0 As Single, unYecran As Single
    Dim unItemAnnuler As Control
    
    'Stockage du X écran correspond à cette sélection
    monXEcranDebModif = unXpick
    monXEcranFinModif = unXpick 'Il sera modifié dans un MouseMove,
                                'donc dans ModifierSelection
    
    'Initialisation pour la conversion de valeurs réelles en écrans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la sélection dans la fenêtre plein écran
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
        'Calcul de la longueur écran de l'axe des ordonnées
        uneLongEcranAxeY = uneZoneDessin.AxeY.Y2 - uneZoneDessin.AxeY.Y1
        'Origine des Y
        unY0 = uneZoneDessin.FrameCarfTC.Top + uneZoneDessin.AxeY.Y2
        'Stockage de l'item Annuler dernière modif graphique
        Set unItemAnnuler = uneZoneDessin.mnuAnnulerModif
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la sélection dans l'onglet Graphique Onde Verte
        'Calcul de la longueur écran de l'axe des ordonnées
        uneLongEcranAxeY = monSite.AxeOrdonnée.Y2 - monSite.AxeOrdonnée.Y1
        'Origine des Y
        unEspacement = 120 'même valeur que dans AffichageOngletVisu
        unY0 = monSite.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
        'niveau que le min des Y
            
        'Affectation de uneZoneDessin à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeTemps.X2 - uneZoneDessin.AxeTemps.X1
        'Stockage de l'item Annuler dernière modif graphique
        Set unItemAnnuler = frmMain.mnuGraphicOndeAnnul
    Else
        MsgBox "Erreur de programmation dans OndeV dans SelectionGraphique", vbCritical
    End If
    
    'Recherche de l'objet graphique cliqué (point de référence, plage de vert
    'ou poignée de sélection)
    '==> Arrêt au premier trouvé
    'Le reste du traitement, la modification interactive intervient dans
    'l'event MouseMove avec bouton gauche enfoncé de la zone de dessin
    'qui appelera la fonction ModifierSelection et l'apparition des poignées
    'apparait dans le MouseUp ainsi que le recalcul et redessin
    'qui appelera la fonction MettreAJourSelection
    
    'Initialisation de la sélection à vide
    unObjTrouv = False
    monTypeObjPick = NoSel
    uneZoneDessin.PlageVert(0).Visible = False
    
    'Recherche parmi les poignées de sélection (gauche et droite)
    If uneZoneDessin.PoigneeGauche.Visible Then
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        unXmilieu = uneZoneDessin.PoigneeDroite.Left + uneZoneDessin.PoigneeDroite.Width / 2
        unYmilieu = uneZoneDessin.PoigneeDroite.Top + uneZoneDessin.PoigneeDroite.Height / 2
        If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
            unObjTrouv = True
            monTypeObjPick = PgDSel
            'Cas de la sélection interactive d'un fin de vert
            unMsg = "Sélection du fin de vert : Carrefour " + unCarf.monNom
            unMsg = unMsg + " et Feu = " + Format(monObjPick.monIndFeu)
        End If
        
        If Not unObjTrouv Then
            unXmilieu = uneZoneDessin.PoigneeGauche.Left + uneZoneDessin.PoigneeGauche.Width / 2
            unYmilieu = uneZoneDessin.PoigneeGauche.Top + uneZoneDessin.PoigneeGauche.Height / 2
            If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
                unObjTrouv = True
                monTypeObjPick = PgGSel
                'Cas de la sélection interactive d'un début de vert
                unMsg = "Sélection du début de vert : Carrefour " + unCarf.monNom
                unMsg = unMsg + " et Feu = " + Format(monObjPick.monIndFeu)
            End If
        End If
    End If
    
    'Masquage des poignées si aucune n'a été sélectionnée
    If Not unObjTrouv Then
        uneZoneDessin.PoigneeDroite.Visible = False
        uneZoneDessin.PoigneeGauche.Visible = False
    End If
    
    'Hauteur en twips de triangle symbolisant le point de référence
    'fixé dans la fonction DessinerOndeVerte
    uneH = 120
    
    'Recherche parmi les points de référence de l'onde verte descendante
    i = 1
    unNbTotal = monSite.maColRefGraphicD.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unRef = monSite.maColRefGraphicD(i)
        'Info dans la barre d'état indiquant la sélection d'un point de référence
        unMsg = "Sélection du point de référence : Carrefour " + monSite.mesCarrefours(unRef.monIndCarf).monNom
        'Récupération du Y minimun des feux du carrefour
        unYreel = DonnerYMinCarf(monSite.mesCarrefours(unRef.monIndCarf))
        'Conversion du Yréel en Y écran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est prés du point de référence
        'suivant la précision PrecPick
        If unRef.monDecal - uneH - PrecPick < unXpick And unXpick < unRef.monDecal + uneH + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + uneH + PrecPick Then
            'Stockage de l'objet trouvé par pick écran
            unObjTrouv = True
            monTypeObjPick = RefSel
            Set monObjPick = unRef
            'Apparition d'une image d'un triangle déplaçable
            'interativement au même emplacement
            uneZoneDessin.PtRef.Left = unRef.monDecal - uneZoneDessin.PtRef.Width / 2
            uneZoneDessin.PtRef.Top = unY0 - unYecran
            uneZoneDessin.PtRef.Visible = True
            
            'Création des plages de vert
            'qui se déplaceront avec le triangle point de référence
            PlacerPlagesVert uneZoneDessin, monObjPick, unY0, uneLongEcranAxeY, unDTEcran
        End If
        i = i + 1
    Loop
    
    'Recherche parmi les points de référence de l'onde verte montante
    i = 1
    unNbTotal = monSite.maColRefGraphicM.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unRef = monSite.maColRefGraphicM(i)
        'Info dans la barre d'état indiquant la sélection d'un point de référence
        unMsg = "Sélection du point de référence : Carrefour " + monSite.mesCarrefours(unRef.monIndCarf).monNom
        'Récupération du Y minimun des feux du carrefour
        unYreel = DonnerYMinCarf(monSite.mesCarrefours(unRef.monIndCarf))
        'Conversion du Yréel en Y écran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est prés du point de référence
        'suivant la précision PrecPick
        If unRef.monDecal - uneH - PrecPick < unXpick And unXpick < unRef.monDecal + uneH + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + uneH + PrecPick Then
            'Stockage de l'objet trouvé par pick écran
            unObjTrouv = True
            monTypeObjPick = RefSel
            Set monObjPick = unRef
            'Apparition d'une image d'un triangle déplaçable
            'interativement au même emplacement
            uneZoneDessin.PtRef.Left = unRef.monDecal - uneZoneDessin.PtRef.Width / 2
            uneZoneDessin.PtRef.Top = unY0 - unYecran
            uneZoneDessin.PtRef.Visible = True
            
            'Création des plages de vert
            'qui se déplaceront avec le triangle point de référence
            PlacerPlagesVert uneZoneDessin, monObjPick, unY0, uneLongEcranAxeY, unDTEcran
        End If
        i = i + 1
    Loop
    
    'Recherche parmi les plages de vert de l'onde verte descendante
    i = 1
    unNbTotal = monSite.maColPlageGraphicD.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unePl = monSite.maColPlageGraphicD(i)
        'Info dans la barre d'état indiquant la sélection d'un début de vert
        unMsg = "Sélection de la plage de vert : Carrefour " + monSite.mesCarrefours(unePl.monIndCarf).monNom
        unMsg = unMsg + " et Feu = " + Format(unePl.monIndFeu)
        'Récupération du Y du feu
        unYreel = monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monOrdonnée
        'Conversion du Yréel en Y écran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est dans la plage de vert
        'suivant la précision PrecPick
        If unePl.monDebVert - PrecPick < unXpick And unXpick < unePl.monFinVert + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + PrecPick Then
            'Stockage de l'objet trouvé par pick écran
            unObjTrouv = True
            monTypeObjPick = PlaSel
            Set monObjPick = unePl
            'Apparition d'une ligne modifiable interativement au même emplacement
            uneZoneDessin.PlageVert(0).X1 = unePl.monDebVert
            uneZoneDessin.PlageVert(0).Y1 = unY0 - unYecran
            uneZoneDessin.PlageVert(0).X2 = unePl.monFinVert
            uneZoneDessin.PlageVert(0).Y2 = unY0 - unYecran
            'Affectation de la couleur onde montante ou descendante
            If monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monSensMontant Then
                uneZoneDessin.PlageVert(0).BorderColor = monSite.mesOptionsAffImp.maCoulBandComM
            Else
                uneZoneDessin.PlageVert(0).BorderColor = monSite.mesOptionsAffImp.maCoulBandComD
            End If
            uneZoneDessin.PlageVert(0).Visible = True
        End If
        i = i + 1
    Loop
    
    'Recherche parmi les plages de vert de l'onde verte montante
    i = 1
    unNbTotal = monSite.maColPlageGraphicM.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unePl = monSite.maColPlageGraphicM(i)
        'Info dans la barre d'état indiquant la sélection d'un fin de vert
        unMsg = "Sélection de la plage de vert : Carrefour " + monSite.mesCarrefours(unePl.monIndCarf).monNom
        unMsg = unMsg + " et Feu = " + Format(unePl.monIndFeu)
        'Récupération du Y du feu
        unYreel = monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monOrdonnée
        'Conversion du Yréel en Y écran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est dans la plage de vert
        'suivant la précision PrecPick
        If unePl.monDebVert - PrecPick < unXpick And unXpick < unePl.monFinVert + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + PrecPick Then
            'Stockage de l'objet trouvé par pick écran
            unObjTrouv = True
            monTypeObjPick = PlaSel
            Set monObjPick = unePl
            'Apparition d'une ligne modifiable interativement au même emplacement
            uneZoneDessin.PlageVert(0).X1 = unePl.monDebVert
            uneZoneDessin.PlageVert(0).Y1 = unY0 - unYecran
            uneZoneDessin.PlageVert(0).X2 = unePl.monFinVert
            uneZoneDessin.PlageVert(0).Y2 = unY0 - unYecran
            'Affectation de la couleur onde montante ou descendante
            If monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monSensMontant Then
                uneZoneDessin.PlageVert(0).BorderColor = monSite.mesOptionsAffImp.maCoulBandComM
            Else
                uneZoneDessin.PlageVert(0).BorderColor = monSite.mesOptionsAffImp.maCoulBandComD
            End If
            uneZoneDessin.PlageVert(0).Visible = True
        End If
        i = i + 1
    Loop
    
    'Affichage dans la 1ère zone de la barre d'état
    'du résultat du pick
    If monTypeObjPick = NoSel Then
        'Cas où la sélection graphique est vide
        unMsg = "Rien de sélectionner"
        'Mise en grisée de l'annulation de la dernère modif graphique
        unItemAnnuler.Enabled = False
    End If
    frmMain.sbStatusBar.Panels.Item(1).Text = unMsg
End Sub


Public Sub MettreAJourSelection(uneZoneDessin As Object, unXpick As Single)
    'Affichage des poignées si on a sélectionné une plage de vert
    'ou Recalcul et redessin de l'onde verte si on a modifie un début ou
    'un fin de vert, un point de référence ou un décalage de carrefour
    Dim unFeu As Feu, uneForm As Object
    Dim unCarf As Carrefour, uneModifDec As Boolean
    Dim unX0 As Long, unY0 As Long, unDTEcran As Long
    Dim uneHt As Long, unDessinOnglet As Boolean
    Dim unOldDecal As Single, unEcartReel As Single
    
    'Initialisation pour la conversion de valeurs réelles en écrans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la sélection dans la fenêtre plein écran
        
        'Affectation à la form mère pour accéder aux controls
        Set uneForm = uneZoneDessin
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneForm.AxeT.X2 - uneForm.AxeT.X1
        'Calcul du cadre où l'on dessine
        unDessinOnglet = False
        unEspacement = 120 'même valeur que dans AffichageOngletVisu
        unX0 = uneForm.AxeT.X1
        unY0 = uneForm.FrameCarfTC.Top + uneForm.AxeY.Y2
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
        'niveau que le min des Y
        uneHt = uneForm.AxeY.Y2 - uneForm.AxeY.Y1
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la sélection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneForm à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        Set uneForm = monSite
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneForm.AxeTemps.X2 - uneForm.AxeTemps.X1
        'Calcul du cadre où l'on dessine
        unDessinOnglet = True
        unEspacement = 120 'même valeur que dans AffichageOngletVisu
        unX0 = uneForm.AxeTemps.X1
        unY0 = uneForm.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
        'niveau que le min des Y
        uneHt = uneForm.AxeOrdonnée.Y2 - uneForm.AxeOrdonnée.Y1
    End If
            
    'Masquage de l'info bulle de modif graphique
    uneForm.InfoModif.Visible = False
    
    'On vide les valeurs de la précédente modif graphique
    ViderCollection maColValPred
    
    'Sauvegarde de l'englobant en temps pour réutilisation dans la
    'fonction AnnulerLastModifGraphic qui annule la dernière modif graphique
    monTmpTotalAvantModif = monSite.monTmpTotal
        
    'Masquage de la plage de vert montrant la modification interactive
    uneForm.PlageVert(0).Visible = False
    
    If monTypeObjPick = PgGSel Then
        'Cas de la modification interactive d'un début de vert
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Sauvegarde de l'ancienne position de référence et l'ancienne durée
        'de vert dans la collection des valeurs précédentes pour un Annuler
        maColValPred.Add unFeu.maPositionPointRef
        maColValPred.Add unFeu.maDuréeDeVert
        'Conversion de l'écart de la modification de la plage de
        'vert d'une valeur écran en valeur réelle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Changement du point de référence du feu sélectionné arrondi en entier
        'si unEcartReel > 0 ==> Diminution du début de vert, donc de la durée
        'de vert et de la position de référence mais qui est < 0 en interne
        'donc il faut l'augmenter. Si unEcartReel < 0 ==> On fait l'inverse
        unFeu.maPositionPointRef = CInt(unFeu.maPositionPointRef + unEcartReel)
        unFeu.maDuréeDeVert = CInt(unFeu.maDuréeDeVert - unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = PgDSel Then
        'Cas de la modification interactive d'un fin de vert
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Conversion de l'écart de la modification de la plage de
        'vert d'une valeur écran en valeur réelle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Sauvegarde de l'ancienne durée de vert dans la
        'collection des valeurs précédentes pour un Annuler
        maColValPred.Add unFeu.maDuréeDeVert
        'Changement de la durée de vert du feu sélectionné arrondie en entier
        'si unEcartReel > 0 ==> Augmentation de la durée de vert
        'Si unEcartReel < 0 ==> On fait l'inverse
        unFeu.maDuréeDeVert = CInt(unFeu.maDuréeDeVert + unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = PlaSel Then
        'Cas de la sélection interactive d'une plage de vert
        '==> Apparition des poignées droite et gauche sélectionnables
        'aux extrémités de la plage qui a été sélectionnée
        uneForm.PoigneeGauche.Left = uneForm.PlageVert(0).X1 - uneForm.PoigneeGauche.Width / 2
        uneForm.PoigneeGauche.Top = uneForm.PlageVert(0).Y1 - uneForm.PoigneeGauche.Height / 2
        uneForm.PoigneeDroite.Left = uneForm.PlageVert(0).X2 - uneForm.PoigneeDroite.Width / 2
        uneForm.PoigneeDroite.Top = uneForm.PlageVert(0).Y2 - uneForm.PoigneeDroite.Height / 2
        uneForm.PoigneeDroite.Visible = True
        uneForm.PoigneeGauche.Visible = True
        'Masquage de la plage de vert montrant la sélection
        uneForm.PlageVert(0).Visible = False
        'Déplacement du point de référence du feu sélectionné
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Conversion de l'écart de la modification de la plage de
        'vert d'une valeur écran en valeur réelle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Sauvegarde de l'ancienne position de référence dans la
        'collection des valeurs précédentes pour un Annuler
        maColValPred.Add unFeu.maPositionPointRef
        'Translation du point de référence du feu sélectionné arrondi en entier
        unFeu.maPositionPointRef = CInt(unFeu.maPositionPointRef + unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = RefSel Then
        'Cas de la modification interactive d'un point de référence
        '==> Recalcul et redessin de l'onde verte
        'Masquage et destruction des plages de vert dépaçable
        'et du triangle
        uneForm.PtRef.Visible = False
        For i = uneForm.PlageVert.Count - 1 To 1 Step -1
            uneForm.PlageVert(i).Visible = False
            Unload uneForm.PlageVert(i)
        Next i
        'Modification du décalage du carrefour sélectionné
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        'Conversion de l'écart de la translation du point de
        'référence d'une valeur écran en valeur réelle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Stockage du décalage avant modification
        unOldDecal = unCarf.monDecModif
        'Sauvegarde de l'ancien décalage dans la collection
        'des valeurs précédentes pour un Annuler
        maColValPred.Add unOldDecal
        'Modification du décalage du carrefour sélectionné modulo cycle
        'On en prend la partie entière pour avoir la même chose si on avait
        'saisi ce nouveau décalage dans l'onglet Tableau Résultat
        unCarf.monDecModif = CIntCorrigé(ModuloZeroCycle(unCarf.monDecModif + unEcartReel, monSite.maDuréeDeCycle))
        'Modif du décalage modifiable du carrefour choisi
        'On ajoute la différence entre la vrai valeur en réelle et l'arrondi
        'en entier pour l'affichage dans l'onglet Tableau Résultat pour ne
        'pas perdre en précision de calcul
        'Exemple : si le calcul trouve un décalage de 29.8 que l'on stocke
        'on affiche par contre 30, si l'utilisateur remet 30 il peut avoir
        'un résultat différent car le 30 qu'il voit, vaut en fait 29.8.
        'En ajoutant la différence du à l'arrondi on retrouve la même chose
        unCarf.monDecModif = unOldDecal - CIntCorrigé(unOldDecal) + unCarf.monDecModif
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = NoSel Then
        'Cas où la sélection graphique est vide
        '==> On en fait rien
    Else
        MsgBox "Erreur de programmation dans OndeV dans MettreAJourSelection", vbCritical
    End If
        
    '==> Recalcul et redessin de l'onde verte
    'Test si l'élément pické a été bougé
    If monXEcranDebModif = monXEcranFinModif Then
        'Aucun mouvement de souris ==> pas de modif
        monSite.maModifDataDes = False
    End If
    
    If monSite.maModifDataDes Then
        'Initialisation des booléens permettant de savoir
        'si les calculs ont réussi
        unCalculOndeFait = False
        unRecalculBandeFait = False
        
        'Masquage des poignées de sélection
        uneForm.PoigneeDroite.Visible = False
        uneForm.PoigneeGauche.Visible = False
        
        If monTypeObjPick = RefSel And unCarf.monDecImp = 0 Then
            'Cas d'une modification de décalage d'un carrefour à décalage
            'non imposé ==> Recalcul de bandes passantes
            unRecalculBandeFait = RecalculerBandesPassantes(monSite)
            If unRecalculBandeFait Then
                'Cas où le recalcul a été possible
                'Stockage d'une modification de valeurs dans les décalages
                'Ceci permettra aussi de demander une sauvegarde à la fermeture
                maModifDataDec = True
                'Indication de fin de modif graphique, ainsi on ne refera
                'pas le calcul de l'onde verte en cas de changement d'onglet
                monSite.maModifDataDes = False
            Else
                'Cas où le recalcul a été impossible
                'On remet la valeur du décalage avant modif
                unCarf.monDecModif = unOldDecal
            End If
        Else
            'Cas d'une modification autre qu'un décalage d'un carrefour à
            'décalage non imposé, pour ce cas il faut recalculer l'onde aussi
            '==> Calcul de l'onde à refaire pour mise à jour
            unCalculOndeFait = True
            If monTypeObjPick = RefSel And unCarf.monDecImp = 1 Then
                'Cas d'une modif graphique du décalage
                'd'un carrefour à décalage imposé
                uneModifDec = True
            Else
                'Autres cas
                uneModifDec = False
            End If
            unCalculOndeFait = CalculerOndeVerte(monSite, uneModifDec)
        End If
        
        If unCalculOndeFait Or unRecalculBandeFait Then
            'Dessiner le graphique de l'onde verte
            uneZoneDessin.Cls
            DessinerTout uneZoneDessin, unX0, unY0, unDTEcran, uneHt, unDessinOnglet
            'Activation du menu pour annuler la dernière modif
            If TypeOf uneZoneDessin Is Form Then
                'Cas de modif interactive en plein écran
                uneForm.mnuAnnulerModif.Enabled = True
            Else
                'Cas de modif interactive dans l'onglet Graphique onde verte
                frmMain.mnuGraphicOndeAnnul.Enabled = True
            End If
        Else
            'Restauration des valeurs précédentes
            MsgBox "Détermination d'onde verte impossible ==> Restauration des valeurs précédentes", vbInformation
            AnnulerLastModifGraphic uneZoneDessin
        End If
    End If
End Sub

Public Sub ChangerPointeurSouris(uneZoneDessin As Object, unXpick As Single, unYpick As Single)
    'Changement du pointeur souris en croix si on passe
    'sur les poignées de sélection si elles sont visibles
    Dim unObjTrouv As Boolean
    
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la sélection dans la fenêtre plein écran
        Set uneForm = uneZoneDessin
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la sélection dans l'onglet Graphique Onde Verte
            
        'Affectation de uneZoneDessin à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneForm = monSite
    End If
    
    unObjTrouv = False
    'Test de la visibilité des poignées de sélection
    If uneForm.PoigneeGauche.Visible Then
        unXmilieu = uneForm.PoigneeDroite.Left + uneForm.PoigneeDroite.Width / 2
        unYmilieu = uneForm.PoigneeDroite.Top + uneForm.PoigneeDroite.Height / 2
        If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
            unObjTrouv = True
        End If
        
        If Not unObjTrouv Then
            unXmilieu = uneForm.PoigneeGauche.Left + uneForm.PoigneeGauche.Width / 2
            unYmilieu = uneForm.PoigneeGauche.Top + uneForm.PoigneeGauche.Height / 2
            If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
                unObjTrouv = True
            End If
        End If
    End If
    
    'Changement du pointeur souris par une croix
    'si on est sur une poignée
    If unObjTrouv Then
        uneZoneDessin.MousePointer = vbCrosshair
    Else
        uneZoneDessin.MousePointer = vbDefault
    End If
End Sub

Public Sub ModifierSelection(uneZoneDessin As Object, unXpick As Single, unYpick As Single)
    'Modification interactive de l'objet sélectionné dans le mouseDown
    'de uneZoneDessin
    Dim uneSecEcran As Long, unDTEcran As Long
    Dim unCycleEcran As Long, unDX As Single, unDX2 As Single
    Dim unCarf As Carrefour, unFeu As Feu
    Dim unXDebVert As Single, unXRef As Single
    
    'Initialisation pour la conversion de valeurs réelles en écrans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la sélection dans la fenêtre plein écran
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la sélection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneZoneDessin à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
        'Calcul de la longueur écran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeTemps.X2 - uneZoneDessin.AxeTemps.X1
    End If
        
    'Conversion d'une seconde et d'une durée de cycle réelles en valeur écran
    uneSecEcran = ConvertirReelEnEcran(1, monSite.monTmpTotal, unDTEcran)
    unCycleEcran = ConvertirReelEnEcran(monSite.maDuréeDeCycle, monSite.monTmpTotal, unDTEcran)
    
    If monTypeObjPick = PgGSel Then
        'Cas de la modification interactive d'un début de vert
        '==> Déplacement horizontale de l'extrémité début de vert de la plage
        uneZoneDessin.PlageVert(0).Visible = True
        'Calcul de la nouvelle durée de vert par Conversion de l'écart de
        'la modification d'une valeur écran en valeur réelle
        unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
        If unFeu.maDuréeDeVert - unEcartReel >= 1 And unFeu.maDuréeDeVert - unEcartReel <= monSite.maDuréeDeCycle - 1 Then
            'La durée de vert doit être >= à 1 et <= durée de cycle - 1
            uneZoneDessin.PlageVert(0).X1 = unXpick
            uneZoneDessin.PoigneeGauche.Left = unXpick - uneZoneDessin.PoigneeGauche.Width / 2
            'Stockage du X écran de fin de modification
            monXEcranFinModif = unXpick
            'Affichage dans l'info bulle de modif de la nouvelle durée de vert
            'et de la nouvelle position de référence
            unePosRef = CInt(unFeu.maPositionPointRef + unEcartReel)
            uneInfoPosRef = "Référence = " + Format(-unePosRef)
            uneDV = CInt(unFeu.maDuréeDeVert - unEcartReel)
            AfficherInfoModif uneZoneDessin, uneInfoPosRef + " Durée de vert = " + Format(uneDV), unXpick, unYpick
        End If
    ElseIf monTypeObjPick = PgDSel Then
        'Cas de la modification interactive d'un fin de vert
        '==> Déplacement horizontale de l'extrémité fin de vert de la plage
        uneZoneDessin.PlageVert(0).Visible = True
        If unXpick >= uneZoneDessin.PlageVert(0).X1 + uneSecEcran And unXpick <= uneZoneDessin.PlageVert(0).X1 + unCycleEcran - uneSecEcran Then
            'Déplacement doit être > au début de vert et la plage < cycle
            uneZoneDessin.PlageVert(0).X2 = unXpick
            uneZoneDessin.PoigneeDroite.Left = unXpick - uneZoneDessin.PoigneeGauche.Width / 2
            'Calcul de la nouvelle durée de vert par Conversion de l'écart de
            'la modification d'une valeur écran en valeur réelle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
            'Stockage du X écran de fin de modification
            monXEcranFinModif = unXpick
            'Affichage dans l'info bulle de modif de la nouvelle durée de vert
            uneDV = CInt(unFeu.maDuréeDeVert + unEcartReel)
            AfficherInfoModif uneZoneDessin, "Durée de vert = " + Format(uneDV), unXpick, unYpick
        End If
    ElseIf monTypeObjPick = PlaSel Then
        'Cas de la sélection interactive d'une plage de vert
        '==> Déplacement horizontale de la plage de vert
        uneZoneDessin.PlageVert(0).Visible = True
        'Calcul du déplacement par rapport au début de la sélection
        unDX = unXpick - monXEcranDebModif
        'Déplacement de la plage de vert sélectionnée en le bloquant
        'horizontalement de telle façon que le début du vert de la plage
        'varie entre la fin de vert de la plage - Cycle et cette fin de vert
        '==> Toutes les modifs possibles sont dans cet intervalle
        unXDebVert = monObjPick.monDebVert + unDX
        If monObjPick.monFinVert - unCycleEcran <= unXDebVert And unXDebVert <= monObjPick.monFinVert Then
            uneZoneDessin.PlageVert(0).X1 = monObjPick.monDebVert + unDX
            uneZoneDessin.PlageVert(0).X2 = monObjPick.monFinVert + unDX
            'Calcul du nouveau point de référence par conversion
            'de l'écart de la modification de la plage de
            'vert d'une valeur écran en valeur réelle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
            'Affichage dans l'info bulle de modif du nouveau point de ref
            unPref = CInt(unFeu.maPositionPointRef + unEcartReel)
            AfficherInfoModif uneZoneDessin, "Référence = " + Format(-unPref), unXpick, unYpick
            'Stockage du X écran de fin de modification
            monXEcranFinModif = unXpick
        End If
    ElseIf monTypeObjPick = RefSel Then
        'Cas de la modification interactive d'un point de référence
        '==> Déplacement horizontale de tous les feux et du point de
        'référence du carrefour, donc du triangle
                
        'Calcul du déplacement par rapport au début de la sélection
        unDX = unXpick - monXEcranDebModif
        'Déplacement du triangle point de référence et de tous les feux
        'du carrefours entre ce point de référence - un cycle et ce point
        'de référence + cycle
        '==> Toutes les modifs possibles sont dans cet intervalle
        unXRef = monObjPick.monDecal + unDX
        If monObjPick.monDecal - unCycleEcran <= unXRef And unXRef <= monObjPick.monDecal + unCycleEcran Then
            uneZoneDessin.PtRef.Left = unXRef - uneZoneDessin.PtRef.Width / 2
            For i = 1 To uneZoneDessin.PlageVert.Count - 1
                'Détermination de l'indice du feu
                Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
                unNbFeux = unCarf.mesFeux.Count
                If i <= unNbFeux Then
                    'Cas des plages du même cycle
                    unInd = i
                    Set unFeu = unCarf.mesFeux(unInd)
                    'Déplacement du début de vert de la plage i
                    unePosRefEcran = ConvertirSingleEnEcran(unFeu.maPositionPointRef, monSite.monTmpTotal, unDTEcran)
                    unDX2 = monObjPick.monDecal + unDX + unePosRefEcran - uneZoneDessin.PlageVert(i).X1
                    'unDX2 est le déplacement relatif par rapport à la dernière position
                    'alors qu'unDX est le déplacement par rapport au début de modification
                    uneZoneDessin.PlageVert(i).X1 = uneZoneDessin.PlageVert(i).X1 + unDX2
                Else
                    'Cas des plages graphiquement sélectionnables
                    unInd = i - unNbFeux
                    Set unFeu = unCarf.mesFeux(unInd)
                    'Déplacement du début de vert de la plage i par + unDX2
                    uneZoneDessin.PlageVert(i).X1 = uneZoneDessin.PlageVert(i).X1 + unDX2
                End If
                
                'Déplacement de la fin de vert de la plage i
                uneDurVertEcran = ConvertirSingleEnEcran(unFeu.maDuréeDeVert, monSite.monTmpTotal, unDTEcran)
                uneZoneDessin.PlageVert(i).X2 = uneZoneDessin.PlageVert(i).X1 + uneDurVertEcran
                'uneZoneDessin.PlageVert(i).Visible = True
            Next i
            'Calcul du décalage par conversion de l'écart de la
            'translation du point de référence d'une valeur écran
            'en valeur réelle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
            'Affichage dans l'info bulle de modif du nouveau point de ref
            unDec = CIntCorrigé(ModuloZeroCycle(unCarf.monDecModif + unEcartReel, monSite.maDuréeDeCycle))
            'On garde la même précision des chiffres après la virgule
            unDec = CIntCorrigé(unCarf.monDecModif - CIntCorrigé(unCarf.monDecModif) + unDec)
            AfficherInfoModif uneZoneDessin, "Décalage = " + Format(unDec), unXpick, unYpick
            'Stockage du X écran de fin de modification
            monXEcranFinModif = unXpick
        End If
    ElseIf monTypeObjPick = NoSel Then
        'Cas où la sélection graphique est vide
        '==> On en fait rien
    Else
        MsgBox "Erreur de programmation dans OndeV dans ModifierSelection", vbCritical
    End If
End Sub

Public Sub PlacerPlagesVert(uneZoneDessin As Object, unObjPick As Object, unY0 As Single, uneLongEcranAxeY As Long, unDTEcran As Long)
    'Création des plages de vert
    'qui se déplaceront avec le triangle point de référence
    Dim uneColPlageGraphic As Collection
    
    'Initialisation pour la conversion de valeurs réelles en écrans
    If TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la sélection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneZoneDessin à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
    End If
    
    'Récupération du carrefour du point de référence sélectionné
    Set unCarf = monSite.mesCarrefours(unObjPick.monIndCarf)
    
    'Création des plages de vert de tous les feux du carrefour
    'se trouvant dans le même cycle que le décalage du carrefour
    unNbFeux = unCarf.mesFeux.Count
    For i = 1 To unNbFeux
        Set unFeu = unCarf.mesFeux(i)
        unY = ConvertirReelEnEcran(unFeu.monOrdonnée - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        
        'Création d'une nouvelle plage de vert déplaçable,
        'celle de même cycle que le décalage du carrefour
        Load uneZoneDessin.PlageVert(i)
        
        'Positionnement en coordonnées écran de cette plage
        uneZoneDessin.PlageVert(i).X1 = unObjPick.monDecal + ConvertirSingleEnEcran(unFeu.maPositionPointRef, monSite.monTmpTotal, unDTEcran)
        uneZoneDessin.PlageVert(i).Y1 = unY0 - unY
        uneZoneDessin.PlageVert(i).X2 = uneZoneDessin.PlageVert(i).X1 + ConvertirSingleEnEcran(unFeu.maDuréeDeVert, monSite.monTmpTotal, unDTEcran)
        uneZoneDessin.PlageVert(i).Y2 = unY0 - unY
        
        'Affectation de la couleur onde montante ou descendante
        If unFeu.monSensMontant Then
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComM
        Else
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComD
        End If
        uneZoneDessin.PlageVert(i).Visible = True
    Next i
    
    'Création des plages de vert de tous les feux du carrefour
    'celles stockées dans la liste de sélection, donc celles
    'sélectionnables graphiquement
    For i = unNbFeux + 1 To 2 * unNbFeux
        Set unFeu = unCarf.mesFeux(i - unNbFeux)
        unY = ConvertirReelEnEcran(unFeu.monOrdonnée - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        
        'Création d'une nouvelle plage de vert déplaçable,
        'celle de même cycle que le décalage du carrefour
        Load uneZoneDessin.PlageVert(i)
    
        'Affectation de la couleur onde montante ou descendante et
        'Utilisation des plages graphiques montantes ou descendantes
        If unFeu.monSensMontant Then
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComM
            Set uneColPlageGraphic = monSite.maColPlageGraphicM
        Else
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComD
            Set uneColPlageGraphic = monSite.maColPlageGraphicD
        End If
            
        'Recherche de la plage graphique sélectionnable représentant ce feu
        'dans une collection contenant cette plage
        unInd = DonnerIndicePlage(uneColPlageGraphic, unObjPick.monIndCarf, i - unNbFeux)
        
        If unInd = 0 Then
            'Aucune plage trouvée
            uneZoneDessin.PlageVert(i).Visible = False
        Else
            'Positionnement en coordonnées écran de cette plage
            uneZoneDessin.PlageVert(i).X1 = uneColPlageGraphic(unInd).monDebVert
            uneZoneDessin.PlageVert(i).Y1 = unY0 - unY
            uneZoneDessin.PlageVert(i).X2 = uneColPlageGraphic(unInd).monFinVert
            uneZoneDessin.PlageVert(i).Y2 = unY0 - unY
            uneZoneDessin.PlageVert(i).Visible = True
        End If
    Next i
End Sub

Public Sub AnnulerLastModifGraphic(uneZoneDessin As Object)
    'Annulation de la dernière modification interactive dans un
    'graphique d'onde verte en récupérant les anciennes valeurs de l'objet
    'graphique sélectionné et stockées lors la modification dans la collection
    'maColValPred pour revenir à l'état précédent.
    'Recalcul de l'onde ou des décalages précédents et redessin du graphique
        
    If monTypeObjPick = PgGSel Then
        'Restauration de la position du point de référence et de la durée
        'de verte précédentes du feu sélectionné
        'Le 1er élément de maColValPred est l'ancienne position
        'et le 2 ème l'ancienne durée de vert
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maPositionPointRef = maColValPred(1)
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maDuréeDeVert = maColValPred(2)
    ElseIf monTypeObjPick = PgDSel Then
        'Restauration de la durée de verte précédente du feu
        'sélectionné, le 1er élément de maColValPred est cette durée
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maDuréeDeVert = maColValPred(1)
    ElseIf monTypeObjPick = PlaSel Then
        'Restauration de la position du point de référence précédente du feu
        'sélectionné, le 1er élément de maColValPred est cette position.
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maPositionPointRef = maColValPred(1)
    ElseIf monTypeObjPick = RefSel Then
        'Restauration du décalage précédent, le 1er élément de
        'maColValPred est le décalage précédent
        monSite.mesCarrefours(monObjPick.monIndCarf).monDecModif = maColValPred(1)
        'Recalcul des bandes passantes, toujours possible car
        'il l'était avant la modif graphique
        unResultat = RecalculerBandesPassantes(monSite)
        'Modification de l'indicateur de changement de données
        'car modification est invalidée
        'pas de recalcul de l'onde verte en cas de changement d'onglet
        monSite.maModifDataDec = False
    Else
        If monTypeObjPick > NoSel Then
            'Cas autre que rien de sélectionner
            MsgBox "Erreur de programmation dans OndeV dans AnnulerLastModifGraphic", vbCritical
        End If
    End If
            
    'Recalcul de l'onde verte pour l'annulation d'une modif
    'interactive autre que celle d'un décalage
    If monTypeObjPick = PgDSel Or monTypeObjPick = PgGSel Or monTypeObjPick = PlaSel Then
        'Indication d'une modif pour recalculer l'onde verte
        monSite.maModifDataCarf = True
        'Recalcul d'onde verte sans tester sur la faisabilité
        'du calcul car il avait été fait avant
        unResultat = CalculerOndeVerte(monSite)
    End If
    
    'Mise à jour du dessin d'ondes vertes et de progression TC
    MettreAJourDessin
End Sub

Public Sub MettreAJourDessin()
    'Redessin du Graphique d'Onde verte, avec les progressions des TC
    'éventuelles dans l'onglet Graphique Onde Verte
    'ou dans la fenêtre plein écran suivant la valeur de uneZoneDessin
    
    Dim unX0 As Long, unY0 As Long
    Dim uneHt As Long, uneLg As Long
    Dim uneZoneDessin As Object
    
    'Initialisation pour la conversion de valeurs réelles en écrans
    If monPleinEcranVisible = False Then
        'Cas de modification dans l'onglet Graphique Onde Verte
        'Affectation de uneForm à monSite pour accéder aux controls
        'd'interaction graphique poignéee, etc...
        Set uneZoneDessin = monSite.ZoneDessin
        'Calcul de la longueur écran de l'axe des temps
        uneLg = monSite.AxeTemps.X2 - monSite.AxeTemps.X1
        'Calcul du cadre où l'on dessine
        unEspacement = 120 'même valeur que dans AffichageOngletVisu
        unX0 = monSite.AxeTemps.X1
        unY0 = monSite.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
        'niveau que le min des Y
        uneHt = monSite.AxeOrdonnée.Y2 - monSite.AxeOrdonnée.Y1
        'Redessin de l'onglet Graphique Onde verte de la
        'fenêtre active si c'est l'onglet en cours d'utilisation
        If monSite.TabFeux.Tab = 4 Then
            uneZoneDessin.Cls 'effacement
            unEspacement = 120 'même valeur que dans AffichageOngletVisu
            DessinerTout uneZoneDessin, unX0, unY0, uneLg, uneHt, True
        End If
    Else
        'Cas de modification dans la fenêtre plein écran
        'Affectation à la form mère pour accéder aux controls
        Set uneZoneDessin = frmPleinEcran
        'Calcul de la longueur écran de l'axe des temps
        uneLg = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
        'Calcul du cadre où l'on dessine
        unX0 = uneZoneDessin.AxeT.X1
        unY0 = uneZoneDessin.FrameCarfTC.Top + uneZoneDessin.AxeY.Y2
        uneHt = uneZoneDessin.AxeY.Y2 - uneZoneDessin.AxeY.Y1
        'Effacement de la zone de dessin
        uneZoneDessin.Cls
        'Redessin
        DessinerTout uneZoneDessin, unX0, unY0, uneLg, uneHt, False
    End If
End Sub


Public Sub AfficherInfoModif(uneForm As Object, unMsg As String, unX As Single, unY As Single)
    'Affichage dans la zone de dessin d'une info bulle affichant le paramètre
    'modifié interactivement et sa valeur
    uneForm.InfoModif.Caption = unMsg
    uneForm.InfoModif.Left = unX - uneForm.InfoModif.Width - 60 'en twips
    uneForm.InfoModif.Top = unY
    uneForm.InfoModif.Visible = True
End Sub

Public Sub TrouverTempsParcoursEtCarrefours(unIndCarfM, unIndCarfD, unTmpM, unTmpD)
    'Recherche du carrefour le plus haut ayant un feu montant et du
    'carrefour le plus haut ayant un feu descendant.
    'Ces carrefours donneront les temps de parcours dans les deux sens
    'Les carrefours  seront retournés dans les variables unIndCarfM et
    'unIndCarfD et les temps de parcours dans unTmpM et unTmpD
    'Valeurs nulles si rien n'est trouvé dans un sens
    Dim unTC As TC
    
    unIndCarfM = 0
    unIndCarfD = 0
    i = UBound(monTabCarfY, 1) + 1
    'On part du carrefour le plus haut pour minimiser la boucle
    Do
        i = i - 1
        Set unCarfRed = monTabCarfY(i).monCarfReduit
        If TypeOf unCarfRed Is CarfReduitSensDouble Then
            If unIndCarfM = 0 Then
                unIndCarfM = i
            End If
            If unIndCarfD = 0 Then
                unIndCarfD = i
            End If
        ElseIf TypeOf unCarfRed Is CarfReduitSensUnique Then
            If unCarfRed.monSensMontant Then
                'Cas d'un carrefour à sens unique montant
                If unIndCarfM = 0 Then
                    unIndCarfM = i
                End If
            Else
                'Cas d'un carrefour à sens unique descendant
                If unIndCarfD = 0 Then
                    unIndCarfD = i
                End If
            End If
        End If
    Loop While (unIndCarfM = 0 Or unIndCarfD = 0) And i > 1
    
    'Récupération des temps de parcours montant et descendant
    'Tous les deux sont > 0, les temps de parcours montant et descendant
    'sont stockés dans les décalages dus à la vitesse dans le carrefour le
    'plus haut en Y, sauf pour le temps de parcours descendant dans le cas
    'd'une onde verte cadrée par un TC descendant
    If unIndCarfM = 0 Then
        unTmpM = 0
    Else
        unTmpM = monTabCarfY(unIndCarfM).monCarfReduit.monCarrefour.monDecVitSensM
    End If
    If unIndCarfD = 0 Then
        unTmpD = 0
    Else
        If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
            'Cas d'une onde cadrée par un TC descendant, le temps de parcours
            'descendant total est donnée par la fin de la dernière du tableau
            'de marche cadrant l'onde moins le début de la 1ère phase
            Set unTC = monSite.mesTC(monSite.monTCD)
            Set uneLastPhase = unTC.mesPhasesTMOnde(unTC.mesPhasesTMOnde.Count)
            unTmpD = uneLastPhase.monTDeb + uneLastPhase.maDureePhase - unTC.mesPhasesTMOnde(1).monTDeb
        Else
            unTmpD = monTabCarfY(unIndCarfD).monCarfReduit.monCarrefour.monDecVitSensD
        End If
    End If
End Sub

Public Sub ImprimerDureeCycle(unX0 As Long, unY0 As Long, unX)
    'Affichage d'un texte contenant 0/Durée du cycle
    'sur chaque trait de cycle
    uneInfoCycle = "0/" + Format(monSite.maDuréeDeCycle)
    Printer.ForeColor = 0
    Printer.CurrentX = unX - Printer.TextWidth(uneInfoCycle) / 2
    Printer.CurrentY = unY0 + Printer.TextHeight("OndeV") * 2
    Printer.Print uneInfoCycle
    Printer.Line (unX, unY0 + Printer.TextHeight(uneInfoCycle) * 2)-(unX, unY0), 0
End Sub

Public Function RemplirFicheResultPourImp() As Boolean
    'Remplissage de la fiche résultats pour l'imprimer
    
    Dim unNomTC As String
    
    'Calcul de l'onde verte si l'onglet courant n'est ni l'onglet
    'Résultat décalages et ni l'onglet Graphique onde verte
    unCalculOndeFait = True
    If monSite.TabFeux.Tab <> 3 And monSite.TabFeux.Tab <> 4 Then
        unCalculOndeFait = CalculerOndeVerte(monSite)
    End If
    
    If unCalculOndeFait Then
        'Remplissage possible car onde verte trouvée
        RemplirFicheResultPourImp = True
        
        'Calcul des vitesses maximun si un des décalages a été changé
        'ou si un nouveau calcul d'onde verte a été fait
        If monSite.maModifDataDec Then
            CalculerVitMax monSite
        End If
        
        'Ajout aux TC dont on cherche la progression de ceux
        'pris en compte pour l'onde verte si on calcule une onde TC,
        'sauf s'ils en font déjà partie
        If monSite.OptionTC.Value Then
            monSite.monTypeOnde = OndeTC
            unTCM = 0
            unTCD = 0
            i = 1
            Do While i <= monSite.mesTCutil.Count And (unTCM = 0 Or unTCD = 0)
                'Recherche de la présence du TC cadrant onde sens montant
                unNomTC = monSite.mesTCutil(i).monNom
                If monTCM <> 0 Then
                    'Cas où un TC cadre l'onde en sens montant
                    If unNomTC = monSite.mesTC(monTCM).monNom Then
                        unTCM = TrouverTCParNom(monSite, unNomTC)
                    End If
                End If
                'Recherche de la présence du TC cadrant onde sens descendant
                If monTCD <> 0 Then
                    'Cas où un TC cadre l'onde en sens descendant
                    If unNomTC = monSite.mesTC(monTCD).monNom Then
                        unTCD = TrouverTCParNom(monSite, unNomTC)
                    End If
                End If
                i = i + 1
            Loop
            If unTCM = 0 And monTCM <> 0 Then monSite.mesTCutil.Add monSite.mesTC(monTCM)
            If unTCD = 0 And monTCD <> 0 Then monSite.mesTCutil.Add monSite.mesTC(monTCD)
        End If
        'Remplir l'onglet Fiche résultat
        RemplirOngletFicheResult monSite
    Else
        'Remplissage impossible car onde verte non trouvée
        RemplirFicheResultPourImp = False
    End If
End Function

Public Sub RendreNulleBandesEtDecalages(unSite As Form)
    'Mise à zéro des bandes passantes et
    'des décalages de tous les carrefours
    unSite.maBandeM = 0
    unSite.maBandeD = 0
    unSite.maBandeModifM = 0
    unSite.maBandeModifD = 0
    
    For i = 1 To unSite.mesCarrefours.Count
        unSite.mesCarrefours(i).monDecCalcul = 0
        unSite.mesCarrefours(i).monDecModif = 0
    Next i
End Sub

Public Sub TrierFeuYCroissant(unTabFeu() As Feu, unSens As Integer)
    'Réorganisation par ordre croissant de leur ordonnée
    'd'un tableau de feu indexé entre 1 et n.
    'unSens permet de changer le signe des Y à classer, c'est utilisé
    'pour le sens descendant (unSens vaut 1 ou -1)
    
    'Algo choisi : Le tri insertion (récupérer sur Internet)
    'Il consiste à comparer successivement un élément
    'à tous les précédents et à décaler les éléments intermédiaires

    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unFeuTmp As Feu
    
    'Tri
    unNbTotal = UBound(unTabFeu, 1)
    For j = 2 To unNbTotal
            uneFinBoucle = False
            Set unFeuTmp = unTabFeu(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                If unTabFeu(i).monOrdonnée * unSens > unFeuTmp.monOrdonnée * unSens Then
                    Set unTabFeu(i + 1) = unTabFeu(i)
                    i = i - 1
                Else
                    'Fin de boucle
                    uneFinBoucle = True
                End If
            Loop
            Set unTabFeu(i + 1) = unFeuTmp
    Next j
End Sub


Public Function VerifierVitessePasseToutVert(unSite As Form, uneV As Single, unTabFeu() As Feu, unSensMontant As Boolean) As Single
    'Vérification du passage à tous les feux verts  à la vitesse uneV
    'd'un tableau de feux dans le même sens et trié par Y croisant.
    '
    'Valeur de retour : 0 si la vitesse ne passe pas
    '                   la bande passante trouvée sinon
    
    Dim unFeu As Feu, unFeuTmp As Feu
    Dim unNbFeux As Integer, unPasseToutVert As Boolean
    Dim unCarf As Carrefour
    Dim uneColCarf As New ColCarrefour
    Dim unCarfRed As Object, uneDureeVert As Single
    Dim uneOrdonnee As Integer, unePosRef As Single
    
    unNbFeux = UBound(unTabFeu, 1)
    
    'Création d'un nouveau carrefour avec ses vitesses M et D non nulles
    'qui contiendra tous les feux du tableau de feux et dont on cherchera
    'le feu équivalent en vitesse constante = uneV passée en paramètre.
    Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
    
    'Ajout à ce carrefour global des feux du tableau unTabFeu
    For i = 1 To unNbFeux
        'Récupération du feu et de ses paramètres
        Set unFeu = unTabFeu(i)
        uneOrdonnee = unFeu.monOrdonnée
        uneDureeVert = unFeu.maDuréeDeVert
        unePosRef = unFeu.maPositionPointRef
        
        'Modif du point de référence pour tenir compte du décalage
        'du à l'onde verte en cours
        unePosRef = unePosRef + unFeu.monCarrefour.monDecModif
        
        'Ajout d'un nouveau feu
        Set unFeuTmp = unCarf.mesFeux.Add(unSensMontant, uneOrdonnee, uneDureeVert, unePosRef)
    Next i
    
    'Sauvegarde des paramètres de calcul d'onde avant modif
    unTypeOnde = unSite.monTypeOnde
    unTypeVit = unSite.monTypeVit
    uneVM = unSite.maVitSensM
    uneVD = unSite.maVitSensD
    
    'Changement des paramètres de calcul d'onde pour être en onde
    'double sens à vitesse const = uneV pour trouver le feu équivalent
    unSite.monTypeOnde = OndeDouble
    unSite.monTypeVit = VitConst
    unSite.maVitSensM = uneV
    unSite.maVitSensD = uneV
    
    'Calcul du feu équivalent éventuel comme dans le cas d'une recherche
    'de bandes passantes.
    'Les paramètres de ce feu équivalent seront retournés
    'dans les variables uneDureeVert, unePosRef et uneOrdonnee
    unPasseToutVert = CalculerFeuEquivalent(unCarf, unSensMontant, uneDureeVert, unePosRef, uneOrdonnee, False, True)
    
    'Retour de la bande passante
    If unPasseToutVert Then
        'Bande passante trouvée
        VerifierVitessePasseToutVert = uneDureeVert
    Else
        'Bande passante non trouvée
        VerifierVitessePasseToutVert = 0
    End If
    
    'Restauration des paramètres de calcul d'onde
    unSite.monTypeOnde = unTypeOnde
    unSite.monTypeVit = unTypeVit
    unSite.maVitSensM = uneVM
    unSite.maVitSensD = uneVD
    
    'Suppression des feux,de la collection ne contenant que le carrefour
    'et ce dernier créé dans cette fonction pour libérer la mémoire.
    For i = 1 To unCarf.mesFeux.Count
        unCarf.mesFeux.Remove 1
    Next i
    uneColCarf.Remove 1
    Set unCarf = Nothing
    Set uneColCarf = Nothing
End Function

Public Sub TriCroissantVMax(unTabV() As Single, unTabDT() As Single, unTabDY() As Single, unTabIndFeu() As Integer)
    'Réorganisation par ordre croissant d'un tableau de vitesses
    'Les tableaux unTabDT et unTabDY sont aussi réorganisés
    'pour rester cohérent avec unTabV
    
    'Algo choisi : Le tri insertion (récupérer sur Internet)
    'Il consiste à comparer successivement un élément
    'à tous les précédents et à décaler les éléments intermédiaires

    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer
    Dim uneVtmp As Single, unDYtmp As Single
    Dim unDTtmp As Single, unIndFeuTmp As Integer
    
    'Tri
    unNbTotal = UBound(unTabV, 1)
    For j = 2 To unNbTotal
            uneFinBoucle = False
            uneVtmp = unTabV(j)
            unDTtmp = unTabDT(j)
            unDYtmp = unTabDY(j)
            unIndFeuTmp = unTabIndFeu(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                If unTabV(i) > uneVtmp Then
                    'Réorganisation
                    unTabV(i + 1) = unTabV(i)
                    unTabDT(i + 1) = unTabDT(i)
                    unTabDY(i + 1) = unTabDY(i)
                    unTabIndFeu(i + 1) = unTabIndFeu(i)
                    i = i - 1
                Else
                    'Fin de boucle
                    uneFinBoucle = True
                End If
            Loop
            unTabV(i + 1) = uneVtmp
            unTabDT(i + 1) = unDTtmp
            unTabDY(i + 1) = unDYtmp
            unTabIndFeu(i + 1) = unIndFeuTmp
    Next j
End Sub

Public Function CalculerVMaxInfVMaxLim(unSite As Form, uneVitMinLim As Single, uneVitMaxLim As Single, unTabFeu() As Feu, unTabV() As Single, unTabIndFeu() As Integer, unTabDT() As Single, unTabDY() As Single, unSens As Integer) As Single
    'Calcul de la vitesse max montante si unSens = 1, ou descendante
    ' si unSens = -1, possible < à la vitesse maxi limite
    
    'Cette vitesse est retournée en km/h
    
    Dim unFeuHaut As Feu, unFeuBas As Feu
    Dim unDebVertHaut As Single, unFinVertBas As Single
    Dim uneDatePassage As Single
    
    'Retaillage dynamique des tableaux des vitesses max,
    'des DY et des DT de tous les couples de feux
    'bas et haut
    unNbFeux = UBound(unTabFeu, 1)
    unNbVMax = unNbFeux * (unNbFeux - 1) / 2
    ReDim Preserve unTabV(1 To unNbVMax)
    ReDim Preserve unTabIndFeu(1 To unNbVMax)
    ReDim Preserve unTabDT(1 To unNbVMax)
    ReDim Preserve unTabDY(1 To unNbVMax)
    
    'Calcul des vitesses max entre deux feux du même sens
    unInd = 0
    unNbFeux = UBound(unTabFeu, 1)
    For i = 1 To unNbFeux - 1
        Set unFeuBas = unTabFeu(i)
        For j = i + 1 To unNbFeux
            Set unFeuHaut = unTabFeu(j)
            unInd = unInd + 1
            unTabDY(unInd) = (unFeuHaut.monOrdonnée - unFeuBas.monOrdonnée) * unSens
            
            unDebVertHaut = unFeuHaut.monCarrefour.monDecModif + unFeuHaut.maPositionPointRef
            unDebVertHaut = ModuloZeroCycle(unDebVertHaut, unSite.maDuréeDeCycle)
            
            unFinVertBas = unFeuBas.monCarrefour.monDecModif + unFeuBas.maPositionPointRef + unFeuBas.maDuréeDeVert
            unFinVertBas = ModuloZeroCycle(unFinVertBas, unSite.maDuréeDeCycle)
            
            'Recherche du premier début de vert du feu haut qui est <
            'à la fin de vert bas plus la durée entre les feux haut et bas
            'à la vitesse maxi limite avec une précision de calcul de 0.001
            'seconde et ceci modulo cycle
            unDebVertHautTrouv = False
            Do
                If unDebVertHaut < unFinVertBas + unTabDY(unInd) / uneVitMaxLim - 0.001 Then
                    unDebVertHaut = unDebVertHaut + unSite.maDuréeDeCycle
                Else
                    unDebVertHautTrouv = True
                End If
            Loop Until unDebVertHautTrouv = True
            
            unTabIndFeu(unInd) = i
            unTabDT(unInd) = unDebVertHaut - unFinVertBas
            unTabV(unInd) = unTabDY(unInd) / unTabDT(unInd)
        Next j
    Next i
    
    'Tri par ordre croissant des vitesses max possibles
    'Les tableaux unTabDT et unTabDY sont aussi réorganisés
    'pour rester cohérent avec unTabV
    TriCroissantVMax unTabV, unTabDT, unTabDY, unTabIndFeu
    
    'On essaye toutes les vitesses max possibles en commençant par
    'la plus grande, donc le dernier élément du tableau de vitesse
    '(Indice = unNbVMax)
   Do
        i = 1
        unIndBas = unTabIndFeu(unNbVMax)
        
        'Date de départ du Feu bas
        Set unFeuBas = unTabFeu(unIndBas)
        unFinVertBas = unFeuBas.monCarrefour.monDecModif + unFeuBas.maPositionPointRef + unFeuBas.maDuréeDeVert
        'Test si on passe à tous les feux vert
        unPasseToutVert = True
        Do While unPasseToutVert And i <= unNbFeux
            unPasseAuVert = False
            If i <> unIndBas Then
                'Récup du feu haut
                Set unFeuHaut = unTabFeu(i)
                
                'Début et fin de vert du feu haut
                unDebVertHaut = unFeuHaut.monCarrefour.monDecModif + unFeuHaut.maPositionPointRef
                unFinVertHaut = unDebVertHaut + unFeuHaut.maDuréeDeVert
                
                'Date de passage au feu haut
                uneDatePassage = unFinVertBas + (unFeuHaut.monOrdonnée - unFeuBas.monOrdonnée) * unSens / unTabV(unNbVMax)
                Do While uneDatePassage <= unDebVertHaut - 0.001
                    'Cas d'une date de passage inférieur au début de vert du feu
                    'haut ==> on lui rajoute la Durée du cycle jusqu'à une valeur
                    'supérieure au début de vert pour la prendre en compte dans
                    'les calculs suivants
                    uneDatePassage = uneDatePassage + unSite.maDuréeDeCycle
                Loop
                
                'Vérification si la date de passage est entre le début et
                'la fin de vert du feu haut modulo cycle à une précision de
                'calcul de 0.001
                Do
                    If uneDatePassage > unDebVertHaut - 0.001 And uneDatePassage < unFinVertHaut + 0.001 Then
                        unPasseAuVert = True
                    Else
                        'Incrémentation suivante
                        unDebVertHaut = unDebVertHaut + unSite.maDuréeDeCycle
                        unFinVertHaut = unFinVertHaut + unSite.maDuréeDeCycle
                    End If
                Loop Until uneDatePassage < unDebVertHaut - 0.001 Or unPasseAuVert = True
                
                unPasseToutVert = unPasseAuVert
            Else
                unPasseToutVert = True
            End If
            
            'Incrémentation suivante
            i = i + 1
        Loop
        
        If Not unPasseToutVert Then
            'Modification des éléments d'indice le dernier
            unTabDT(unNbVMax) = unTabDT(unNbVMax) + unSite.maDuréeDeCycle
            unTabV(unNbVMax) = unTabDY(unNbVMax) / unTabDT(unNbVMax)
            'Re-triage par vitesse croissante
            TriCroissantVMax unTabV, unTabDT, unTabDY, unTabIndFeu
        End If
        
        'Boucle jusqu'au passage à tous les verts ou si vitesse est
        '< une vitesse limite mini pour avoir une condition d'arrêt
    Loop Until unPasseToutVert Or unTabV(unNbVMax) < uneVitMinLim
    
    'Retour de la vitesse montante maxi trouvée en km/h
    If unTabV(unNbVMax) < uneVitMinLim Then
        CalculerVMaxInfVMaxLim = uneVitMinLim * 3.6
    Else
        CalculerVMaxInfVMaxLim = unTabV(unNbVMax) * 3.6
    End If
End Function
        
Public Function DonnerIndicePlage(uneColPlageGraphic As Collection, unIndCarf, unIndFeu) As Integer
    'Recherche de la plage graphique sélectionnable représentant le feu
    'd'indice unIndFeu du carrefour d'indice unIndCarf
    'dans une collection contenant cette plage
    'Retourne l'indice trouvé ou 0 si aucune plage ne correspond
    Dim unePlageTrouv As Boolean
    Dim unePlage As PlageGraphic
    
    i = 0
    unePlageTrouv = False
    Do While unePlageTrouv = False And i < uneColPlageGraphic.Count
        i = i + 1
        Set unePlage = uneColPlageGraphic(i)
        If unePlage.monIndCarf = unIndCarf And unePlage.monIndFeu = unIndFeu Then
            unePlageTrouv = True
        End If
    Loop
    
    If unePlageTrouv Then
        DonnerIndicePlage = i
    Else
        DonnerIndicePlage = 0
    End If
End Function

Public Function EstModifierManuel() As Boolean
    'Retourne si les décalages ont été obtenus par modification
    'manuelle
    Dim unCarf As Carrefour
    
    i = 1
    uneModif = False
    Do While i <= monSite.mesCarrefours.Count And uneModif = False
        Set unCarf = monSite.mesCarrefours(i)
        If CInt(unCarf.monDecCalcul) <> CInt(unCarf.monDecModif) Then
            uneModif = True
        End If
        i = i + 1
    Loop
    EstModifierManuel = uneModif
End Function

Public Function CalculerDateDansPhase(unePhase As PhaseTabMarche, unY As Single) As Single
    'Calcul de la date dans une phase donnée suivant son type
    'La phase ne doit pas être de type Arret.
    Dim uneVal As Single
    
    If unePhase.monType = VConst Then
        CalculerDateDansPhase = unePhase.monTDeb + Abs((unY - unePhase.monYDeb) / unePhase.maVitPhase)
    ElseIf unePhase.monType = Accel Then
        'Arrondi de valeur à 1 par rapport à la précision 0.001 pour
        'éviter d'avoir un Sqr de uneVal avec uneVal voisin de 0 mais < 0
        '(unY - unePhase.monYDeb) est forcément > ou = 0
        '==> Problème si = ça doit faire 0 et pas 0.000000023 ou -0.000000002
        uneVal = (unY - unePhase.monYDeb) / unePhase.maLongPhase
        If Abs(uneVal) < 0.001 Then uneVal = 0
        'Calcul de la date
        CalculerDateDansPhase = unePhase.monTDeb + unePhase.maDureePhase * Sqr(uneVal)
    ElseIf unePhase.monType = Decel Then
        'Arrondi de valeur à 1 par rapport à la précision 0.001 pour
        'éviter d'avoir un Sqr de (1-uneVal) avec uneVal voisin de 1 mais < 1
        '(unY - unePhase.monYDeb) est forcément < ou = unePhase.maLongPhase
        '==> Problème si = ça doit faire 1 et pas 1.00023 ou 0.9999982
        uneVal = (unY - unePhase.monYDeb) / unePhase.maLongPhase
        If Abs(uneVal - 1) < 0.001 Then uneVal = 1
        'Calcul de la date
        CalculerDateDansPhase = unePhase.monTDeb + unePhase.maDureePhase * (1 - Sqr(1 - uneVal))
    Else
        MsgBox "ERREUR de programmation dans OndeV dans CalculerDateDansPhase", vbCritical
    End If
End Function

Public Function CalculerYDansPhaseParabole(unePhase As PhaseTabMarche, unT As Single) As Single
    'Calcul du Y à l'instant unT dans une phase donnée se dessinant en
    'discrétisant une parabole
    'La phase ne peut être que du type Accel ou Decel.
    
    'Formules de calcul obtenues par inversion de celles
    'de la fonction CalculerDateDansPhase
    Dim unDT As Single
    
    'Calcul de l'écart en temps par rapport au début de la phase
    unDT = unT - unePhase.monTDeb
    
    If unePhase.monType = Accel Then
        'Calcul du Y
        CalculerYDansPhaseParabole = unePhase.monYDeb + unePhase.maLongPhase * unDT * unDT / unePhase.maDureePhase / unePhase.maDureePhase
    ElseIf unePhase.monType = Decel Then
        CalculerYDansPhaseParabole = unePhase.monYDeb + unePhase.maLongPhase * (1 - (1 - unDT / unePhase.maDureePhase) * (1 - unDT / unePhase.maDureePhase))
    Else
        MsgBox "ERREUR de programmation dans OndeV dans CalculerYDansPhaseParabole", vbCritical
    End If
End Function


Public Sub AfficherPropsObjPick()
    If monTypeObjPick = NoSel Then
        MsgBox "Aucun objet n'a été sélectionné", vbInformation
    Else
        frmPropsObjPick.Show vbModal
    End If
End Sub

Public Sub ViderObjPick()
    'Mise à vide de l'objet sélectionné graphiquement
    monTypeObjPick = NoSel
    Set monObjPick = Nothing
End Sub

Public Function DonnerObjPick() As Object
    Set DonnerObjPick = monObjPick
End Function

Public Sub CalculerB1B2(K As Single, A1 As Single, A2 As Single, B1 As Single, B2 As Single, unBB1 As Single, unBB2 As Single)
    'Cette procédure alimente les variables unBB1 et unBB2 passées en paramètres
    If K < 0 Then
        unBB1 = A1
        unBB2 = A2
    ElseIf K > 1 Then
        unBB1 = B1
        unBB2 = B2
    Else
        unBB1 = (1 - K) * A1 + K * B1
        unBB2 = (1 - K) * A2 + K * B2
    End If
End Sub

Public Sub IndiquerModifTC()
    'Indique si une modif a eu lieu dans les données TC ne cadrant pas l'onde
    'ou les données des TC cadrant l'onde
    Dim unIndTC As Integer
    
    'Récupération du TC modifié
    unIndTC = monSite.ComboTC.ListIndex + 1
    
    'Indication de la bonne modif
    If monSite.monTypeOnde = OndeTC And (monSite.monTCM = unIndTC Or monSite.monTCD = unIndTC) Then
        monSite.maModifDataOndeTC = True
    Else
        monSite.maModifDataTC = True
    End If
End Sub

Public Function UtiliserDecalagesImposes(unIndUniqCarfImp, unNbFeuDateImpSensM, unNbFeuDateImpSensD) As Object
    'Fonction préparant le calcul d'onde verte avec des décalages
    'imposés à certains carrefours
    'Elle retourne le carrefour réduit réduisant tous les carrefours
    'à date imposée et alimente unIndUniqCarfImp avec l'indice du seul
    'carrefour à date imposé si c'est le cas, sinon il vaut 0
    'Elle retourne aussi le nombre de feux à date imposés dans le sens montant
    'et descendant
    
    'Variables locales
    Dim uneColCarf As ColCarrefour
    Dim unCarfRed As Object, unCarfRedTmp As Object
    Dim unCarf As New Carrefour
    Dim unCarfRedU As CarfReduitSensUnique
    Dim unCarfRed2 As CarfReduitSensDouble
    
    Set unCarfRed = Nothing
    Set uneColCarf = TrouverLesCarfsAvecDateImp
    
    If uneColCarf.Count = 1 Then
        'Cas d'un seul carrefour à date imposée
        unIndUniqCarfImp = uneColCarf(1).maPosition
    Else
        'Autres cas
        unIndUniqCarfImp = 0
    End If
    
    If uneColCarf.Count Then
        'Cas où il y a des carrefours à date imposée
        
        'Initialisation du message d'erreur
        unMsg = "Impossible de calculer les ondes vertes avec les valeurs actuelles des décalages imposés." + Chr(13)
        unMsg = unMsg + Chr(13) + "Le calcul des ondes vertes ne tiendra pas compte des décalages imposés aux carrefours."
        
        'Calcul des temps de parcours sans s'occuper des dates imposées
        CalculerTempsParcours monSite

        'Réduction de tous les carrefours réduits à date imposée en un seul
        'L'ajout du décalage des carrefours à date imposée à leur point de
        'référence de leur carrefour réduit est fait dans CalculerFeuEquivalent
        'appelée dans ReduireCarfsEnUn avec le paramètre unDecalModif = TRUE
        'Cet ajout est décrit dans le dossier de spécifs, partie Date imposée
        Set unCarfRed = ReduireCarfsEnUn(uneColCarf, unCarf, unIndCarfBas, unIndCarfHaut, unNbFeuDateImpSensM, unNbFeuDateImpSensD)
        
        'Calcul des décalages en temps aux carrefours le plus haut
        'pour le sens montant et le plus bas pour le sens descendant
        unCarf.monDecVitSensD = monSite.mesCarrefours(unIndCarfBas).monDecVitSensD
        unCarf.monDecVitSensM = monSite.mesCarrefours(unIndCarfHaut).monDecVitSensM
        
        If unCarfRed Is Nothing Then
            'Cas où la réduction des carrefours réduits des carrefours à date
            'imposée n'a pas marché
            MsgBox unMsg, vbInformation
        Else
            'Lien à un nouveau carrefour du carrefour réduit créé ci-dessus
            'Ce carrefour servira à stocker le décalage calculé
            'et il est différent des carrefours à date imposée
            Set unCarf.monCarfRed = unCarfRed
            unCarf.monNom = ""
            'Aucun carrefour créé par saisie ne peut avoir de nom vide
            
            If TypeOf unCarfRed Is CarfReduitSensDouble Then
                'Cas double sens
                
                'Calcul de l'écart
                'l'écart est le temps s'écoulant entre les événements "passage au vert
                'dans le sens montant" et "fin du vert dns le sens descendant" après
                'projection sur une référence commune à l'ensemble des carrefours
                '(cf Dossier de programmation et spécifs)
                'On utilise des décalages dus aux vitesses variables ou
                'constantes de chaque carrefour
                unCarfRed.monEcart = unCarfRed.maPosRefD + unCarf.monDecVitSensD + unCarfRed.maDureeVertD
                unCarfRed.monEcart = unCarfRed.monEcart - (unCarfRed.maPosRefM - unCarf.monDecVitSensM)
                'On ramène l'écart modulo entre [0, duréee du cycle[
                unCarfRed.monEcart = ModuloZeroCycle(unCarfRed.monEcart, monSite.maDuréeDeCycle)
            End If
            
            'Suppression des carrefours réduits du site courant issus
            'd'un carrefour à date imposée, ces derniers sont dans uneColCarf
            unN = monSite.mesCarfReduitsSens2.Count
            For i = unN To 1 Step -1
                Set unCarfRedTmp = monSite.mesCarfReduitsSens2(i)
                If uneColCarf.HasCarfRed(unCarfRedTmp) Then
                    monSite.mesCarfReduitsSens2.Remove i
                End If
            Next i
            unN = monSite.mesCarfReduitsSensM.Count
            For i = unN To 1 Step -1
                Set unCarfRedTmp = monSite.mesCarfReduitsSensM(i)
                If uneColCarf.HasCarfRed(unCarfRedTmp) Then
                    monSite.mesCarfReduitsSensM.Remove i
                End If
            Next i
            unN = monSite.mesCarfReduitsSensD.Count
            For i = unN To 1 Step -1
                Set unCarfRedTmp = monSite.mesCarfReduitsSensD(i)
                If uneColCarf.HasCarfRed(unCarfRedTmp) Then
                    monSite.mesCarfReduitsSensD.Remove i
                End If
            Next i
        End If
    End If

    'Stockage de la valeur de retour
    Set UtiliserDecalagesImposes = unCarfRed
    
    'Libération de la mémoire
    For i = 1 To uneColCarf.Count
        uneColCarf.Remove 1
    Next i
    Set uneColCarf = Nothing
End Function


Public Function ReduireCarfsEnUn(uneColCarf As ColCarrefour, unCarf As Carrefour, unIndCarfBas, unIndCarfHaut, unNbFeuDateImpSensM, unNbFeuDateImpSensD) As Object
    'Réduction des carrefours de la collection uneColCarf qui contient tous
    'les feux équivalents montant et descendant des carrefours réduits
    
    'Elle retourne un carrefour réduit valant :
    '   - nothing si aucun feu équivalent trouvé
    '   - de type CarfReduitSensUnique si un feu équivalent trouvé (montant ou descendant)
    '   - de type CarfReduitSensDouble si deux feux équivalents trouvés (montant et descendant)
    '
    'Les variables unIndCarfBas et unIndCarfHaut, donnant les indices des
    'carrefours le plus bas et le plus haut de uneColCarf, sont renseignés
    'par cette fonction pour ainsi récupérer leurs décalages en temps respectif
    '
    'Elle Retourne aussi le nombre de feux à date imposée dans les sens M et D
    
    Dim unFeu As Feu, unTCM As TC, unTCD As TC
    Dim unCarfRed As Object
    Dim unNbFeuxM As Integer, unNbFeuxD As Integer
    Dim unCarfRedU As CarfReduitSensUnique
    Dim unCarfRed2 As CarfReduitSensDouble
    Dim uneDureeVertM As Single, unPosRefM As Single, uneOrdonneeM As Integer
    Dim uneDureeVertD As Single, unPosRefD As Single, uneOrdonneeD As Integer
    
    'Initialisation du nombre de feux montants et descendants du carrefour
    'réduit que l'on va créer
    unNbFeuxM = 0
    unNbFeuxD = 0
    
    'Création d'une collection de feu pour le carrefour temporaire stockant
    'le carrefour réduisant touts les carrefours à date imposée
    Set unCarf.mesFeux = New ColFeu
    
    'Initialisation pour la recherche du carrefour le plus haut et le plus bas
    'Y dans OndeV compris entre -9999 et 9999 mètres
    unYMin = 100000
    unYMax = -100000
    
    'Initialisation des indices des carf le plus haut et le plus bas
    unNbCarfImp = uneColCarf.Count
    unIndCarfHaut = uneColCarf(unNbCarfImp).maPosition
    unIndCarfBas = unIndCarfHaut
    
    'Ajout à ce carrefour global des feux équivalents
    'des carrefours réduits double sens
    For i = 1 To unNbCarfImp
        'Récup du carrefour réduit
        Set unCarfRed = uneColCarf(i).monCarfRed
        
        'Recherche du carrefour le plus haut ayant un feu montant
        'et du carrefour le plus bas ayant un feu descendant
        unY = DonnerYCarrefour(uneColCarf(i))
        If unY > unYMax And unCarfRed.HasFeuMontant Then
            unYMax = unY
            unIndCarfHaut = uneColCarf(i).maPosition
        End If
        If unY < unYMin And unCarfRed.HasFeuDescendant Then
            unYMin = unY
            unIndCarfBas = uneColCarf(i).maPosition
        End If
        
        'Création des feux du carrefour dont la réduction, donc recherche des
        'feux équivalent montant et descendant donne le carrefour réduit
        'réduisant tous les carrefours à date imposée
        
        'On stockera dans le feu créé, le carrefour correspondant à celui réduit
        'car c'est son décalage modifié (= monDecModif) et celui en temps du
        'à sa vitesse (= monDecVitSensM ou D) qui est utilisé dans le calcul du
        'feu équivalent montant ou descendant
        If TypeOf unCarfRed Is CarfReduitSensDouble Then
            'Ajout d'un nouveau feu montant
            unNbFeuxM = unNbFeuxM + 1
            'Stockage de l'indice du carrefour de ce feu M
            unIndCarfM = i
            Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonneeM, unCarfRed.maDureeVertM, unCarfRed.maPosRefM)
            'Stockage du carrefour de celui réduit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Ajout d'un nouveau feu descendant
            unNbFeuxD = unNbFeuxD + 1
            Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonneeD, unCarfRed.maDureeVertD, unCarfRed.maPosRefD)
            'Stockage du carrefour de celui réduit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu D
            unIndCarfD = i
        ElseIf TypeOf unCarfRed Is CarfReduitSensUnique And unCarfRed.HasFeuMontant Then
            'Ajout d'un nouveau feu montant
            unNbFeuxM = unNbFeuxM + 1
            Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
            'Stockage du carrefour de celui réduit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu M
            unIndCarfM = i
        ElseIf TypeOf unCarfRed Is CarfReduitSensUnique And unCarfRed.HasFeuDescendant Then
            'Ajout d'un nouveau feu descendant
            unNbFeuxD = unNbFeuxD + 1
            Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
            'Stockage du carrefour de celui réduit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu D
            unIndCarfD = i
        Else
            MsgBox "ERREUR de programmation dans OndeV dans ReduireCarfsEnUn", vbCritical
        End If
    Next i
          
    'Retour du nombre de feux à date imposée dans les sens M et D
    unNbFeuDateImpSensM = unNbFeuxM
    unNbFeuDateImpSensD = unNbFeuxD
    
    'Prise en compte d'une onde cadrée par un TC montant et/ou descendant
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde cadrée par unTC dans le sens montant
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde non cadrée par unTC dans le sens montant
        Set unTCM = Nothing
    End If
    
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde cadrée par unTC dans le sens descendant
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde non cadrée par unTC dans le sens descendant
        Set unTCD = Nothing
    End If
        
    'Récupération de la vitesse d'arrivée sur les carrefours le plus
    'haut pour le sens montant et le plus bas pour le sens descendant,
    'car elles servent dans le calcul du point de référence dans la
    'fonction CalculerFeuEquivalent, dans le carrefour du carrefour
    'réduisant tous les carrefours à date imposée
    unCarf.maVitSensD = monSite.mesCarrefours(unIndCarfBas).maVitSensD
    unCarf.maVitSensM = monSite.mesCarrefours(unIndCarfHaut).maVitSensM
                
    'Calcul du feu équivalent montant éventuel
    'avec unDecalModif = True (6ème paramètre)
    unFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVertM, unPosRefM, uneOrdonneeM, True, , , , unTCM)
    If unNbFeuxM = 1 And (unNbFeuxD > 1 Or (unNbFeuxD = 1 And unIndCarfM <> unIndCarfD)) Then
        'Cas où le feu montant équivalent trouvé est l'équivalent d'un seul feu
        'montant mais avec un feu descendant équivalent qui est celui de plusieurs
        'feux descendants ou si le seul feu descendant est celui d'un autre carrefour à sens unique D
        '==> Rajout du décalage du carrefour le plus haut pour
        'lier ces deux feux équivalents car le feu M contrairement au feu D n'a
        'pas ce décalage intégré
        unPosRefM = unPosRefM + monSite.mesCarrefours(unIndCarfHaut).monDecModif
    End If
                      
    'Calcul du feu équivalent descendant éventuel
    'avec unDecalModif = True (6ème paramètre)
    unFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unPosRefD, uneOrdonneeD, True, , , , unTCD)
    If unNbFeuxD = 1 And (unNbFeuxM > 1 Or (unNbFeuxM = 1 And unIndCarfM <> unIndCarfD)) Then
        'Cas où le feu descendant équivalent trouvé est l'équivalent d'un seul feu
        'descendant mais avec un feu montant équivalent qui est celui de plusieurs
        'feux montants ou si le seul feu montant est celui d'un autre carrefour à sens unique M
        '==> Rajout du décalage du carrefour le plus bas pour
        'lier ces deux feux équivalents car le feu D contrairement au feu M n'a
        'pas ce décalage intégré
        unPosRefD = unPosRefD + monSite.mesCarrefours(unIndCarfBas).monDecModif
    End If
    
    If unFeuEquivSensMExist And unFeuEquivSensDExist Then
        'Ajout aux carrefours réduits double sens du site courant
        Set unCarfRed2 = monSite.mesCarfReduitsSens2.Add(unCarf)
        'Alimentation d'un carrefour réduit double sens
        unCarfRed2.SetPropsSensM uneDureeVertM, unPosRefM, uneOrdonneeM
        unCarfRed2.SetPropsSensD uneDureeVertD, unPosRefD, uneOrdonneeD
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRed2
    ElseIf unFeuEquivSensMExist Then
        'Ajout aux carrefours réduits sens unique montant du site courant
        Set unCarfRedU = monSite.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVertM, unPosRefM, uneOrdonneeM)
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRedU
    ElseIf unFeuEquivSensDExist Then
        'Ajout aux carrefours réduits sens unique descendant du site courant
        Set unCarfRedU = monSite.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVertD, unPosRefD, uneOrdonneeD)
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRedU
    Else
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = Nothing
    End If
        
    'Libération de la mémoire
    For i = 1 To unCarf.mesFeux.Count
        unCarf.mesFeux.Remove 1
    Next i
    Set unCarf.mesFeux = Nothing
End Function

Public Function TrouverLesCarfsAvecDateImp() As ColCarrefour
    'Retourne une collection contenant les carrefours à décalage imposé
    Dim uneColCarf As New ColCarrefour
    Dim unCarf As Carrefour
    
    For i = 1 To monSite.mesCarrefours.Count
        Set unCarf = monSite.mesCarrefours(i)
        If unCarf.monDecImp = 1 Then uneColCarf.AjouterCarf unCarf
    Next i

    Set TrouverLesCarfsAvecDateImp = uneColCarf
End Function

Public Sub TestForDebug()
    Debug.Print
    Debug.Print "********************* TestForDebug *****************"
    unNbFeuxSens2 = monSite.mesCarfReduitsSens2.Count
    For i = 1 To unNbFeuxSens2
        'Récup du carrefour réduit
        Set unCarfRed = monSite.mesCarfReduitsSens2(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : Décalage = "; unCarfRed.monCarrefour.monDecModif
        Debug.Print Tab; "==> Sens M : "; unCarfRed.monOrdonneeM; unCarfRed.maDureeVertM; unCarfRed.maPosRefM; unCarfRed.monCarrefour.monDecVitSensM
        Debug.Print Tab; "==> Sens D : "; unCarfRed.monOrdonneeD; unCarfRed.maDureeVertD; unCarfRed.maPosRefD; unCarfRed.monCarrefour.monDecVitSensD
    Next i
    
    unNbFeuxSensM = monSite.mesCarfReduitsSensM.Count
    For i = 1 To unNbFeuxSensM
        'Récup du carrefour réduit
        Set unCarfRed = monSite.mesCarfReduitsSensM(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : Décalage = "; unCarfRed.monCarrefour.monDecModif; unCarfRed.monOrdonnee; unCarfRed.maDureeVert; unCarfRed.maPosRef
    Next i
    
    'Ajout à ce carrefour global des feux équivalents
    'des carrefours réduits à sens unique descendant
    unNbFeuxSensD = monSite.mesCarfReduitsSensD.Count
    For i = 1 To unNbFeuxSensD
        'Récup du carrefour réduit
        Set unCarfRed = monSite.mesCarfReduitsSensD(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : Décalage = "; unCarfRed.monCarrefour.monDecModif; unCarfRed.monOrdonnee; unCarfRed.maDureeVert; unCarfRed.maPosRef
    Next i
    Debug.Print "****************** Fin de TestForDebug **************"
End Sub

Public Sub RecalculerAvecDateImp(unCarf As Carrefour, unText As String)
    'Lance le recalcul avec les dates imposées si la valeur de la variable
    'unText (valant Oui pour date imposé ou Non sinon ) est différent du
    'type de décalage imposé (1 = imposé, sinon 0)
    
    If unCarf.monDecImp = 1 And unText = "Non" Then
        'Modification du type de décalage (0 pour non imposé)
        unCarf.monDecImp = 0
        'Indication d'un changement pour lancer le calcul
        monSite.maModifDataOnde = True
        'Calcul d'onde verte en tenant compte des décalages imposés
        CalculerOndeVerte monSite, True
    ElseIf unCarf.monDecImp = 0 And unText = "Oui" Then
        'Modification du type de décalage (1 pour imposé)
        unCarf.monDecImp = 1
        'Indication d'un changement pour lancer le calcul
        monSite.maModifDataOnde = True
        'Calcul d'onde verte en tenant compte des décalages imposés
        CalculerOndeVerte monSite, True
    End If
End Sub

Public Sub CorrectionDateImposée(uneForm As Form, unCarfRed As Object, unB1 As Single, unB2 As Single, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD)
    'Correction de la solution à date imposée si on obtient un carrefour
    'réduisant tous les carrefours à date imposée qui est à sens unique
    'alors que ces carrefours ont des feux dans les 2 sens, donc il n'y a pas
    'de bandes passantes dans le sens opposé à celui du carrefour réduit
    
    If TypeOf unCarfRed Is CarfReduitSensUnique Then
        If unCarfRed.monSensMontant And unNbFeuxDateImpSensD > 0 Then
            unB2 = 0
            uneForm.monOndeDoubleTrouve = False
            unMsg = "Une solution a été trouvée dans le sens Montant, mais pas dans le sens Descendant"
            MsgBox unMsg, vbInformation
        ElseIf unCarfRed.monSensMontant = False And unNbFeuxDateImpSensM > 0 Then
            unB1 = 0
            uneForm.monOndeDoubleTrouve = False
            unMsg = "Une solution a été trouvée dans le sens Descendant, mais pas dans le sens Montant"
            MsgBox unMsg, vbInformation
        End If
    End If
End Sub

Public Sub DessinerBandesInterCarfVP(uneZoneDessin As Object, unTM1 As Single, unTD1 As Single, unY0 As Long, uneHt As Long, unDY As Long, unT As Long, unCarfRedM1, unCarfRedD1, unYMin As Long, unIndCarfBasM, unIndCarfD, uneLg As Long, unNewX0, unMaxT)
    'Dessin des bandes inter-carrefours dans le sens M ou D suivant le cas.
    'Utilisation des Y sur vitesse pour avoir les décalages inter-carrefours
    'Utiliser si choix coché dans le menu Montrer les bandes inter-carrefours
    Dim unY As Long
    Dim unCarf As Carrefour, unCarfRed As Object
    Dim unCarfPred As Object
    Dim unDebVertM As Single, unDebVertD As Single
    Dim unFinVertM As Single, unFinVertD As Single
    Dim unMaxDebVert As Single, unMinFinVert As Single
    Dim unDebVertMPred As Single, unDebVertDPred As Single
    Dim unFinVertMPred As Single, unFinVertDPred As Single
    Dim uneDuréeVertMPred As Single, uneDuréeVertDPred As Single
    Dim unTmpInterCarf As Single
    Dim unLastDebVertM As Single, unLastDebVertD As Single
    
    'Sauvegarde du style de dessin
    unDrawStyleSave = uneZoneDessin.DrawStyle
    'Dessin en pointillé
    uneZoneDessin.DrawStyle = vbDash
    
    unNbCarf = UBound(monTabCarfY, 1)
    With monSite
        For i = 1 To unNbCarf
            'Parcours des carrefours dans le sens des Y croissants
            'pour l'onde verte montante car on dessine à partir du
            'carrefour le plus bas ayant un feu montant
            Set unCarfRed = monTabCarfY(i).monCarfReduit
            Set unCarf = unCarfRed.monCarrefour
            If unIndCarfBasM > 0 Then
                'Cas d'une onde verte montante possible ==> Dessin
                'unIndCarfBasM > 0 dit qu'on a trouvé des carrefours montants
                If unCarfRed.HasFeuMontant = True Then
                    'Cas d'un carrefour contraignant de l'onde verte montante
                    'donc ayant un feu de sens montant
                    
                    'Calcul du temps de parcours inter-carrefours qui vaut
                    '(YcarfCourant - YcarfPred)/ vitesse du carf courant
                    If i = unIndCarfBasM Then
                        unTmpInterCarf = 0
                        unTM2 = unTM1
                    Else
                        Set unCarfPred = monTabCarfY(unIndCarfMPred).monCarfReduit
                        unTmpInterCarf = (unCarfRed.DonnerYSens(True) - unCarfPred.DonnerYSens(True)) / unCarf.DonnerVitCarfSens(True)
                    End If
                    
                    'Abscisse du point suivant de l'onde montante vaut
                    'l'abscisse du point précédent plus le décalage en
                    'temps entre le carrefour courant et le précédent montant
                    unTM2 = unDebVertMPred + unTmpInterCarf
                    
                    'Ordonnée égale à l'ordonnée du carrefour réduit courant
                    'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                    unYM2 = unCarfRed.DonnerYSens(True)
                    
                    'Conversion en coordonnées écran de unYM2
                    unY = ConvertirReelEnEcran(unYM2 - unYMin, unDY, uneHt)
                    unY = unY0 - unY
                    
                    'Dessin de l'onde verte montante inter-carrefours donc
                    'entre ce carrefour réduit et son précédent en Y
                    'et si ce n'est pas le 1er carrefour montant le + bas
                    
                    If i <> unIndCarfBasM Then
                        'Cas d'une onde montante
                        'Calcul du début de vert de ce carrefour réduit
                        unDebVertM = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(True)
                        unDebVertM = ModuloZeroCycle(unDebVertM, .maDuréeDeCycle)
                        'Calcul du nombre de cycle séparant le début de
                        'vert du début de l'onde verte montante
                        unNbCycle = Fix((0.001 + unTM2 - unDebVertM) / .maDuréeDeCycle)
                        If unNbCycle < 0 And .maBandeModifM = 0 Then
                            'Si pas de bande commune, on ne peut pas être en retard
                            'unTM2 ne doit pas être corrigé s'il est < unDebVertM
                            unNbCycle = 0
                        End If
                        If unTM2 < unDebVertM - 0.001 Then
                            'Début de vert > T de onde montante
                            '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                            unDebVertM = unDebVertM + unNbCycle * .maDuréeDeCycle
                        ElseIf unTM2 > unDebVertM + unCarfRed.DonnerDureeVertSens(True) + 0.001 Then
                            'Fin de vert < T de départ onde montante
                            '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                            unDebVertM = unDebVertM + unNbCycle * .maDuréeDeCycle
                        End If
                                                  
                        'Calcul de la fin de vert de ce carrefour réduit
                        unFinVertM = unDebVertM + unCarfRed.DonnerDureeVertSens(True)
                        RecalculerDebEtFinVert unCarfRed, True, .maDuréeDeCycle, unCarf.DonnerVitCarfSens(True), unDebVertM, unFinVertM
                        unFinVertMPred = unDebVertMPred + uneDuréeVertMPred
    
                        unI = 0
                        unJ = 0
                        unLastDebVertDéjàStocké = False
                        'projection du début et fin de vert du carrefour
                        'précédent sur la droite Y = Y du carrefour courant
                        unDebVertMPred = unDebVertMPred + unTmpInterCarf
                        unFinVertMPred = unFinVertMPred + unTmpInterCarf
                        'Initialisation du LastDebVert au cas où aucun
                        'bande inter-carf trouvée
                        unLastDebVertM = unDebVertM
                        Do
                            'La boucle sert pour afficher toutes les bandes
                            'inter-carrefour pour cela on regarde dans
                            'le cycle en cours et le suivant et on prend la
                            'bande inter-carf maximale
                            
                            'On prend le minimun des fins de vert projeté
                            'sur la droite Y = Y du carrefour courant
                            If (unFinVertMPred + unJ * .maDuréeDeCycle) < (unFinVertM + unI * .maDuréeDeCycle) Then
                                unMinFinVert = unFinVertMPred + unJ * .maDuréeDeCycle
                            Else
                                unMinFinVert = unFinVertM + unI * .maDuréeDeCycle
                            End If
                            
                            'On prend le maximun des débuts de vert projeté
                            'sur la droite Y = Y du carrefour courant
                            If (unDebVertMPred + unJ * .maDuréeDeCycle) > (unDebVertM + unI * .maDuréeDeCycle) Then
                                unMaxDebVert = unDebVertMPred + unJ * .maDuréeDeCycle
                            Else
                                unMaxDebVert = unDebVertM + unI * .maDuréeDeCycle
                            End If
                            
                            'Test de l'existence d'une bande inter-carrefour
                            'supérieure à 1 seconde
                            uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                            
                            If uneBandeInterCarfExist Then
                                If unLastDebVertDéjàStocké = False Then
                                    'Stockage du dernier debvert ayant une bande
                                    'inter-carrefour, ce stockage est fait une fois et
                                    'une seule entre deux carrefours
                                    unLastDebVertDéjàStocké = True
                                    unLastDebVertM = unDebVertM + unI * .maDuréeDeCycle
                                End If
                                'Remise dans l'englobant total si
                                'MinFinVert en sort
                                If unMinFinVert > unMaxT + 0.01 Then
                                    unMinFinVert = unMinFinVert - .maDuréeDeCycle
                                    unMaxDebVert = unMaxDebVert - .maDuréeDeCycle
                                End If
                            End If
                            
                            unI = unI + 1
                            If unI = 2 Then
                                'On se place pour essayer les début et fin de vert
                                'du carrefour courant dans le cycle courant avec les
                                'début et fin de vert du carrefour précédent dans le
                                'cycle suivant
                                unI = 0
                                unJ = 1
                            End If
    
                            'Dessin de bande inter-carrefour
                            If uneBandeInterCarfExist Then
                                'Cas du dessin des bandes inter-carrefours
                                'voitures d'une onde TC
                                
                                'Conversion en coordonnées écran
                                unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 1ère partie bande montante inter-carrefours
                                uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                
                                'Conversion en coordonnées écran
                                unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 2ème partie bande montante inter-carrefours
                                uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                            End If
                            'Boucle fait trois fois pour trouver toutes
                            'les bande inter-carrefour
                        Loop Until unI = 1 And unJ = 1
                    End If 'Fin du dessin de bande verte montante inter-carrefours
                             
                   'Stockage du début de vert précédent
                    If i = unIndCarfBasM Then
                        'Calcul spécial pour le carrefour le + bas montant
                        unDebVertMPred = unCarfRedM1.monCarrefour.monDecModif + unCarfRedM1.DonnerPosRefSens(True)
                        unDebVertMPred = ModuloZeroCycle(unDebVertMPred, .maDuréeDeCycle)
                        If unTM1 < unDebVertMPred - 0.001 Then
                            'Début de vert > T de départ onde montante
                            '==> Recul d'un cycle
                            unDebVertMPred = unDebVertMPred - .maDuréeDeCycle
                        ElseIf unTM1 > unDebVertMPred + unCarfRedM1.DonnerDureeVertSens(True) + 0.001 Then
                            'Fin de vert < T de départ onde montante
                            '==> Avancé d'un cycle
                            unDebVertMPred = unDebVertMPred + .maDuréeDeCycle
                        End If
                        'Initialisation de la durée de vert du premier
                        'carrefour descendant
                        uneDuréeVertMPred = unCarfRedM1.DonnerDureeVertSens(True)
                    Else
                        'Affectation du début de vert du carf réduit précédent
                        unDebVertMPred = unLastDebVertM
                        'Affectation de la durée de vert du carf réduit précédent
                        'à faire car modifs possibles dans RecalculerDebEtFinVert
                        uneDuréeVertMPred = unFinVertM - unDebVertM
                    End If
                    
                    'Stockage de l'indice de ce carrefour
                    unIndCarfMPred = i
                    'Stockage du Y écran  du point précédent pour le coup suivant
                    unYMpred = unY
                End If
            End If
            
            'Parcours des carrefours dans le sens des Y décroissants
            'pour l'onde verte descendante car on dessine à partir du
            'carrefour le plus haut ayant un feu descendant
            Set unCarfRed = monTabCarfY(unNbCarf + 1 - i).monCarfReduit
            Set unCarf = unCarfRed.monCarrefour
            
            If unIndCarfD > 0 Then
                'Cas d'une onde verte descendante possible ==> Dessin
                'unIndCarfD > 0 dit qu'on a trouvé des carrefours descendants
                If unCarfRed.HasFeuDescendant = True Then
                    'Cas d'un carrefour contraignant de l'onde verte descendante
                    'donc ayant un feu de sens descendant
                
                    'Calcul du temps de parcours inter-carrefours qui vaut
                    '(Y carf courant - Y carf précédent) / vitesse carf courant
                    'Ce temps est > car la différence des Y est < 0
                    '(car les carrefours sont parcourus dans le sens des Y
                    'décroissants pour le sens descendant et les Vitesses
                    'descendantes < 0
                    If (unNbCarf + 1 - i) = unIndCarfD Then
                        unTmpInterCarf = 0
                        unTD2 = unTD1
                    Else
                        'Calcul du tmps de parcours inter-carrefour
                        Set unCarfPred = monTabCarfY(unIndCarfDPred).monCarfReduit
                        unTmpInterCarf = (unCarfRed.DonnerYSens(False) - unCarfPred.DonnerYSens(False)) / unCarf.DonnerVitCarfSens(False)
                   End If
                    
                    'Abscisse du point suivant de l'onde descendante vaut
                    'l'abscisse du point précédent plus le décalage en
                    'temps entre le carrefour courant et le précédent descendant
                    unTD2 = unDebVertDPred + unTmpInterCarf
                  
                    'Ordonnée égale à l'ordonnée du carrefour réduit courant
                    'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                    unYD2 = unCarfRed.DonnerYSens(False)
                    
                    'Conversion en coordonnées écran de unYD2
                    unY = ConvertirReelEnEcran(unYD2 - unYMin, unDY, uneHt)
                    unY = unY0 - unY
                    
                    'Dessin de l'onde verte descendante inter-carrefours donc
                    'entre ce carrefour réduit et son précédent en Y
                    'et si ce n'est pas le 1er carrefour descendante le + haut
                    
                    If (unNbCarf + 1 - i) <> unIndCarfD Then
                        'Calcul du début de vert de ce carrefour réduit
                        unDebVertD = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(False)
                        unDebVertD = ModuloZeroCycle(unDebVertD, .maDuréeDeCycle)
                        'Calcul du nombre de cycle séparant le début de
                        'vert du début de l'onde verte descendante
                        unNbCycle = Fix((0.001 + unTD2 - unDebVertD) / .maDuréeDeCycle)
                        If unNbCycle < 0 And .maBandeModifD = 0 Then
                            'Si pas de bande commune, on ne peut pas être en retard
                            'unTD2 ne doit pas être corrigé si il est < unDebVertD
                            unNbCycle = 0
                        End If
                        If unTD2 < unDebVertD - 0.001 Then
                            'Début de vert > T de onde descendante
                            '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                            unDebVertD = unDebVertD + unNbCycle * .maDuréeDeCycle
                        ElseIf unTD2 > unDebVertD + unCarfRed.DonnerDureeVertSens(False) + 0.001 Then
                            'Fin de vert < T de départ onde descendante
                            '==> Recul ou Avancé d'un nombre entier cycle dépendant du temps de parcours
                            unDebVertD = unDebVertD + unNbCycle * .maDuréeDeCycle
                        End If
                                                  
                        'Calcul de la fin de vert de ce carrefour réduit
                        unFinVertD = unDebVertD + unCarfRed.DonnerDureeVertSens(False)
                        RecalculerDebEtFinVert unCarfRed, False, .maDuréeDeCycle, unCarf.DonnerVitCarfSens(False), unDebVertD, unFinVertD
                        unFinVertDPred = unDebVertDPred + uneDuréeVertDPred
    
                        unI = 0
                        unJ = 0
                        unLastDebVertDéjàStocké = False
                        'projection du début et fin de vert du carrefour
                        'précédent sur la droite Y = Y du carrefour courant
                        unDebVertDPred = unDebVertDPred + unTmpInterCarf
                        unFinVertDPred = unFinVertDPred + unTmpInterCarf
                        'Initialisation du LastDebVert au cas où aucun
                        'bande inter-carf trouvée
                        unLastDebVertD = unDebVertD
                    
                        Do
                            'La boucle sert pour afficher toutes les bandes
                            'inter-carrefour pour cela on regarde dans
                            'le cycle en cours et le suivant et on prend la
                            'bande inter-carf maximale
                            
                            'On prend le minimun des fins de vert projeté
                            'sur la droite Y = Y du carrefour courant
                            If (unFinVertDPred + unJ * .maDuréeDeCycle) < (unFinVertD + unI * .maDuréeDeCycle) Then
                                unMinFinVert = unFinVertDPred + unJ * .maDuréeDeCycle
                            Else
                                unMinFinVert = unFinVertD + unI * .maDuréeDeCycle
                            End If
                            
                            'On prend le maximun des débuts de vert projeté
                            'sur la droite Y = Y du carrefour courant
                            If (unDebVertDPred + unJ * .maDuréeDeCycle) > (unDebVertD + unI * .maDuréeDeCycle) Then
                                unMaxDebVert = unDebVertDPred + unJ * .maDuréeDeCycle
                            Else
                                unMaxDebVert = unDebVertD + unI * .maDuréeDeCycle
                            End If
                            
                            'Test de l'existence d'une bande inter-carrefour
                            'supérieure à 1 seconde
                            uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                            
                            If uneBandeInterCarfExist Then
                                'If unTD2 - 0.01 < unMinFinVert And unLastDebVertDéjàStocké = False Then
                                If unLastDebVertDéjàStocké = False Then
                                    'Stockage du dernier debvert ayant une bande
                                    'inter-carrefour, ce stochage est fait une fois
                                    'et une seule entre deux carrefours
                                    unLastDebVertDéjàStocké = True
                                    unLastDebVertD = unDebVertD + unI * .maDuréeDeCycle
                                End If
                                'Remise dans l'englobant total si
                                'MinFinVert en sort
                                If unMinFinVert > unMaxT + 0.01 Then
                                    unMinFinVert = unMinFinVert - .maDuréeDeCycle
                                    unMaxDebVert = unMaxDebVert - .maDuréeDeCycle
                                End If
                            End If
                            
                            unI = unI + 1
                            If unI = 2 Then
                                'On se place pour essayer les début et fin de vert
                                'du carrefour courant dans le cycle courant avec les
                                'début et fin de vert du carrefour précédent dans le
                                'cycle suivant
                                unI = 0
                                unJ = 1
                            End If
                                                                                
                            'Dessin de bande inter-carrefour
                            If uneBandeInterCarfExist Then
                                'Cas du dessin des bandes inter-carrefours
                                'voitures d'une onde TC
                                
                                'Conversion en coordonnées écran
                                unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 1ère partie bande descendante inter-carrefours
                                uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                
                                'Conversion en coordonnées écran
                                unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 2ème partie bande descendante inter-carrefours
                                uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                            End If
                            'Boucle fait trois fois pour trouver toutes
                            'les bande inter-carrefour
                        Loop Until unI = 1 And unJ = 1
                            
                    End If 'Fin du dessin de bande verte descendante inter-carrefours
                                                                                                 
                    'Stockage du début de vert précédent
                    If unNbCarf + 1 - i = unIndCarfD Then
                        'Calcul spécial pour le carrefour le + haut descendant
                        unDebVertDPred = unCarfRedD1.monCarrefour.monDecModif + unCarfRedD1.DonnerPosRefSens(False)
                        unDebVertDPred = ModuloZeroCycle(unDebVertDPred, .maDuréeDeCycle)
                        If unTD1 < unDebVertDPred - 0.001 Then
                            'Début de vert > T de départ onde descendante
                            '==> Recul d'un cycle
                            unDebVertDPred = unDebVertDPred - .maDuréeDeCycle
                        ElseIf unTD1 > unDebVertDPred + unCarfRedD1.DonnerDureeVertSens(False) + 0.001 Then
                            'Fin de vert < T de départ onde descendante
                            '==> Avancé d'un cycle
                            unDebVertDPred = unDebVertDPred + .maDuréeDeCycle
                        End If
                        'Initialisation de la durée de vert du premier
                        'carrefour descendant
                        uneDuréeVertDPred = unCarfRedD1.DonnerDureeVertSens(False)
                    Else
                        'Affectation du début de vert du carf réduit précédent
                        unDebVertDPred = unLastDebVertD
                        'Affectation de la durée de vert du carf réduit précédent
                        'à faire car modifs possibles dans RecalculerDebEtFinVert
                        uneDuréeVertDPred = unFinVertD - unDebVertD
                    End If
                    
                    'Stockage de l'indice de ce carrefour
                    unIndCarfDPred = unNbCarf + 1 - i
                    'Stockage du Y écran  du point précédent pour le coup suivant
                    unYDpred = unY
                End If
            End If
        Next i
    End With
    
    'Restauration du style de dessin précédent cette fonction
    uneZoneDessin.DrawStyle = unDrawStyleSave
End Sub

Public Sub RecalculerDebEtFinVert(unCarfRed As Object, unSensMontant As Boolean, uneDuréeDeCycle As Integer, uneVitesse As Single, unDebVert As Single, unFinVert As Single)
    'Recalcul des début et fin de vert du carrefour réduit passé en
    'paramètre car la réduction en onde TC est différente de celle
    'à vitesse variable ou constante
    'Les nouveaux début et fin de vert sont retournés et modifiés dans les
    'variables unDebVert et unFinVert. De plus leurs valeurs passées en
    'paramètre servent à initialiser pour les recherches de min et de max
    'décrite ci-dessous.
    
    'On garde le max des début de vert projeté sur le feu le plus haut en
    'montant (le plus bas en descendant) et du début de vert du carf réduit
    
    'On garde le min des fin de vert projeté sur le feu le plus haut en
    'montant (le plus bas en descendant) et du fin de vert du carf réduit
    
    Dim unYCarfRed As Integer, unYFeu As Integer
    Dim unCarf As Carrefour, unNbFeux As Integer
    Dim unDebVertFeu As Single, unFinVertFeu As Single
    Dim unMaxDebVert As Single, unMinFinVert As Single
    Dim unFeu As Feu
    
    'Le Y du feu le plus haut en montant ou le plus bas en descendant est
    'celui du carrefour réduit
    unYCarfRed = unCarfRed.DonnerYSens(unSensMontant)
    
    'Le début de vert et du fin de vert à trouver
    'sont initialisés avec ceux du carrefour réduit
    'Ce sont ceux passé en paramètre
    unMaxDebVert = unDebVert
    unMinFinVert = unFinVert
    
    'Recherche du max début de vert et du min fin de vert
    Set unCarf = unCarfRed.monCarrefour
    unNbFeux = unCarf.mesFeux.Count
    For i = 1 To unNbFeux
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monSensMontant = unSensMontant Then
            'Calcul du début de vert ramené entre 0 et cycle
            unDebVertFeu = ModuloZeroCycle(unCarf.monDecModif + unFeu.maPositionPointRef, uneDuréeDeCycle)
            'On le ramène dans le cycle du DebVert (= le min actuel)
            If unDebVert - unDebVertFeu >= 0 Then
                unePrec = 0.001
            Else
                unePrec = -0.001
            End If
            unNbCycle = Fix((unDebVert - unDebVertFeu + unePrec) / uneDuréeDeCycle)
            unDebVertFeu = unDebVertFeu + unNbCycle * uneDuréeDeCycle
            'Calcul du début de vert du feu projeté sur le Y du carrefour réduit
            unDebVertFeu = unDebVertFeu + (unYCarfRed - unFeu.monOrdonnée) / uneVitesse
            'En sens M , Diff des Y > 0 et Vitesse M > 0 ==> tout > 0 donc OK
            'En sens D , Diff des Y < 0 et Vitesse D < 0 ==> tout > 0 donc OK
            If unDebVertFeu > unMaxDebVert Then unMaxDebVert = unDebVertFeu
            
            'Calcul du fin de vert du feu projeté sur le Y du carrefour réduit
            unFinVertFeu = unDebVertFeu + unFeu.maDuréeDeVert
            If unFinVertFeu < unMinFinVert Then unMinFinVert = unFinVertFeu
        End If
    Next i
    
    'Retour des nouveaux débuts et fin de vert
    unDebVert = unMaxDebVert
    unFinVert = unMinFinVert
End Sub
