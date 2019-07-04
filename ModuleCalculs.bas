Attribute VB_Name = "ModuleCalculs"
'Variable globale et public � ce module permettant de pointer sur les
'carrefours r�duit du site courant avec un Y correspondant � la moyenne des
'ordonn�es des feux �quivalents du carrefour r�duit, ce tableau sera
'r�organiser en classant les carrefours r�duits par ordonn�e croissante
Public monTabCarfY() As CarfY

'Variable priv�e stockant le d�but de vert lors du calcul du feu �quivalent
'Elle n'est utilis�e que pour le dessin de l'onde verte
'(Proc�dure DessinerOndeVerte de ce module)
Private monDebutVert As Single
Private monTStartVert As Single

'Constantes donnant les r�sultats possibles des bandes passantes
Public Const AucuneSolution As Integer = 0
Public Const DoubleSensPossible As Integer = 1
Public Const DoubleSensImpossible As Integer = 2

'Constantes donnant le type de dessin � r�aliser dans la fonction
'TracerProgressionTC
Public Const DessinProgTC As Integer = 0  'Dessin de la progression du TC
Public Const DessinOndeTCM As Integer = 1 'Dessin de l'onde montante cadr�e TC
Public Const DessinOndeTCD As Integer = 2 'Dessin de l'onde descendante cadr�e TC

'Constante pour la variable indiquant la coh�rence entre les donn�es
'et les r�sultats du calcul d'onde
Public Const OK As Integer = 0
Public Const CalculImpossible As Integer = 1
Public Const IncoherenceDonneeCalcul As Integer = 2

'Constantes pour la s�lection graphique dans
'l'onglet Graphique Onde Verte et dans la fen�tre frmPleinEcran
Public Const NoSel As Integer = 0  'Aucune S�lection graphique trouv�e
Public Const PgGSel As Integer = 1 'S�lection graphique de la poign�e gauche
Public Const PgDSel As Integer = 2 'S�lection graphique de la poign�e droite
Public Const PlaSel As Integer = 3 'S�lection graphique d'une plage de vert
Public Const RefSel As Integer = 4 'S�lection graphique d'un point de r�f�rence

'Constante pour la pr�cision du pick �cran en Twips
Public Const PrecPick As Integer = 60

'Variables stockant les X �cran de d�but et de fin d'une modification
'graphique interactive, Initialis�e dans la fonction SelectionGraphique.
'Ainsi la diff�rence entre ce d�but et cette fin permet d'avoir la valeur
'de la modification.
Public monXEcranDebModif As Single
Public monXEcranFinModif As Single

'Collection stockant les valeurs avant une modif graphique du dernier
'objet graphique s�lectionn�, pour les restaurer si besoin est
Private maColValPred As New Collection

'Variable priv�e stockant l'objet pick� et son type dans
'Onglet Graphique ou frmPleinEcran
Private monTypeObjPick As Integer
Private monObjPick As Object

'Variable priv�e stockant le Temps total qui est l'englobant
'en temps donc suivant les X en coordonn�es r�elles
'Lors d'une annulation modif graphique il faut r�utiliser ce nombre
'pour retrouver la m�me �chelle en X car la modif a pu la changer
'car elle lance un redessin d'onde
Private monTmpTotalAvantModif As Long

Public Function ReduireCarrefourSite(uneForm As Form, uneColCarf As ColCarrefour, unTypeOnde As Integer) As Boolean
    'Procedure r�duisant tous les carrefours du site
    'en cr�ant les carrefours r�duits � sens unique ou double
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
        
    'Mise � vide des listes de Carrefour r�duit
    uneForm.mesCarfReduitsSensM.Vider
    uneForm.mesCarfReduitsSensD.Vider
    uneForm.mesCarfReduitsSens2.Vider
    unNbCarf = uneColCarf.Count
    unNbCarfUtil = 0
    
    'Recherche des Y des feux d'Y min et d'Y max pour une onde TC
    If unTypeOnde = OndeTC Then
        If uneForm.monTCM > 0 Then
            'Pour une onde TC montante on ne prend que les feux dont
            'le Y est compris entre le Y min des feux du carrefour de d�part
            'et le Y max des feux du carrefour d'arriv�e.
            'Dans ce cas, le d�part a un Y < � celui de l'arriv�e
            Set unTCM = uneForm.mesTC(uneForm.monTCM)
            unYFeuMinM = DonnerYMinCarfSens(unTCM.monCarfDep, True, unIndFeu)
            unYFeuMaxM = DonnerYMaxCarfSens(unTCM.monCarfArr, True, unIndFeu)
            'Calcul du tableau de marche prenant en compte les arr�ts
            'mais pas les feux
            unTCM.CalculerTableauMarcheOnde
        End If
        
        If uneForm.monTCD > 0 Then
            'Pour une onde TC descendante on ne prend que les feux dont
            'le Y est compris entre le Y min des feux du carrefour d'arriv�e
            'et le Y max des feux du carrefour de d�part
            'Dans ce cas, le d�part a un Y > � celui de l'arriv�e
            Set unTCD = uneForm.mesTC(uneForm.monTCD)
            unYFeuMinD = DonnerYMinCarfSens(unTCD.monCarfArr, False, unIndFeu)
            unYFeuMaxD = DonnerYMaxCarfSens(unTCD.monCarfDep, False, unIndFeu)
            'Calcul du tableau de marche prenant en compte les arr�ts
            'mais pas les feux
            unTCD.CalculerTableauMarcheOnde
        End If
    End If
    
    'Initialisation de la valeur de retour de ReduireCarrefourSite
    ReduireCarrefourSite = True
    
    'Parcours de tous les carrefours pass�s en param�tre
    For i = 1 To unNbCarf
        Set unCarf = uneColCarf(i)
        'On ne travaille que sur les carrefours choisis par l'utilisateur
        If unCarf.monIsUtil Then
            'Parcours de tous les feux du carrefour pour voir s'ils sont
            'tous dans le m�me sens ou dans deux sens diff�rents
            j = 2
            unIsSensDouble = False
            'Test des sens de deux feux cons�cutifs
            'Sortie si les sens sont diff�rents ==> Carrefour � double sens
            'Sinon Carrefour � sens unique celui du feu 1 par exemple
            'Si un seul feu dans le carrefour, on ne rentre pas dans la boucle
            '==> Carrefour � sens unique, celui du seul feu
            Do While j <= unCarf.mesFeux.Count And unIsSensDouble = False
                If unCarf.mesFeux(j - 1).monSensMontant <> unCarf.mesFeux(j).monSensMontant Then
                    unIsSensDouble = True
                End If
                j = j + 1
            Loop
            
            'Alimentation des listes de Carrefour r�duit
            If Not unIsSensDouble And unCarf.mesFeux(1).monSensMontant Then
                'Cas d'un carrefour ayant tous ses feux dans le sens montant
                'Calcul du feu �quivalent dans le sens unique montant
                unIsFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinM, unYFeuMaxM, unTCM)
                If unIsFeuEquivSensMExist Then
                    'Retaillage dynamique du tableau des carrefours r�duits avec ordonn�e
                    'On ne stocke que les carrefours utilis�s
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout � la liste des Carrefours r�duits � sens unique montant
                    Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
                    'Alimentation du tableau des carrefours r�duits avec ordonn�e
                    AjouterCarfY unNbCarfUtil, unCarfRedSensU
                End If
                '(unIsFeuEquivSensMExist Or monNbFeuxMpris = 1) est VRAI si on
                'a trouv� un feu �quivalent montant ou s'il n'y a pas de feu
                '�quivalent mais que tous les feux du carrefour sont en dehors
                'de unYFeuMin et unYFeuMax donc ce n'est pas une erreur
                '==> pas de changement
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensMExist Or uneForm.monNbFeuxMpris = 1)
            ElseIf Not unIsSensDouble And Not unCarf.mesFeux(1).monSensMontant Then
                'Cas d'un carrefour ayant tous ses feux dans le sens descendant
                'Calcul du feu �quivalent dans le sens unique descendant
                unIsFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinD, unYFeuMaxD, unTCD)
                If unIsFeuEquivSensDExist Then
                    'Retaillage dynamique du tableau des carrefours r�duits avec ordonn�e
                    'On ne stocke que les carrefours utilis�s
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout � la liste des Carrefours r�duits � sens unique descendant
                    Set unCarfRedSensU = uneForm.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVert, unePosRef, uneOrdonnee)
                    'Alimentation du tableau des carrefours r�duits avec ordonn�e
                    AjouterCarfY unNbCarfUtil, unCarfRedSensU
                End If
                '(unIsFeuEquivSensDExist Or monNbFeuxDpris = 1) est VRAI si on
                'a trouv� un feu �quivalent descendant ou s'il n'y a pas de feu
                '�quivalent mais que tous les feux du carrefour sont en dehors
                'de unYFeuMin et unYFeuMax donc ce n'est pas une erreur
                '==> pas de changement
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensDExist Or uneForm.monNbFeuxDpris = 1)
            Else
                'Cas d'un carrefour ayant des feux dans les deux sens
                'Calcul du feu �quivalent dans le sens montant
                unIsFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee, , , unYFeuMinM, unYFeuMaxM, unTCM)
                'Calcul du feu �quivalent dans le sens descendant
                unIsFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unePosRefD, uneOrdonneeD, , , unYFeuMinD, unYFeuMaxD, unTCD)
                If unIsFeuEquivSensMExist And unIsFeuEquivSensDExist Then
                    'Cas d'un carrefour r�duit � double sens
                    'Pour toutes les ondes sauf celle TC on doit passer l�
                    'sinon calcul impossible car double feu �quivalent imposible
                    
                    'Retaillage dynamique du tableau des carrefours r�duits avec ordonn�e
                    'On ne stocke que les carrefours utilis�s
                    unNbCarfUtil = unNbCarfUtil + 1
                    ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                    'Ajout � la liste des Carrefours r�duits � double sens
                    Set unCarfRedSens2 = uneForm.mesCarfReduitsSens2.Add(unCarf)
                    'Mise � jour des propri�t�s dans le sens montant du carrefour r�duit
                    unCarfRedSens2.SetPropsSensM uneDureeVert, unePosRef, uneOrdonnee
                    'Mise � jour des propri�t�s dans le sens descendant du carrefour r�duit
                    unCarfRedSens2.SetPropsSensD uneDureeVertD, unePosRefD, uneOrdonneeD
                    'Alimentation du tableau des carrefours r�duits avec ordonn�e
                    AjouterCarfY unNbCarfUtil, unCarfRedSens2
                Else
                    'Ce cas n'est pas une erreur uniquement si onde TC
                    'On aura un carrefour r�duit � sens unique
                    If unIsFeuEquivSensMExist Then
                        'Cas du sens unique montant
                        'Retaillage dynamique du tableau des carrefours r�duits avec ordonn�e
                        'On ne stocke que les carrefours utilis�s
                        unNbCarfUtil = unNbCarfUtil + 1
                        ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                        'Ajout � la liste des Carrefours r�duits � sens unique montant
                        Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
                        'Alimentation du tableau des carrefours r�duits avec ordonn�e
                        AjouterCarfY unNbCarfUtil, unCarfRedSensU
                    End If
                    If unIsFeuEquivSensDExist Then
                        'Cas du sens unique descendant
                        'Retaillage dynamique du tableau des carrefours r�duits avec ordonn�e
                        'On ne stocke que les carrefours utilis�s
                        unNbCarfUtil = unNbCarfUtil + 1
                        ReDim Preserve monTabCarfY(1 To unNbCarfUtil)
                        'Ajout � la liste des Carrefours r�duits � sens unique descendant
                        Set unCarfRedSensU = uneForm.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVertD, unePosRefD, uneOrdonneeD)
                        'Alimentation du tableau des carrefours r�duits avec ordonn�e
                        AjouterCarfY unNbCarfUtil, unCarfRedSensU
                    End If
                End If
                
                'Affectation de la valeur de retour de cette fonction
                'Retour Vrai si on a trouv� un feu �quivalent
                'dans les deux sens ou que le carrfour n'a aucun feux
                'entre Ymin et Ymax, faux sinon
                ReduireCarrefourSite = ReduireCarrefourSite And (unIsFeuEquivSensMExist Or uneForm.monNbFeuxMpris = 1) And (unIsFeuEquivSensDExist Or uneForm.monNbFeuxDpris = 1)
            End If
        End If
    Next i
    
    'V�rification de la coh�rence entre le type d'onde verte choisi
    'et les carrefours r�duits trouv�s
    If uneForm.monTypeOnde = OndeSensM And uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
        ReduireCarrefourSite = False
        MsgBox "Impossible de privil�gier le sens montant. Aucun Carrefour n'est dans ce sens", vbCritical
    ElseIf uneForm.monTypeOnde = OndeSensD And uneForm.mesCarfReduitsSensD.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
        ReduireCarrefourSite = False
        MsgBox "Impossible de privil�gier le sens descendant. Aucun Carrefour n'est dans ce sens", vbCritical
    End If

    If uneForm.monTypeOnde = OndeTC And uneForm.monTCM > 0 Then
        'Cas d'une onde cadr�e par un TC montant
        If uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
            ReduireCarrefourSite = False
            MsgBox "Impossible de cadrer le sens montant par le TC : " + unTCM.monNom + Chr(13) + "Aucun Carrefour n'a de feu dans ce sens entre le d�part et l'arriv�e de ce TC", vbCritical
        End If
    End If
    
    If uneForm.monTypeOnde = OndeTC And uneForm.monTCD > 0 Then
        'Cas d'une onde cadr�e par un TC descendant
        If uneForm.mesCarfReduitsSensD.Count = 0 And uneForm.mesCarfReduitsSens2.Count = 0 Then
            ReduireCarrefourSite = False
            MsgBox "Impossible de cadrer le sens descendant par le TC : " + unTCD.monNom + Chr(13) + "Aucun Carrefour n'a de feu dans ce sens entre le d�part et l'arriv�e de ce TC", vbCritical
        End If
    End If
    
    If unNbCarfUtil = 0 Then
        'Cas o� aucun carrefour r�duit cr��, on n'en cr�e un fictif
        'pour que ubound(monTabCarfY,1) ne plante pas (cf CalculerTempsparcours)
        'Ce carrefour r�duit ne sert � rien d'autre
        ReDim Preserve monTabCarfY(1 To 1)
        'Ajout � la liste des Carrefours r�duits � sens unique montant
        Set unCarfRedSensU = uneForm.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
        'Alimentation du tableau des carrefours r�duits avec ordonn�e
        AjouterCarfY 1, unCarfRedSensU
    End If
End Function

Public Function CalculerFeuEquivalent(unCarf As Carrefour, unSensMontant As Boolean, uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer, Optional unDecalModif As Boolean = False, Optional unSansMsgErreur As Boolean = False, Optional unYFeuMin As Integer = 0, Optional unYFeuMax As Integer = 0, Optional unTC As TC = Nothing) As Boolean
    'Calcul, lors de la r�duction d'un carrefour, le feu �quivalent
    'd'un carrefour qui sera utilis� dans le carrefour r�duit
    'Si unSensMontant est vrai, calcul du feu �quivalent dans le sens montant
    'Sinon calcul du feu �quivalent dans le sens descendant
    'Cette fonction modifie les valeurs de ses trois derniers
    'param�tres d'entr�e
    
    Dim uneVitesse As Single, unDecalCauseVitesse As Single
    Dim unYExtremun As Single, unCoefSens As Integer
    Dim uneColFeuSens1 As New Collection
    Dim unFeuSens1 As FeuSens1
    Dim unFeu As Feu, unFeuExt As Feu
    Dim unDebutVert As Single, unDecalage As Single
    Dim unTabBorne() As Single, unFeuPris As Boolean
    Dim uneColPeriodeVert As New Collection
    Dim unePeriodeVert As PeriodeVert
    
    'Initialisation du nombre de feux du sens choisi (FeuSens1) cr��
    unNbFeu = 1
    
    'Affectation de la vitesse en m/s du carrefour dans le sens �tudi�
    'Elle d�pend du type d'onde verte, du type de vitesse choisi (constante
    'ou variable) et des TC cadrant l'onde verte en sens montant et/ou
    'descendant
    uneVitesse = unCarf.DonnerVitSens(unSensMontant)
    If unSensMontant Then
        'Cas de recherche du feu �quivalent dans le sens montant
        unSens = "montant"
        unCoefSens = 1
        'Initialisation de l'extremun des Y en sens montant
        '==> c'est un maximun, donc valeur initiale petite
        unYExtremun = -10000
    Else
        'Cas de recherche du feu �quivalent dans le sens descendant
        unSens = "descendant"
        unCoefSens = -1
        'Initialisation de l'extremun des Y en sens descendant
        '==> c'est un minimun, donc valeur initiale grande
        unYExtremun = 10000
    End If
        
    'Calcul des plages de p�riode de vert des feux
    'du carrefour entre t=0 et t<dur�e du cycle dans
    'le sens montant si unSensmontant = True, descendant sinon
    For i = 1 To unCarf.mesFeux.Count
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monSensMontant = unSensMontant Then
            'Cas d'un feu ayant le sens cherch� pour le calcul
            
            If Not (unTC Is Nothing) And unDecalModif = False Then
                'Cas d'une onde verte cadr�e par un TC dans le sens donn�
                'par unSensMontant celui du feu �quivalent recherch�
                If unFeu.monOrdonn�e >= unYFeuMin And unFeu.monOrdonn�e <= unYFeuMax Then
                    'Cas o� le feu est entre le d�part et l'arriv�e de ce TC
                    '==> Feu pris en compte pour le calcul du feu �quivalent,
                    'sinon non
                    unFeuPris = True
                Else
                    unFeuPris = False
                End If
            Else
                'Cas d'une onde verte qui n'est pas cadr�e par un TC
                'dans le sens du feu �quivalent recherch�
                'Tous les feux dans ce sens sont pris
                unFeuPris = True
            End If
            
            If unFeuPris Then
                'Cas d'un feu intervenant dans le calcul du feu �quivalent
                
                'Calcul du maximun des Y en sens montant ou
                'du minimun des Y en sens descendant
                If (unSensMontant And unFeu.monOrdonn�e > unYExtremun) Or (Not unSensMontant And unFeu.monOrdonn�e < unYExtremun) Then
                    unYExtremun = unFeu.monOrdonn�e
                    Set unFeuExt = unFeu 'Stockage du feu d'Y extremun
                End If
                
                'Calcul de son d�but de vert
                If unDecalModif Then
                    'Cas o� l'on recalcule les bandes passantes
                    'apr�s une modification des d�calages
                    'unCarf est un carrefour r�duit global en prenant
                    'tous les feux equivalents des carrefours r�duits
                    '==> la vitesse n'est plus local au carrefour
                    'si vitesse non constante
                    '==> Utilisation du d�calage induit par les vitesses
                    'constantes ou variables �ventuelles
                    
                    'Le feu du carrefour global est le feu �quivalent montant
                    'ou descendant d'un carrefour r�duit, mais son champ
                    'carrefour pointe sur le carrefour dont il est issu
                    'apr�s r�duction
                    'Le d�calage du aux vitesses doit �tre multipli� par le sens
                    '(1 = montant et -1 = descendant)
                    If unSensMontant Then
                        unDecalCauseVitesse = unFeu.monCarrefour.monDecVitSensM
                    Else
                        unDecalCauseVitesse = -unFeu.monCarrefour.monDecVitSensD
                    End If
                    'R�cup�ration du d�calage modifi� du carrefour dont
                    'le feu unFeu est le feu �quivalent
                    unDecalage = unFeu.monCarrefour.monDecModif
                Else
                    'Cas o� l'on calcule les bandes passantes,
                    'les d�calages seront calcul�s plus tard
                    '==> Mis � z�ro pour avoir une formule de
                    'calcul du d�but vert valable pour tous les cas
                    unDecalage = 0
                    'Dans ce cas, m�me en vitesse variable, le d�calage du
                    'aux vitesses est du uniquement � la vitesse locale du
                    'carrefour
                    If unTC Is Nothing Then
                        'Cas d'une onde non cadr�e par unTC dans le sens du feu �quivalent
                        unDecalCauseVitesse = unFeu.monOrdonn�e / uneVitesse
                    Else
                        'Cas d'une onde cadr�e par un TC dans le sens du feu �quivalent
                        'Le coefsens permet d'inverser le signe des Y pour le cas descendant
                        'et de rendre le d�calage n�gatif dans le cas descendant pour respecter
                        'l'analogie avec les vitesses variables ci-dessus,
                        'uneVitesse > 0 si montant, < 0 si descendant
                        unDecalCauseVitesse = unCoefSens * unTC.CalculerDecalCauseProgTC(unTC.mesPhasesTMOnde, unFeu.monOrdonn�e, unCoefSens)
                    End If
                End If 'Fin du cas de modif des d�calages
                
                unDebutVert = unDecalage + unFeu.maPositionPointRef - unDecalCauseVitesse
                'On ram�ne modulo entre [0, dur�ee du cycle[
                unDebutVert = ModuloZeroCycle(unDebutVert, monSite.maDur�eDeCycle)
                
                'Cr�ation d'un feu avec ses plages de vert
                Set unFeuSens1 = New FeuSens1
                Set unFeuSens1.monFeu = unFeu
                If unDebutVert + unFeu.maDur�eDeVert > monSite.maDur�eDeCycle Then
                    'Cas de l'existence de deux plages de vert
                    unFeuSens1.monNbPlageVert = 2
                    unFeuSens1.maBorneVert1 = unDebutVert + unFeu.maDur�eDeVert - monSite.maDur�eDeCycle
                    unFeuSens1.maBorneVert2 = unDebutVert
                Else
                    'Cas de l'existence d'une seule plage de vert
                    unFeuSens1.monNbPlageVert = 1
                    unFeuSens1.maBorneVert1 = unDebutVert
                    unFeuSens1.maBorneVert2 = unDebutVert + unFeu.maDur�eDeVert
                End If
                'Augmentation du tableau dynamique pour stocker
                'les 2 bornes de vert en preservant les pr�c�dents
                ReDim Preserve unTabBorne(unNbFeu + 1)
                'Ajout au tableau des bornes de vert
                unTabBorne(unNbFeu) = unFeuSens1.maBorneVert1
                unTabBorne(unNbFeu + 1) = unFeuSens1.maBorneVert2
                unNbFeu = unNbFeu + 2
                'Cr�ation d'un feu avec ses plages de vert
                uneColFeuSens1.Add unFeuSens1
            End If 'Fin de if unFeuPris
        End If 'Fin de if Feu.sensmontant = sensmontant
    Next i
        
    'Calcul des caract�ristiques feu �quivalent
    '(ordonn�e, dur�e de vert, position du point de r�ference)
    If uneColFeuSens1.Count = 0 Then
        'Cas d'une onde cadr�e par un TC dans le sens du feu �quivalent
        'cherch� mais n'ayant aucun feu situ� entre le d�part et l'arriv�e
        'de ce TC
        CalculerFeuEquivalent = False
    ElseIf uneColFeuSens1.Count = 1 Then
        'Cas d'un carrefour n'ayant qu'un seul feu dans le sens �tudi�
        'Le feu �quivalent trouv� sera cet unique feu
        CalculerFeuEquivalent = True
        Set unFeu = uneColFeuSens1(1).monFeu
        uneDureeVert = unFeu.maDur�eDeVert
        unePosRef = unFeu.maPositionPointRef
        uneOrdonnee = unFeu.monOrdonn�e
    Else
        'Cas d'un carrefour ayant plusieurs feux dans le sens �tudi�
        'Tri dans l'ordre croissant des diff�rentes bornes de vert trouv�es
        TrierOrdreCroissant unTabBorne
        'Rajout de la derni�re borne de vert valant Dur�e du cycle
        ReDim Preserve unTabBorne(unNbFeu)
        unTabBorne(unNbFeu) = monSite.maDur�eDeCycle
        'Test de l'�tat des feux entre les bornes ordonn�es
        'avec la borne d'indice 0 valant 0
        For i = 0 To UBound(unTabBorne, 1) - 1
            If unTabBorne(i) < unTabBorne(i + 1) Then
                'Cas o� deux bornes successives sont diff�rentes
                'Car tri pr�c�dent par ordre croissant sans virer les doublons
                j = 1
                IsTousFeuxVert = True
                Do
                    'On regarde si tous les feux du sens �tudi�s sont verts
                    'entre deux bornes successives diff�rentes en regardant
                    '� la borne inf de la p�riode, unTabBorne(i)
                    If Not uneColFeuSens1(j).IsVert(unTabBorne(i)) Then
                        'Cas o� le feu rouge � cet instant
                        IsTousFeuxVert = False
                    End If
                    j = j + 1
                'Fin de boucle sur les feux du sens �tudi�
                Loop While j <= uneColFeuSens1.Count And IsTousFeuxVert = True
                
                'Ajout � la collection des p�riodes de vert trouv�e
                'si tous les feux sont verts dans cette p�riode
                If IsTousFeuxVert Then
                    Set unePeriodeVert = New PeriodeVert
                    unePeriodeVert.monDebutVert = unTabBorne(i)
                    unePeriodeVert.maDuree = unTabBorne(i + 1) - unTabBorne(i)
                    uneColPeriodeVert.Add unePeriodeVert
                End If
            End If 'Fin du if entre 2 bornes cons�cutives
        Next i 'Fin de boucle sur les bornes de vert
        
        If uneColPeriodeVert.Count = 0 Then
            'Cas o� aucun p�riode de vert n'a �t� trouv�
            '==> Aucun feu du carrefour vert en m�me temps
            'dans le sens �tudi�, donc pas de feu �quivalent trouv�
            CalculerFeuEquivalent = False
            
            If Not unDecalModif And unSansMsgErreur = False Then
                'Affichage d'un message d'erreur cibl�e dans le cas o�
                'l'on ne recalcule pas les bandes apr�s une modification
                'd'un d�calage
                unMsg = "Impossible de trouver une plage de vert commune pour "
                unMsg = unMsg + "les feux du carrefour " + unCarf.monNom
                unMsg = unMsg + " dans le sens " + unSens + "." + Chr(13) + Chr(13)
                unMsg = unMsg + "Modifier un ou plusieurs des param�tres de ce carrefour "
                unMsg = unMsg + "dans le sens " + unSens + " :" + Chr(13)
                unMsg = unMsg + "    sa dur�e de vert, sa position du point de r�f�rence, sa vitesse"
                MsgBox unMsg, vbCritical
            End If
        Else
            'Cas o� le feu �quivalent existe, car une p�riode de vert trouv�
            CalculerFeuEquivalent = True
            
            'Si la premi�re et derni�re p�riode a tous ses feux verts et que la
            'premi�re p�riode commence � 0 et que la derni�re finit � la dur�e du cycle
            'la dur�e de la derni�re devient la somme des dur�es de ces 2 p�riodes
            'et la date de d�but de vert reste celle de la derni�re p�riode et
            'on supprime la premi�re p�riode
            unNbPeriodeVert = uneColPeriodeVert.Count
            If unNbPeriodeVert <> 1 And uneColPeriodeVert(1).monDebutVert = 0 And (uneColPeriodeVert(unNbPeriodeVert).monDebutVert + uneColPeriodeVert(unNbPeriodeVert).maDuree = monSite.maDur�eDeCycle) Then
                'Calcul des caract�ristiques feu �quivalent
                '(ordonn�e, dur�e de vert, position du point de r�ference)
                uneColPeriodeVert(unNbPeriodeVert).maDuree = uneColPeriodeVert(unNbPeriodeVert).maDuree + uneColPeriodeVert(1).maDuree
                'Suppression de la premi�re p�riode de vert
                uneColPeriodeVert.Remove 1
                unNbPeriodeVert = uneColPeriodeVert.Count
            End If
            'Recherche de la p�riode de vert la plus grande
            uneDureeMax = 0
            unIndPeriodeMax = 0
            For i = 1 To unNbPeriodeVert
                If uneColPeriodeVert(i).maDuree > uneDureeMax Then
                    uneDureeMax = uneColPeriodeVert(i).maDuree
                    unIndPeriodeMax = i
                End If
            Next i
            'Calcul des caract�ristiques feu �quivalent
            '(ordonn�e, dur�e de vert, position du point de r�ference)
            uneDureeVert = uneDureeMax
            If unTC Is Nothing Then
                'Cas d'une onde non cadr�e par unTC dans le sens du feu �quivalent
                If unDecalModif Then
                    'Cas de la r�duction pour le calcul des bandes passantes lors
                    'de la modif manuelle d'un d�calage calcul� ou lors de la
                    'r�duction des carrefours � date impos�e ==> Projection sur le
                    'carrefour le plus bas en Y pour le cas montant et sur le
                    'carrefour le plus haut en Y pour le cas descendant
                    If unSensMontant Then
                        unDecalCauseVitesse = unFeuExt.monCarrefour.monDecVitSensM
                    Else
                        unDecalCauseVitesse = -unFeuExt.monCarrefour.monDecVitSensD
                    End If
                    unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unDecalCauseVitesse
                Else
                    'Cas de la r�duction d'un carrefour pour calculer les
                    'd�calages ==> Projection sur Y = 0
                    unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unYExtremun / uneVitesse
                End If
            Else
                'Cas d'une onde cadr�e par un TC dans le sens du feu �quivalent
                'Le coefsens permet d'inverser le signe des Y pour le cas descendant
                'et de rendre le d�calage n�gatif dans le cas descendant pour respecter
                'l'analogie avec les vitesses variables ci-dessus,
                'uneVitesse > 0 si montant, < 0 si descendant
                unePosRef = uneColPeriodeVert(unIndPeriodeMax).monDebutVert + unCoefSens * unTC.CalculerDecalCauseProgTC(unTC.mesPhasesTMOnde, unYExtremun, unCoefSens)
            End If
            uneOrdonnee = unYExtremun
            'Stockage du d�but de vert dans une variable priv�e
            '� ce module ModuleCalculs (cf D�clarations de ce module)
            monDebutVert = uneColPeriodeVert(unIndPeriodeMax).monDebutVert
        End If 'Fin du if p�riode vert trouv�
    End If 'Fin du calcul des caract�ristiques du feu �quivalent
    
    'Stockage du nombre de feux pris pour le calcul du feu �quivalent
    'dans le sens choisi
    If unSensMontant Then
        monSite.monNbFeuxMpris = unNbFeu
    Else
        monSite.monNbFeuxDpris = unNbFeu
    End If
    
    'Lib�ration de la m�moire
    Set uneColFeuSens1 = Nothing
    Set uneColPeriodeVert = Nothing
End Function

Public Function ModuloZeroCycle(unReel As Single, uneDureeCycle As Integer) As Single
    'Fonction ramenant un nombre r�el dans l'intervalle [0, dur�e du cycle[
    If unReel >= 0 Then
        ModuloZeroCycle = unReel - uneDureeCycle * Fix(unReel / uneDureeCycle)
    Else
        ModuloZeroCycle = unReel + uneDureeCycle * (1 - Fix(unReel / uneDureeCycle))
    End If
    
    If ModuloZeroCycle = uneDureeCycle Then ModuloZeroCycle = 0
End Function

Public Sub TrierOrdreCroissant(unTabBorne() As Single)
    'R�organisation par ordre croissant d'un tableau de r�els
    'index� entre 1 et n avec l'indice z�ro nul mais qui ne sert pas.
    'Algo choisi : Le tri insertion (r�cup�rer sur Internet)
    'Il consiste � comparer successivement un �l�ment
    '� tous les pr�c�dents et � d�caler les �l�ments interm�diaires

    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unTmp As Single
    
    'Mise � z�ro du contenu d'indice 0
    unTabBorne(0) = 0
    
    'Tri
    unNbTotal = UBound(unTabBorne, 1)
    For j = 2 To unNbTotal
            unTmp = unTabBorne(j)
            i = j - 1
            Do While i > 0 And unTabBorne(i) > unTmp 'Indice z�ro nul, �a �vite le plantage du And pour i = 0
                unTabBorne(i + 1) = unTabBorne(i)
                i = i - 1
            Loop
            unTabBorne(i + 1) = unTmp
    Next j
End Sub

Public Function CalculerOndeVerte(uneForm As Form, Optional uneModifDec As Boolean = False) As Boolean
    'Proc�dure essayant de calculer l'onde verte
    'Si uneModifDec = true c'est que CalculerOndeVerte a �t� appel�e
    'apr�s une modif manuelle dans l'onglet Tableau D�calage
    'ou apr�s une modif graphique � la souris d'un d�calage d'un carrefour
    '� d�calage impos� ==> on ne change que le champ monDecModif de ce carrefour
    
    'Variables locales caract�ristiques des bandes passantes cherch�es
    'Sens 1 = sens montant pour onde verte double sens ou le sens
    '         privil�gi� pour une onde verte � sens privil�gi�
    'Sens 2 = sens descendant pour onde verte double sens ou l'autre sens
    '         que celui privil�gi� pour une onde verte � sens privil�gi�
    Dim unB1 As Single 'valeur de la bande passante de vert du sens 1
    Dim unB2 As Single 'valeur de la bande passante de vert du sens 2
    Dim unH As Single  'Temps �coul� entre les �v�nements
                       '"Passage au vert montant" et "Fin de vert descendant"
    Dim unCarfRed As Object
    Dim unCarf As Carrefour
    Dim unDecImp As Single
    
    'Affectation � vrai de la r�alisation du calcul de l'onde
    'On mettra faux chaque fois que le calcul de l'onde est impossible
    CalculerOndeVerte = True
        
    With uneForm
        'Calcul � double sens des ondes vertes TC
        'si aucun TC montant et descendant
        If .monTypeOnde = 3 And .ComboTCM.Text = "Aucun" And .ComboTCD.Text = "Aucun" Then
                MsgBox "Dans l'onglet Cadrage Onde Verte, aucun TC montant et/ou descendant n'ont �t� choisis." + Chr(13) + Chr(13) + "Calcul d'onde verte prenant en compte les TC impossible", vbCritical
                CalculerOndeVerte = False
                monSite.maCoherenceDataCalc = CalculImpossible
                Exit Function
        End If
                
        'Le calcul n'est pas effectu� s'il n'y a pas eu une modif dans les
        'donn�es carrefours, TC qui cadre l'onde verte, calculs d'onde et de
        'modifications graphiques
        If Not .maModifDataCarf And Not .maModifDataOndeTC And Not .maModifDataOnde And Not .maModifDataDes Then
            If .maCoherenceDataCalc = CalculImpossible Then
                MsgBox "Le calcul d'onde verte est impossible avec les donn�es de ce site", vbCritical
                'Mise � z�ro des bandes et des d�calages
                RendreNulleBandesEtDecalages uneForm
            End If
            If .maCoherenceDataCalc <> IncoherenceDonneeCalcul Then Exit Function
            'Si les donn�es sont incoh�rentes avec les
            'r�sultats (modif de donn�es sans recalcul) ont refait le calcul
        End If
                
        'Mise en gris�e du menu Annuler derni�re modif graphique si on
        'a fait une modif par saisie et pas par interaction graphique
        If .maModifDataDes = False Then
            frmMain.mnuGraphicOndeAnnul.Enabled = False
        End If
        
        'Affectation � vrai de la r�alisation d'une onde � sens privil�gi�
        'mais possible � cadrer � double sens
        'On mettra faux si le cadrage � double sens de l'onde � sens privil�gi�
        'est impossible
        .monOndeDoubleTrouve = True
        
        'Stockage d'une modification de valeurs dans les d�calages
        'Ceci permettra aussi de demander une sauvegarde � la fermeture
        .maModifDataDec = True
        
        'Remise � FALSE des autres indicateurs pour pouvoir relancer un
        'calcul d'onde verte s'il repasse � TRUE
        .maModifDataCarf = False
        .maModifDataOndeTC = False
        .maModifDataOnde = False
        .maModifDataDes = False
    End With
    
    'Remise � z�ro d'une translation pr�c�dente globale � tous les
    'carrefours
    uneForm.maTransDec = 0
    uneForm.TextTransDec.Text = Format(uneForm.maTransDec)
    
    'R�duction des carrefours
    If ReduireCarrefourSite(uneForm, uneForm.mesCarrefours, uneForm.monTypeOnde) Then
        'Cas o� tous les carrefours ont pu �tre r�duits
        '==> tous les feux �quivalents ont pu �tre calcul�s
        
        'Initialisation des d�calages � -99 des carrefours � d�calages non impos�s
        'Cas des carrefours non pris en compte dans le calcul (monIsUtil = False
        'ou Y carrefour pas entre Ymin et Ymax pour les ondes TC)
        'Seule possibilit� d'avoir cette valeur qui reste � -99
        'car sinon les d�calages sont entre 0 et la dur�e du cycle
        For i = 1 To uneForm.mesCarrefours.Count
            'Sauvegarde des d�calages calcul�s et modifi�s
            'avant un calcul avec d�calage impos�
            uneForm.mesCarrefours(i).monDecCalculSave = uneForm.mesCarrefours(i).monDecCalcul
            uneForm.mesCarrefours(i).monDecModifSave = uneForm.mesCarrefours(i).monDecModif
        
        
            ' LCHAMMBON Correction
            If uneForm.mesCarrefours(i).monDecImp = 1 And Not uneForm.mesCarrefours(i).monIsUtil Then
                uneForm.mesCarrefours(i).monDecImp = 0
            End If
            
            If uneForm.mesCarrefours(i).monDecImp = 0 Then
                'Cas d'un carrefour � d�calage non impos�
                'Initialisation � -99
                uneForm.mesCarrefours(i).monDecCalcul = -99
                uneForm.mesCarrefours(i).monDecModif = -99
            End If
        Next i
                                
        'Lancement du cas o� des d�calages sont impos�s
        'unCarfRed est diff�rent de nothing si le calcul
        'a eu lieu avec des dates impos�es
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
                'Cas o� aucune solution de bande passante n'a �t� trouv�
                MsgBox "Aucune solution d'onde verte � double sens n'a pu �tre trouv�e", vbCritical
                If unCarfRed Is Nothing Then
                    'Cas d'un calcul sans aucun d�calage impos�
                    'Le calcul n'a rien trouv� donc on signale l'impossibilit�
                    'on ne pourra pas voir les dessins des plages de vert car
                    'les d�calages sont inconnus
                    CalculerOndeVerte = False
                Else
                    'Cas d'un calcul avec des d�calages impos�s
                    'Le calcul n'a rien trouv� donc on signale l'impossibilit�
                    'mais on trouve des bandes nulles et on garde les d�calages
                    'saisis donnant cette impossibilit�
                    'mais on peut voir le dessin des plages de vert car les
                    'd�calages sont connus
                    CalculerOndeVerte = True
                End If
            Case DoubleSensPossible
                'Cas o� les bandes passantes existent
                'avec une solution � double sens
                
                If Not (unCarfRed Is Nothing) Then
                    'Correction de la solution � date impos�e
                    '(cf commentaires dans la fonction CorrectionDateImpos�e)
                    CorrectionDateImpos�e uneForm, unCarfRed, unB1, unB2, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD
                End If
                
                'Stockage et affichage des bandes passantes calcul�es
                StockerEtAfficherBandes uneForm, unB1, unB2
                'Calcul des d�calages pour chaque carrefour
                CalculerDecalageDoubleSens uneForm, unB2, unH
            Case DoubleSensImpossible
                'Cas o� les bandes passantes existent
                'mais pas de solution � double sens
                unTousCarfSensUniqueM = uneForm.mesCarfReduitsSens2.Count = 0 And uneForm.mesCarfReduitsSensM.Count > 0 And uneForm.mesCarfReduitsSensD.Count = 0
                unTousCarfSensUniqueD = uneForm.mesCarfReduitsSens2.Count = 0 And uneForm.mesCarfReduitsSensM.Count = 0 And uneForm.mesCarfReduitsSensD.Count > 0
                If uneForm.monTypeOnde <> OndeTC And unTousCarfSensUniqueM = False And unTousCarfSensUniqueD = False Then
                    'Message affich� si on n'est pas en onde cadr�e par TC et s'il
                    'n'y a pas que des carrefours � sens unique dans le m�me sens
                    MsgBox "Une solution pour le sens privil�gi� a �t� calcul�e, mais sans arriver � trouver une solution pour l'autre sens", vbInformation
                End If
                
                If Not (unCarfRed Is Nothing) Then
                    'Correction de la solution � date impos�e
                    '(cf commentaires dans la fonction CorrectionDateImpos�e)
                    CorrectionDateImpos�e uneForm, unCarfRed, unB1, unB2, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD
                End If
                
                'Stockage et affichage des bandes passantes calcul�es
                StockerEtAfficherBandes uneForm, unB1, unB2
                'Calcul des d�calages pour chaque carrefour
                CalculerDecalageSansDoubleSens uneForm
                'Cadrage � double sens de l'onde � sens privil�gi� impossible
                '==> On ne dessinera que l'onde dans le sens privil�gi�
                uneForm.monOndeDoubleTrouve = False
            Case Else
                CalculerOndeVerte = False
                MsgBox "Erreur de programmation dans OndeV dans CalculerOndeVerte", vbCritical
        End Select
    
        If Not (unCarfRed Is Nothing) And CalculerOndeVerte Then
            'Cas o� le calcul a eu lieu avec des dates impos�es
            'et il s'est bien pass�
            
            'R�cup�ration du d�calage calcul� du carrefour r�duisant
            'tous les carrefours � date impos�e
            If unIndUniqCarfImp = 0 Then
                'Cas avec plusieurs carrefours � d�calage impos�
                unDecImp = unCarfRed.monCarrefour.monDecCalcul
            Else
                'Cas particulier d'un seul carrefour avec d�calage impos�
                'Nouveau d�calage moins l'ancien
                unDecImp = unCarfRed.monCarrefour.monDecCalcul - uneForm.mesCarrefours(unIndUniqCarfImp).monDecModif
            End If
            
            'Lib�ration m�moire du carrefour de unCarfRed et de lui-m�me
            Set unCarfRed.monCarrefour = Nothing
            Set unCarfRed = Nothing
            
            If unResCalculBandes Then
                'Cas o� des solutions avec des d�calages impos�s existent
            
                'Soustraction de ce d�calage � tous les d�calages des carrefours
                'sans date impos�e et mise de ces d�calages entre 0 et dur�e du cycle
                For i = 1 To uneForm.mesCarrefours.Count
                    Set unCarf = uneForm.mesCarrefours(i)
                    If unCarf.monDecImp = 0 And unCarf.monDecCalcul <> -99 Then
                        'Cas d'un carrefour � d�calage non impos�
                        'et utilis� dans le calcul
                        unCarf.monDecCalcul = ModuloZeroCycle(unCarf.monDecCalcul - unDecImp, uneForm.maDur�eDeCycle)
                        unCarf.monDecModif = unCarf.monDecCalcul
                        If uneModifDec Then unCarf.monDecCalcul = unCarf.monDecCalculSave
                    ElseIf unCarf.monDecImp = 1 And unCarf.monDecCalcul <> -99 Then
                        'Cas d'un carrefour � d�calage impos�
                        'et utilis� dans le calcul
                        If uneModifDec = False Then unCarf.monDecCalcul = unCarf.monDecModif
                    End If
                    'Si uneModifDec = true c'est que CalculerOndeVerte a �t� appel�e
                    'apr�s une modif manuelle dans l'onglet Tableau D�calage (cf proc�dure TabDecal_EditMode de frmDocument et RecalculerAvecDateImp dans ModuleCalculs)
                    'ou apr�s une modif graphique (cf proc�dure MettreAJourSelection dans ModuleCalculs) � la souris d'un d�calage d'un carrefour
                    '� d�calage impos� ==> on remet la valeur du champ monDecCalcul de ce carrefour avant le calcul � date impos�e
                Next i
                    
                'Restauration des bandes passantes avant leur modif si on a fait
                'une modif manuelle ou graphique d'un d�calage
                If uneModifDec Then
                    uneForm.maBandeM = uneBMsave
                    uneForm.maBandeD = uneBDsave
                End If
            Else
                'Cas o� aucune solution trouv�e avec des d�calages impos�s
                
                'Restauration des d�calages pr�c�dent le calcul
                For i = 1 To uneForm.mesCarrefours.Count
                    uneForm.mesCarrefours(i).monDecCalcul = uneForm.mesCarrefours(i).monDecCalculSave
                    uneForm.mesCarrefours(i).monDecModif = uneForm.mesCarrefours(i).monDecModifSave
                Next i
                'Restauration des bandes passantes pr�c�dent le calcul
                uneForm.maBandeM = uneBMsave
                uneForm.maBandeD = uneBDsave
                uneForm.maBandeModifM = unB1
                uneForm.maBandeModifD = unB2
            End If
            
            'Affichage dans l'onglet Tableau de r�sultat
            RemplirOngletTabDecalage uneForm
            
            'Remise � jour de la r�duction des carrefours du site
            'et des temps de parcours pour le dessin des ondes
            ReduireCarrefourSite uneForm, uneForm.mesCarrefours, uneForm.monTypeOnde
            CalculerTempsParcours uneForm
        End If
    Else
        'R�duction de tous les carrefours impossible
        CalculerOndeVerte = False
    End If

    'Indication du niveau de coh�rence entre les donn�es
    'et le r�sultat du calcul d'onde verte
    If CalculerOndeVerte = False Then
        monSite.maCoherenceDataCalc = CalculImpossible
    Else
        monSite.maCoherenceDataCalc = OK
    End If
End Function

Public Function CalculerBandesPassantesMaxi(uneForm As Form, unB1 As Single, unB2 As Single, unH As Single, unCasDateImp As Boolean) As Integer
    'Fonction cherchant les bandes passantes maximales
    'dans les deux sens ou en privil�giant un sens
    
    Dim unCarfRedSens2 As CarfReduitSensDouble
    'Variables locales donnant les p�riodes de verte minimales
    'dans les sens montant et descendant
    Dim unMinVertSensM As Single
    Dim unMinVertSensD As Single
    
    Dim unS As Single 'Correspond au S des sp�cifs
    Dim unMin As Single, unMinLoc As Single
    Dim unPMsurPD As Single
    Dim A1 As Single, A2 As Single, K As Single
    Dim B1 As Single, B2 As Single
    Dim unB1M As Single, unB1D As Single
    Dim unB2M As Single, unB2D As Single

    'Initialisation
    unMinVertSensM = uneForm.maDur�eDeCycle
    unMinVertSensD = uneForm.maDur�eDeCycle
    unNbCarfRedSens2 = uneForm.mesCarfReduitsSens2.Count
    
    'Calcul des temps de parcours dans chaque sens � chaque carrefour
    'en prenant comme origine le premier carrefour dans chaque sens
    'consid�r� et en ayant trier par ordonn�e croissante les carrefours
    'r�duits.
    'De plus dans CalculerTempsParcours on calcul les �carts
    'des carrefours r�duits � double sens
    'Calcul temps de parcours si on n'est pas dans le cas date impos�e
    If unCasDateImp = False Then CalculerTempsParcours uneForm
        
    '1�re condition sur l'onde verte
    'La bande passante <= � la plus petite p�riode de
    'verte rencontr�e dans le sens consid�r�
    '==> Calcul des p�riodes vertes minimun unMinVertSensM et unMinVertSensD
    
    'Recherche sur tous les carrefours r�duits � sens unique montant
    'pour la dur�e de vert minimale dans ce sens
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        If uneForm.mesCarfReduitsSensM(i).maDureeVert < unMinVertSensM Then
            unMinVertSensM = uneForm.mesCarfReduitsSensM(i).maDureeVert
        End If
    Next i
    
    'Recherche sur tous les carrefours r�duits � sens unique descendant
    'pour la dur�e de vert minimale dans ce sens
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        If uneForm.mesCarfReduitsSensD(i).maDureeVert < unMinVertSensD Then
            unMinVertSensD = uneForm.mesCarfReduitsSensD(i).maDureeVert
        End If
    Next i
    
    'Recherche sur tous les carrefours r�duits � double sens
    'pour les dur�es de vert minimales dans chaque sens
    
    'Initialisation du unS qui est un maximun, pour le trouver dans
    'le code plus bas explicant la 2�me condition
    unS = -3 * uneForm.maDur�eDeCycle 'Car toutes les valeurs sont dans [0,dur�e du cycle[
    For i = 1 To unNbCarfRedSens2
        'R�cup�ration du carrefour r�duit i
        Set unCarfRedSens2 = uneForm.mesCarfReduitsSens2(i)
        '1�re condition sur l'onde verte
        If unCarfRedSens2.maDureeVertM < unMinVertSensM Then
            unMinVertSensM = unCarfRedSens2.maDureeVertM
        End If
        If unCarfRedSens2.maDureeVertD < unMinVertSensD Then
            unMinVertSensD = unCarfRedSens2.maDureeVertD
        End If
                
        '2�me condition sur l'onde verte s'il y a
        'des carrefours r�duits � double sens
                
        'Calcul sur tous les carrefours r�duits double sens de la fonction :
        'Min(Z) = Minimun(DureeVertSensM_Carf_i + DureeVertSensD_Carf_i - Ecart_Carf_i(Z) + Z)
        'pour tout i variant de 1 � nombre de carrefours r�duits double sens
        'et Z variant de monEcart du premier carrefour r�duit double sens �
        'monEcart du dernier carrefour r�duit double sens
        'En m�me temps on cherche unS = Max de ces min et on stocke dans unH
        'le Z correspondant au Max
        unMin = 3 * uneForm.maDur�eDeCycle 'Car toutes les valeurs sont dans [0,dur�e du cycle[
        For j = 1 To unNbCarfRedSens2
            unMinLoc = uneForm.mesCarfReduitsSens2(j).maDureeVertM + uneForm.mesCarfReduitsSens2(j).maDureeVertD - Ecart(uneForm.mesCarfReduitsSens2(j).monEcart, unCarfRedSens2.monEcart, uneForm.maDur�eDeCycle) + unCarfRedSens2.monEcart
            If unMinLoc < unMin Then
                'Stockage du minimun
                unMin = unMinLoc
            End If
        Next j
        
        'Stockage du maximun des minimuns sur tous les
        'carrefours r�duits et l'�cart r�alisant ce max
        If unMin > unS Then
            unS = unMin
            unH = unCarfRedSens2.monEcart
        End If
    Next i
    
    'D�termination des bandes passantes maximales
    If uneForm.mesCarfReduitsSens2.Count = 0 Then
        'Cas o� tous les carrefours sont � sens unique
        CalculerBandesPassantesMaxi = DoubleSensImpossible
        If uneForm.mesCarfReduitsSensM.Count = 0 Then
            'Cas o� tous les carrefours � sens unique descendant
            unB1 = 0
            unB2 = unMinVertSensD
        ElseIf uneForm.mesCarfReduitsSensD.Count = 0 Then
            'Cas o� tous les carrefours � sens unique montant
            unB1 = unMinVertSensM
            unB2 = 0
        Else
            unB1 = unMinVertSensM
            unB2 = unMinVertSensD
            unH = 0 'Tous les H dans [0, Dur�e du cycle[ sont possibles
        End If
    Else
        'Cas o� il y a des carrefours r�duits � double sens
        
        If uneForm.monTypeOnde = OndeTC And (uneForm.monTCM > 0 Or uneForm.monTCD > 0) Then
            'Cas d'une onde cadr�e par un TC
                
            'Coordonn�es des points A et B segment sur lequel se trouve la
            ' solution (cf Dossier de programmation / Solution temps impos�)
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
            
            'Calcul des bandes passantes par temps impos�s
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
                'Cas o� le temps est impos� par un TC dans les deux sens
                '(cf Dossier programmation)
                unTimp = (unS - uneForm.maBandeTCD + uneForm.maBandeTCM) / 2
                K = (unTimp - A1) / (B1 - A1)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            ElseIf uneForm.monTCM > 0 Then
                'Cas o� le temps est impos� par un TC dans le sens montant
                K = (uneForm.maBandeTCM - A1) / (B1 - A1)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            ElseIf uneForm.monTCD > 0 Then
                'Cas o� le temps est impos� par un TC dans le sens desendant
                'Solution sens descendant
                K = (uneForm.maBandeTCD - A2) / (B2 - A2)
                Call CalculerB1B2(K, A1, A2, B1, B2, unB1, unB2)
                CalculerBandesPassantesMaxi = DoubleSensPossible
            Else
                MsgBox "ERREUR de programmation dans OndeV dans CalculerBandesPassantesMaxi", vbCritical
            End If
            
            'Sortie pour �viter de faire le code qui suit
            Exit Function
        End If
        
        'Cas d'une onde non cadr�e par un TC
        If uneForm.monTypeOnde = OndeDouble Then
            'Cas de l'onde double
            CalculerBandesPassantesMaxi = DoubleSensPossible
            'La valeur ci-dessus sera modifi�e uniquement
            'si aucune solution trouv�e
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
            'Cas de l'onde � sens privil�gi� montant
            CalculerBandesPassantesMaxi = CalculerBandeSensPrivi(unS, unB1, unB2, unMinVertSensM, unMinVertSensD)
        ElseIf uneForm.monTypeOnde = OndeSensD Then
            'Cas de l'onde � sens privil�gi� descendant
            CalculerBandesPassantesMaxi = CalculerBandeSensPrivi(unS, unB2, unB1, unMinVertSensD, unMinVertSensM)
        Else
            'Cas d'une erreur de programmation
            MsgBox "Erreur dans le calcul des bandes passantes maximales", vbCritical
        End If
    End If
End Function

Public Function Ecart(unEi As Single, unH As Single, uneDureeCycle As Integer) As Single
    'Fonction correspondant � la fonction ECART_i(h) = Ei+Ai*Cycle des sp�cifs
    If unH > unEi Then
        Ecart = unEi + uneDureeCycle
    Else
        Ecart = unEi
    End If
End Function

Public Sub CalculerTempsParcours(uneForm As Form)
    'Calcul des temps de parcours cumul�s de chaque carrefour dans le
    'sens montant (respectivement descendant) � partir du premier carrefour
    'dans ce sens montant (respectivement descendant), les carrefours ayant
    '�t� au pr�alable class�s par ordre croissant gr�ce � la moyenne des
    'ordonn�es des feux �quivalents du carrefour r�duit.
    
    'De plus, on en profite pour faire le Calcul des �carts de chaque
    'carrefour r�duit � double sens.
    'L'�cart est le temps s'�coulant entre les �v�nements "passage au vert
    'dans le sens montant" et "fin du vert dns le sens descendant" apr�s
    'projection sur une r�f�rence commune � l'ensemble des carrefours
    '(cf Dossier de programmation et sp�cifs)
    
    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unCarfTmp As CarfY
    Dim unCarfRedSensU As CarfReduitSensUnique
    Dim unCarfRedSens2 As CarfReduitSensDouble
    Dim unDecVitSensM As Single, unDecVitSensD As Single
    Dim unIndCarfM As Integer, unIndCarfD As Integer
    Dim uneOndeTCM As Boolean, uneOndeTCD As Boolean
    Dim unIndPhase As Integer, unTCM As TC, unTCD As TC
    
    'D�termination du type d'onde verte � calculer
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde verte cadr�e par un TC dans le sens montant
        uneOndeTCM = True
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde verte non cadr�e par un TC dans le sens montant
        uneOndeTCM = False
        Set unTCM = Nothing
    End If
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde verte cadr�e par un TC dans le sens descendant
        uneOndeTCD = True
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde verte non cadr�e par un TC dans le sens descendant
        uneOndeTCD = False
        Set unTCD = Nothing
    End If
    
    'R�organisation du tableau de carrefours r�duits avec leur ordonn�e
    'par classement suivant les ordonn�es croissantes
    'Algo choisi : Le tri insertion (r�cup�rer sur Internet)
    'Il consiste � comparer successivement un �l�ment
    '� tous les pr�c�dents et � d�caler les �l�ments interm�diaires
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

    'Calcul des temps de parcours cumul�s de chaque carrefour dans le
    'sens montant (respectivement descendant) � partir du premier carrefour
    'dans ce sens montant (respectivement descendant), les carrefours ayant
    '�t� r�organis� ci-dessus
    unIndCarfM = 0
    unIndCarfD = 0
    For i = 1 To UBound(monTabCarfY, 1)
        'Initialisation � 0 des d�calages dues aux vitesses du carrefour
        monTabCarfY(i).monCarfReduit.monCarrefour.monDecVitSensM = 0
        monTabCarfY(i).monCarfReduit.monCarrefour.monDecVitSensD = 0
        
        If TypeOf monTabCarfY(i).monCarfReduit Is CarfReduitSensUnique Then
            'Cas o� le carrefour r�duit est � sens unique
            Set unCarfRedSensU = monTabCarfY(i).monCarfReduit
            If unCarfRedSensU.monSensMontant Then
                'Cas d'un carrefour � sens unique montant
                If uneOndeTCM Then
                    unDecVitSensM = unTCM.CalculerDecalCauseProgTC(unTCM.mesPhasesTMOnde, unCarfRedSensU.monOrdonnee, 1)
                Else
                    CalculerDecVitesse unCarfRedSensU, True, i, unIndCarfM, unDecVitSensM
                End If
                'Stockage du d�calage cumul� au carrefour
                unCarfRedSensU.monCarrefour.monDecVitSensM = unDecVitSensM
            Else
                'Cas d'un carrefour � sens unique descendant
                If uneOndeTCD Then
                    '-1 pour inversion du signe du Y pour le cas descendant
                    unDecVitSensD = unTCD.CalculerDecalCauseProgTC(unTCD.mesPhasesTMOnde, unCarfRedSensU.monOrdonnee, -1)
                Else
                    CalculerDecVitesse unCarfRedSensU, False, i, unIndCarfD, unDecVitSensD
                End If
                'Stockage du d�calage cumul� au carrefour
                unCarfRedSensU.monCarrefour.monDecVitSensD = unDecVitSensD
           End If
        Else
            'Cas o� le carrefour r�duit est � double sens
            Set unCarfRedSens2 = monTabCarfY(i).monCarfReduit
            
            'Calcul du d�calage cumul� au carrefour dans le sens montant
            If uneOndeTCM Then
                unDecVitSensM = unTCM.CalculerDecalCauseProgTC(unTCM.mesPhasesTMOnde, unCarfRedSens2.monOrdonneeM, 1)
            Else
                CalculerDecVitesse unCarfRedSens2, True, i, unIndCarfM, unDecVitSensM
            End If
            unCarfRedSens2.monCarrefour.monDecVitSensM = unDecVitSensM
            
            'Calcul du d�calage cumul� au carrefour dans le sens descendant
            If uneOndeTCD Then
                '-1 pour Inversion du signe des Y pour le cas descendant
                unDecVitSensD = unTCD.CalculerDecalCauseProgTC(unTCD.mesPhasesTMOnde, unCarfRedSens2.monOrdonneeD, -1)
            Else
                CalculerDecVitesse unCarfRedSens2, False, i, unIndCarfD, unDecVitSensD
            End If
            unCarfRedSens2.monCarrefour.monDecVitSensD = unDecVitSensD
            
            'Calcul des �carts de chaque carrefour r�duit � double sens
            'l'�cart est le temps s'�coulant entre les �v�nements "passage au vert
            'dans le sens montant" et "fin du vert dns le sens descendant" apr�s
            'projection sur une r�f�rence commune � l'ensemble des carrefours
            '(cf Dossier de programmation et sp�cifs)
            'On utilise des d�calages dus aux vitesses variables ou
            'constantes de chaque carrefour
            unCarfRedSens2.monEcart = unCarfRedSens2.maPosRefD + unCarfRedSens2.monCarrefour.monDecVitSensD + unCarfRedSens2.maDureeVertD
            unCarfRedSens2.monEcart = unCarfRedSens2.monEcart - (unCarfRedSens2.maPosRefM - unCarfRedSens2.monCarrefour.monDecVitSensM)
            'On ram�ne l'�cart modulo entre [0, dur�ee du cycle[
            unCarfRedSens2.monEcart = ModuloZeroCycle(unCarfRedSens2.monEcart, uneForm.maDur�eDeCycle)
        End If
    Next i
End Sub

Public Sub AjouterCarfY(unIndex As Integer, unCarfReduit As Object)
    'Alimentation du tableau des carrefours r�duits avec ordonn�e
    'Cette ordonn�e est calcul�e dans cette proc�dure et elle correspond
    '� la moyenne des ordonn�es des feux �quivalents du carrefour r�duit
    Dim unCarfY As New CarfY
    
    Set unCarfY.monCarfReduit = unCarfReduit
    'Calcul du Y
    If TypeOf unCarfReduit Is CarfReduitSensUnique Then
            'Cas o� le carrefour r�duit est � sens unique
            unCarfY.monY = unCarfReduit.monOrdonnee
    Else
            'Cas o� le carrefour r�duit est � double sens
            unCarfY.monY = (unCarfReduit.monOrdonneeM + unCarfReduit.monOrdonneeD) / 2
    End If
    'Ajout dans le tableau des carrefours r�duits avec ordonn�e
    Set monTabCarfY(unIndex) = unCarfY
    'Stockage dans le carrefour non r�duit de son carrefour r�duit
    Set unCarfReduit.monCarrefour.monCarfRed = unCarfReduit
End Sub

Public Sub CalculerDecVitesse(unCarfReduit As Object, unSensMontant As Boolean, unIndCarf As Integer, unIndCarfPred As Integer, unDecVitesse As Single)
    'Procedure appel� par CalculerTempsParcours
    'Elle calcule dans un sens consid�r�, le d�calage due aux vitesses
    'variable entre deux carrefours de m�me sens
    'Le d�calage unDecVitesse et l'indice du dernier carrefour dans le m�me
    'sens sont modifi�s pour �tre utilis�s au prochain appel de cette proc�dure
       
    If unIndCarfPred = 0 Then
        'Cas du premier carrefour dans le sens consid�r�
        '==> son d�calage est nul car il sert d'origine aux autres
        unDecVitesse = 0
    Else
        'Cumul du d�calage en ajoutant le d�calage due � la
        'vitesse variable entre les deux derniers carrefours
        'du sens consid�r� avec d�calage > 0 en sens montant, < 0 sinon
        
        'Calcul de la distance entre les deux derniers carrefours dans
        'le sens consid�r�, celui donn� par la valeur de unSensMontant
        uneDistance = unCarfReduit.DonnerYSens(unSensMontant) - monTabCarfY(unIndCarfPred).monCarfReduit.DonnerYSens(unSensMontant)
        
        'Calcul de la vitesse entre les deux derniers carrefours
        'On prend la vitesse variable du carrefour d'arriv�e dans le sens consid�r�
        'Si sens montant, c'est la vitesse du carrefour donn� par unCarfReduit
        'Si sens descendant, c'est la vitesse du carrefour pr�c�dent dans ce sens
        'c'est � dire le carrefour r�duit d'indice unIndCarfPred
        If unSensMontant Then
            uneVitesse = unCarfReduit.DonnerVitSens(unSensMontant)
        Else
            uneVitesse = monTabCarfY(unIndCarfPred).monCarfReduit.DonnerVitSens(unSensMontant)
        End If
        'Explication du code ci-dessus : Par polymorphisme, les m�thodes
        'DonnerVitSens et DonnerYSens sont appel�es sur les bonnes
        'instances des classes CarfReduitSensUnique et CarfReduitSensDouble
        
        'Cumul du d�calage entre les deux derniers carrefours
        'dans le sens consid�r�, celui donn� par unSensMontant
        'Le d�calage en temps est toujours > 0, on le multiplie ailleurs
        'dans le code par -1 pour le sens descendant et par 1 sinon
        unDecVitesse = unDecVitesse + Abs(uneDistance / uneVitesse)
    End If
    
    'Stockage du dernier carrefour rencontr� dans le sens consid�r�
    unIndCarfPred = unIndCarf
    
End Sub

Public Function CalculerBandeSensPrivi(unS As Single, unB1 As Single, unB2 As Single, unMinVert1 As Single, unMinVert2 As Single) As Integer
    'Calcul de la bande passante dans le sens privil�gi� qui est le sens 1
    If unS <= 0 Then
        'Cas sans solution � double sens
        CalculerBandeSensPrivi = DoubleSensImpossible
        unB1 = unMinVert1
    ElseIf unS >= unMinVert1 + unMinVert2 Then
        'Cas avec solution � double sens
        CalculerBandeSensPrivi = DoubleSensPossible
        unB1 = unMinVert1
        unB2 = unMinVert2
    ElseIf unS <= unMinVert1 Then
        'Cas sans solution � double sens
        CalculerBandeSensPrivi = DoubleSensImpossible
        unB1 = unMinVert1
    ElseIf unS > unMinVert1 Then
        'Cas avec solution � double sens
        CalculerBandeSensPrivi = DoubleSensPossible
        unB1 = unMinVert1
        unB2 = unS - unMinVert1
    End If
End Function

Public Sub StockerEtAfficherBandes(uneForm As Form, unB1 As Single, unB2 As Single, Optional unDecalModif As Boolean = False)
    'Stockage et affichage des bandes passantes calcul�es
    'si undecalModif est vrai on ne stocke et n'affiche que
    'les bandes modifiables
    
    'Arrondi au deuxi�me chiffre apr�s la virgule
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
    'Calcul des d�calages des carrefours si une solution � double sens
    'pour les bandes passantes a �t� trouv�e
    Dim unCarfRed As Object
    Dim unK1 As Single, unK2 As Single
    
    'Parcours des carrefours  r�duits � double sens
    For i = 1 To uneForm.mesCarfReduitsSens2.Count
        'On prend pour chaque carrefour le unK1 suivant (cf Dossier programmation)
        'Min (0, Dur�e de vert sens Desc - unB2 + unH - Ecart du carrefour - Ai * dur�e du cycle)
        'avec Ai =1 si unH > Ecart du carrefour, 0 sinon
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        unK1 = unCarfRed.maDureeVertD - unB2 + unH - unCarfRed.monEcart
        If unH > unCarfRed.monEcart Then
            unK1 = unK1 - uneForm.maDur�eDeCycle
        End If
        If unK1 > 0 Then unK1 = 0
        'Calcul et affichage des d�calages calcul�s et modifiables
        'ramen�s modulo dur�e du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRefM, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours r�duits � sens montant
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Valeurs possibles de unK1 = tout l'intervalle
        '[Largeur bande sens montante - Dur�e de vert sens Montant, 0]
        'On prend unK1 = 0
        unK1 = 0
        'Calcul et affichage des d�calages calcul�s et modifiables
        'ramen�s modulo dur�e du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRef, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours r�duits � sens descendant
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Valeurs possibles de unK2 = tout l'intervalle
        '[unH - Dur�e de vert sens descendant, unH - Largeur bande sens descendante]
        'On prend unK2 = unH - Largeur bande sens descendante
        unK2 = unH - unB2
        'Calcul et affichage des d�calages calcul�s et modifiables
        'ramen�s modulo dur�e du cycle
        CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRef, -unCarfRed.monCarrefour.monDecVitSensD
    Next i
End Sub

Public Sub CalculerDecalageSansDoubleSens(uneForm As Form)
    'Calcul des d�calages des carrefours si aucune solution � double sens
    'pour les bandes passantes n'a �t� trouv�e
    'Ceci ne se produit que pour une onde verte � sens privil�gi�
    Dim unCarfRed As Object
    Dim unK1 As Single, unK2 As Single
    
    'Parcours des carrefours  r�duits � double sens
    'Ils ont une ligne de feux dans le sens privil�gi�
    For i = 1 To uneForm.mesCarfReduitsSens2.Count
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        If monSite.monTypeOnde = OndeSensM Or (monSite.monTypeOnde = OndeTC And monSite.monTCM > 0) Then
            'Cas d'une onde � sens M privi mais incadrable dans le sens D
            'Les valeurs possibles pour k1 = tout l'intervalle
            '[Dur�e de vert sens Montant - unB1, 0] (cf Dossier programmation)
            'On prend pour chaque carrefour unK1 = 0
            unK1 = 0
            'Calcul et affichage des d�calages calcul�s et modifiables
            'ramen�s modulo dur�e du cycle
            CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRefM, unCarfRed.monCarrefour.monDecVitSensM
        ElseIf monSite.monTypeOnde = OndeSensD Or (monSite.monTypeOnde = OndeTC And monSite.monTCD > 0) Then
            'Cas d'une onde � sens D privi mais incadrable dans le sens M
            'Les valeurs possibles pour k2 = tout l'intervalle
            '[Dur�e de vert sens Descendant - unB2, 0] (cf Dossier programmation)
            'On prend pour chaque carrefour unK2 = 0
            unK2 = 0
            'Calcul et affichage des d�calages calcul�s et modifiables
            'ramen�s modulo dur�e du cycle
            CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRefD, -unCarfRed.monCarrefour.monDecVitSensD
        Else
            MsgBox "Erreur de programmation dans OndeV dans CalculerDecalageSansDoubleSens", vbCritical
        End If
    Next i
    
    'Parcours des carrefours r�duits � sens montant
    For i = 1 To uneForm.mesCarfReduitsSensM.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Cas o� le sens privil�gi� est le montant
        '==> Valeurs possibles de unK1 = tout l'intervalle
        '[Largeur bande sens montante - Dur�e de vert sens Montant, 0]
        'Si le sens privil�gi� est le descendant k1 = 0 marche aussi
        'don on prend unK1 = 0 (cf Dossier programmation)
        unK1 = 0
        'Calcul et affichage des d�calages calcul�s et modifiables
        'ramen�s modulo dur�e du cycle
        CalculerDecalage uneForm, unCarfRed, unK1, unCarfRed.maPosRef, unCarfRed.monCarrefour.monDecVitSensM
    Next i
    
    'Parcours des carrefours r�duits � sens descendant
    For i = 1 To uneForm.mesCarfReduitsSensD.Count
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Cas o� le sens privil�gi� est le descendant
        '==> Valeurs possibles de unK2 = tout l'intervalle
        '[Largeur bande sens descendante - Dur�e de vert sens descendant, 0]
        'Si le sens privil�gi� est le montant k2 = 0 marche aussi
        'donc on prend unK2 = 0 (cf Dossier programmation)
        unK2 = 0
        'Calcul et affichage des d�calages calcul�s et modifiables
        'ramen�s modulo dur�e du cycle
        CalculerDecalage uneForm, unCarfRed, unK2, unCarfRed.maPosRef, -unCarfRed.monCarrefour.monDecVitSensD
    Next i
End Sub


Public Sub CalculerDecalage(uneForm As Form, unCarfRed As Object, unK As Single, unePosRef As Single, unDecVit As Single)
    'Calcul et Affichage des d�calages calcul�s et modifiables du
    'carrefour obetenu gr�ce � son carrefour r�duit
    
    'Calcul des d�calages calcul�s et modifiables
    unCarfRed.monCarrefour.monDecCalcul = unK - unePosRef + unDecVit
    unCarfRed.monCarrefour.monDecCalcul = ModuloZeroCycle(unCarfRed.monCarrefour.monDecCalcul, uneForm.maDur�eDeCycle)
    unCarfRed.monCarrefour.monDecModif = unCarfRed.monCarrefour.monDecCalcul
    
    'Si le carrefour r�duit est celui cr�� pour les dates impos�es o� n'
    'affiche pas sa valeur
    If unCarfRed.monCarrefour.maPosition <= 0 Then Exit Sub
    
    'Affichage dans l'onglet Tableau de r�sultat en arrondissant � l'entier
    'le plus proche, d'o� l'utilisation de la fonction VB5 CInt
    uneForm.TabDecal.Row = unCarfRed.monCarrefour.maPosition
    uneForm.TabDecal.Col = 2
    If CIntCorrig�(unCarfRed.monCarrefour.monDecCalcul) = uneForm.maDur�eDeCycle Then
        'Une valeur valant dur�e du cycle s'affiche 0
        uneForm.TabDecal.Text = "0"
    Else
        uneForm.TabDecal.Text = CIntCorrig�(unCarfRed.monCarrefour.monDecCalcul)
    End If
    uneForm.TabDecal.Col = 3
    If CIntCorrig�(unCarfRed.monCarrefour.monDecModif) = uneForm.maDur�eDeCycle Then
        'Une valeur valant dur�e du cycle s'affiche 0
        uneForm.TabDecal.Text = "0"
    Else
        uneForm.TabDecal.Text = CIntCorrig�(unCarfRed.monCarrefour.monDecModif)
    End If
End Sub

Public Function RecalculerBandesPassantes(uneForm As Form) As Boolean
    'Recalcul des bandes passantes du site donn� par uneForm
    'apr�s la modification d'un d�calage dans l'onglet Tableau
    'des r�sultats.
    'Retour : VRAI si recalcul a �t� possible, FAUX sinon
    
    'Cr�ation d'un nouveau carrefour qui contiendra tous les feux
    '�quivalents montant et descendant des carrefours r�duits dont on
    'cherchera le feu �quivalent montant et descendant, les nouvelles bandes
    'passantes correspondront aux plages de vert maximales trouv�es
    Dim unCarf As Carrefour
    Dim uneColCarf As New ColCarrefour
    Dim unB1 As Single, unB2 As Single
    Dim unDebVertM As Single, unDebVertD As Single
       
    'Cr�ation d'un nouveau carrefour qui contiendra tous les feux
    '�quivalents montant et descendant des carrefours r�duits
    'avec des vitesses non nulles.
    Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
    
    'R�duction des carrefours r�duits
    unResultat = ReduireCarfReduits(uneForm, unCarf, unB1, unB2, unDebVertM, unDebVertD)
    If unResultat >= 0 Then
        '> 0 pour les cas o� il y a r�duction r�ussi et 0 sinon
        'avec >=0 on permet l'affichage et le dessin des bandes communes
        'nulles et de voir les bandes inter-carrefours (demande sites pilotes)
        RecalculerBandesPassantes = True
    Else
        RecalculerBandesPassantes = False
    End If
    
    If RecalculerBandesPassantes Then
        'Cas de r�ussite de la r�duction des carrefours r�duits
        'Stocker et afficher les nouvelles bandes passantes
        StockerEtAfficherBandes uneForm, unB1, unB2, True
    End If
    
    'Suppression de la collection ne contenant que le carrefour cr�� au d�but
    'pour lib�rer la m�moire.
    '(les events Terminate sont d�clench�s sur les classes ColCarrefour,
    'Carrefour et ColFeu)
    uneColCarf.Remove 1
    Set unCarf = Nothing
    Set uneColCarf = Nothing
End Function
    
Public Function ReduireCarfReduits(uneForm As Form, unCarf As Carrefour, uneDureeVertM As Single, uneDureeVertD As Single, unDebVertM As Single, unDebVertD As Single) As Integer
    'R�duction du carrefour qui contient tous les feux
    '�quivalents montant et descendant des carrefours r�duits
    
    'Elle retourne un entier valant :
    '   - 0 si aucun feu �quivalent trouv�
    '   - 1 si un feu �quivalent trouv� (montant ou descendant)
    '   - 2 si deux feux �quivalents trouv�s (montant et descendant)
    
    Dim unFeu As Feu, unTC As TC
    Dim unCarfRed As Object
    Dim uneOrdonnee As Integer, unPosRef As Single
    Dim unNbFeuxSensM As Integer, unNbFeuxSensD As Integer
    Dim unNbFeuxSens2 As Integer
    
    'Ajout � ce carrefour global des feux �quivalents
    'des carrefours r�duits double sens
    unNbFeuxSens2 = uneForm.mesCarfReduitsSens2.Count
    For i = 1 To unNbFeuxSens2
        'R�cup du carrefour r�duit
        Set unCarfRed = uneForm.mesCarfReduitsSens2(i)
        'Ajout d'un nouveau feu montant
        Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonneeM, unCarfRed.maDureeVertM, unCarfRed.maPosRefM)
        'Stockage du carrefour du feu cr��
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
        'Ajout d'un nouveau feu descendant
        Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonneeD, unCarfRed.maDureeVertD, unCarfRed.maPosRefD)
        'Stockage dans le feu cr�� du carrefour correspondant � celui r�duit
        'car c'est son d�calage en temps du � sa vitesse qui est utilis� dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
    
    'Ajout � ce carrefour global des feux �quivalents
    'des carrefours r�duits � sens unique montant
    unNbFeuxSensM = uneForm.mesCarfReduitsSensM.Count
    For i = 1 To unNbFeuxSensM
        'R�cup du carrefour r�duit
        Set unCarfRed = uneForm.mesCarfReduitsSensM(i)
        'Ajout d'un nouveau feu montant
        Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
        'Stockage dans le feu cr�� du carrefour correspondant � celui r�duit
        'car c'est son d�calage en temps du � sa vitesse qui est utilis� dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
    
    'Ajout � ce carrefour global des feux �quivalents
    'des carrefours r�duits � sens unique descendant
    unNbFeuxSensD = uneForm.mesCarfReduitsSensD.Count
    For i = 1 To unNbFeuxSensD
        'R�cup du carrefour r�duit
        Set unCarfRed = uneForm.mesCarfReduitsSensD(i)
        'Ajout d'un nouveau feu descendant
        Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
        'Stockage dans le feu cr�� du carrefour correspondant � celui r�duit
        'car c'est son d�calage en temps du � sa vitesse qui est utilis� dans
        'le calcul des bandes passantes
        Set unFeu.monCarrefour = unCarfRed.monCarrefour
    Next i
      
    'Initialisation de la valeur de retour de cette fonction
    ReduireCarfReduits = 0
    
    'Initialisation des r�sultats des feux �quivalents montant et descendant
     unFeuEquivSensMExist = True
     unFeuEquivSensDExist = True
    
    'Calcul du feu �quivalent montant �ventuel
    If unNbFeuxSens2 > 0 Or unNbFeuxSensM > 0 Then
        If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
            'Cas d'une onde cadr�e par unTC dans le sens montant
            Set unTC = monSite.mesTC(monSite.monTCM)
        Else
            'Cas d'une onde non cadr�e par unTC dans le sens montant
            Set unTC = Nothing
        End If
        
        unFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVertM, unPosRef, uneOrdonnee, True, , , , unTC)
        If unFeuEquivSensMExist Then
            'Stockage du d�but de la plus grande plage de
            'vert trouv� lors du calcul du feu �quivalent
            unDebVertM = monDebutVert
            'Calcul de la valeur de retour de cette fonction
            ReduireCarfReduits = ReduireCarfReduits + 1
        Else
            'Cas o� aucune solution de bande passante montante n'a �t� trouv�
            unMsg = "Aucune solution de bande passante montante n'a pu �tre trouv�e"
        End If
    End If
                      
    'Calcul du feu �quivalent descendant �ventuel
    If unNbFeuxSens2 > 0 Or unNbFeuxSensD > 0 Then
        If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
            'Cas d'une onde cadr�e par unTC dans le sens descendant
            Set unTC = monSite.mesTC(monSite.monTCD)
        Else
            'Cas d'une onde non cadr�e par unTC dans le sens descendant
            Set unTC = Nothing
        End If
        
        unFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unPosRef, uneOrdonnee, True, , , , unTC)
        If unFeuEquivSensDExist Then
            'Stockage du d�but de la plus grande plage de
            'vert trouv� lors du calcul du feu �quivalent
            unDebVertD = monDebutVert
            'Calcul de la valeur de retour de cette fonction
            ReduireCarfReduits = ReduireCarfReduits + 1
        Else
            'Cas o� aucune solution de bande passante descendante n'a �t� trouv�
            unMsg = unMsg + Chr(13) + "Aucune solution de bande passante descendante n'a pu �tre trouv�e"
        End If
    End If
    
    If ReduireCarfReduits = 0 Then
        'Cas d'�chec de la r�duction des carrefours r�duits
        'MsgBox unMsg, vbCritical
        'On n'affiche plus de message d'erreur ainsi on affichera et
        'stockera les bandes M et D m�me si elles sont nulles
    End If
    
    'Lib�ration de la m�moire en supprimant tous les feux cr��s
    For j = 1 To unCarf.mesFeux.Count
    'Suppression du 1er dans une collection
    '==> Suppression de tout car le 2�me devient 1er, etc...
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
    
    'Valeurs par d�faut avant r�alisation de l'algo final
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
    
    'Cas de pr�sence de feux montants
    If unNbFeuxM <> 0 Then
        'R�organisation par ordonn�e croissant des feux montants
        TrierFeuYCroissant unTabFeuM, 1
    
        'V�rification du passage � tous les verts montants
        'avec la vitesse maxi limite dans le sens montant
        uneVerif = VerifierVitessePasseToutVert(unSite, CInt(uneVitMaxLim * 3.6), unTabFeuM, True)
                              
        If uneVerif <> 0 Then
            'Cas o� la vitesse maxi montante est > � la vitesse maxi limite
            unSite.maVitMaxM = "> " + Format(CInt(uneVitMaxLim * 3.6))
        Else
            'Cas o� une vitesse max est < � la vitesse maxi limite
            
            'Calcul de la vitesse max montante possible en km/h
            '< � la vitesse maxi limite
            uneVMax = CalculerVMaxInfVMaxLim(unSite, uneVitMinLim, uneVitMaxLim, unTabFeuM, unTabVM, unTabIndFeuM, unTabDTM, unTabDYM, 1)
            
            'Stockage de la vitesse montante maxi trouv�e
            If uneVMax < uneVitMinLim * 3.6 + 0.001 Then
                unSite.maVitMaxM = "< " + Format(CInt(uneVitMinLim * 3.6))
            Else
                unSite.maVitMaxM = Format(uneVMax)
            End If
        End If
    Else
        unSite.maVitMaxM = ""
    End If
    
    'Cas de pr�sence de feux descendants
    If unNbFeuxD <> 0 Then
        'R�organisation par ordonn�e croissant des feux descendants
        TrierFeuYCroissant unTabFeuD, -1
    
        'V�rification du passage � tous les verts descendants
        'avec la vitesse maxi limite dans le sens descendant
        uneVerif = VerifierVitessePasseToutVert(unSite, CInt(uneVitMaxLim * 3.6), unTabFeuD, False)
                              
        If uneVerif <> 0 Then
            'Cas o� la vitesse maxi descendante est > � la vitesse maxi limite
            unSite.maVitMaxD = "> " + Format(CInt(uneVitMaxLim * 3.6))
        Else
            'Cas o� une vitesse max est < � la vitesse maxi limite
            
            'Calcul de la vitesse max descendante possible en km/h
            '< � la vitesse maxi limite remise en valeur positive
            uneVMax = CalculerVMaxInfVMaxLim(unSite, uneVitMinLim, uneVitMaxLim, unTabFeuD, unTabVD, unTabIndFeuD, unTabDTD, unTabDYD, -1)
            
            'Stockage de la vitesse descendante maxi trouv�e
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
        'Affectation d'une couleur pour les cellules lock�es
        'Cette onglet n'est qu'une �dition ==> pas de saisie
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
        
        'Remplissage des r�sultats carrefours
        .TabFicheCarf.MaxRows = .mesCarrefours.Count
        For i = 1 To .mesCarrefours.Count
            Set unCarf = .mesCarrefours(i)
            .TabFicheCarf.Row = i
            .TabFicheCarf.Col = 1
            .TabFicheCarf.Text = unCarf.monNom
            If unCarf.monDecCalcul = -99 Then
                'Cas des carrefours inutilis�s ou non compris entre
                'Ymin et Ymax d'une onde cadr�e par un TC
                For j = 2 To 7
                    .TabFicheCarf.Col = j
                    .TabFicheCarf.Text = ""
                Next j
            Else
                'Affichage du d�calage en arrondissant � l'entier le plus
                'proche, d'o� l'utilisation de la fonction VB5 CInt
                .TabFicheCarf.Col = 2
                If CIntCorrig�(unCarf.monDecModif) = .maDur�eDeCycle Then
                    'Une valeur valant dur�e du cycle s'affiche 0
                    .TabFicheCarf.Text = "0"
                Else
                    .TabFicheCarf.Text = CIntCorrig�(unCarf.monDecModif)
                End If
                'Affichage en fonction du type de carrefour
                'r�duit (double sens ou sens unique)
                If TypeOf unCarf.monCarfRed Is CarfReduitSensDouble Then
                    .TabFicheCarf.Col = 3
                    uneRCap = unCarf.monCarfRed.maDureeVertM / .maDur�eDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                    'Mise en rouge des r�serves de capacit� n�gatives
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
                    uneRCap = unCarf.monCarfRed.maDureeVertD / .maDur�eDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                    'Mise en rouge des r�serves de capacit� n�gatives
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
                    .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par d�faut
                    .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(True) * 3.6)
                    .TabFicheCarf.Col = 6
                    .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(False) * -3.6)
                    'Affichage du D�calage � l'ouverture
                    'Il est ind�termin� si plusieurs lignes de feux dans le
                    'm�me sens (Carrefour <> Carf r�duit)==> Affichage "Ind�fini"
                    unCarf.DonnerNbFeuxMetD unNbFeuxM, unNbFeuxD
                    .TabFicheCarf.Col = 7
                    If unNbFeuxM = 1 And unNbFeuxD = 1 Then
                        .TabFicheCarf.Text = CInt(unCarf.monCarfRed.maPosRefM - unCarf.monCarfRed.maPosRefD)
                    Else
                        .TabFicheCarf.Text = "Ind�fini"
                    End If
                Else
                    If unCarf.monCarfRed.monSensMontant Then
                        'Cas d'un carrefour � sens unique montant
                        .TabFicheCarf.Col = 3
                        uneRCap = unCarf.monCarfRed.maDureeVert / .maDur�eDeCycle * unCarf.monDebSatM - unCarf.maDemandeM
                        'Mise en rouge des r�serves de capacit� n�gatives
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
                        .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par d�faut
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 5
                        .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(True) * 3.6)
                        .TabFicheCarf.Col = 6
                        .TabFicheCarf.Text = ""
                    Else
                        'Cas d'un carrefour � sens unique descendant
                        .TabFicheCarf.Col = 3
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 4
                        uneRCap = unCarf.monCarfRed.maDureeVert / .maDur�eDeCycle * unCarf.monDebSatD - unCarf.maDemandeD
                        'Mise en rouge des r�serves de capacit� n�gatives
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
                        .TabFicheCarf.ForeColor = .ForeColor 'On remet la couleur par d�faut
                        .TabFicheCarf.Text = ""
                        .TabFicheCarf.Col = 6
                        .TabFicheCarf.Text = CInt(unCarf.DonnerVitSens(False) * -3.6)
                    End If
                    'D�calage � l'ouverture ind�termin� ==> Affichage "Ind�fini"
                    .TabFicheCarf.Col = 7
                    .TabFicheCarf.Text = "Ind�fini"
                End If
            End If
        Next i
        
        'Remplissage des r�sultats des TC utilis�s
        .TabFicheTC.MaxRows = .mesTCutil.Count
        For i = 1 To .mesTCutil.Count
            Set unTC = .mesTCutil(i)
            If .maModifDataTC Or .maModifDataOndeTC Then
                'Recalcul du tableau de marche de progression s'il y a eu une
                'modif dans les donn�es TC, de plus cela donne le sens du TC
                unSens = unTC.CalculerTableauMarcheProg()
            Else
                'D�termination du sens du TC
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
    'Dessin de l'onde verte, des plages de vert et les points de r�f�rence
    'des feux des carrefours et les parcours des TC choisis
    'unX0 et unY0 sont les coordonn�es du point bas gauche
    
    'Elle r�alise aussi le dessin des progressions des TC si
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
    
    'Stockage pour savoir si on dessine � l'�cran ou sur imprimante
    'uneSortieImprimante = vrai si dessin sur imprimante faux si dessin �cran
    uneSortieImprimante = (TypeOf uneZoneDessin Is Printer)
    
    'On vide les collections contenant les �l�ments graphics des ondes
    If uneSortieImprimante = False Then
        ViderCollection monSite.maColPlageGraphicD
        ViderCollection monSite.maColPlageGraphicM
        ViderCollection monSite.maColRefGraphicD
        ViderCollection monSite.maColRefGraphicM
    End If
    
    'D�termination du type d'onde verte � dessiner
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde verte cadr�e par un TC dans le sens montant
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde verte non cadr�e par un TC dans le sens montant
        Set unTCM = Nothing
    End If
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde verte cadr�e par un TC dans le sens descendant
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde verte non cadr�e par un TC dans le sens descendant
        Set unTCD = Nothing
    End If
    
    With monSite
        'D�termination de la hauteur englobante du dessin = Temps englobant
        TrouverTempsParcoursEtCarrefours unIndCarfM, unIndCarfD, unTmpM, unTmpD
                
        'Dessin des ondes vertes montantes et descendantes
        
        'Cr�ation d'un nouveau carrefour qui contiendra tous les feux
        '�quivalents montant et descendant des carrefours r�duits
        'avec des vitesses non nulles.
        'Ainsi on obtient le d�but de vert montant et descendant de la plus
        'plage de vert dans ces sens qui sert pour le dessin
        '(cf Dossier programmation, R�presentation graphique)
        Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
        
        'R�duction des carrefours r�duits pour avoir unDebVertM et unDebVertD
        unResultat = ReduireCarfReduits(monSite, unCarf, uneDureeVertM, uneDureeVertD, unDebVertM, unDebVertD)
        
        'Suppression pour lib�rer la m�moire.
        uneColCarf.Remove 1
        Set unCarf = Nothing
        Set uneColCarf = Nothing

        If unResultat >= 0 Then
            '> 0 ==> Cas de r�ussite de la r�duction des carrefours r�duits
            'donc Dessin des ondes communes possibles
            '=0 dessin des ondes inter-carrefours m�me si l'onde commune
            'est impossible, donc >=0 tous les cas de figures
            '> 0 remplac� par >=0 apr�s demande sites pilotes pour le cas = 0
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
                        'Cas d'un carrefour � sens unique montant
                        If unIndCarfBasM = 0 Then
                            unIndCarfBasM = i
                        End If
                    Else
                         'Cas d'un carrefour � sens unique descendant
                        If unIndCarfBasD = 0 Then
                            unIndCarfBasD = i
                        End If
                   End If
                End If
            Loop While (unIndCarfBasM = 0 Or unIndCarfBasD = 0) And i < unNbCarf
        
            'D�termination du premier point de l'onde verte montante
            '(cf Dossier programmation : Repr�sentation graphique)
            'Ce point correspond au carrefour le plus bas ayant un feu
            'montant, indice=unIndCarfBasM calcul� juste avant
            If unIndCarfBasM > 0 Then
                'Cas o� un carrefour ayant un feu montant a �t� trouv�
                Set unCarfRedM1 = monTabCarfY(unIndCarfBasM).monCarfReduit
                Set unCarf = unCarfRedM1.monCarrefour
                'Calcul de la r�f�rence et du Y du carrefour r�duit
                'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                unePosRef = unCarfRedM1.DonnerPosRefSens(True)
                unYM1 = unCarfRedM1.DonnerYSens(True)
                'Calcul du d�but de vert du carrefour r�duit
                unDebVertM0 = unCarf.monDecModif + unePosRef + unCarf.monDecVitSensM
                If .maBandeModifM = 0 Then
                    'Cas d'impossibilit� d'avoir une bande montante commune
                    unK = 0
                Else
                    'Cas d'existence d'une bande passante montante commune
                    unK = ModuloZeroCycle(unDebVertM - unDebVertM0, .maDur�eDeCycle)
                End If
                'Abscisse du premier point, donc un temps ramen� modulo cycle
                unTM1 = ModuloZeroCycle(unCarf.monDecModif + unePosRef + unK, .maDur�eDeCycle)
            
                'Correction du TM1 s'il n'y a qu'un carrefour ayant des feux
                'montants, sinon le dessin d'onde verte est erron� la bande
                'passante montante ne relie pas les plages des verts montants
                If monSite.mesCarfReduitsSens2.Count + monSite.mesCarfReduitsSensM.Count = 1 Then
                    unTM1 = 0
                    unTmpM = monSite.maDur�eDeCycle 'pour �viter une valeur nulle
                End If
            End If
            
            'D�termination du premier point de l'onde verte descendante
            '(cf Dossier programmation : Repr�sentation graphique)
            'Ce point correspond au carrefour le plus haut ayant un feu
            'descendant, indice=unIndCarfD calcul� en d�but de cette fonction
            If unIndCarfD > 0 Then
                'Cas o� un carrefour ayant un feu descendant a �t� trouv�
                Set unCarfRedD1 = monTabCarfY(unIndCarfD).monCarfReduit
                Set unCarf = unCarfRedD1.monCarrefour
                'Calcul de la r�f�rence et du Y du carrefour r�duit
                'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                unePosRef = unCarfRedD1.DonnerPosRefSens(False)
                unYD1 = unCarfRedD1.DonnerYSens(False)
                'Calcul du d�but de vert du carrefour r�duit
                unDebVertD0 = unCarf.monDecModif + unePosRef + unCarf.monDecVitSensD
                If .maBandeModifD = 0 Then
                    'Cas d'impossibilit� d'avoir une bande descendante commune
                    unK = 0
                Else
                    'Cas d'existence d'une bande passante descendante commune
                    unK = ModuloZeroCycle(unDebVertD - unDebVertD0, .maDur�eDeCycle)
                End If
                'Abscisse du premier point, donc un temps ramen� modulo cycle
                unTD1 = ModuloZeroCycle(unCarf.monDecModif + unePosRef + unK, .maDur�eDeCycle)
                
                'Correction du TD1 s'il n'y a qu'un carrefour ayant des feux
                'descendants, sinon le dessin d'onde verte est erron� la bande
                'passante descendante ne relie pas les plages des verts descendants
                If monSite.mesCarfReduitsSens2.Count + monSite.mesCarfReduitsSensD.Count = 1 Then
                    unTD1 = unDebVertD0
                    unTmpD = monSite.maDur�eDeCycle 'pour �viter une valeur nulle
                End If
            End If
            
            
            If unIndCarfM > 0 Then
                'Calcul du nombre entier de cycle parcouru par l'onde montante
                unNbCycleM = Int((unTmpM + unTM1) / .maDur�eDeCycle)
                'Calcul du fin de vert maximun des feux du carrefour montant le + haut
                Set unCarf = monTabCarfY(unIndCarfM).monCarfReduit.monCarrefour
                unMaxFinVertHaut = unNbCycleM * .maDur�eDeCycle + TrouverMaxFinVert(unCarf) Mod .maDur�eDeCycle
                If unTmpM + unTM1 > unMaxFinVertHaut + 0.001 Then unMaxFinVertHaut = unMaxFinVertHaut + .maDur�eDeCycle
            Else
                unMaxFinVertHaut = -.maDur�eDeCycle '==> Max le + petit
            End If
            
            'Calcul du d�but de vert minimun des feux du
            'carrefour descendant le + haut
            If unIndCarfD > 0 Then
                Set unCarf = monTabCarfY(unIndCarfD).monCarfReduit.monCarrefour
                unMinDebVertHaut = TrouverMinDebVert(unCarf)
                'Cadrage dans le cycle
                If unTD1 < unMinDebVertHaut - 0.001 Then unMinDebVertHaut = unMinDebVertHaut - .maDur�eDeCycle
                If unTD1 > TrouverMaxFinVert(unCarf) + 0.001 Then unMinDebVertHaut = unMinDebVertHaut + .maDur�eDeCycle
            Else
                unMinDebVertHaut = .maDur�eDeCycle '==> Min le + grand
            End If
            
            'Calcul du nombre entier de cycle parcouru par l'onde descendante
            unNbCycleD = Int((unTmpD + unTD1) / .maDur�eDeCycle)
            'Calcul du fin de vert maximun des feux du carrefour descendant le + bas
            If unIndCarfBasD > 0 Then
                Set unCarf = monTabCarfY(unIndCarfBasD).monCarfReduit.monCarrefour
                unMaxFinVerBas = unNbCycleD * .maDur�eDeCycle + TrouverMaxFinVert(unCarf) Mod .maDur�eDeCycle
                If unTmpD + unTD1 > unMaxFinVerBas + 0.001 Then unMaxFinVerBas = unMaxFinVerBas + .maDur�eDeCycle
            Else
                unMaxFinVerBas = -.maDur�eDeCycle  '==> Max le + petit
            End If
            
            'Calcul du d�but de vert minimun des feux du
            'carrefour montant le + bas
            If unIndCarfBasM > 0 Then
                Set unCarf = monTabCarfY(unIndCarfBasM).monCarfReduit.monCarrefour
                unMinDebVertBas = TrouverMinDebVert(unCarf)
                'Cadrage dans le cycle
                If unTM1 < unMinDebVertBas - 0.001 Then
                    unMinDebVertBas = unMinDebVertBas - .maDur�eDeCycle
                End If
                If unTM1 > TrouverMaxFinVert(unCarf) + 0.001 Then
                    unMinDebVertBas = unMinDebVertBas + .maDur�eDeCycle
                End If
            Else
                unMinDebVertBas = .maDur�eDeCycle '==> Min le + grand
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
            
            'D�termination de la hauteur englobante du dessin = Temps englobant
            unT = unMaxT - unMinT
            'Stockage pour utilisation ailleurs
            monSite.monTmpTotal = unT
            monSite.monTMin = unMinT
            
            'D�termination de l'�cart en Y englobant le dessin
            If unDessinDansOnglet Then
                'Cas o� l'on dessine l'onde verte dans
                'l'onglet Graphique onde verte
                unYMin = .monYMinFeu
                unYMax = .monYMaxFeu
            Else
                'Cas o� l'on dessine l'onde verte dans
                'la fen�tre pleine �cran ou sur imprimante
                'On ne prend en compte que les carrefours utilis�s
                '==> niveau de zoom diff�rent
                TrouverMinYMaxY unYMin, unYMax
                If unYMax = unYMin Then
                    'Cas d'un seul carrefour avec un seul
                    'Pour �viter unDY = unYMax - unYMin = 0
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
                'Stockage dans une variable priv�e de ce module
                'pour utilisation ailleurs
                monSite.monOrigX = unNewX0
                
                'Calcul de l'englobant en temps des progressions de TC
                's�lectionn�s pour connaitre l'englobant en coordonn�es �cran
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
                
                'Englobant en Temps total en valeur �cran
                unTmpTotalEcran = uneLg
                If monSite.monTDepTCMin < unX0 Then
                    'Cas o� l'englobant des progressions TC commence
                    'avant l'englobant des ondes
                    '==> Changement du d�but de l'englobant
                    'en convertissant valeur �cran en r�elle
                    unMinT = (monSite.monTDepTCMin - monSite.monOrigX) / unTmpTotalEcran * monSite.monTmpTotal
                    'Modification du min en temps pour avoir des lignes
                    'de rappels toutes les 10 secondes englobant ce min
                    ModifierMinTempsPourVisu unMinT
                    monSite.monTMin = unMinT
                    'D�termination de la nouvelle largeur englobante
                    'du dessin = Temps englobant
                    monSite.monTmpTotal = unMaxT - unMinT
                End If
                
                If unTmpTotalEcran + unX0 < monSite.monTFinTCMax Then
                    'Cas o� l'englobant en temps des progressions TC est plus
                    'grand que l'englobant en temps des ondes
                    '==> Changement de l'englobant en temps
                    'en convertissant valeur �cran en r�elle
                    unMaxT = (monSite.monTFinTCMax - unX0) / unTmpTotalEcran * monSite.monTmpTotal + monSite.monTMin
                    'Modification du min en temps pour avoir des lignes
                    'de rappels toutes les 10 secondes englobant ce min
                    ModifierMaxTempsPourVisu unMaxT
                    'D�termination de la nouvelle largeur englobante
                    'du dessin = Temps englobant
                    monSite.monTmpTotal = unMaxT - unMinT
                End If
                
                'Stockage pour utilisation ailleurs
                unT = monSite.monTmpTotal
                
                'Calcul de l'origine pour les temps
                monSite.monOrigX = unX0 - ConvertirReelEnEcran(CLng(unMinT), unT, uneLg)
                
                'Dessin des progressions des TC pour quelles soient
                'affich�es avant les plages de vert des feux, qu'ainsi
                'elles ne masqueront pas
                unNbTCUtil = monSite.mesTCutil.Count
                For i = 1 To unNbTCUtil
                    Set unTC = monSite.mesTCutil(i)
                    TracerProgressionTC uneZoneDessin, unTC, unX0, unY0, uneLg, uneHt
                Next i
            End If
            
            'Conversion de la dur�e du cycle en valeur �cran (twips)
            'Cette variable locale est utilis�e plus bas
            uneLongCycle = ConvertirReelEnEcran(.maDur�eDeCycle, unT, uneLg)
        
            'Calcul de l'origine pour les temps
            unNewX0 = unX0 - ConvertirReelEnEcran(CLng(unMinT), unT, uneLg)
            'Stockage dans une variable priv�e de ce module
            'pour utilisation ailleurs
            monSite.monOrigX = unNewX0
            
            'Conversion en coordonn�es �cran du point (unTD1, unYD1)
            'Sert de premier point � l'onde verte descendante
            unXDpred = ConvertirSingleEnEcran(unTD1, unT, uneLg)
            unXDpred = unXDpred + unNewX0
            unYDpred = ConvertirReelEnEcran(unYD1 - unYMin, unDY, uneHt)
            unYDpred = unY0 - unYDpred
            
            'Conversion en coordonn�es �cran du point (unTM1, unYM1)
            'Sert de premier point � l'onde verte montante
            unXMpred = ConvertirSingleEnEcran(unTM1, unT, uneLg)
            unXMpred = unXMpred + unNewX0
            unYMpred = ConvertirReelEnEcran(unYM1 - unYMin, unDY, uneHt)
            unYMpred = unY0 - unYMpred
            
            'Dessin des ondes vertes
            unMsgDessinOnde = ""
            'unNoDessinOndeM = (.monOndeDoubleTrouve = False And .monTypeOnde = OndeSensD) Or .maBandeModifM = 0
            unNoDessinOndeM = (.maBandeModifM = 0)
            If unNoDessinOndeM And monSite.mesCarfReduitsSensM.Count > 0 And uneSortieImprimante = False Then
                'Cas d'une onde � sens privil�gi� descendant mais ayant des feux
                'montants mais ne pouvant pas �tre cadrer dans le sens montant
                '==> Pas d'onde verte montante, d'o� pas de dessin.
                unMsgDessinOnde = "Pas de dessin de l'onde verte MONTANTE car elle n'existe pas."
            End If
            
            'unNoDessinOndeD = (.monOndeDoubleTrouve = False And .monTypeOnde = OndeSensM) Or .maBandeModifD = 0
            unNoDessinOndeD = (.maBandeModifD = 0)
            If unNoDessinOndeD And monSite.mesCarfReduitsSensD.Count > 0 And uneSortieImprimante = False Then
                'Cas d'une onde � sens privil�gi� montant mais ayant des feux
                'descendants ne pouvant pas �tre cadrer dans le sens descendant
                '==> Pas d'onde verte descendant, d'o� pas de dessin.
                If unMsgDessinOnde = "" Then
                    unMsgDessinOnde = "Pas de dessin de l'onde verte DESCENDANTE car elle n'existe pas."
                Else
                    unMsgDessinOnde = unMsgDessinOnde + Chr(13) + Chr(13) + "Pas de dessin de l'onde verte DESCENDANTE car elle n'existe pas."
                End If
            End If
            
            'Affichage du non dessin d'onde �ventuelle
            If unMsgDessinOnde <> "" Then MsgBox unMsgDessinOnde, vbInformation
                        
            'Initialisation de variables donnant un num�ro de phases
            'elles commencent � 1
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
                'pour l'onde verte montante car on dessine � partir du
                'carrefour le plus bas ayant un feu montant
                Set unCarfRed = monTabCarfY(i).monCarfReduit
                Set unCarf = unCarfRed.monCarrefour
                'Conversion en valeur �cran de la largeur de bande montante
                uneLBM = ConvertirSingleEnEcran(.maBandeModifM, unT, uneLg)
                If unIndCarfM > 0 Then
                    'Cas d'une onde verte montante possible ==> Dessin
                    'unIndCarfM > 0 dit qu'on a trouv� des carrefours montants
                    If unCarfRed.HasFeuMontant = True Then
                        'Cas d'un carrefour contraignant de l'onde verte montante
                        'donc ayant un feu de sens montant
                        
                        'Abscisse du point suivant de l'onde montante vaut
                        'l'abscisse du point pr�c�dent plus le d�calage en
                        'temps entre le carrefour courant et le premier montant
                        unTM2 = unTM1 + unCarf.monDecVitSensM - unCarfRedM1.monCarrefour.monDecVitSensM
                        'Stockage dans d�but onde pour trouver les plages
                        's�lectionnables graphiquement
                        unCarfRed.AffecterDebOndeSens unTM2, True
                            
                        'Ordonn�e �gale � l'ordonn�e du carrefour r�duit courant
                        'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                        unYM2 = unCarfRed.DonnerYSens(True)
                        
                        'Conversion en coordonn�es �cran de unYM2
                        unY = ConvertirReelEnEcran(unYM2 - unYMin, unDY, uneHt)
                        unY = unY0 - unY
                        
                        'Dessin de l'onde verte montante inter-carrefours donc
                        'entre ce carrefour r�duit et son pr�c�dent en Y
                        'si choix coch� dans les options d'affichage et d'impression
                        'et si ce n'est pas le 1er carrefour montant le + bas
                        
                        'Fait avant la bande commune pour ne voir que la bande
                        'commune si superposition avec la bande inter-carrefours
                        If .mesOptionsAffImp.maVisuBandInterCarfM And i <> unIndCarfBasM Then
                            'Cas d'une onde montante pas cadr�e par un TC montant
                            'Calcul du d�but de vert de ce carrefour r�duit
                            unDebVertM = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(True)
                            unDebVertM = ModuloZeroCycle(unDebVertM, .maDur�eDeCycle)
                            'Calcul du nombre de cycle s�parant le d�but de
                            'vert du d�but de l'onde verte montante
                            unNbCycle = Fix((0.001 + unTM2 - unDebVertM) / .maDur�eDeCycle)
                            If unNbCycle < 0 And .maBandeModifM = 0 Then
                                'Si pas de bande commune, on ne peut pas �tre en retard
                                'unTM2 ne doit pas �tre corrig� si il est < unDebVertM
                                unNbCycle = 0
                            End If
                            If unTM2 < unDebVertM - 0.001 Then
                                'D�but de vert > T de onde montante
                                '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                                unDebVertM = unDebVertM + unNbCycle * .maDur�eDeCycle
                            ElseIf unTM2 > unDebVertM + unCarfRed.DonnerDureeVertSens(True) + 0.001 Then
                                'Fin de vert < T de d�part onde montante
                                '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                                unDebVertM = unDebVertM + unNbCycle * .maDur�eDeCycle
                            End If
                                                      
                           'Calcul du temps de parcours inter-carrefours
                            unTmpInterCarf = unCarf.monDecVitSensM - monTabCarfY(unIndCarfMPred).monCarfReduit.monCarrefour.monDecVitSensM
                            'Calcul de la fin de vert de ce carrefour r�duit
                            unFinVertM = unDebVertM + unCarfRed.DonnerDureeVertSens(True)
                            unFinVertMPred = unDebVertMPred + monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerDureeVertSens(True)

                            unI = 0
                            unJ = 0
                            unDebOndeD�j�Stock� = False
                            unLastDebVertD�j�Stock� = False
                            'projection du d�but et fin de vert du carrefour
                            'pr�c�dent sur la droite Y = Y du carrefour courant
                            unDebVertMPred = unDebVertMPred + unTmpInterCarf
                            unFinVertMPred = unFinVertMPred + unTmpInterCarf
                            'Initialisation du LastDebVert au cas o� aucun
                            'bande inter-carf trouv�e
                            unLastDebVertM = unDebVertM
                            Do
                                'La boucle sert pour afficher toutes les bandes
                                'inter-carrefour pour cela on regarde dans
                                'le cycle en cours et le suivant et on prend la
                                'bande inter-carf maximale
                                
                                'On prend le minimun des fins de vert projet�
                                'sur la droite Y = Y du carrefour courant
                                If (unFinVertMPred + unJ * .maDur�eDeCycle) < (unFinVertM + unI * .maDur�eDeCycle) Then
                                    unMinFinVert = unFinVertMPred + unJ * .maDur�eDeCycle
                                Else
                                    unMinFinVert = unFinVertM + unI * .maDur�eDeCycle
                                End If
                                
                                'On prend le maximun des d�buts de vert projet�
                                'sur la droite Y = Y du carrefour courant
                                If (unDebVertMPred + unJ * .maDur�eDeCycle) > (unDebVertM + unI * .maDur�eDeCycle) Then
                                    unMaxDebVert = unDebVertMPred + unJ * .maDur�eDeCycle
                                Else
                                    unMaxDebVert = unDebVertM + unI * .maDur�eDeCycle
                                End If
                                
                                'Test de l'existence d'une bande inter-carrefour
                                'sup�rieure � 1 seconde
                                uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                                
                                If uneBandeInterCarfExist Then
                                    If unTM2 - 0.01 < unMinFinVert And unLastDebVertD�j�Stock� = False Then
                                        'Stockage du dernier debvert ayant une bande
                                        'inter-carrefour, ce stockage est fait une fois et
                                        'une seule entre deux carrefours
                                        unLastDebVertD�j�Stock� = True
                                        unLastDebVertM = unDebVertM + unI * .maDur�eDeCycle
                                    End If
                                    'Stockage dans d�but onde pour trouver les
                                    'plages s�lectionnables graphiquement,
                                    'la 1�re fois seulement
                                    If unDebOndeD�j�Stock� = False Then
                                        unDebOndeD�j�Stock� = True
                                        unCarfRed.AffecterDebOndeSens (unMaxDebVert + unMinFinVert) / 2, True
                                    End If
                                    'Remise dans l'englobant total si
                                    'MinFinVert en sort
                                    If unMinFinVert > unMaxT + 0.01 Then
                                        unMinFinVert = unMinFinVert - .maDur�eDeCycle
                                        unMaxDebVert = unMaxDebVert - .maDur�eDeCycle
                                    End If
                                End If
                                
                                unI = unI + 1
                                If unI = 2 Then
                                    'On se place pour essayer les d�but et fin de vert
                                    'du carrefour courant dans le cycle courant avec les
                                    'd�but et fin de vert du carrefour pr�c�dent dans le
                                    'cycle suivant
                                    unI = 0
                                    unJ = 1
                                End If

                                'Dessin de bande inter-carrefour
                                If unTCM Is Nothing And uneBandeInterCarfExist Then
                                    'Cas d'une onde montante non cadr�e par un TC ou du
                                    'dessin des bandes inter-carrefours voitures d'une onde TC
                                    
                                    'Conversion en coordonn�es �cran
                                    unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 1�re partie bande montante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                    
                                    'Conversion en coordonn�es �cran
                                    unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 2�me partie bande montante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                ElseIf uneBandeInterCarfExist Then
                                    'Cas d'une onde montante cadr�e par un TC
                                    'Recup du Y du carf r�duit pr�c�dent
                                    unYTmp! = monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True)
                                    'Calcul de la date en Y = unYTmp%
                                    unDecT0! = unTCM.CalculerDateDansTabMarche(unTCM.mesPhasesTMOnde, unYTmp!, unIndPM%, unIndPM%)
                                    'Calcul du d�calage en date pour avoir la prog partielle du TC
                                    'qui commence � la date absolue unMaxDebVert - unTmpInterCarf
                                    unDecT! = unMaxDebVert - unTmpInterCarf - unDecT0!
                                    'Dessin 1�re ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True), unCarf.monCarfRed.DonnerYSens(True), unDecT!, unIndPM%
                                    'Le dernier param�tre unDecT! sert � se cadrer au d�part de l'onde sinon le TC d�bute � son T d�part
                                    
                                    'Calcul du d�calage en date pour avoir la prog partielle du TC
                                    'qui commence � la date absolue unMinFinVert - unTmpInterCarf
                                    unDecT! = unMinFinVert - unTmpInterCarf - unDecT0!
                                    'Dessin 2�me ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfMPred).monCarfReduit.DonnerYSens(True), unCarf.monCarfRed.DonnerYSens(True), unDecT!, unIndPM%
                                    'Le dernier param�tre unDecT! sert � se cadrer au d�part de l'onde sinon le TC d�bute � son T d�part
                                End If
                                'Boucle fait trois fois pour trouver toutes
                                'les bande inter-carrefour
                            Loop Until unI = 1 And unJ = 1
                        End If 'Fin du dessin de bande verte montante inter-carrefours
                        
                        'Dessin de la bande verte montante commune aux carrefours
                        'si choix coch� dans les options d'affichage et d'impression
                        'et si l'onde montante n'est pas cadr�e par un TC
                        '(dessin fait avant ce code dans ce cas, cf for i= 1 to unNbCarf)
                        If Not unNoDessinOndeM And .mesOptionsAffImp.maVisuBandComM And unTCM Is Nothing Then
                            'Stockage dans d�but onde pour trouver les plages
                            's�lectionnables graphiquement
                            unCarfRed.AffecterDebOndeSens unTM2, True
                            
                            'Conversion en coordonn�es �cran de unTM2
                            unX = ConvertirSingleEnEcran(unTM2, unT, uneLg)
                            unX = unX + unNewX0
                            
                            'Dessin en coordonn�es �cran de la ligne entre les
                            'points (unTM1, unYM1) et (unTM2, unYM2) et d'une
                            'ligne // � une largeur de bande montante
                            uneZoneDessin.Line (unXMpred, unYMpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComM
                            uneZoneDessin.Line (unXMpred + uneLBM, unYMpred)-(unX + uneLBM, unY), .mesOptionsAffImp.maCoulBandComM
                        
                            'Dessin de l'onde verte dans les feux du premier
                            'carrefour dans le sens montant, sinon le dessin
                            'ne commence qu'au feu de Y maximun (cf r�duction carrefour)
                            'Ce dessin va de unYMin jusqu'� Max Y 1er carrefour montant
                            If i = unIndCarfBasM Then
                                'Conversion en coordonn�es �cran du 1er point
                                unYFeu = DonnerYMinCarfSens(unCarf, True, unIndFeu)
                                unXMpred = unTM1 - (unYM1 - unYFeu) / unCarfRed.DonnerVitSens(True)
                                unXMpred = ConvertirSingleEnEcran(unXMpred, unT, uneLg)
                                unXMpred = unXMpred + unNewX0
                                unYMpred = ConvertirReelEnEcran(CLng(unYFeu) - unYMin, unDY, uneHt)
                                unYMpred = unY0 - unYMpred
                                uneZoneDessin.Line (unXMpred, unYMpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComM
                                uneZoneDessin.Line (unXMpred + uneLBM, unYMpred)-(unX + uneLBM, unY), .mesOptionsAffImp.maCoulBandComM
                            End If
                            
                            'Stockage du X �cran  du point pr�c�dent pour le coup suivant
                            unXMpred = unX
                        End If 'Fin du dessin de bande verte montante commune
                                 
                       'Stockage du d�but de vert pr�c�dent
                        If i = unIndCarfBasM Then
                            'Calcul sp�cial pour le carrefour le + bas montant
                            unDebVertMPred = unCarfRedM1.monCarrefour.monDecModif + unCarfRedM1.DonnerPosRefSens(True)
                            unDebVertMPred = ModuloZeroCycle(unDebVertMPred, .maDur�eDeCycle)
                            If unTM1 < unDebVertMPred - 0.001 Then
                                'D�but de vert > T de d�part onde montante
                                '==> Recul d'un cycle
                                unDebVertMPred = unDebVertMPred - .maDur�eDeCycle
                            ElseIf unTM1 > unDebVertMPred + unCarfRedM1.DonnerDureeVertSens(True) + 0.001 Then
                                'Fin de vert < T de d�part onde montante
                                '==> Avanc� d'un cycle
                                unDebVertMPred = unDebVertMPred + .maDur�eDeCycle
                            End If
                        Else
                            unDebVertMPred = unLastDebVertM
                        End If
                        
                        'Stockage de l'indice de ce carrefour
                        unIndCarfMPred = i
                        'Stockage du Y �cran  du point pr�c�dent pour le coup suivant
                        unYMpred = unY
                    End If
                End If
                
                'Parcours des carrefours dans le sens des Y d�croissants
                'pour l'onde verte descendante car on dessine � partir du
                'carrefour le plus haut ayant un feu descendant
                Set unCarfRed = monTabCarfY(unNbCarf + 1 - i).monCarfReduit
                Set unCarf = unCarfRed.monCarrefour
                'Conversion en valeur �cran de la largeur de bande descendante
                uneLBD = ConvertirSingleEnEcran(.maBandeModifD, unT, uneLg)
                
                If unIndCarfD > 0 Then
                    'Cas d'une onde verte descendante possible ==> Dessin
                    'unIndCarfD > 0 dit qu'on a trouv� des carrefours descendants
                    If unCarfRed.HasFeuDescendant = True Then
                        'Cas d'un carrefour contraignant de l'onde verte descendante
                        'donc ayant un feu de sens descendant
                        
                        'Abscisse du point suivant de l'onde descendante vaut
                        'l'abscisse du point pr�c�dent plus le d�calage en
                        'temps entre le carrefour courant et le premier descendant
                        unTD2 = unTD1 - unCarf.monDecVitSensD + unCarfRedD1.monCarrefour.monDecVitSensD
                        'Signe - et + inverse par rapport au sens montant car les
                        'd�calages en temps sont > 0 m�me en sens descendant
                        
                        'Stockage dans d�but onde pour trouver les plages
                        's�lectionnables graphiquement
                        unCarfRed.AffecterDebOndeSens unTD2, False
                        
                        'Ordonn�e �gale � l'ordonn�e du carrefour r�duit courant
                        'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                        unYD2 = unCarfRed.DonnerYSens(False)
                        
                        'Conversion en coordonn�es �cran de unYD2
                        unY = ConvertirReelEnEcran(unYD2 - unYMin, unDY, uneHt)
                        unY = unY0 - unY
                        
                        'Dessin de l'onde verte descendante inter-carrefours donc
                        'entre ce carrefour r�duit et son pr�c�dent en Y
                        'si choix coch� dans les options d'affichage et d'impression
                        'et si ce n'est pas le 1er carrefour descendante le + haut
                        
                        'Fait avant la bande commune pour ne voir que la bande
                        'commune si superposition avaec la bande inter-carrefours
                        If .mesOptionsAffImp.maVisuBandInterCarfD And (unNbCarf + 1 - i) <> unIndCarfD Then
                            'Calcul du d�but de vert de ce carrefour r�duit
                            unDebVertD = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(False)
                            unDebVertD = ModuloZeroCycle(unDebVertD, .maDur�eDeCycle)
                            'Calcul du nombre de cycle s�parant le d�but de
                            'vert du d�but de l'onde verte descendante
                            unNbCycle = Fix((0.001 + unTD2 - unDebVertD) / .maDur�eDeCycle)
                            If unNbCycle < 0 And .maBandeModifD = 0 Then
                                'Si pas de bande commune, on ne peut pas �tre en retard
                                'unTD2 ne doit pas �tre corrig� si il est < unDebVertD
                                unNbCycle = 0
                            End If
                            If unTD2 < unDebVertD - 0.001 Then
                                'D�but de vert > T de onde descendante
                                '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                                unDebVertD = unDebVertD + unNbCycle * .maDur�eDeCycle
                            ElseIf unTD2 > unDebVertD + unCarfRed.DonnerDureeVertSens(False) + 0.001 Then
                                'Fin de vert < T de d�part onde descendante
                                '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                                unDebVertD = unDebVertD + unNbCycle * .maDur�eDeCycle
                            End If
                                                      
                           'Calcul du temps de parcours inter-carrefours
                           'permutation des arguments de la soustraction par rapport
                           'au sens montant car les d�calages compt�s > 0 par rapport
                           'au carrefour descendant le plus bas
                            unTmpInterCarf = monTabCarfY(unIndCarfDPred).monCarfReduit.monCarrefour.monDecVitSensD - unCarf.monDecVitSensD
                            'Calcul de la fin de vert de ce carrefour r�duit
                            unFinVertD = unDebVertD + unCarfRed.DonnerDureeVertSens(False)
                            unFinVertDPred = unDebVertDPred + monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerDureeVertSens(False)

                            unI = 0
                            unJ = 0
                            unDebOndeD�j�Stock� = False
                            unLastDebVertD�j�Stock� = False
                            'projection du d�but et fin de vert du carrefour
                            'pr�c�dent sur la droite Y = Y du carrefour courant
                            unDebVertDPred = unDebVertDPred + unTmpInterCarf
                            unFinVertDPred = unFinVertDPred + unTmpInterCarf
                            'Initialisation du LastDebVert au cas o� aucun
                            'bande inter-carf trouv�e
                            unLastDebVertD = unDebVertD
                            Do
                                'La boucle sert pour afficher toutes les bandes
                                'inter-carrefour pour cela on regarde dans
                                'le cycle en cours et le suivant et on prend la
                                'bande inter-carf maximale
                                
                                'On prend le minimun des fins de vert projet�
                                'sur la droite Y = Y du carrefour courant
                                If (unFinVertDPred + unJ * .maDur�eDeCycle) < (unFinVertD + unI * .maDur�eDeCycle) Then
                                    unMinFinVert = unFinVertDPred + unJ * .maDur�eDeCycle
                                Else
                                    unMinFinVert = unFinVertD + unI * .maDur�eDeCycle
                                End If
                                
                                'On prend le maximun des d�buts de vert projet�
                                'sur la droite Y = Y du carrefour courant
                                If (unDebVertDPred + unJ * .maDur�eDeCycle) > (unDebVertD + unI * .maDur�eDeCycle) Then
                                    unMaxDebVert = unDebVertDPred + unJ * .maDur�eDeCycle
                                Else
                                    unMaxDebVert = unDebVertD + unI * .maDur�eDeCycle
                                End If
                                
                                'Test de l'existence d'une bande inter-carrefour
                                'sup�rieure � 1 seconde
                                uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                                
                                If uneBandeInterCarfExist Then
                                    If unTD2 - 0.01 < unMinFinVert And unLastDebVertD�j�Stock� = False Then
                                        'Stockage du dernier debvert ayant une bande
                                        'inter-carrefour, ce stochage est fait une fois
                                        'et une seule entre deux carrefours
                                        unLastDebVertD�j�Stock� = True
                                        unLastDebVertD = unDebVertD + unI * .maDur�eDeCycle
                                    End If
                                    'Stockage dans d�but onde pour trouver les
                                    'plages s�lectionnables graphiquement,
                                    'la 1�re fois seulement
                                    If unDebOndeD�j�Stock� = False Then
                                        unDebOndeD�j�Stock� = True
                                        unCarfRed.AffecterDebOndeSens (unMaxDebVert + unMinFinVert) / 2, False
                                    End If
                                    'Remise dans l'englobant total si
                                    'MinFinVert en sort
                                    If unMinFinVert > unMaxT + 0.01 Then
                                        unMinFinVert = unMinFinVert - .maDur�eDeCycle
                                        unMaxDebVert = unMaxDebVert - .maDur�eDeCycle
                                    End If
                                End If
                                
                                unI = unI + 1
                                If unI = 2 Then
                                    'On se place pour essayer les d�but et fin de vert
                                    'du carrefour courant dans le cycle courant avec les
                                    'd�but et fin de vert du carrefour pr�c�dent dans le
                                    'cycle suivant
                                    unI = 0
                                    unJ = 1
                                End If
                                                                                    
                                'Dessin de bande inter-carrefour
                                If unTCD Is Nothing And uneBandeInterCarfExist Then
                                    'Cas d'une onde descendante non cadr�e par un TC
                                    
                                    'Conversion en coordonn�es �cran
                                    unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 1�re partie bande descendante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                    
                                    'Conversion en coordonn�es �cran
                                    unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                    unX1 = unX1 + unNewX0
                                    unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                    unX2 = unX2 + unNewX0
                                    'Dessin 2�me partie bande descendante inter-carrefours
                                    uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                ElseIf uneBandeInterCarfExist Then
                                    'Cas d'une onde descendante cadr�e par un TC
                                    'Recup du Y du carf r�duit pr�c�dent
                                    unYTmp! = monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False)
                                    'Calcul de la date en Y = -unYTmp% car inversion des Y en sens descendant
                                    unDecT0! = unTCD.CalculerDateDansTabMarche(unTCD.mesPhasesTMOnde, -unYTmp!, unIndPD%, unIndPD%)
                                    'Calcul du d�calage en date pour avoir la prog partielle du TC
                                    'qui commence � la date absolue unMaxDebVert - unTmpInterCarf
                                    unDecT! = unMaxDebVert - unTmpInterCarf - unDecT0!
                                    'Dessin 1�re ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False), unCarf.monCarfRed.DonnerYSens(False), unDecT!, unIndPD%
                                    'Le dernier param�tre unDecT! sert � se cadrer au d�part de l'onde sinon le TC d�bute � son T d�part
                                    
                                    'Calcul du d�calage en date pour avoir la prog partielle du TC
                                    'qui commence � la date absolue unMinFinVert - unTmpInterCarf
                                    unDecT! = unMinFinVert - unTmpInterCarf - unDecT0!
                                    'Dessin 2�me ligne de la bande montante inter-carf
                                    TracerProgPartielleTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, monTabCarfY(unIndCarfDPred).monCarfReduit.DonnerYSens(False), unCarf.monCarfRed.DonnerYSens(False), unDecT!, unIndPD%
                                    'Le dernier param�tre unDecT! sert � se cadrer au d�part de l'onde sinon le TC d�bute � son T d�part
                                End If
                                'Boucle fait trois fois pour trouver toutes
                                'les bande inter-carrefour
                            Loop Until unI = 1 And unJ = 1
                                
                        End If 'Fin du dessin de bande verte descendante inter-carrefours
                        
                        'Dessin de la bande verte descendante commune aux carrefours
                        'si choix coch� dans les options d'affichage et d'impression
                        If Not unNoDessinOndeD And .mesOptionsAffImp.maVisuBandComD And unTCD Is Nothing Then
                            'Stockage dans d�but onde pour trouver les plages
                            's�lectionnables graphiquement
                            unCarfRed.AffecterDebOndeSens unTD2, False
                        
                            'Conversion en coordonn�es �cran de unTD2
                            unX = ConvertirSingleEnEcran(unTD2, unT, uneLg)
                            unX = unX + unNewX0
                            
                            'Dessin en coordonn�es �cran de la ligne entre les
                            'points (unTD1, unYD1) et (unTD2, unYD2) et d'une
                            'ligne // � une largeur de bande descendante
                            uneZoneDessin.Line (unXDpred, unYDpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComD
                            uneZoneDessin.Line (unXDpred + uneLBD, unYDpred)-(unX + uneLBD, unY), .mesOptionsAffImp.maCoulBandComD
                        
                            'Dessin de l'onde verte dans les feux du premier
                            'carrefour dans le sens descendant, sinon le dessin
                            'ne commence qu'au feu de Y minimun (cf r�duction carrefour)
                            'Ce dessin va de unYMax jusqu'� Min Y 1er carrefour descendant
                            If unIndCarfD = (unNbCarf + 1 - i) Then
                                'Conversion en coordonn�es �cran du 1er point
                                unYFeu = DonnerYMaxCarfSens(unCarf, False, unIndFeu)
                                unXDpred = unTD1 - (unYD1 - unYFeu) / unCarfRed.DonnerVitSens(False)
                                unXDpred = ConvertirSingleEnEcran(unXDpred, unT, uneLg)
                                unXDpred = unXDpred + unNewX0
                                unYDpred = ConvertirReelEnEcran(CLng(unYFeu) - unYMin, unDY, uneHt)
                                unYDpred = unY0 - unYDpred
                                uneZoneDessin.Line (unXDpred, unYDpred)-(unX, unY), .mesOptionsAffImp.maCoulBandComD
                                uneZoneDessin.Line (unXDpred + uneLBD, unYDpred)-(unX + uneLBD, unY), .mesOptionsAffImp.maCoulBandComD
                            End If
                            'Stockage du X �cran du point pr�c�dent pour le coup suivant
                            unXDpred = unX
                        End If 'Fin du dessin de la bande verte descendante commune
                                                                                 
                        'Stockage du d�but de vert pr�c�dent
                        If unNbCarf + 1 - i = unIndCarfD Then
                            'Calcul sp�cial pour le carrefour le + haut descendant
                            unDebVertDPred = unCarfRedD1.monCarrefour.monDecModif + unCarfRedD1.DonnerPosRefSens(False)
                            unDebVertDPred = ModuloZeroCycle(unDebVertDPred, .maDur�eDeCycle)
                            If unTD1 < unDebVertDPred - 0.001 Then
                                'D�but de vert > T de d�part onde descendante
                                '==> Recul d'un cycle
                                unDebVertDPred = unDebVertDPred - .maDur�eDeCycle
                            ElseIf unTD1 > unDebVertDPred + unCarfRedD1.DonnerDureeVertSens(False) + 0.001 Then
                                'Fin de vert < T de d�part onde descendante
                                '==> Avanc� d'un cycle
                                unDebVertDPred = unDebVertDPred + .maDur�eDeCycle
                            End If
                        Else
                            unDebVertDPred = unLastDebVertD
                        End If
                        
                        'Stockage de l'indice de ce carrefour
                        unIndCarfDPred = unNbCarf + 1 - i
                        'Stockage du Y �cran  du point pr�c�dent pour le coup suivant
                        unYDpred = unY
                    End If
                End If
            Next i
            
            'Dessin des ondes vertes cadr�es par des TC, bandes communes
            'on le fait apr�s les bandes inter-carrefours pour les bandes
            'communes ne soient pas �cras�es par les bandes inter-carrefour
            
            'Dessin de l'onde verte commune montante dans le cas
            'd'un cadrage par TC si le dessin est possible et si
            'choisi dans les options d'affichage
            If Not (unTCM Is Nothing) And Not unNoDessinOndeM And unIndCarfM > 0 And monSite.mesOptionsAffImp.maVisuBandComM Then
                'Dessin de la 1�re ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, DessinOndeTCM, unTM1
                'Dessin de la 2�me ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCM, unX0, unY0, uneLg, uneHt, DessinOndeTCM, unTM1 + .maBandeModifM
            End If
            
            'Dessin de l'onde verte commune descendante dans le cas
            'd'un cadrage par TC si le dessin est possible et si
            'choisi dans les options d'affichage
            If Not (unTCD Is Nothing) And Not unNoDessinOndeD And unIndCarfD > 0 And monSite.mesOptionsAffImp.maVisuBandComD Then
                'Dessin de la 1�re ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, DessinOndeTCD, unTD1
                'Dessin de la 2�me ligne de la bande passante montante
                TracerProgressionTC uneZoneDessin, unTCD, unX0, unY0, uneLg, uneHt, DessinOndeTCD, unTD1 + .maBandeModifD
            End If
        Else
            'R�duction globale impossible
            If Not TypeOf uneZoneDessin Is Printer Then uneZoneDessin.Cls
            'Initialisation pour �viter division par 0
            unT = 300
            monSite.monTmpTotal = unT
            unDY = 500
            monSite.monDYTotal = unDY
            'Idem pour �viter plantage dans PleinEcran
            monSite.monYMaxFeuUtil = 1
            monSite.monYMinFeuUtil = -1
            Exit Sub
        End If
                
        'Dessin des traits de rappels et des cycles en �paisseur 1
        'dans le cas d'une impression uniquement.
        'Raison : bug VB5 en impression les pointill�s d'�paisseur > 1
        'font des traits continus sur certaines imprimantes jet d'encre
        'couleurs.
        'Restauration de l'�paisseur de trait pour imprimer plus bas
        If TypeOf uneZoneDessin Is Printer Then
            Printer.DrawWidth = 1
            
            'Si demand� dessin des lignes de rappels toutes les n secondes pour
            'une impression, n choisi dans la fen�tre d'impression (n entre 1 et
            '10) d'o� :
            'unN = monSite.mesOptionsAffImp.monNbSecondesRappel en impression
            'Dessin de ses sous-divisions en trait pointill� si unN ne vaut pas
            '10 car les traits des dizaines sont faits apr�s
            uneZoneDessin.DrawStyle = vbDash
            unN = monSite.mesOptionsAffImp.monNbSecondesRappel
            If .mesOptionsAffImp.maVisuLigne And unN <> 10 Then
                For unI = unMinT To unMaxT Step unN
                    If unI Mod .maDur�eDeCycle <> 0 Then
                        unX = ConvertirReelEnEcran(unI, unT, uneLg)
                        uneZoneDessin.Line (unNewX0 + unX, unY0)-(unNewX0 + unX, unY0 - uneHt), .mesOptionsAffImp.maCoulLigne
                    End If
                Next unI
            End If
        End If
        
        'Dessin des lignes de rappels toutes les 10 secondes en trait
        'plein si demand�
        uneZoneDessin.DrawStyle = vbSolid
        If .mesOptionsAffImp.maVisuLigne Then
            For unI = unMinT To unMaxT Step 10
                If unI Mod .maDur�eDeCycle <> 0 Then
                    unX = ConvertirReelEnEcran(unI, unT, uneLg)
                    uneZoneDessin.Line (unNewX0 + unX, unY0)-(unNewX0 + unX, unY0 - uneHt), .mesOptionsAffImp.maCoulLigne
                End If
            Next unI
        End If
                
        'Dessin des Traits de s�paration de cycle en tiret-point (trait mixte)
        uneZoneDessin.DrawStyle = vbDashDot
        unNbCycle = Int(unT / .maDur�eDeCycle)
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
        'la dur�e du cycle dans la fen�tre plein �cran
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
                'Affichage de la dur�e du cycle sur imprimante
                ImprimerDureeCycle unX0, unY0, unNewX0 + i * unDT
            ElseIf Not unDessinDansOnglet Then
                'Affichage de la dur�e du cycle en fen�tre Pleine �cran
                uneZoneDessin.AfficherDureeCycle (unNewX0 + i * unDT)
            End If
        Next i
        'Remise du dessin de ligne pleine
        uneZoneDessin.DrawStyle = vbSolid
        
        'Restauration de l'�paisseur de trait si on �tait en impression
        If TypeOf uneZoneDessin Is Printer Then
            Printer.DrawWidth = monSite.mesOptionsAffImp.monEpaisseurLigne
        End If
        
        'Calcul des limites de la zone de dessin sur �cran ou sur imprimante
        If TypeOf uneZoneDessin Is Printer Then
            unDebZone = unX0
            unFinZone = unX0 + uneLg
        Else
            unDebZone = 0
            unFinZone = uneZoneDessin.Width
        End If
        
        'Calcul d'une pr�cision montante et descendante pour les tests < et >
        '0n prend la dur�e du cycle convertie en largeur �cran comme
        'pr�cision
        unePrecM = ConvertirSingleEnEcran(monSite.maDur�eDeCycle / 2, unT, uneLg)
        unePrecD = unePrecM
        
        'Dessin des plages de vert montant et descendantes, et des points
        'de r�f�rence de tous feux de tous les carrefours
        For i = 1 To .mesCarrefours.Count
            Set unCarf = .mesCarrefours(i)
            If unCarf.monDecCalcul <> -99 Then
                unMin = 10000 ' Les ordonn�es sont <= 9999 dans OndeV
                'r�cup�ration du point milieu de l'onde Mont ou Desc
                'converti en valeur �cran
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
                    'Calcul du d�but de vert ramen� entre 0 et +cycle
                    'pour dessiner les plages de vert si FinVert (= DebVert
                    'modulo cycle + la dur�e de vert) - le d�calage est >= au
                    'cycle (donc le point de r�f�rence et cette plage de vert
                    'sont dans le m�me cycle) ou ramen� entre -cycle et 0 sinon
                    'm�me si les position de r�f�rence sont tr�s grandes
                    '+ PosRef et pas - car les position point r�f�rence
                    'sont entr�es avec un moins en interne dans OndeV
                    unDebVert = unCarf.monDecModif + unFeu.maPositionPointRef
                    unDebVertMod = ModuloZeroCycle(unDebVert, .maDur�eDeCycle)
                    If unDebVertMod + unFeu.maDur�eDeVert - unCarf.monDecModif > .maDur�eDeCycle - 0.01 Then
                        unDebVert = unDebVertMod - .maDur�eDeCycle
                    Else
                        unDebVert = unDebVertMod
                    End If
                    'Si la borne inf de l'englobant du graphic est < 0
                    'on enl�ve un cycle pour dessiner la partie T n�gative
                    If unMinT < 0 Then unDebVert = unDebVert - .maDur�eDeCycle
                    'Conversion de la plage de vert en coordonn�es �cran
                    unTmpDebVert = ConvertirSingleEnEcran(unDebVert, unT, uneLg)
                    unTmpFinVert = unTmpDebVert + ConvertirSingleEnEcran(CSng(unFeu.maDur�eDeVert), unT, uneLg)
                    'Conversion de l'ordonn�e en coordonn�es �cran
                    unY = ConvertirReelEnEcran(unFeu.monOrdonn�e - unYMin, unDY, uneHt)
                    
                    'Choix de la couleur du trait montant ou descendant
                    If unFeu.monSensMontant Then
                        uneCouleur = .mesOptionsAffImp.maCoulBandComM
                    Else
                        uneCouleur = .mesOptionsAffImp.maCoulBandComD
                    End If
                    
                    'Dessin de la plage de vert pour tous les cycles
                    For K = 0 To unNbCycle + 1
                        'Calcul des X �crans
                        unDebVert = unTmpDebVert + K * uneLongCycle
                        unFinVert = unTmpFinVert + K * uneLongCycle
                        unX = unDebVert + unNewX0
                        unXf = unFinVert + unNewX0
                        
                        'Stockage des lignes symbolisant la plage de
                        'vert contenant l'onde verte montante et descendante,
                        'cette ligne sera s�lectionnable interactivement
                        'Pr�cision = unePrecM ou unePrecD twips suivant le sens
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
                    
                    'Recherche du feu ayant l'ordonn�e minimale du carrefour
                    If unFeu.monOrdonn�e < unMin Then
                        unMin = unFeu.monOrdonn�e
                    End If
                                    
                    If j = unCarf.mesFeux.Count Then
                        'Dessin du point de r�f�rence unique du carrefour au Y
                        'du feu le plus bas en ordonn�e
                        'Conversion de l'ordonn�e en coordonn�es �cran
                        unY = ConvertirReelEnEcran(unMin - unYMin, unDY, uneHt)
                        'Conversion du point de r�f�rence en coordonn�es �cran
                        unTmpPtRef = ConvertirSingleEnEcran(unCarf.monDecModif, unT, uneLg)
                        'Si la borne inf de l'englobant du graphic est < 0 on enl�ve
                        'un cycle en valeur �cran pour dessiner la partie T < 0
                        If unMinT < 0 Then unTmpPtRef = unTmpPtRef - uneLongCycle
                        'Dessin du point de r�f�rence (triangle de hauteur �cran
                        'valant 120 twips) pour tous les cycles
                        uneH = 120
                        'Indication pour l'impression du d�calage
                        unDejaImprimer = False
                        For K = 0 To unNbCycle + 1
                            unTPtRef = unTmpPtRef + K * uneLongCycle + unNewX0
                            'On ne dessine dans les limites de la zone de
                            'dessin
                            If unDebZone < unTPtRef And unTPtRef < unFinZone Then
                                uneZoneDessin.Line (unTPtRef, unY0 - unY)-(unTPtRef - uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                uneZoneDessin.Line (unTPtRef, unY0 - unY)-(unTPtRef + uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                uneZoneDessin.Line (unTPtRef - uneH, unY0 - unY + uneH)-(unTPtRef + uneH, unY0 - unY + uneH), .mesOptionsAffImp.maCoulPtRef
                                'Affichage d'un cercle � l'int�rieur du triangle
                                'pour les carrefours � d�calages impos�s
                                If unCarf.monDecImp = 1 Then
                                    uneZoneDessin.FillColor = .mesOptionsAffImp.maCoulPtRef
                                    uneZoneDessin.FillStyle = vbFSSolid
                                    uneZoneDessin.Circle (unTPtRef, unY0 - unY + uneH * 2 / 3), uneH / 3, .mesOptionsAffImp.maCoulPtRef
                                End If
                                
                                'Affichage du d�calage une seule fois en impression
                                If TypeOf uneZoneDessin Is Printer And unDejaImprimer = False Then
                                    unDejaImprimer = True 'impression une seule fois
                                    Printer.ForeColor = .mesOptionsAffImp.maCoulPtRef
                                    Printer.CurrentX = unTPtRef + 1.1 * uneH
                                    Printer.CurrentY = unY0 - unY + 0.1 * uneH
                                    Printer.Print Format(CIntCorrig�(unCarf.monDecModif))
                                End If
                            End If
                            
                            'Stockage du point symbolisant la valeur du d�calage
                            'de l'onde verte montante et descendante, ce point
                            'sera s�lectionnable interactivement
                            'Pr�cision = unePrecM ou unePrecD twips suivant le sens
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
    'Modification du param�tre unMaxT pour l'arrondi � la dizaine sup�rieure
    unReste = unMaxT - 10 * Fix(unMaxT / 10)
    If unMaxT > 0 Then
        unMaxT = unMaxT + 10 - unReste 'Ici unReste > 0
    Else
        unMaxT = unMaxT - unReste 'Ici unReste < 0
    End If
End Sub

Public Sub ModifierMinTempsPourVisu(unMinT)
    'Modification du param�tre unMinT pour l'arrondi � la dizaine inf�rieure
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
        unFinVert = unCarf.monDecModif + unCarf.mesFeux(i).maPositionPointRef + unCarf.mesFeux(i).maDur�eDeVert
        If unFinVert > TrouverMaxFinVert Then TrouverMaxFinVert = unFinVert
    Next i
End Function

Public Function TrouverMinDebVert(unCarf As Carrefour) As Single
    'Retourne la date de d�but de vert minimun parmi tous
    'les feux du carrefour unCarf
    'On ne la ram�ne pas modulo cycle car le d�but de vert minimun
    'd'o� commence l'onde verte peut �tre n�gatif
    TrouverMinDebVert = 1000
    For i = 1 To unCarf.mesFeux.Count
        unDebVert = unCarf.monDecModif + unCarf.mesFeux(i).maPositionPointRef
        If unDebVert < TrouverMinDebVert Then TrouverMinDebVert = unDebVert
    Next i
End Function


Public Sub TrouverMinYMaxY(unYMin As Long, unYMax As Long)
    'Recherche du max et du min en Y des feux des carrefours
    'utilis�s dans le calcul de l'onde verte
    'Les variables unYMin et unYMax sont modifi�s par cette
    'proc�dure
    
    Dim unCarf As Carrefour
    Dim unFeu As Feu
    
    unYMin = 10000
    unYMax = -10000 'Les Y dans OndeV sont entre -9999 et 9999
    
    unNbCarf = monSite.mesCarrefours.Count
    For i = 1 To unNbCarf
        Set unCarf = monSite.mesCarrefours(i)
        If unCarf.monDecCalcul <> -99 Then
            'Cas d'un carrefour utilis� dans le calcul de l'onde
            For j = 1 To unCarf.mesFeux.Count
                Set unFeu = unCarf.mesFeux(j)
                If unFeu.monOrdonn�e < unYMin Then
                    unYMin = unFeu.monOrdonn�e
                End If
                If unFeu.monOrdonn�e > unYMax Then
                    unYMax = unFeu.monOrdonn�e
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
    
    'Information � l'utilisateur si les carrefours de d�part et d'arriv�e
    'n'ont pas de feu dans le sens du TC.
    'C'est juste une information les calculs continueront mais l'affichage
    'des graphiques feront apparaitre ces incoh�rences.
    unMsg = ""
    unTC.monCarfDep.DonnerNbFeuxMetD unNbFeuxMdep, unNbFeuxDdep
    unTC.monCarfArr.DonnerNbFeuxMetD unNbFeuxMarr, unNbFeuxDarr
    If DonnerYCarrefour(unTC.monCarfDep) <= DonnerYCarrefour(unTC.monCarfArr) Then
        'Cas d'un TC montant
        If unNbFeuxMdep = 0 Then
            unMsg = unMsg + "Le TC : " + unTC.monNom + " est de sens montant mais son carrefour de d�part " + unTC.monCarfDep.monNom + " n'a aucun feu dans ce sens."
        End If
        'If unTC.monCarfArr.monCarfRed.HasFeuMontant = False Then
        If unNbFeuxMarr = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens montant mais son carrefour d'arriv�e " + unTC.monCarfArr.monNom + " n'a aucun feu dans ce sens."
        End If
    Else
        'Cas d'un TC descendant
        If unNbFeuxDdep = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens descendant mais son carrefour de d�part " + unTC.monCarfDep.monNom + " n'a aucun feu dans ce sens."
        End If
        'If unTC.monCarfArr.monCarfRed.HasFeuDescendant = False Then
        If unNbFeuxDarr = 0 Then
            unMsg = unMsg + Chr(13) + "Le TC : " + unTC.monNom + " est de sens descendant mais son carrefour d'arriv�e " + unTC.monCarfArr.monNom + " n'a aucun feu dans ce sens."
        End If
    End If
    If unMsg <> "" Then MsgBox unMsg, vbInformation, "Information OndeV pour correction"
    
    'D�termination du type de dessin TC � r�aliser
    'et donc de sa couleur d'affichage et de la liste des phases,
    'donc du tableau de marche � dessiner (progression ou onde, cf classe TC)
    If unTypeDessin = DessinProgTC Then
        'Dessin du tableau de marche de progression du TC
        uneCouleur = unTC.maCouleur
        Set uneColPhases = unTC.mesPhasesTMProg
        'D�calage en temps � rajouter
        unDecT = 0
    ElseIf unTypeDessin = DessinOndeTCM Then
        'Dessin du tableau de marche du TC cadrant l'onde montante
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandComM
        Set uneColPhases = unTC.mesPhasesTMOnde
        'D�calage en temps entre le Y (= Y du feu montant le plus haut)
        'du feu �quivalent du carrefour r�duit et
        'du feu montant le plus bas du carrefour de d�part
        uneDateYext = unTC.CalculerDateDansTabMarche(uneColPhases, unTC.monCarfDep.monCarfRed.DonnerYSens(True), unIndPhase, 1)
        uneDateYdep = unTC.CalculerDateDansTabMarche(uneColPhases, DonnerYMinCarfSens(unTC.monCarfDep, True, unIndFeu), unIndPhase, 1)
        unDecT = unDecIniT + uneDateYext - uneDateYdep
        'D�calage en temps � rajouter pour commencer au d�but de l'onde montante
        unDecT = unDecT - unTC.mesPhasesTMOnde(1).monTDeb
    ElseIf unTypeDessin = DessinOndeTCD Then
        'Dessin du tableau de marche du TC cadrant l'onde descendante
        uneCouleur = monSite.mesOptionsAffImp.maCoulBandComD
        Set uneColPhases = unTC.mesPhasesTMOnde
        'D�calage en temps entre le Y (= Y du feu descendant le plus bas)
        'du feu �quivalent du carrefour r�duit et
        'du feu descendant le plus haut du carrefour de d�part
        'en inversant le signe des Y pour le sens descendant
        uneDateYext = unTC.CalculerDateDansTabMarche(uneColPhases, -unTC.monCarfDep.monCarfRed.DonnerYSens(False), unIndPhase, 1)
        uneDateYdep = unTC.CalculerDateDansTabMarche(uneColPhases, -DonnerYMaxCarfSens(unTC.monCarfDep, False, unIndFeu), unIndPhase, 1)
        unDecT = unDecIniT - uneDateYext + uneDateYdep
        'D�calage en temps � rajouter pour commencer au d�but de l'onde descendante
        unDecT = unDecT - unTC.mesPhasesTMOnde(1).monTDeb
    Else
        MsgBox "ERREUR de programmation dans OndeV dans TracerProgressionTC", vbCritical
    End If
    
    'D�termination du sens du TC car les Y des phases sont invers�s
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
        
        'Conversion �cran du point d�but de phase en inversant
        'les signes des Y d�but de phase si TC descendant (unSens = -1)
        unTE1 = ConvertirSingleEnEcran(unePhase.monTDeb + unDecT, monSite.monTmpTotal, uneLg)
        unTE1 = unTE1 + monSite.monOrigX
        
        unYE1 = ConvertirSingleEnEcran(unePhase.monYDeb * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
        unYE1 = unY0 - unYE1
        
        If unePhase.monType = Accel Or unePhase.monType = Decel Then
            'Cas d'une phase d'acc�l�ration ou de d�c�l�ration, on dessine
            'la parabole gr�ce unNbPoints + 2 segments
            
            'Calcul du d�calage en Temps entre chaque point de la parabole
            unDT = unePhase.maDureePhase / unNbPoints
            
            'Dessin des autres points de la phase d'acc ou de d�c�l�ration
            For j = 1 To unNbPoints - 1
                'Calcul d'un point courant de la parabole
                unT = unePhase.monTDeb + j * unDT
                unY = CalculerYDansPhaseParabole(unePhase, unT)
                
                'Conversion �cran du point courant de la parabole
                unTE2 = ConvertirSingleEnEcran(unT + unDecT, monSite.monTmpTotal, uneLg)
                unTE2 = unTE2 + monSite.monOrigX
                
                unYE2 = ConvertirSingleEnEcran(unY * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
                unYE2 = unY0 - unYE2
                
                'Dessin de la ligne entre le point courant et le pr�c�dent
                uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
                
                'Stockage du point �cran courant pour l'incr�mentation suivante
                unTE1 = unTE2
                unYE1 = unYE2
            Next j
        End If
        
        'Conversion �cran du point fin de phase en inversant
        'les signes des Y d�but de phase si TC descendant (unSens = -1)
        unTE2 = ConvertirSingleEnEcran(unePhase.monTDeb + unDecT + unePhase.maDureePhase, monSite.monTmpTotal, uneLg)
        unTE2 = unTE2 + monSite.monOrigX
        
        unYE2 = ConvertirSingleEnEcran((unePhase.monYDeb + unePhase.maLongPhase) * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
        unYE2 = unY0 - unYE2
        
        'Dessin de la ligne entre le point pr�c�dent et le point fin de phase
        uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
    Next i
End Sub

Public Sub TracerProgPartielleTC(uneZoneDessin As Object, unTC As TC, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, unYDeb As Integer, unYFin As Integer, unDecT As Single, unIndPhaseDep As Integer)
    'Dessin de la progression partielle du TC unTC de la liste des TC du site
    'entre les ordonn�es unYDeb et unYFin
    'Utiliser pour dessiner les bandes inter-carrefours d'ondes cadr�es par
    'un TC montant et/ou un TC descendant
    
    Dim unePhase As PhaseTabMarche
    Dim unSens As Integer, i As Integer
    Dim unY As Single, unDT As Single
    Dim unY1 As Single, unY2 As Single
    Dim unTDeb As Single, unTFin As Single
    Dim unT As Single, unIndPhase As Integer
    Dim uneCouleur As Long
    Dim uneColPhases As ColPhaseTM
        
    'D�termination du sens du TC car les Y des phases sont invers�s
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
    uneFirstPhase = True 'Signale si on se trouve dans la 1�re phase
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
                'Cas de la 1�re phase de la prog partielle
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
                'Cas o� la phase courante sort de la prog partielle
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
            
            'Conversion �cran du point d�but de phase en inversant
            'les signes des Y d�but de phase si TC descendant (unSens = -1)
            unTE1 = ConvertirSingleEnEcran(unTDeb + unDecT, monSite.monTmpTotal, uneLg)
            unTE1 = unTE1 + monSite.monOrigX
            
            unYE1 = ConvertirSingleEnEcran(unY1 * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
            unYE1 = unY0 - unYE1
            
            If unePhase.monType = Accel Or unePhase.monType = Decel Then
                'Cas d'une phase d'acc�l�ration ou de d�c�l�ration, on dessine
                'la parabole gr�ce unNbPoints + 2 segments
                
                'Calcul du d�calage en Temps entre chaque point de la parabole
                unDT = (unTFin - unTDeb) / unNbPoints
                
                'Dessin des autres points de la phase d'acc ou de d�c�l�ration
                For j = 1 To unNbPoints - 1
                    'Calcul d'un point courant de la parabole
                    unT = unTDeb + j * unDT
                    unY = CalculerYDansPhaseParabole(unePhase, unT)
                    
                    If unY > unYFin * unSens + 0.001 Then Exit For
                    
                    'Conversion �cran du point courant de la parabole
                    unTE2 = ConvertirSingleEnEcran(unT + unDecT, monSite.monTmpTotal, uneLg)
                    unTE2 = unTE2 + monSite.monOrigX
                    
                    unYE2 = ConvertirSingleEnEcran(unY * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
                    unYE2 = unY0 - unYE2
                    
                    'Dessin de la ligne entre le point courant et le pr�c�dent
                    uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
                    
                    'Stockage du point �cran courant pour l'incr�mentation suivante
                    unTE1 = unTE2
                    unYE1 = unYE2
                Next j
            End If
            
            'Conversion �cran du point fin de phase en inversant
            'les signes des Y si TC descendant (unSens = -1)
            unTE2 = ConvertirSingleEnEcran(unTFin + unDecT, monSite.monTmpTotal, uneLg)
            unTE2 = unTE2 + monSite.monOrigX
            
            unYE2 = ConvertirSingleEnEcran(unY2 * unSens - monSite.monYMin, monSite.monDYTotal, uneHt)
            unYE2 = unY0 - unYE2
            
            'Dessin de la ligne entre le point pr�c�dent et le point fin de phase
            uneZoneDessin.Line (unTE1, unYE1)-(unTE2, unYE2), uneCouleur
        End If
    Loop Until uneSortieBoucle
End Sub

Public Sub DonnerEnglobantTC(unTC As TC, unX0 As Long, unY0 As Long, uneLg As Long, uneHt As Long, unTDep As Long, unTFin As Long)
    'alimentation des dates de d�but et de fin du parcours unTDep et unTFin pour
    'calculer l'englobant �cran en temps , donc le niveau de zoom du graphique onde
    
    Dim unePhase As PhaseTabMarche
    Dim unSens As Integer
    
    'Calcul du tableau de marche de progression du TC
    'avec r�cup�ration de son sens (1 montant, -1 descendant)
    unSens = unTC.CalculerTableauMarcheProg()
        
    'Calcul du d�but de l'englobant �cran en temps par conversion en
    'coordonn�es �cran du point de d�part du TC, donc du T d�but de la
    'premi�re phase
    Set unePhase = unTC.mesPhasesTMProg(1)
    unTDep = ConvertirSingleEnEcran(unePhase.monTDeb, monSite.monTmpTotal, uneLg)
    unTDep = unTDep + monSite.monOrigX
    
    'Calcul de la fin de l'englobant �cran en temps par conversion en
    'coordonn�es �cran du point de fin du TC, donc du T d�but de la
    'derni�re phase plus sa dur�e
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
    'Ainsi un pick souris n'est pas chang� par un texte print�
    'contrairement � un control label qui trappe les clicks souris
    'Mais ce texte print� doit �tre rafraichit comme un dessin de ligne
    'd'o� ce code en fin de dessin total
    If unDessinDansOnglet Then
        uneZoneDessin.CurrentX = monSite.LabelFleche.Left - uneZoneDessin.TextWidth("t en secondes")
        uneZoneDessin.CurrentY = monSite.AxeTemps.Y1 - monSite.LabelFleche.Height / 2 - uneZoneDessin.TextHeight("t en secondes")
        uneForeColor = uneZoneDessin.ForeColor 'Stockage pour restaurer apr�s
        uneZoneDessin.ForeColor = 0 'Mise en noir
        uneZoneDessin.Print "t en secondes"
        uneZoneDessin.ForeColor = uneForeColor 'Restauration couleur initiale
    End If
End Sub

Public Function EstTCUtil(unIndTC) As Boolean
    'Fonction retournant :
    '   - vrai si le TC d'index unIndTC fait partie des TC utilis�s, dont
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
    'Proc�dure de s�lection graphique d'une plage de vert, d'une poign�e
    'ou d'un point de r�f�rence
    'Lanc� par l'event MouseDown sur le bouton gauche de la souris en
    'un point (unXpick, unYpick) en coordonn�es �cran
    'd'une Zone de dessin (PictureBox d'une fenetre site = frmDocument)
    'ou par la Form frmPleinEcran
    Dim unYreel As Long, unY As Long, unDTEcran As Long
    Dim uneLongEcranAxeY As Long
    Dim unePl As PlageGraphic, unRef As RefGraphic
    Dim unObjTrouv As Boolean
    Dim unY0 As Single, unYecran As Single
    Dim unItemAnnuler As Control
    
    'Stockage du X �cran correspond � cette s�lection
    monXEcranDebModif = unXpick
    monXEcranFinModif = unXpick 'Il sera modifi� dans un MouseMove,
                                'donc dans ModifierSelection
    
    'Initialisation pour la conversion de valeurs r�elles en �crans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la s�lection dans la fen�tre plein �cran
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
        'Calcul de la longueur �cran de l'axe des ordonn�es
        uneLongEcranAxeY = uneZoneDessin.AxeY.Y2 - uneZoneDessin.AxeY.Y1
        'Origine des Y
        unY0 = uneZoneDessin.FrameCarfTC.Top + uneZoneDessin.AxeY.Y2
        'Stockage de l'item Annuler derni�re modif graphique
        Set unItemAnnuler = uneZoneDessin.mnuAnnulerModif
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la s�lection dans l'onglet Graphique Onde Verte
        'Calcul de la longueur �cran de l'axe des ordonn�es
        uneLongEcranAxeY = monSite.AxeOrdonn�e.Y2 - monSite.AxeOrdonn�e.Y1
        'Origine des Y
        unEspacement = 120 'm�me valeur que dans AffichageOngletVisu
        unY0 = monSite.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au m�me
        'niveau que le min des Y
            
        'Affectation de uneZoneDessin � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeTemps.X2 - uneZoneDessin.AxeTemps.X1
        'Stockage de l'item Annuler derni�re modif graphique
        Set unItemAnnuler = frmMain.mnuGraphicOndeAnnul
    Else
        MsgBox "Erreur de programmation dans OndeV dans SelectionGraphique", vbCritical
    End If
    
    'Recherche de l'objet graphique cliqu� (point de r�f�rence, plage de vert
    'ou poign�e de s�lection)
    '==> Arr�t au premier trouv�
    'Le reste du traitement, la modification interactive intervient dans
    'l'event MouseMove avec bouton gauche enfonc� de la zone de dessin
    'qui appelera la fonction ModifierSelection et l'apparition des poign�es
    'apparait dans le MouseUp ainsi que le recalcul et redessin
    'qui appelera la fonction MettreAJourSelection
    
    'Initialisation de la s�lection � vide
    unObjTrouv = False
    monTypeObjPick = NoSel
    uneZoneDessin.PlageVert(0).Visible = False
    
    'Recherche parmi les poign�es de s�lection (gauche et droite)
    If uneZoneDessin.PoigneeGauche.Visible Then
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        unXmilieu = uneZoneDessin.PoigneeDroite.Left + uneZoneDessin.PoigneeDroite.Width / 2
        unYmilieu = uneZoneDessin.PoigneeDroite.Top + uneZoneDessin.PoigneeDroite.Height / 2
        If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
            unObjTrouv = True
            monTypeObjPick = PgDSel
            'Cas de la s�lection interactive d'un fin de vert
            unMsg = "S�lection du fin de vert : Carrefour " + unCarf.monNom
            unMsg = unMsg + " et Feu = " + Format(monObjPick.monIndFeu)
        End If
        
        If Not unObjTrouv Then
            unXmilieu = uneZoneDessin.PoigneeGauche.Left + uneZoneDessin.PoigneeGauche.Width / 2
            unYmilieu = uneZoneDessin.PoigneeGauche.Top + uneZoneDessin.PoigneeGauche.Height / 2
            If Abs(unXpick - unXmilieu) <= PrecPick And Abs(unYpick - unYmilieu) <= PrecPick Then
                unObjTrouv = True
                monTypeObjPick = PgGSel
                'Cas de la s�lection interactive d'un d�but de vert
                unMsg = "S�lection du d�but de vert : Carrefour " + unCarf.monNom
                unMsg = unMsg + " et Feu = " + Format(monObjPick.monIndFeu)
            End If
        End If
    End If
    
    'Masquage des poign�es si aucune n'a �t� s�lectionn�e
    If Not unObjTrouv Then
        uneZoneDessin.PoigneeDroite.Visible = False
        uneZoneDessin.PoigneeGauche.Visible = False
    End If
    
    'Hauteur en twips de triangle symbolisant le point de r�f�rence
    'fix� dans la fonction DessinerOndeVerte
    uneH = 120
    
    'Recherche parmi les points de r�f�rence de l'onde verte descendante
    i = 1
    unNbTotal = monSite.maColRefGraphicD.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unRef = monSite.maColRefGraphicD(i)
        'Info dans la barre d'�tat indiquant la s�lection d'un point de r�f�rence
        unMsg = "S�lection du point de r�f�rence : Carrefour " + monSite.mesCarrefours(unRef.monIndCarf).monNom
        'R�cup�ration du Y minimun des feux du carrefour
        unYreel = DonnerYMinCarf(monSite.mesCarrefours(unRef.monIndCarf))
        'Conversion du Yr�el en Y �cran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est pr�s du point de r�f�rence
        'suivant la pr�cision PrecPick
        If unRef.monDecal - uneH - PrecPick < unXpick And unXpick < unRef.monDecal + uneH + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + uneH + PrecPick Then
            'Stockage de l'objet trouv� par pick �cran
            unObjTrouv = True
            monTypeObjPick = RefSel
            Set monObjPick = unRef
            'Apparition d'une image d'un triangle d�pla�able
            'interativement au m�me emplacement
            uneZoneDessin.PtRef.Left = unRef.monDecal - uneZoneDessin.PtRef.Width / 2
            uneZoneDessin.PtRef.Top = unY0 - unYecran
            uneZoneDessin.PtRef.Visible = True
            
            'Cr�ation des plages de vert
            'qui se d�placeront avec le triangle point de r�f�rence
            PlacerPlagesVert uneZoneDessin, monObjPick, unY0, uneLongEcranAxeY, unDTEcran
        End If
        i = i + 1
    Loop
    
    'Recherche parmi les points de r�f�rence de l'onde verte montante
    i = 1
    unNbTotal = monSite.maColRefGraphicM.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unRef = monSite.maColRefGraphicM(i)
        'Info dans la barre d'�tat indiquant la s�lection d'un point de r�f�rence
        unMsg = "S�lection du point de r�f�rence : Carrefour " + monSite.mesCarrefours(unRef.monIndCarf).monNom
        'R�cup�ration du Y minimun des feux du carrefour
        unYreel = DonnerYMinCarf(monSite.mesCarrefours(unRef.monIndCarf))
        'Conversion du Yr�el en Y �cran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est pr�s du point de r�f�rence
        'suivant la pr�cision PrecPick
        If unRef.monDecal - uneH - PrecPick < unXpick And unXpick < unRef.monDecal + uneH + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + uneH + PrecPick Then
            'Stockage de l'objet trouv� par pick �cran
            unObjTrouv = True
            monTypeObjPick = RefSel
            Set monObjPick = unRef
            'Apparition d'une image d'un triangle d�pla�able
            'interativement au m�me emplacement
            uneZoneDessin.PtRef.Left = unRef.monDecal - uneZoneDessin.PtRef.Width / 2
            uneZoneDessin.PtRef.Top = unY0 - unYecran
            uneZoneDessin.PtRef.Visible = True
            
            'Cr�ation des plages de vert
            'qui se d�placeront avec le triangle point de r�f�rence
            PlacerPlagesVert uneZoneDessin, monObjPick, unY0, uneLongEcranAxeY, unDTEcran
        End If
        i = i + 1
    Loop
    
    'Recherche parmi les plages de vert de l'onde verte descendante
    i = 1
    unNbTotal = monSite.maColPlageGraphicD.Count
    Do While i <= unNbTotal And unObjTrouv = False
        Set unePl = monSite.maColPlageGraphicD(i)
        'Info dans la barre d'�tat indiquant la s�lection d'un d�but de vert
        unMsg = "S�lection de la plage de vert : Carrefour " + monSite.mesCarrefours(unePl.monIndCarf).monNom
        unMsg = unMsg + " et Feu = " + Format(unePl.monIndFeu)
        'R�cup�ration du Y du feu
        unYreel = monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monOrdonn�e
        'Conversion du Yr�el en Y �cran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est dans la plage de vert
        'suivant la pr�cision PrecPick
        If unePl.monDebVert - PrecPick < unXpick And unXpick < unePl.monFinVert + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + PrecPick Then
            'Stockage de l'objet trouv� par pick �cran
            unObjTrouv = True
            monTypeObjPick = PlaSel
            Set monObjPick = unePl
            'Apparition d'une ligne modifiable interativement au m�me emplacement
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
        'Info dans la barre d'�tat indiquant la s�lection d'un fin de vert
        unMsg = "S�lection de la plage de vert : Carrefour " + monSite.mesCarrefours(unePl.monIndCarf).monNom
        unMsg = unMsg + " et Feu = " + Format(unePl.monIndFeu)
        'R�cup�ration du Y du feu
        unYreel = monSite.mesCarrefours(unePl.monIndCarf).mesFeux(unePl.monIndFeu).monOrdonn�e
        'Conversion du Yr�el en Y �cran
        unYecran = ConvertirReelEnEcran(unYreel - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        'Test si le point du pick est dans la plage de vert
        'suivant la pr�cision PrecPick
        If unePl.monDebVert - PrecPick < unXpick And unXpick < unePl.monFinVert + PrecPick And unY0 - unYecran - PrecPick < unYpick And unYpick < unY0 - unYecran + PrecPick Then
            'Stockage de l'objet trouv� par pick �cran
            unObjTrouv = True
            monTypeObjPick = PlaSel
            Set monObjPick = unePl
            'Apparition d'une ligne modifiable interativement au m�me emplacement
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
    
    'Affichage dans la 1�re zone de la barre d'�tat
    'du r�sultat du pick
    If monTypeObjPick = NoSel Then
        'Cas o� la s�lection graphique est vide
        unMsg = "Rien de s�lectionner"
        'Mise en gris�e de l'annulation de la dern�re modif graphique
        unItemAnnuler.Enabled = False
    End If
    frmMain.sbStatusBar.Panels.Item(1).Text = unMsg
End Sub


Public Sub MettreAJourSelection(uneZoneDessin As Object, unXpick As Single)
    'Affichage des poign�es si on a s�lectionn� une plage de vert
    'ou Recalcul et redessin de l'onde verte si on a modifie un d�but ou
    'un fin de vert, un point de r�f�rence ou un d�calage de carrefour
    Dim unFeu As Feu, uneForm As Object
    Dim unCarf As Carrefour, uneModifDec As Boolean
    Dim unX0 As Long, unY0 As Long, unDTEcran As Long
    Dim uneHt As Long, unDessinOnglet As Boolean
    Dim unOldDecal As Single, unEcartReel As Single
    
    'Initialisation pour la conversion de valeurs r�elles en �crans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la s�lection dans la fen�tre plein �cran
        
        'Affectation � la form m�re pour acc�der aux controls
        Set uneForm = uneZoneDessin
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneForm.AxeT.X2 - uneForm.AxeT.X1
        'Calcul du cadre o� l'on dessine
        unDessinOnglet = False
        unEspacement = 120 'm�me valeur que dans AffichageOngletVisu
        unX0 = uneForm.AxeT.X1
        unY0 = uneForm.FrameCarfTC.Top + uneForm.AxeY.Y2
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au m�me
        'niveau que le min des Y
        uneHt = uneForm.AxeY.Y2 - uneForm.AxeY.Y1
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la s�lection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneForm � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        Set uneForm = monSite
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneForm.AxeTemps.X2 - uneForm.AxeTemps.X1
        'Calcul du cadre o� l'on dessine
        unDessinOnglet = True
        unEspacement = 120 'm�me valeur que dans AffichageOngletVisu
        unX0 = uneForm.AxeTemps.X1
        unY0 = uneForm.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au m�me
        'niveau que le min des Y
        uneHt = uneForm.AxeOrdonn�e.Y2 - uneForm.AxeOrdonn�e.Y1
    End If
            
    'Masquage de l'info bulle de modif graphique
    uneForm.InfoModif.Visible = False
    
    'On vide les valeurs de la pr�c�dente modif graphique
    ViderCollection maColValPred
    
    'Sauvegarde de l'englobant en temps pour r�utilisation dans la
    'fonction AnnulerLastModifGraphic qui annule la derni�re modif graphique
    monTmpTotalAvantModif = monSite.monTmpTotal
        
    'Masquage de la plage de vert montrant la modification interactive
    uneForm.PlageVert(0).Visible = False
    
    If monTypeObjPick = PgGSel Then
        'Cas de la modification interactive d'un d�but de vert
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Sauvegarde de l'ancienne position de r�f�rence et l'ancienne dur�e
        'de vert dans la collection des valeurs pr�c�dentes pour un Annuler
        maColValPred.Add unFeu.maPositionPointRef
        maColValPred.Add unFeu.maDur�eDeVert
        'Conversion de l'�cart de la modification de la plage de
        'vert d'une valeur �cran en valeur r�elle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Changement du point de r�f�rence du feu s�lectionn� arrondi en entier
        'si unEcartReel > 0 ==> Diminution du d�but de vert, donc de la dur�e
        'de vert et de la position de r�f�rence mais qui est < 0 en interne
        'donc il faut l'augmenter. Si unEcartReel < 0 ==> On fait l'inverse
        unFeu.maPositionPointRef = CInt(unFeu.maPositionPointRef + unEcartReel)
        unFeu.maDur�eDeVert = CInt(unFeu.maDur�eDeVert - unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = PgDSel Then
        'Cas de la modification interactive d'un fin de vert
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Conversion de l'�cart de la modification de la plage de
        'vert d'une valeur �cran en valeur r�elle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Sauvegarde de l'ancienne dur�e de vert dans la
        'collection des valeurs pr�c�dentes pour un Annuler
        maColValPred.Add unFeu.maDur�eDeVert
        'Changement de la dur�e de vert du feu s�lectionn� arrondie en entier
        'si unEcartReel > 0 ==> Augmentation de la dur�e de vert
        'Si unEcartReel < 0 ==> On fait l'inverse
        unFeu.maDur�eDeVert = CInt(unFeu.maDur�eDeVert + unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = PlaSel Then
        'Cas de la s�lection interactive d'une plage de vert
        '==> Apparition des poign�es droite et gauche s�lectionnables
        'aux extr�mit�s de la plage qui a �t� s�lectionn�e
        uneForm.PoigneeGauche.Left = uneForm.PlageVert(0).X1 - uneForm.PoigneeGauche.Width / 2
        uneForm.PoigneeGauche.Top = uneForm.PlageVert(0).Y1 - uneForm.PoigneeGauche.Height / 2
        uneForm.PoigneeDroite.Left = uneForm.PlageVert(0).X2 - uneForm.PoigneeDroite.Width / 2
        uneForm.PoigneeDroite.Top = uneForm.PlageVert(0).Y2 - uneForm.PoigneeDroite.Height / 2
        uneForm.PoigneeDroite.Visible = True
        uneForm.PoigneeGauche.Visible = True
        'Masquage de la plage de vert montrant la s�lection
        uneForm.PlageVert(0).Visible = False
        'D�placement du point de r�f�rence du feu s�lectionn�
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        Set unFeu = unCarf.mesFeux(monObjPick.monIndFeu)
        'Conversion de l'�cart de la modification de la plage de
        'vert d'une valeur �cran en valeur r�elle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Sauvegarde de l'ancienne position de r�f�rence dans la
        'collection des valeurs pr�c�dentes pour un Annuler
        maColValPred.Add unFeu.maPositionPointRef
        'Translation du point de r�f�rence du feu s�lectionn� arrondi en entier
        unFeu.maPositionPointRef = CInt(unFeu.maPositionPointRef + unEcartReel)
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = RefSel Then
        'Cas de la modification interactive d'un point de r�f�rence
        '==> Recalcul et redessin de l'onde verte
        'Masquage et destruction des plages de vert d�pa�able
        'et du triangle
        uneForm.PtRef.Visible = False
        For i = uneForm.PlageVert.Count - 1 To 1 Step -1
            uneForm.PlageVert(i).Visible = False
            Unload uneForm.PlageVert(i)
        Next i
        'Modification du d�calage du carrefour s�lectionn�
        Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
        'Conversion de l'�cart de la translation du point de
        'r�f�rence d'une valeur �cran en valeur r�elle
        unEcartReel = (monXEcranFinModif - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        'Stockage du d�calage avant modification
        unOldDecal = unCarf.monDecModif
        'Sauvegarde de l'ancien d�calage dans la collection
        'des valeurs pr�c�dentes pour un Annuler
        maColValPred.Add unOldDecal
        'Modification du d�calage du carrefour s�lectionn� modulo cycle
        'On en prend la partie enti�re pour avoir la m�me chose si on avait
        'saisi ce nouveau d�calage dans l'onglet Tableau R�sultat
        unCarf.monDecModif = CIntCorrig�(ModuloZeroCycle(unCarf.monDecModif + unEcartReel, monSite.maDur�eDeCycle))
        'Modif du d�calage modifiable du carrefour choisi
        'On ajoute la diff�rence entre la vrai valeur en r�elle et l'arrondi
        'en entier pour l'affichage dans l'onglet Tableau R�sultat pour ne
        'pas perdre en pr�cision de calcul
        'Exemple : si le calcul trouve un d�calage de 29.8 que l'on stocke
        'on affiche par contre 30, si l'utilisateur remet 30 il peut avoir
        'un r�sultat diff�rent car le 30 qu'il voit, vaut en fait 29.8.
        'En ajoutant la diff�rence du � l'arrondi on retrouve la m�me chose
        unCarf.monDecModif = unOldDecal - CIntCorrig�(unOldDecal) + unCarf.monDecModif
        'Indication de modification pour un recalcul
        monSite.maModifDataDes = True
    ElseIf monTypeObjPick = NoSel Then
        'Cas o� la s�lection graphique est vide
        '==> On en fait rien
    Else
        MsgBox "Erreur de programmation dans OndeV dans MettreAJourSelection", vbCritical
    End If
        
    '==> Recalcul et redessin de l'onde verte
    'Test si l'�l�ment pick� a �t� boug�
    If monXEcranDebModif = monXEcranFinModif Then
        'Aucun mouvement de souris ==> pas de modif
        monSite.maModifDataDes = False
    End If
    
    If monSite.maModifDataDes Then
        'Initialisation des bool�ens permettant de savoir
        'si les calculs ont r�ussi
        unCalculOndeFait = False
        unRecalculBandeFait = False
        
        'Masquage des poign�es de s�lection
        uneForm.PoigneeDroite.Visible = False
        uneForm.PoigneeGauche.Visible = False
        
        If monTypeObjPick = RefSel And unCarf.monDecImp = 0 Then
            'Cas d'une modification de d�calage d'un carrefour � d�calage
            'non impos� ==> Recalcul de bandes passantes
            unRecalculBandeFait = RecalculerBandesPassantes(monSite)
            If unRecalculBandeFait Then
                'Cas o� le recalcul a �t� possible
                'Stockage d'une modification de valeurs dans les d�calages
                'Ceci permettra aussi de demander une sauvegarde � la fermeture
                maModifDataDec = True
                'Indication de fin de modif graphique, ainsi on ne refera
                'pas le calcul de l'onde verte en cas de changement d'onglet
                monSite.maModifDataDes = False
            Else
                'Cas o� le recalcul a �t� impossible
                'On remet la valeur du d�calage avant modif
                unCarf.monDecModif = unOldDecal
            End If
        Else
            'Cas d'une modification autre qu'un d�calage d'un carrefour �
            'd�calage non impos�, pour ce cas il faut recalculer l'onde aussi
            '==> Calcul de l'onde � refaire pour mise � jour
            unCalculOndeFait = True
            If monTypeObjPick = RefSel And unCarf.monDecImp = 1 Then
                'Cas d'une modif graphique du d�calage
                'd'un carrefour � d�calage impos�
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
            'Activation du menu pour annuler la derni�re modif
            If TypeOf uneZoneDessin Is Form Then
                'Cas de modif interactive en plein �cran
                uneForm.mnuAnnulerModif.Enabled = True
            Else
                'Cas de modif interactive dans l'onglet Graphique onde verte
                frmMain.mnuGraphicOndeAnnul.Enabled = True
            End If
        Else
            'Restauration des valeurs pr�c�dentes
            MsgBox "D�termination d'onde verte impossible ==> Restauration des valeurs pr�c�dentes", vbInformation
            AnnulerLastModifGraphic uneZoneDessin
        End If
    End If
End Sub

Public Sub ChangerPointeurSouris(uneZoneDessin As Object, unXpick As Single, unYpick As Single)
    'Changement du pointeur souris en croix si on passe
    'sur les poign�es de s�lection si elles sont visibles
    Dim unObjTrouv As Boolean
    
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la s�lection dans la fen�tre plein �cran
        Set uneForm = uneZoneDessin
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la s�lection dans l'onglet Graphique Onde Verte
            
        'Affectation de uneZoneDessin � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneForm = monSite
    End If
    
    unObjTrouv = False
    'Test de la visibilit� des poign�es de s�lection
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
    'si on est sur une poign�e
    If unObjTrouv Then
        uneZoneDessin.MousePointer = vbCrosshair
    Else
        uneZoneDessin.MousePointer = vbDefault
    End If
End Sub

Public Sub ModifierSelection(uneZoneDessin As Object, unXpick As Single, unYpick As Single)
    'Modification interactive de l'objet s�lectionn� dans le mouseDown
    'de uneZoneDessin
    Dim uneSecEcran As Long, unDTEcran As Long
    Dim unCycleEcran As Long, unDX As Single, unDX2 As Single
    Dim unCarf As Carrefour, unFeu As Feu
    Dim unXDebVert As Single, unXRef As Single
    
    'Initialisation pour la conversion de valeurs r�elles en �crans
    If TypeOf uneZoneDessin Is Form Then
        'Cas de la s�lection dans la fen�tre plein �cran
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la s�lection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneZoneDessin � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
        'Calcul de la longueur �cran de l'axe des temps
        unDTEcran = uneZoneDessin.AxeTemps.X2 - uneZoneDessin.AxeTemps.X1
    End If
        
    'Conversion d'une seconde et d'une dur�e de cycle r�elles en valeur �cran
    uneSecEcran = ConvertirReelEnEcran(1, monSite.monTmpTotal, unDTEcran)
    unCycleEcran = ConvertirReelEnEcran(monSite.maDur�eDeCycle, monSite.monTmpTotal, unDTEcran)
    
    If monTypeObjPick = PgGSel Then
        'Cas de la modification interactive d'un d�but de vert
        '==> D�placement horizontale de l'extr�mit� d�but de vert de la plage
        uneZoneDessin.PlageVert(0).Visible = True
        'Calcul de la nouvelle dur�e de vert par Conversion de l'�cart de
        'la modification d'une valeur �cran en valeur r�elle
        unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
        Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
        If unFeu.maDur�eDeVert - unEcartReel >= 1 And unFeu.maDur�eDeVert - unEcartReel <= monSite.maDur�eDeCycle - 1 Then
            'La dur�e de vert doit �tre >= � 1 et <= dur�e de cycle - 1
            uneZoneDessin.PlageVert(0).X1 = unXpick
            uneZoneDessin.PoigneeGauche.Left = unXpick - uneZoneDessin.PoigneeGauche.Width / 2
            'Stockage du X �cran de fin de modification
            monXEcranFinModif = unXpick
            'Affichage dans l'info bulle de modif de la nouvelle dur�e de vert
            'et de la nouvelle position de r�f�rence
            unePosRef = CInt(unFeu.maPositionPointRef + unEcartReel)
            uneInfoPosRef = "R�f�rence = " + Format(-unePosRef)
            uneDV = CInt(unFeu.maDur�eDeVert - unEcartReel)
            AfficherInfoModif uneZoneDessin, uneInfoPosRef + " Dur�e de vert = " + Format(uneDV), unXpick, unYpick
        End If
    ElseIf monTypeObjPick = PgDSel Then
        'Cas de la modification interactive d'un fin de vert
        '==> D�placement horizontale de l'extr�mit� fin de vert de la plage
        uneZoneDessin.PlageVert(0).Visible = True
        If unXpick >= uneZoneDessin.PlageVert(0).X1 + uneSecEcran And unXpick <= uneZoneDessin.PlageVert(0).X1 + unCycleEcran - uneSecEcran Then
            'D�placement doit �tre > au d�but de vert et la plage < cycle
            uneZoneDessin.PlageVert(0).X2 = unXpick
            uneZoneDessin.PoigneeDroite.Left = unXpick - uneZoneDessin.PoigneeGauche.Width / 2
            'Calcul de la nouvelle dur�e de vert par Conversion de l'�cart de
            'la modification d'une valeur �cran en valeur r�elle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
            'Stockage du X �cran de fin de modification
            monXEcranFinModif = unXpick
            'Affichage dans l'info bulle de modif de la nouvelle dur�e de vert
            uneDV = CInt(unFeu.maDur�eDeVert + unEcartReel)
            AfficherInfoModif uneZoneDessin, "Dur�e de vert = " + Format(uneDV), unXpick, unYpick
        End If
    ElseIf monTypeObjPick = PlaSel Then
        'Cas de la s�lection interactive d'une plage de vert
        '==> D�placement horizontale de la plage de vert
        uneZoneDessin.PlageVert(0).Visible = True
        'Calcul du d�placement par rapport au d�but de la s�lection
        unDX = unXpick - monXEcranDebModif
        'D�placement de la plage de vert s�lectionn�e en le bloquant
        'horizontalement de telle fa�on que le d�but du vert de la plage
        'varie entre la fin de vert de la plage - Cycle et cette fin de vert
        '==> Toutes les modifs possibles sont dans cet intervalle
        unXDebVert = monObjPick.monDebVert + unDX
        If monObjPick.monFinVert - unCycleEcran <= unXDebVert And unXDebVert <= monObjPick.monFinVert Then
            uneZoneDessin.PlageVert(0).X1 = monObjPick.monDebVert + unDX
            uneZoneDessin.PlageVert(0).X2 = monObjPick.monFinVert + unDX
            'Calcul du nouveau point de r�f�rence par conversion
            'de l'�cart de la modification de la plage de
            'vert d'une valeur �cran en valeur r�elle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unFeu = monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu)
            'Affichage dans l'info bulle de modif du nouveau point de ref
            unPref = CInt(unFeu.maPositionPointRef + unEcartReel)
            AfficherInfoModif uneZoneDessin, "R�f�rence = " + Format(-unPref), unXpick, unYpick
            'Stockage du X �cran de fin de modification
            monXEcranFinModif = unXpick
        End If
    ElseIf monTypeObjPick = RefSel Then
        'Cas de la modification interactive d'un point de r�f�rence
        '==> D�placement horizontale de tous les feux et du point de
        'r�f�rence du carrefour, donc du triangle
                
        'Calcul du d�placement par rapport au d�but de la s�lection
        unDX = unXpick - monXEcranDebModif
        'D�placement du triangle point de r�f�rence et de tous les feux
        'du carrefours entre ce point de r�f�rence - un cycle et ce point
        'de r�f�rence + cycle
        '==> Toutes les modifs possibles sont dans cet intervalle
        unXRef = monObjPick.monDecal + unDX
        If monObjPick.monDecal - unCycleEcran <= unXRef And unXRef <= monObjPick.monDecal + unCycleEcran Then
            uneZoneDessin.PtRef.Left = unXRef - uneZoneDessin.PtRef.Width / 2
            For i = 1 To uneZoneDessin.PlageVert.Count - 1
                'D�termination de l'indice du feu
                Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
                unNbFeux = unCarf.mesFeux.Count
                If i <= unNbFeux Then
                    'Cas des plages du m�me cycle
                    unInd = i
                    Set unFeu = unCarf.mesFeux(unInd)
                    'D�placement du d�but de vert de la plage i
                    unePosRefEcran = ConvertirSingleEnEcran(unFeu.maPositionPointRef, monSite.monTmpTotal, unDTEcran)
                    unDX2 = monObjPick.monDecal + unDX + unePosRefEcran - uneZoneDessin.PlageVert(i).X1
                    'unDX2 est le d�placement relatif par rapport � la derni�re position
                    'alors qu'unDX est le d�placement par rapport au d�but de modification
                    uneZoneDessin.PlageVert(i).X1 = uneZoneDessin.PlageVert(i).X1 + unDX2
                Else
                    'Cas des plages graphiquement s�lectionnables
                    unInd = i - unNbFeux
                    Set unFeu = unCarf.mesFeux(unInd)
                    'D�placement du d�but de vert de la plage i par + unDX2
                    uneZoneDessin.PlageVert(i).X1 = uneZoneDessin.PlageVert(i).X1 + unDX2
                End If
                
                'D�placement de la fin de vert de la plage i
                uneDurVertEcran = ConvertirSingleEnEcran(unFeu.maDur�eDeVert, monSite.monTmpTotal, unDTEcran)
                uneZoneDessin.PlageVert(i).X2 = uneZoneDessin.PlageVert(i).X1 + uneDurVertEcran
                'uneZoneDessin.PlageVert(i).Visible = True
            Next i
            'Calcul du d�calage par conversion de l'�cart de la
            'translation du point de r�f�rence d'une valeur �cran
            'en valeur r�elle
            unEcartReel = (unXpick - monXEcranDebModif) * monSite.monTmpTotal / unDTEcran
            Set unCarf = monSite.mesCarrefours(monObjPick.monIndCarf)
            'Affichage dans l'info bulle de modif du nouveau point de ref
            unDec = CIntCorrig�(ModuloZeroCycle(unCarf.monDecModif + unEcartReel, monSite.maDur�eDeCycle))
            'On garde la m�me pr�cision des chiffres apr�s la virgule
            unDec = CIntCorrig�(unCarf.monDecModif - CIntCorrig�(unCarf.monDecModif) + unDec)
            AfficherInfoModif uneZoneDessin, "D�calage = " + Format(unDec), unXpick, unYpick
            'Stockage du X �cran de fin de modification
            monXEcranFinModif = unXpick
        End If
    ElseIf monTypeObjPick = NoSel Then
        'Cas o� la s�lection graphique est vide
        '==> On en fait rien
    Else
        MsgBox "Erreur de programmation dans OndeV dans ModifierSelection", vbCritical
    End If
End Sub

Public Sub PlacerPlagesVert(uneZoneDessin As Object, unObjPick As Object, unY0 As Single, uneLongEcranAxeY As Long, unDTEcran As Long)
    'Cr�ation des plages de vert
    'qui se d�placeront avec le triangle point de r�f�rence
    Dim uneColPlageGraphic As Collection
    
    'Initialisation pour la conversion de valeurs r�elles en �crans
    If TypeOf uneZoneDessin Is PictureBox Then
        'Cas de la s�lection dans l'onglet Graphique Onde Verte
        
        'Affectation de uneZoneDessin � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        'Il faut qu'uneZoneDessin soit un Form
        Set uneZoneDessin = monSite
    End If
    
    'R�cup�ration du carrefour du point de r�f�rence s�lectionn�
    Set unCarf = monSite.mesCarrefours(unObjPick.monIndCarf)
    
    'Cr�ation des plages de vert de tous les feux du carrefour
    'se trouvant dans le m�me cycle que le d�calage du carrefour
    unNbFeux = unCarf.mesFeux.Count
    For i = 1 To unNbFeux
        Set unFeu = unCarf.mesFeux(i)
        unY = ConvertirReelEnEcran(unFeu.monOrdonn�e - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        
        'Cr�ation d'une nouvelle plage de vert d�pla�able,
        'celle de m�me cycle que le d�calage du carrefour
        Load uneZoneDessin.PlageVert(i)
        
        'Positionnement en coordonn�es �cran de cette plage
        uneZoneDessin.PlageVert(i).X1 = unObjPick.monDecal + ConvertirSingleEnEcran(unFeu.maPositionPointRef, monSite.monTmpTotal, unDTEcran)
        uneZoneDessin.PlageVert(i).Y1 = unY0 - unY
        uneZoneDessin.PlageVert(i).X2 = uneZoneDessin.PlageVert(i).X1 + ConvertirSingleEnEcran(unFeu.maDur�eDeVert, monSite.monTmpTotal, unDTEcran)
        uneZoneDessin.PlageVert(i).Y2 = unY0 - unY
        
        'Affectation de la couleur onde montante ou descendante
        If unFeu.monSensMontant Then
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComM
        Else
            uneZoneDessin.PlageVert(i).BorderColor = monSite.mesOptionsAffImp.maCoulBandComD
        End If
        uneZoneDessin.PlageVert(i).Visible = True
    Next i
    
    'Cr�ation des plages de vert de tous les feux du carrefour
    'celles stock�es dans la liste de s�lection, donc celles
    's�lectionnables graphiquement
    For i = unNbFeux + 1 To 2 * unNbFeux
        Set unFeu = unCarf.mesFeux(i - unNbFeux)
        unY = ConvertirReelEnEcran(unFeu.monOrdonn�e - monSite.monYMin, monSite.monDYTotal, uneLongEcranAxeY)
        
        'Cr�ation d'une nouvelle plage de vert d�pla�able,
        'celle de m�me cycle que le d�calage du carrefour
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
            
        'Recherche de la plage graphique s�lectionnable repr�sentant ce feu
        'dans une collection contenant cette plage
        unInd = DonnerIndicePlage(uneColPlageGraphic, unObjPick.monIndCarf, i - unNbFeux)
        
        If unInd = 0 Then
            'Aucune plage trouv�e
            uneZoneDessin.PlageVert(i).Visible = False
        Else
            'Positionnement en coordonn�es �cran de cette plage
            uneZoneDessin.PlageVert(i).X1 = uneColPlageGraphic(unInd).monDebVert
            uneZoneDessin.PlageVert(i).Y1 = unY0 - unY
            uneZoneDessin.PlageVert(i).X2 = uneColPlageGraphic(unInd).monFinVert
            uneZoneDessin.PlageVert(i).Y2 = unY0 - unY
            uneZoneDessin.PlageVert(i).Visible = True
        End If
    Next i
End Sub

Public Sub AnnulerLastModifGraphic(uneZoneDessin As Object)
    'Annulation de la derni�re modification interactive dans un
    'graphique d'onde verte en r�cup�rant les anciennes valeurs de l'objet
    'graphique s�lectionn� et stock�es lors la modification dans la collection
    'maColValPred pour revenir � l'�tat pr�c�dent.
    'Recalcul de l'onde ou des d�calages pr�c�dents et redessin du graphique
        
    If monTypeObjPick = PgGSel Then
        'Restauration de la position du point de r�f�rence et de la dur�e
        'de verte pr�c�dentes du feu s�lectionn�
        'Le 1er �l�ment de maColValPred est l'ancienne position
        'et le 2 �me l'ancienne dur�e de vert
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maPositionPointRef = maColValPred(1)
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maDur�eDeVert = maColValPred(2)
    ElseIf monTypeObjPick = PgDSel Then
        'Restauration de la dur�e de verte pr�c�dente du feu
        's�lectionn�, le 1er �l�ment de maColValPred est cette dur�e
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maDur�eDeVert = maColValPred(1)
    ElseIf monTypeObjPick = PlaSel Then
        'Restauration de la position du point de r�f�rence pr�c�dente du feu
        's�lectionn�, le 1er �l�ment de maColValPred est cette position.
        monSite.mesCarrefours(monObjPick.monIndCarf).mesFeux(monObjPick.monIndFeu).maPositionPointRef = maColValPred(1)
    ElseIf monTypeObjPick = RefSel Then
        'Restauration du d�calage pr�c�dent, le 1er �l�ment de
        'maColValPred est le d�calage pr�c�dent
        monSite.mesCarrefours(monObjPick.monIndCarf).monDecModif = maColValPred(1)
        'Recalcul des bandes passantes, toujours possible car
        'il l'�tait avant la modif graphique
        unResultat = RecalculerBandesPassantes(monSite)
        'Modification de l'indicateur de changement de donn�es
        'car modification est invalid�e
        'pas de recalcul de l'onde verte en cas de changement d'onglet
        monSite.maModifDataDec = False
    Else
        If monTypeObjPick > NoSel Then
            'Cas autre que rien de s�lectionner
            MsgBox "Erreur de programmation dans OndeV dans AnnulerLastModifGraphic", vbCritical
        End If
    End If
            
    'Recalcul de l'onde verte pour l'annulation d'une modif
    'interactive autre que celle d'un d�calage
    If monTypeObjPick = PgDSel Or monTypeObjPick = PgGSel Or monTypeObjPick = PlaSel Then
        'Indication d'une modif pour recalculer l'onde verte
        monSite.maModifDataCarf = True
        'Recalcul d'onde verte sans tester sur la faisabilit�
        'du calcul car il avait �t� fait avant
        unResultat = CalculerOndeVerte(monSite)
    End If
    
    'Mise � jour du dessin d'ondes vertes et de progression TC
    MettreAJourDessin
End Sub

Public Sub MettreAJourDessin()
    'Redessin du Graphique d'Onde verte, avec les progressions des TC
    '�ventuelles dans l'onglet Graphique Onde Verte
    'ou dans la fen�tre plein �cran suivant la valeur de uneZoneDessin
    
    Dim unX0 As Long, unY0 As Long
    Dim uneHt As Long, uneLg As Long
    Dim uneZoneDessin As Object
    
    'Initialisation pour la conversion de valeurs r�elles en �crans
    If monPleinEcranVisible = False Then
        'Cas de modification dans l'onglet Graphique Onde Verte
        'Affectation de uneForm � monSite pour acc�der aux controls
        'd'interaction graphique poign�ee, etc...
        Set uneZoneDessin = monSite.ZoneDessin
        'Calcul de la longueur �cran de l'axe des temps
        uneLg = monSite.AxeTemps.X2 - monSite.AxeTemps.X1
        'Calcul du cadre o� l'on dessine
        unEspacement = 120 'm�me valeur que dans AffichageOngletVisu
        unX0 = monSite.AxeTemps.X1
        unY0 = monSite.AxeTemps.Y1 - unEspacement / 4
        'le - unEsp/4 pour avoir l'origine de l'axe des temps au m�me
        'niveau que le min des Y
        uneHt = monSite.AxeOrdonn�e.Y2 - monSite.AxeOrdonn�e.Y1
        'Redessin de l'onglet Graphique Onde verte de la
        'fen�tre active si c'est l'onglet en cours d'utilisation
        If monSite.TabFeux.Tab = 4 Then
            uneZoneDessin.Cls 'effacement
            unEspacement = 120 'm�me valeur que dans AffichageOngletVisu
            DessinerTout uneZoneDessin, unX0, unY0, uneLg, uneHt, True
        End If
    Else
        'Cas de modification dans la fen�tre plein �cran
        'Affectation � la form m�re pour acc�der aux controls
        Set uneZoneDessin = frmPleinEcran
        'Calcul de la longueur �cran de l'axe des temps
        uneLg = uneZoneDessin.AxeT.X2 - uneZoneDessin.AxeT.X1
        'Calcul du cadre o� l'on dessine
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
    'Affichage dans la zone de dessin d'une info bulle affichant le param�tre
    'modifi� interactivement et sa valeur
    uneForm.InfoModif.Caption = unMsg
    uneForm.InfoModif.Left = unX - uneForm.InfoModif.Width - 60 'en twips
    uneForm.InfoModif.Top = unY
    uneForm.InfoModif.Visible = True
End Sub

Public Sub TrouverTempsParcoursEtCarrefours(unIndCarfM, unIndCarfD, unTmpM, unTmpD)
    'Recherche du carrefour le plus haut ayant un feu montant et du
    'carrefour le plus haut ayant un feu descendant.
    'Ces carrefours donneront les temps de parcours dans les deux sens
    'Les carrefours  seront retourn�s dans les variables unIndCarfM et
    'unIndCarfD et les temps de parcours dans unTmpM et unTmpD
    'Valeurs nulles si rien n'est trouv� dans un sens
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
                'Cas d'un carrefour � sens unique montant
                If unIndCarfM = 0 Then
                    unIndCarfM = i
                End If
            Else
                'Cas d'un carrefour � sens unique descendant
                If unIndCarfD = 0 Then
                    unIndCarfD = i
                End If
            End If
        End If
    Loop While (unIndCarfM = 0 Or unIndCarfD = 0) And i > 1
    
    'R�cup�ration des temps de parcours montant et descendant
    'Tous les deux sont > 0, les temps de parcours montant et descendant
    'sont stock�s dans les d�calages dus � la vitesse dans le carrefour le
    'plus haut en Y, sauf pour le temps de parcours descendant dans le cas
    'd'une onde verte cadr�e par un TC descendant
    If unIndCarfM = 0 Then
        unTmpM = 0
    Else
        unTmpM = monTabCarfY(unIndCarfM).monCarfReduit.monCarrefour.monDecVitSensM
    End If
    If unIndCarfD = 0 Then
        unTmpD = 0
    Else
        If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
            'Cas d'une onde cadr�e par un TC descendant, le temps de parcours
            'descendant total est donn�e par la fin de la derni�re du tableau
            'de marche cadrant l'onde moins le d�but de la 1�re phase
            Set unTC = monSite.mesTC(monSite.monTCD)
            Set uneLastPhase = unTC.mesPhasesTMOnde(unTC.mesPhasesTMOnde.Count)
            unTmpD = uneLastPhase.monTDeb + uneLastPhase.maDureePhase - unTC.mesPhasesTMOnde(1).monTDeb
        Else
            unTmpD = monTabCarfY(unIndCarfD).monCarfReduit.monCarrefour.monDecVitSensD
        End If
    End If
End Sub

Public Sub ImprimerDureeCycle(unX0 As Long, unY0 As Long, unX)
    'Affichage d'un texte contenant 0/Dur�e du cycle
    'sur chaque trait de cycle
    uneInfoCycle = "0/" + Format(monSite.maDur�eDeCycle)
    Printer.ForeColor = 0
    Printer.CurrentX = unX - Printer.TextWidth(uneInfoCycle) / 2
    Printer.CurrentY = unY0 + Printer.TextHeight("OndeV") * 2
    Printer.Print uneInfoCycle
    Printer.Line (unX, unY0 + Printer.TextHeight(uneInfoCycle) * 2)-(unX, unY0), 0
End Sub

Public Function RemplirFicheResultPourImp() As Boolean
    'Remplissage de la fiche r�sultats pour l'imprimer
    
    Dim unNomTC As String
    
    'Calcul de l'onde verte si l'onglet courant n'est ni l'onglet
    'R�sultat d�calages et ni l'onglet Graphique onde verte
    unCalculOndeFait = True
    If monSite.TabFeux.Tab <> 3 And monSite.TabFeux.Tab <> 4 Then
        unCalculOndeFait = CalculerOndeVerte(monSite)
    End If
    
    If unCalculOndeFait Then
        'Remplissage possible car onde verte trouv�e
        RemplirFicheResultPourImp = True
        
        'Calcul des vitesses maximun si un des d�calages a �t� chang�
        'ou si un nouveau calcul d'onde verte a �t� fait
        If monSite.maModifDataDec Then
            CalculerVitMax monSite
        End If
        
        'Ajout aux TC dont on cherche la progression de ceux
        'pris en compte pour l'onde verte si on calcule une onde TC,
        'sauf s'ils en font d�j� partie
        If monSite.OptionTC.Value Then
            monSite.monTypeOnde = OndeTC
            unTCM = 0
            unTCD = 0
            i = 1
            Do While i <= monSite.mesTCutil.Count And (unTCM = 0 Or unTCD = 0)
                'Recherche de la pr�sence du TC cadrant onde sens montant
                unNomTC = monSite.mesTCutil(i).monNom
                If monTCM <> 0 Then
                    'Cas o� un TC cadre l'onde en sens montant
                    If unNomTC = monSite.mesTC(monTCM).monNom Then
                        unTCM = TrouverTCParNom(monSite, unNomTC)
                    End If
                End If
                'Recherche de la pr�sence du TC cadrant onde sens descendant
                If monTCD <> 0 Then
                    'Cas o� un TC cadre l'onde en sens descendant
                    If unNomTC = monSite.mesTC(monTCD).monNom Then
                        unTCD = TrouverTCParNom(monSite, unNomTC)
                    End If
                End If
                i = i + 1
            Loop
            If unTCM = 0 And monTCM <> 0 Then monSite.mesTCutil.Add monSite.mesTC(monTCM)
            If unTCD = 0 And monTCD <> 0 Then monSite.mesTCutil.Add monSite.mesTC(monTCD)
        End If
        'Remplir l'onglet Fiche r�sultat
        RemplirOngletFicheResult monSite
    Else
        'Remplissage impossible car onde verte non trouv�e
        RemplirFicheResultPourImp = False
    End If
End Function

Public Sub RendreNulleBandesEtDecalages(unSite As Form)
    'Mise � z�ro des bandes passantes et
    'des d�calages de tous les carrefours
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
    'R�organisation par ordre croissant de leur ordonn�e
    'd'un tableau de feu index� entre 1 et n.
    'unSens permet de changer le signe des Y � classer, c'est utilis�
    'pour le sens descendant (unSens vaut 1 ou -1)
    
    'Algo choisi : Le tri insertion (r�cup�rer sur Internet)
    'Il consiste � comparer successivement un �l�ment
    '� tous les pr�c�dents et � d�caler les �l�ments interm�diaires

    Dim i As Integer, j As Integer
    Dim unNbTotal As Integer, unFeuTmp As Feu
    
    'Tri
    unNbTotal = UBound(unTabFeu, 1)
    For j = 2 To unNbTotal
            uneFinBoucle = False
            Set unFeuTmp = unTabFeu(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                If unTabFeu(i).monOrdonn�e * unSens > unFeuTmp.monOrdonn�e * unSens Then
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
    'V�rification du passage � tous les feux verts  � la vitesse uneV
    'd'un tableau de feux dans le m�me sens et tri� par Y croisant.
    '
    'Valeur de retour : 0 si la vitesse ne passe pas
    '                   la bande passante trouv�e sinon
    
    Dim unFeu As Feu, unFeuTmp As Feu
    Dim unNbFeux As Integer, unPasseToutVert As Boolean
    Dim unCarf As Carrefour
    Dim uneColCarf As New ColCarrefour
    Dim unCarfRed As Object, uneDureeVert As Single
    Dim uneOrdonnee As Integer, unePosRef As Single
    
    unNbFeux = UBound(unTabFeu, 1)
    
    'Cr�ation d'un nouveau carrefour avec ses vitesses M et D non nulles
    'qui contiendra tous les feux du tableau de feux et dont on cherchera
    'le feu �quivalent en vitesse constante = uneV pass�e en param�tre.
    Set unCarf = uneColCarf.Add("Carrefour global", 30, 30)
    
    'Ajout � ce carrefour global des feux du tableau unTabFeu
    For i = 1 To unNbFeux
        'R�cup�ration du feu et de ses param�tres
        Set unFeu = unTabFeu(i)
        uneOrdonnee = unFeu.monOrdonn�e
        uneDureeVert = unFeu.maDur�eDeVert
        unePosRef = unFeu.maPositionPointRef
        
        'Modif du point de r�f�rence pour tenir compte du d�calage
        'du � l'onde verte en cours
        unePosRef = unePosRef + unFeu.monCarrefour.monDecModif
        
        'Ajout d'un nouveau feu
        Set unFeuTmp = unCarf.mesFeux.Add(unSensMontant, uneOrdonnee, uneDureeVert, unePosRef)
    Next i
    
    'Sauvegarde des param�tres de calcul d'onde avant modif
    unTypeOnde = unSite.monTypeOnde
    unTypeVit = unSite.monTypeVit
    uneVM = unSite.maVitSensM
    uneVD = unSite.maVitSensD
    
    'Changement des param�tres de calcul d'onde pour �tre en onde
    'double sens � vitesse const = uneV pour trouver le feu �quivalent
    unSite.monTypeOnde = OndeDouble
    unSite.monTypeVit = VitConst
    unSite.maVitSensM = uneV
    unSite.maVitSensD = uneV
    
    'Calcul du feu �quivalent �ventuel comme dans le cas d'une recherche
    'de bandes passantes.
    'Les param�tres de ce feu �quivalent seront retourn�s
    'dans les variables uneDureeVert, unePosRef et uneOrdonnee
    unPasseToutVert = CalculerFeuEquivalent(unCarf, unSensMontant, uneDureeVert, unePosRef, uneOrdonnee, False, True)
    
    'Retour de la bande passante
    If unPasseToutVert Then
        'Bande passante trouv�e
        VerifierVitessePasseToutVert = uneDureeVert
    Else
        'Bande passante non trouv�e
        VerifierVitessePasseToutVert = 0
    End If
    
    'Restauration des param�tres de calcul d'onde
    unSite.monTypeOnde = unTypeOnde
    unSite.monTypeVit = unTypeVit
    unSite.maVitSensM = uneVM
    unSite.maVitSensD = uneVD
    
    'Suppression des feux,de la collection ne contenant que le carrefour
    'et ce dernier cr�� dans cette fonction pour lib�rer la m�moire.
    For i = 1 To unCarf.mesFeux.Count
        unCarf.mesFeux.Remove 1
    Next i
    uneColCarf.Remove 1
    Set unCarf = Nothing
    Set uneColCarf = Nothing
End Function

Public Sub TriCroissantVMax(unTabV() As Single, unTabDT() As Single, unTabDY() As Single, unTabIndFeu() As Integer)
    'R�organisation par ordre croissant d'un tableau de vitesses
    'Les tableaux unTabDT et unTabDY sont aussi r�organis�s
    'pour rester coh�rent avec unTabV
    
    'Algo choisi : Le tri insertion (r�cup�rer sur Internet)
    'Il consiste � comparer successivement un �l�ment
    '� tous les pr�c�dents et � d�caler les �l�ments interm�diaires

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
                    'R�organisation
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
    ' si unSens = -1, possible < � la vitesse maxi limite
    
    'Cette vitesse est retourn�e en km/h
    
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
    
    'Calcul des vitesses max entre deux feux du m�me sens
    unInd = 0
    unNbFeux = UBound(unTabFeu, 1)
    For i = 1 To unNbFeux - 1
        Set unFeuBas = unTabFeu(i)
        For j = i + 1 To unNbFeux
            Set unFeuHaut = unTabFeu(j)
            unInd = unInd + 1
            unTabDY(unInd) = (unFeuHaut.monOrdonn�e - unFeuBas.monOrdonn�e) * unSens
            
            unDebVertHaut = unFeuHaut.monCarrefour.monDecModif + unFeuHaut.maPositionPointRef
            unDebVertHaut = ModuloZeroCycle(unDebVertHaut, unSite.maDur�eDeCycle)
            
            unFinVertBas = unFeuBas.monCarrefour.monDecModif + unFeuBas.maPositionPointRef + unFeuBas.maDur�eDeVert
            unFinVertBas = ModuloZeroCycle(unFinVertBas, unSite.maDur�eDeCycle)
            
            'Recherche du premier d�but de vert du feu haut qui est <
            '� la fin de vert bas plus la dur�e entre les feux haut et bas
            '� la vitesse maxi limite avec une pr�cision de calcul de 0.001
            'seconde et ceci modulo cycle
            unDebVertHautTrouv = False
            Do
                If unDebVertHaut < unFinVertBas + unTabDY(unInd) / uneVitMaxLim - 0.001 Then
                    unDebVertHaut = unDebVertHaut + unSite.maDur�eDeCycle
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
    'Les tableaux unTabDT et unTabDY sont aussi r�organis�s
    'pour rester coh�rent avec unTabV
    TriCroissantVMax unTabV, unTabDT, unTabDY, unTabIndFeu
    
    'On essaye toutes les vitesses max possibles en commen�ant par
    'la plus grande, donc le dernier �l�ment du tableau de vitesse
    '(Indice = unNbVMax)
   Do
        i = 1
        unIndBas = unTabIndFeu(unNbVMax)
        
        'Date de d�part du Feu bas
        Set unFeuBas = unTabFeu(unIndBas)
        unFinVertBas = unFeuBas.monCarrefour.monDecModif + unFeuBas.maPositionPointRef + unFeuBas.maDur�eDeVert
        'Test si on passe � tous les feux vert
        unPasseToutVert = True
        Do While unPasseToutVert And i <= unNbFeux
            unPasseAuVert = False
            If i <> unIndBas Then
                'R�cup du feu haut
                Set unFeuHaut = unTabFeu(i)
                
                'D�but et fin de vert du feu haut
                unDebVertHaut = unFeuHaut.monCarrefour.monDecModif + unFeuHaut.maPositionPointRef
                unFinVertHaut = unDebVertHaut + unFeuHaut.maDur�eDeVert
                
                'Date de passage au feu haut
                uneDatePassage = unFinVertBas + (unFeuHaut.monOrdonn�e - unFeuBas.monOrdonn�e) * unSens / unTabV(unNbVMax)
                Do While uneDatePassage <= unDebVertHaut - 0.001
                    'Cas d'une date de passage inf�rieur au d�but de vert du feu
                    'haut ==> on lui rajoute la Dur�e du cycle jusqu'� une valeur
                    'sup�rieure au d�but de vert pour la prendre en compte dans
                    'les calculs suivants
                    uneDatePassage = uneDatePassage + unSite.maDur�eDeCycle
                Loop
                
                'V�rification si la date de passage est entre le d�but et
                'la fin de vert du feu haut modulo cycle � une pr�cision de
                'calcul de 0.001
                Do
                    If uneDatePassage > unDebVertHaut - 0.001 And uneDatePassage < unFinVertHaut + 0.001 Then
                        unPasseAuVert = True
                    Else
                        'Incr�mentation suivante
                        unDebVertHaut = unDebVertHaut + unSite.maDur�eDeCycle
                        unFinVertHaut = unFinVertHaut + unSite.maDur�eDeCycle
                    End If
                Loop Until uneDatePassage < unDebVertHaut - 0.001 Or unPasseAuVert = True
                
                unPasseToutVert = unPasseAuVert
            Else
                unPasseToutVert = True
            End If
            
            'Incr�mentation suivante
            i = i + 1
        Loop
        
        If Not unPasseToutVert Then
            'Modification des �l�ments d'indice le dernier
            unTabDT(unNbVMax) = unTabDT(unNbVMax) + unSite.maDur�eDeCycle
            unTabV(unNbVMax) = unTabDY(unNbVMax) / unTabDT(unNbVMax)
            'Re-triage par vitesse croissante
            TriCroissantVMax unTabV, unTabDT, unTabDY, unTabIndFeu
        End If
        
        'Boucle jusqu'au passage � tous les verts ou si vitesse est
        '< une vitesse limite mini pour avoir une condition d'arr�t
    Loop Until unPasseToutVert Or unTabV(unNbVMax) < uneVitMinLim
    
    'Retour de la vitesse montante maxi trouv�e en km/h
    If unTabV(unNbVMax) < uneVitMinLim Then
        CalculerVMaxInfVMaxLim = uneVitMinLim * 3.6
    Else
        CalculerVMaxInfVMaxLim = unTabV(unNbVMax) * 3.6
    End If
End Function
        
Public Function DonnerIndicePlage(uneColPlageGraphic As Collection, unIndCarf, unIndFeu) As Integer
    'Recherche de la plage graphique s�lectionnable repr�sentant le feu
    'd'indice unIndFeu du carrefour d'indice unIndCarf
    'dans une collection contenant cette plage
    'Retourne l'indice trouv� ou 0 si aucune plage ne correspond
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
    'Retourne si les d�calages ont �t� obtenus par modification
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
    'Calcul de la date dans une phase donn�e suivant son type
    'La phase ne doit pas �tre de type Arret.
    Dim uneVal As Single
    
    If unePhase.monType = VConst Then
        CalculerDateDansPhase = unePhase.monTDeb + Abs((unY - unePhase.monYDeb) / unePhase.maVitPhase)
    ElseIf unePhase.monType = Accel Then
        'Arrondi de valeur � 1 par rapport � la pr�cision 0.001 pour
        '�viter d'avoir un Sqr de uneVal avec uneVal voisin de 0 mais < 0
        '(unY - unePhase.monYDeb) est forc�ment > ou = 0
        '==> Probl�me si = �a doit faire 0 et pas 0.000000023 ou -0.000000002
        uneVal = (unY - unePhase.monYDeb) / unePhase.maLongPhase
        If Abs(uneVal) < 0.001 Then uneVal = 0
        'Calcul de la date
        CalculerDateDansPhase = unePhase.monTDeb + unePhase.maDureePhase * Sqr(uneVal)
    ElseIf unePhase.monType = Decel Then
        'Arrondi de valeur � 1 par rapport � la pr�cision 0.001 pour
        '�viter d'avoir un Sqr de (1-uneVal) avec uneVal voisin de 1 mais < 1
        '(unY - unePhase.monYDeb) est forc�ment < ou = unePhase.maLongPhase
        '==> Probl�me si = �a doit faire 1 et pas 1.00023 ou 0.9999982
        uneVal = (unY - unePhase.monYDeb) / unePhase.maLongPhase
        If Abs(uneVal - 1) < 0.001 Then uneVal = 1
        'Calcul de la date
        CalculerDateDansPhase = unePhase.monTDeb + unePhase.maDureePhase * (1 - Sqr(1 - uneVal))
    Else
        MsgBox "ERREUR de programmation dans OndeV dans CalculerDateDansPhase", vbCritical
    End If
End Function

Public Function CalculerYDansPhaseParabole(unePhase As PhaseTabMarche, unT As Single) As Single
    'Calcul du Y � l'instant unT dans une phase donn�e se dessinant en
    'discr�tisant une parabole
    'La phase ne peut �tre que du type Accel ou Decel.
    
    'Formules de calcul obtenues par inversion de celles
    'de la fonction CalculerDateDansPhase
    Dim unDT As Single
    
    'Calcul de l'�cart en temps par rapport au d�but de la phase
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
        MsgBox "Aucun objet n'a �t� s�lectionn�", vbInformation
    Else
        frmPropsObjPick.Show vbModal
    End If
End Sub

Public Sub ViderObjPick()
    'Mise � vide de l'objet s�lectionn� graphiquement
    monTypeObjPick = NoSel
    Set monObjPick = Nothing
End Sub

Public Function DonnerObjPick() As Object
    Set DonnerObjPick = monObjPick
End Function

Public Sub CalculerB1B2(K As Single, A1 As Single, A2 As Single, B1 As Single, B2 As Single, unBB1 As Single, unBB2 As Single)
    'Cette proc�dure alimente les variables unBB1 et unBB2 pass�es en param�tres
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
    'Indique si une modif a eu lieu dans les donn�es TC ne cadrant pas l'onde
    'ou les donn�es des TC cadrant l'onde
    Dim unIndTC As Integer
    
    'R�cup�ration du TC modifi�
    unIndTC = monSite.ComboTC.ListIndex + 1
    
    'Indication de la bonne modif
    If monSite.monTypeOnde = OndeTC And (monSite.monTCM = unIndTC Or monSite.monTCD = unIndTC) Then
        monSite.maModifDataOndeTC = True
    Else
        monSite.maModifDataTC = True
    End If
End Sub

Public Function UtiliserDecalagesImposes(unIndUniqCarfImp, unNbFeuDateImpSensM, unNbFeuDateImpSensD) As Object
    'Fonction pr�parant le calcul d'onde verte avec des d�calages
    'impos�s � certains carrefours
    'Elle retourne le carrefour r�duit r�duisant tous les carrefours
    '� date impos�e et alimente unIndUniqCarfImp avec l'indice du seul
    'carrefour � date impos� si c'est le cas, sinon il vaut 0
    'Elle retourne aussi le nombre de feux � date impos�s dans le sens montant
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
        'Cas d'un seul carrefour � date impos�e
        unIndUniqCarfImp = uneColCarf(1).maPosition
    Else
        'Autres cas
        unIndUniqCarfImp = 0
    End If
    
    If uneColCarf.Count Then
        'Cas o� il y a des carrefours � date impos�e
        
        'Initialisation du message d'erreur
        unMsg = "Impossible de calculer les ondes vertes avec les valeurs actuelles des d�calages impos�s." + Chr(13)
        unMsg = unMsg + Chr(13) + "Le calcul des ondes vertes ne tiendra pas compte des d�calages impos�s aux carrefours."
        
        'Calcul des temps de parcours sans s'occuper des dates impos�es
        CalculerTempsParcours monSite

        'R�duction de tous les carrefours r�duits � date impos�e en un seul
        'L'ajout du d�calage des carrefours � date impos�e � leur point de
        'r�f�rence de leur carrefour r�duit est fait dans CalculerFeuEquivalent
        'appel�e dans ReduireCarfsEnUn avec le param�tre unDecalModif = TRUE
        'Cet ajout est d�crit dans le dossier de sp�cifs, partie Date impos�e
        Set unCarfRed = ReduireCarfsEnUn(uneColCarf, unCarf, unIndCarfBas, unIndCarfHaut, unNbFeuDateImpSensM, unNbFeuDateImpSensD)
        
        'Calcul des d�calages en temps aux carrefours le plus haut
        'pour le sens montant et le plus bas pour le sens descendant
        unCarf.monDecVitSensD = monSite.mesCarrefours(unIndCarfBas).monDecVitSensD
        unCarf.monDecVitSensM = monSite.mesCarrefours(unIndCarfHaut).monDecVitSensM
        
        If unCarfRed Is Nothing Then
            'Cas o� la r�duction des carrefours r�duits des carrefours � date
            'impos�e n'a pas march�
            MsgBox unMsg, vbInformation
        Else
            'Lien � un nouveau carrefour du carrefour r�duit cr�� ci-dessus
            'Ce carrefour servira � stocker le d�calage calcul�
            'et il est diff�rent des carrefours � date impos�e
            Set unCarf.monCarfRed = unCarfRed
            unCarf.monNom = ""
            'Aucun carrefour cr�� par saisie ne peut avoir de nom vide
            
            If TypeOf unCarfRed Is CarfReduitSensDouble Then
                'Cas double sens
                
                'Calcul de l'�cart
                'l'�cart est le temps s'�coulant entre les �v�nements "passage au vert
                'dans le sens montant" et "fin du vert dns le sens descendant" apr�s
                'projection sur une r�f�rence commune � l'ensemble des carrefours
                '(cf Dossier de programmation et sp�cifs)
                'On utilise des d�calages dus aux vitesses variables ou
                'constantes de chaque carrefour
                unCarfRed.monEcart = unCarfRed.maPosRefD + unCarf.monDecVitSensD + unCarfRed.maDureeVertD
                unCarfRed.monEcart = unCarfRed.monEcart - (unCarfRed.maPosRefM - unCarf.monDecVitSensM)
                'On ram�ne l'�cart modulo entre [0, dur�ee du cycle[
                unCarfRed.monEcart = ModuloZeroCycle(unCarfRed.monEcart, monSite.maDur�eDeCycle)
            End If
            
            'Suppression des carrefours r�duits du site courant issus
            'd'un carrefour � date impos�e, ces derniers sont dans uneColCarf
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
    
    'Lib�ration de la m�moire
    For i = 1 To uneColCarf.Count
        uneColCarf.Remove 1
    Next i
    Set uneColCarf = Nothing
End Function


Public Function ReduireCarfsEnUn(uneColCarf As ColCarrefour, unCarf As Carrefour, unIndCarfBas, unIndCarfHaut, unNbFeuDateImpSensM, unNbFeuDateImpSensD) As Object
    'R�duction des carrefours de la collection uneColCarf qui contient tous
    'les feux �quivalents montant et descendant des carrefours r�duits
    
    'Elle retourne un carrefour r�duit valant :
    '   - nothing si aucun feu �quivalent trouv�
    '   - de type CarfReduitSensUnique si un feu �quivalent trouv� (montant ou descendant)
    '   - de type CarfReduitSensDouble si deux feux �quivalents trouv�s (montant et descendant)
    '
    'Les variables unIndCarfBas et unIndCarfHaut, donnant les indices des
    'carrefours le plus bas et le plus haut de uneColCarf, sont renseign�s
    'par cette fonction pour ainsi r�cup�rer leurs d�calages en temps respectif
    '
    'Elle Retourne aussi le nombre de feux � date impos�e dans les sens M et D
    
    Dim unFeu As Feu, unTCM As TC, unTCD As TC
    Dim unCarfRed As Object
    Dim unNbFeuxM As Integer, unNbFeuxD As Integer
    Dim unCarfRedU As CarfReduitSensUnique
    Dim unCarfRed2 As CarfReduitSensDouble
    Dim uneDureeVertM As Single, unPosRefM As Single, uneOrdonneeM As Integer
    Dim uneDureeVertD As Single, unPosRefD As Single, uneOrdonneeD As Integer
    
    'Initialisation du nombre de feux montants et descendants du carrefour
    'r�duit que l'on va cr�er
    unNbFeuxM = 0
    unNbFeuxD = 0
    
    'Cr�ation d'une collection de feu pour le carrefour temporaire stockant
    'le carrefour r�duisant touts les carrefours � date impos�e
    Set unCarf.mesFeux = New ColFeu
    
    'Initialisation pour la recherche du carrefour le plus haut et le plus bas
    'Y dans OndeV compris entre -9999 et 9999 m�tres
    unYMin = 100000
    unYMax = -100000
    
    'Initialisation des indices des carf le plus haut et le plus bas
    unNbCarfImp = uneColCarf.Count
    unIndCarfHaut = uneColCarf(unNbCarfImp).maPosition
    unIndCarfBas = unIndCarfHaut
    
    'Ajout � ce carrefour global des feux �quivalents
    'des carrefours r�duits double sens
    For i = 1 To unNbCarfImp
        'R�cup du carrefour r�duit
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
        
        'Cr�ation des feux du carrefour dont la r�duction, donc recherche des
        'feux �quivalent montant et descendant donne le carrefour r�duit
        'r�duisant tous les carrefours � date impos�e
        
        'On stockera dans le feu cr��, le carrefour correspondant � celui r�duit
        'car c'est son d�calage modifi� (= monDecModif) et celui en temps du
        '� sa vitesse (= monDecVitSensM ou D) qui est utilis� dans le calcul du
        'feu �quivalent montant ou descendant
        If TypeOf unCarfRed Is CarfReduitSensDouble Then
            'Ajout d'un nouveau feu montant
            unNbFeuxM = unNbFeuxM + 1
            'Stockage de l'indice du carrefour de ce feu M
            unIndCarfM = i
            Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonneeM, unCarfRed.maDureeVertM, unCarfRed.maPosRefM)
            'Stockage du carrefour de celui r�duit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Ajout d'un nouveau feu descendant
            unNbFeuxD = unNbFeuxD + 1
            Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonneeD, unCarfRed.maDureeVertD, unCarfRed.maPosRefD)
            'Stockage du carrefour de celui r�duit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu D
            unIndCarfD = i
        ElseIf TypeOf unCarfRed Is CarfReduitSensUnique And unCarfRed.HasFeuMontant Then
            'Ajout d'un nouveau feu montant
            unNbFeuxM = unNbFeuxM + 1
            Set unFeu = unCarf.mesFeux.Add(True, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
            'Stockage du carrefour de celui r�duit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu M
            unIndCarfM = i
        ElseIf TypeOf unCarfRed Is CarfReduitSensUnique And unCarfRed.HasFeuDescendant Then
            'Ajout d'un nouveau feu descendant
            unNbFeuxD = unNbFeuxD + 1
            Set unFeu = unCarf.mesFeux.Add(False, unCarfRed.monOrdonnee, unCarfRed.maDureeVert, unCarfRed.maPosRef)
            'Stockage du carrefour de celui r�duit (cf commentaires avant le TypeOf)
            Set unFeu.monCarrefour = unCarfRed.monCarrefour
            'Stockage de l'indice du carrefour de ce feu D
            unIndCarfD = i
        Else
            MsgBox "ERREUR de programmation dans OndeV dans ReduireCarfsEnUn", vbCritical
        End If
    Next i
          
    'Retour du nombre de feux � date impos�e dans les sens M et D
    unNbFeuDateImpSensM = unNbFeuxM
    unNbFeuDateImpSensD = unNbFeuxD
    
    'Prise en compte d'une onde cadr�e par un TC montant et/ou descendant
    If monSite.monTypeOnde = OndeTC And monSite.monTCM > 0 Then
        'Cas d'une onde cadr�e par unTC dans le sens montant
        Set unTCM = monSite.mesTC(monSite.monTCM)
    Else
        'Cas d'une onde non cadr�e par unTC dans le sens montant
        Set unTCM = Nothing
    End If
    
    If monSite.monTypeOnde = OndeTC And monSite.monTCD > 0 Then
        'Cas d'une onde cadr�e par unTC dans le sens descendant
        Set unTCD = monSite.mesTC(monSite.monTCD)
    Else
        'Cas d'une onde non cadr�e par unTC dans le sens descendant
        Set unTCD = Nothing
    End If
        
    'R�cup�ration de la vitesse d'arriv�e sur les carrefours le plus
    'haut pour le sens montant et le plus bas pour le sens descendant,
    'car elles servent dans le calcul du point de r�f�rence dans la
    'fonction CalculerFeuEquivalent, dans le carrefour du carrefour
    'r�duisant tous les carrefours � date impos�e
    unCarf.maVitSensD = monSite.mesCarrefours(unIndCarfBas).maVitSensD
    unCarf.maVitSensM = monSite.mesCarrefours(unIndCarfHaut).maVitSensM
                
    'Calcul du feu �quivalent montant �ventuel
    'avec unDecalModif = True (6�me param�tre)
    unFeuEquivSensMExist = CalculerFeuEquivalent(unCarf, True, uneDureeVertM, unPosRefM, uneOrdonneeM, True, , , , unTCM)
    If unNbFeuxM = 1 And (unNbFeuxD > 1 Or (unNbFeuxD = 1 And unIndCarfM <> unIndCarfD)) Then
        'Cas o� le feu montant �quivalent trouv� est l'�quivalent d'un seul feu
        'montant mais avec un feu descendant �quivalent qui est celui de plusieurs
        'feux descendants ou si le seul feu descendant est celui d'un autre carrefour � sens unique D
        '==> Rajout du d�calage du carrefour le plus haut pour
        'lier ces deux feux �quivalents car le feu M contrairement au feu D n'a
        'pas ce d�calage int�gr�
        unPosRefM = unPosRefM + monSite.mesCarrefours(unIndCarfHaut).monDecModif
    End If
                      
    'Calcul du feu �quivalent descendant �ventuel
    'avec unDecalModif = True (6�me param�tre)
    unFeuEquivSensDExist = CalculerFeuEquivalent(unCarf, False, uneDureeVertD, unPosRefD, uneOrdonneeD, True, , , , unTCD)
    If unNbFeuxD = 1 And (unNbFeuxM > 1 Or (unNbFeuxM = 1 And unIndCarfM <> unIndCarfD)) Then
        'Cas o� le feu descendant �quivalent trouv� est l'�quivalent d'un seul feu
        'descendant mais avec un feu montant �quivalent qui est celui de plusieurs
        'feux montants ou si le seul feu montant est celui d'un autre carrefour � sens unique M
        '==> Rajout du d�calage du carrefour le plus bas pour
        'lier ces deux feux �quivalents car le feu D contrairement au feu M n'a
        'pas ce d�calage int�gr�
        unPosRefD = unPosRefD + monSite.mesCarrefours(unIndCarfBas).monDecModif
    End If
    
    If unFeuEquivSensMExist And unFeuEquivSensDExist Then
        'Ajout aux carrefours r�duits double sens du site courant
        Set unCarfRed2 = monSite.mesCarfReduitsSens2.Add(unCarf)
        'Alimentation d'un carrefour r�duit double sens
        unCarfRed2.SetPropsSensM uneDureeVertM, unPosRefM, uneOrdonneeM
        unCarfRed2.SetPropsSensD uneDureeVertD, unPosRefD, uneOrdonneeD
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRed2
    ElseIf unFeuEquivSensMExist Then
        'Ajout aux carrefours r�duits sens unique montant du site courant
        Set unCarfRedU = monSite.mesCarfReduitsSensM.Add(unCarf, True, uneDureeVertM, unPosRefM, uneOrdonneeM)
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRedU
    ElseIf unFeuEquivSensDExist Then
        'Ajout aux carrefours r�duits sens unique descendant du site courant
        Set unCarfRedU = monSite.mesCarfReduitsSensD.Add(unCarf, False, uneDureeVertD, unPosRefD, uneOrdonneeD)
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = unCarfRedU
    Else
        'Affectation de la valeur de retour de cette fonction
        Set ReduireCarfsEnUn = Nothing
    End If
        
    'Lib�ration de la m�moire
    For i = 1 To unCarf.mesFeux.Count
        unCarf.mesFeux.Remove 1
    Next i
    Set unCarf.mesFeux = Nothing
End Function

Public Function TrouverLesCarfsAvecDateImp() As ColCarrefour
    'Retourne une collection contenant les carrefours � d�calage impos�
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
        'R�cup du carrefour r�duit
        Set unCarfRed = monSite.mesCarfReduitsSens2(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : D�calage = "; unCarfRed.monCarrefour.monDecModif
        Debug.Print Tab; "==> Sens M : "; unCarfRed.monOrdonneeM; unCarfRed.maDureeVertM; unCarfRed.maPosRefM; unCarfRed.monCarrefour.monDecVitSensM
        Debug.Print Tab; "==> Sens D : "; unCarfRed.monOrdonneeD; unCarfRed.maDureeVertD; unCarfRed.maPosRefD; unCarfRed.monCarrefour.monDecVitSensD
    Next i
    
    unNbFeuxSensM = monSite.mesCarfReduitsSensM.Count
    For i = 1 To unNbFeuxSensM
        'R�cup du carrefour r�duit
        Set unCarfRed = monSite.mesCarfReduitsSensM(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : D�calage = "; unCarfRed.monCarrefour.monDecModif; unCarfRed.monOrdonnee; unCarfRed.maDureeVert; unCarfRed.maPosRef
    Next i
    
    'Ajout � ce carrefour global des feux �quivalents
    'des carrefours r�duits � sens unique descendant
    unNbFeuxSensD = monSite.mesCarfReduitsSensD.Count
    For i = 1 To unNbFeuxSensD
        'R�cup du carrefour r�duit
        Set unCarfRed = monSite.mesCarfReduitsSensD(i)
        Debug.Print unCarfRed.monCarrefour.monNom; " : D�calage = "; unCarfRed.monCarrefour.monDecModif; unCarfRed.monOrdonnee; unCarfRed.maDureeVert; unCarfRed.maPosRef
    Next i
    Debug.Print "****************** Fin de TestForDebug **************"
End Sub

Public Sub RecalculerAvecDateImp(unCarf As Carrefour, unText As String)
    'Lance le recalcul avec les dates impos�es si la valeur de la variable
    'unText (valant Oui pour date impos� ou Non sinon ) est diff�rent du
    'type de d�calage impos� (1 = impos�, sinon 0)
    
    If unCarf.monDecImp = 1 And unText = "Non" Then
        'Modification du type de d�calage (0 pour non impos�)
        unCarf.monDecImp = 0
        'Indication d'un changement pour lancer le calcul
        monSite.maModifDataOnde = True
        'Calcul d'onde verte en tenant compte des d�calages impos�s
        CalculerOndeVerte monSite, True
    ElseIf unCarf.monDecImp = 0 And unText = "Oui" Then
        'Modification du type de d�calage (1 pour impos�)
        unCarf.monDecImp = 1
        'Indication d'un changement pour lancer le calcul
        monSite.maModifDataOnde = True
        'Calcul d'onde verte en tenant compte des d�calages impos�s
        CalculerOndeVerte monSite, True
    End If
End Sub

Public Sub CorrectionDateImpos�e(uneForm As Form, unCarfRed As Object, unB1 As Single, unB2 As Single, unNbFeuxDateImpSensM, unNbFeuxDateImpSensD)
    'Correction de la solution � date impos�e si on obtient un carrefour
    'r�duisant tous les carrefours � date impos�e qui est � sens unique
    'alors que ces carrefours ont des feux dans les 2 sens, donc il n'y a pas
    'de bandes passantes dans le sens oppos� � celui du carrefour r�duit
    
    If TypeOf unCarfRed Is CarfReduitSensUnique Then
        If unCarfRed.monSensMontant And unNbFeuxDateImpSensD > 0 Then
            unB2 = 0
            uneForm.monOndeDoubleTrouve = False
            unMsg = "Une solution a �t� trouv�e dans le sens Montant, mais pas dans le sens Descendant"
            MsgBox unMsg, vbInformation
        ElseIf unCarfRed.monSensMontant = False And unNbFeuxDateImpSensM > 0 Then
            unB1 = 0
            uneForm.monOndeDoubleTrouve = False
            unMsg = "Une solution a �t� trouv�e dans le sens Descendant, mais pas dans le sens Montant"
            MsgBox unMsg, vbInformation
        End If
    End If
End Sub

Public Sub DessinerBandesInterCarfVP(uneZoneDessin As Object, unTM1 As Single, unTD1 As Single, unY0 As Long, uneHt As Long, unDY As Long, unT As Long, unCarfRedM1, unCarfRedD1, unYMin As Long, unIndCarfBasM, unIndCarfD, uneLg As Long, unNewX0, unMaxT)
    'Dessin des bandes inter-carrefours dans le sens M ou D suivant le cas.
    'Utilisation des Y sur vitesse pour avoir les d�calages inter-carrefours
    'Utiliser si choix coch� dans le menu Montrer les bandes inter-carrefours
    Dim unY As Long
    Dim unCarf As Carrefour, unCarfRed As Object
    Dim unCarfPred As Object
    Dim unDebVertM As Single, unDebVertD As Single
    Dim unFinVertM As Single, unFinVertD As Single
    Dim unMaxDebVert As Single, unMinFinVert As Single
    Dim unDebVertMPred As Single, unDebVertDPred As Single
    Dim unFinVertMPred As Single, unFinVertDPred As Single
    Dim uneDur�eVertMPred As Single, uneDur�eVertDPred As Single
    Dim unTmpInterCarf As Single
    Dim unLastDebVertM As Single, unLastDebVertD As Single
    
    'Sauvegarde du style de dessin
    unDrawStyleSave = uneZoneDessin.DrawStyle
    'Dessin en pointill�
    uneZoneDessin.DrawStyle = vbDash
    
    unNbCarf = UBound(monTabCarfY, 1)
    With monSite
        For i = 1 To unNbCarf
            'Parcours des carrefours dans le sens des Y croissants
            'pour l'onde verte montante car on dessine � partir du
            'carrefour le plus bas ayant un feu montant
            Set unCarfRed = monTabCarfY(i).monCarfReduit
            Set unCarf = unCarfRed.monCarrefour
            If unIndCarfBasM > 0 Then
                'Cas d'une onde verte montante possible ==> Dessin
                'unIndCarfBasM > 0 dit qu'on a trouv� des carrefours montants
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
                    'l'abscisse du point pr�c�dent plus le d�calage en
                    'temps entre le carrefour courant et le pr�c�dent montant
                    unTM2 = unDebVertMPred + unTmpInterCarf
                    
                    'Ordonn�e �gale � l'ordonn�e du carrefour r�duit courant
                    'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                    unYM2 = unCarfRed.DonnerYSens(True)
                    
                    'Conversion en coordonn�es �cran de unYM2
                    unY = ConvertirReelEnEcran(unYM2 - unYMin, unDY, uneHt)
                    unY = unY0 - unY
                    
                    'Dessin de l'onde verte montante inter-carrefours donc
                    'entre ce carrefour r�duit et son pr�c�dent en Y
                    'et si ce n'est pas le 1er carrefour montant le + bas
                    
                    If i <> unIndCarfBasM Then
                        'Cas d'une onde montante
                        'Calcul du d�but de vert de ce carrefour r�duit
                        unDebVertM = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(True)
                        unDebVertM = ModuloZeroCycle(unDebVertM, .maDur�eDeCycle)
                        'Calcul du nombre de cycle s�parant le d�but de
                        'vert du d�but de l'onde verte montante
                        unNbCycle = Fix((0.001 + unTM2 - unDebVertM) / .maDur�eDeCycle)
                        If unNbCycle < 0 And .maBandeModifM = 0 Then
                            'Si pas de bande commune, on ne peut pas �tre en retard
                            'unTM2 ne doit pas �tre corrig� s'il est < unDebVertM
                            unNbCycle = 0
                        End If
                        If unTM2 < unDebVertM - 0.001 Then
                            'D�but de vert > T de onde montante
                            '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                            unDebVertM = unDebVertM + unNbCycle * .maDur�eDeCycle
                        ElseIf unTM2 > unDebVertM + unCarfRed.DonnerDureeVertSens(True) + 0.001 Then
                            'Fin de vert < T de d�part onde montante
                            '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                            unDebVertM = unDebVertM + unNbCycle * .maDur�eDeCycle
                        End If
                                                  
                        'Calcul de la fin de vert de ce carrefour r�duit
                        unFinVertM = unDebVertM + unCarfRed.DonnerDureeVertSens(True)
                        RecalculerDebEtFinVert unCarfRed, True, .maDur�eDeCycle, unCarf.DonnerVitCarfSens(True), unDebVertM, unFinVertM
                        unFinVertMPred = unDebVertMPred + uneDur�eVertMPred
    
                        unI = 0
                        unJ = 0
                        unLastDebVertD�j�Stock� = False
                        'projection du d�but et fin de vert du carrefour
                        'pr�c�dent sur la droite Y = Y du carrefour courant
                        unDebVertMPred = unDebVertMPred + unTmpInterCarf
                        unFinVertMPred = unFinVertMPred + unTmpInterCarf
                        'Initialisation du LastDebVert au cas o� aucun
                        'bande inter-carf trouv�e
                        unLastDebVertM = unDebVertM
                        Do
                            'La boucle sert pour afficher toutes les bandes
                            'inter-carrefour pour cela on regarde dans
                            'le cycle en cours et le suivant et on prend la
                            'bande inter-carf maximale
                            
                            'On prend le minimun des fins de vert projet�
                            'sur la droite Y = Y du carrefour courant
                            If (unFinVertMPred + unJ * .maDur�eDeCycle) < (unFinVertM + unI * .maDur�eDeCycle) Then
                                unMinFinVert = unFinVertMPred + unJ * .maDur�eDeCycle
                            Else
                                unMinFinVert = unFinVertM + unI * .maDur�eDeCycle
                            End If
                            
                            'On prend le maximun des d�buts de vert projet�
                            'sur la droite Y = Y du carrefour courant
                            If (unDebVertMPred + unJ * .maDur�eDeCycle) > (unDebVertM + unI * .maDur�eDeCycle) Then
                                unMaxDebVert = unDebVertMPred + unJ * .maDur�eDeCycle
                            Else
                                unMaxDebVert = unDebVertM + unI * .maDur�eDeCycle
                            End If
                            
                            'Test de l'existence d'une bande inter-carrefour
                            'sup�rieure � 1 seconde
                            uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                            
                            If uneBandeInterCarfExist Then
                                If unLastDebVertD�j�Stock� = False Then
                                    'Stockage du dernier debvert ayant une bande
                                    'inter-carrefour, ce stockage est fait une fois et
                                    'une seule entre deux carrefours
                                    unLastDebVertD�j�Stock� = True
                                    unLastDebVertM = unDebVertM + unI * .maDur�eDeCycle
                                End If
                                'Remise dans l'englobant total si
                                'MinFinVert en sort
                                If unMinFinVert > unMaxT + 0.01 Then
                                    unMinFinVert = unMinFinVert - .maDur�eDeCycle
                                    unMaxDebVert = unMaxDebVert - .maDur�eDeCycle
                                End If
                            End If
                            
                            unI = unI + 1
                            If unI = 2 Then
                                'On se place pour essayer les d�but et fin de vert
                                'du carrefour courant dans le cycle courant avec les
                                'd�but et fin de vert du carrefour pr�c�dent dans le
                                'cycle suivant
                                unI = 0
                                unJ = 1
                            End If
    
                            'Dessin de bande inter-carrefour
                            If uneBandeInterCarfExist Then
                                'Cas du dessin des bandes inter-carrefours
                                'voitures d'une onde TC
                                
                                'Conversion en coordonn�es �cran
                                unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 1�re partie bande montante inter-carrefours
                                uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                                
                                'Conversion en coordonn�es �cran
                                unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 2�me partie bande montante inter-carrefours
                                uneZoneDessin.Line (unX2, unYMpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfM
                            End If
                            'Boucle fait trois fois pour trouver toutes
                            'les bande inter-carrefour
                        Loop Until unI = 1 And unJ = 1
                    End If 'Fin du dessin de bande verte montante inter-carrefours
                             
                   'Stockage du d�but de vert pr�c�dent
                    If i = unIndCarfBasM Then
                        'Calcul sp�cial pour le carrefour le + bas montant
                        unDebVertMPred = unCarfRedM1.monCarrefour.monDecModif + unCarfRedM1.DonnerPosRefSens(True)
                        unDebVertMPred = ModuloZeroCycle(unDebVertMPred, .maDur�eDeCycle)
                        If unTM1 < unDebVertMPred - 0.001 Then
                            'D�but de vert > T de d�part onde montante
                            '==> Recul d'un cycle
                            unDebVertMPred = unDebVertMPred - .maDur�eDeCycle
                        ElseIf unTM1 > unDebVertMPred + unCarfRedM1.DonnerDureeVertSens(True) + 0.001 Then
                            'Fin de vert < T de d�part onde montante
                            '==> Avanc� d'un cycle
                            unDebVertMPred = unDebVertMPred + .maDur�eDeCycle
                        End If
                        'Initialisation de la dur�e de vert du premier
                        'carrefour descendant
                        uneDur�eVertMPred = unCarfRedM1.DonnerDureeVertSens(True)
                    Else
                        'Affectation du d�but de vert du carf r�duit pr�c�dent
                        unDebVertMPred = unLastDebVertM
                        'Affectation de la dur�e de vert du carf r�duit pr�c�dent
                        '� faire car modifs possibles dans RecalculerDebEtFinVert
                        uneDur�eVertMPred = unFinVertM - unDebVertM
                    End If
                    
                    'Stockage de l'indice de ce carrefour
                    unIndCarfMPred = i
                    'Stockage du Y �cran  du point pr�c�dent pour le coup suivant
                    unYMpred = unY
                End If
            End If
            
            'Parcours des carrefours dans le sens des Y d�croissants
            'pour l'onde verte descendante car on dessine � partir du
            'carrefour le plus haut ayant un feu descendant
            Set unCarfRed = monTabCarfY(unNbCarf + 1 - i).monCarfReduit
            Set unCarf = unCarfRed.monCarrefour
            
            If unIndCarfD > 0 Then
                'Cas d'une onde verte descendante possible ==> Dessin
                'unIndCarfD > 0 dit qu'on a trouv� des carrefours descendants
                If unCarfRed.HasFeuDescendant = True Then
                    'Cas d'un carrefour contraignant de l'onde verte descendante
                    'donc ayant un feu de sens descendant
                
                    'Calcul du temps de parcours inter-carrefours qui vaut
                    '(Y carf courant - Y carf pr�c�dent) / vitesse carf courant
                    'Ce temps est > car la diff�rence des Y est < 0
                    '(car les carrefours sont parcourus dans le sens des Y
                    'd�croissants pour le sens descendant et les Vitesses
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
                    'l'abscisse du point pr�c�dent plus le d�calage en
                    'temps entre le carrefour courant et le pr�c�dent descendant
                    unTD2 = unDebVertDPred + unTmpInterCarf
                  
                    'Ordonn�e �gale � l'ordonn�e du carrefour r�duit courant
                    'par polymorphisme entre les classes CarfReduitSensDouble et CarfReduitSensUnique
                    unYD2 = unCarfRed.DonnerYSens(False)
                    
                    'Conversion en coordonn�es �cran de unYD2
                    unY = ConvertirReelEnEcran(unYD2 - unYMin, unDY, uneHt)
                    unY = unY0 - unY
                    
                    'Dessin de l'onde verte descendante inter-carrefours donc
                    'entre ce carrefour r�duit et son pr�c�dent en Y
                    'et si ce n'est pas le 1er carrefour descendante le + haut
                    
                    If (unNbCarf + 1 - i) <> unIndCarfD Then
                        'Calcul du d�but de vert de ce carrefour r�duit
                        unDebVertD = unCarf.monDecModif + unCarfRed.DonnerPosRefSens(False)
                        unDebVertD = ModuloZeroCycle(unDebVertD, .maDur�eDeCycle)
                        'Calcul du nombre de cycle s�parant le d�but de
                        'vert du d�but de l'onde verte descendante
                        unNbCycle = Fix((0.001 + unTD2 - unDebVertD) / .maDur�eDeCycle)
                        If unNbCycle < 0 And .maBandeModifD = 0 Then
                            'Si pas de bande commune, on ne peut pas �tre en retard
                            'unTD2 ne doit pas �tre corrig� si il est < unDebVertD
                            unNbCycle = 0
                        End If
                        If unTD2 < unDebVertD - 0.001 Then
                            'D�but de vert > T de onde descendante
                            '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                            unDebVertD = unDebVertD + unNbCycle * .maDur�eDeCycle
                        ElseIf unTD2 > unDebVertD + unCarfRed.DonnerDureeVertSens(False) + 0.001 Then
                            'Fin de vert < T de d�part onde descendante
                            '==> Recul ou Avanc� d'un nombre entier cycle d�pendant du temps de parcours
                            unDebVertD = unDebVertD + unNbCycle * .maDur�eDeCycle
                        End If
                                                  
                        'Calcul de la fin de vert de ce carrefour r�duit
                        unFinVertD = unDebVertD + unCarfRed.DonnerDureeVertSens(False)
                        RecalculerDebEtFinVert unCarfRed, False, .maDur�eDeCycle, unCarf.DonnerVitCarfSens(False), unDebVertD, unFinVertD
                        unFinVertDPred = unDebVertDPred + uneDur�eVertDPred
    
                        unI = 0
                        unJ = 0
                        unLastDebVertD�j�Stock� = False
                        'projection du d�but et fin de vert du carrefour
                        'pr�c�dent sur la droite Y = Y du carrefour courant
                        unDebVertDPred = unDebVertDPred + unTmpInterCarf
                        unFinVertDPred = unFinVertDPred + unTmpInterCarf
                        'Initialisation du LastDebVert au cas o� aucun
                        'bande inter-carf trouv�e
                        unLastDebVertD = unDebVertD
                    
                        Do
                            'La boucle sert pour afficher toutes les bandes
                            'inter-carrefour pour cela on regarde dans
                            'le cycle en cours et le suivant et on prend la
                            'bande inter-carf maximale
                            
                            'On prend le minimun des fins de vert projet�
                            'sur la droite Y = Y du carrefour courant
                            If (unFinVertDPred + unJ * .maDur�eDeCycle) < (unFinVertD + unI * .maDur�eDeCycle) Then
                                unMinFinVert = unFinVertDPred + unJ * .maDur�eDeCycle
                            Else
                                unMinFinVert = unFinVertD + unI * .maDur�eDeCycle
                            End If
                            
                            'On prend le maximun des d�buts de vert projet�
                            'sur la droite Y = Y du carrefour courant
                            If (unDebVertDPred + unJ * .maDur�eDeCycle) > (unDebVertD + unI * .maDur�eDeCycle) Then
                                unMaxDebVert = unDebVertDPred + unJ * .maDur�eDeCycle
                            Else
                                unMaxDebVert = unDebVertD + unI * .maDur�eDeCycle
                            End If
                            
                            'Test de l'existence d'une bande inter-carrefour
                            'sup�rieure � 1 seconde
                            uneBandeInterCarfExist = (unMinFinVert > unMaxDebVert + 1)
                            
                            If uneBandeInterCarfExist Then
                                'If unTD2 - 0.01 < unMinFinVert And unLastDebVertD�j�Stock� = False Then
                                If unLastDebVertD�j�Stock� = False Then
                                    'Stockage du dernier debvert ayant une bande
                                    'inter-carrefour, ce stochage est fait une fois
                                    'et une seule entre deux carrefours
                                    unLastDebVertD�j�Stock� = True
                                    unLastDebVertD = unDebVertD + unI * .maDur�eDeCycle
                                End If
                                'Remise dans l'englobant total si
                                'MinFinVert en sort
                                If unMinFinVert > unMaxT + 0.01 Then
                                    unMinFinVert = unMinFinVert - .maDur�eDeCycle
                                    unMaxDebVert = unMaxDebVert - .maDur�eDeCycle
                                End If
                            End If
                            
                            unI = unI + 1
                            If unI = 2 Then
                                'On se place pour essayer les d�but et fin de vert
                                'du carrefour courant dans le cycle courant avec les
                                'd�but et fin de vert du carrefour pr�c�dent dans le
                                'cycle suivant
                                unI = 0
                                unJ = 1
                            End If
                                                                                
                            'Dessin de bande inter-carrefour
                            If uneBandeInterCarfExist Then
                                'Cas du dessin des bandes inter-carrefours
                                'voitures d'une onde TC
                                
                                'Conversion en coordonn�es �cran
                                unX1 = ConvertirSingleEnEcran(unMaxDebVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMaxDebVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 1�re partie bande descendante inter-carrefours
                                uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                                
                                'Conversion en coordonn�es �cran
                                unX1 = ConvertirSingleEnEcran(unMinFinVert, unT, uneLg)
                                unX1 = unX1 + unNewX0
                                unX2 = ConvertirSingleEnEcran(unMinFinVert - unTmpInterCarf, unT, uneLg)
                                unX2 = unX2 + unNewX0
                                'Dessin 2�me partie bande descendante inter-carrefours
                                uneZoneDessin.Line (unX2, unYDpred)-(unX1, unY), .mesOptionsAffImp.maCoulBandInterCarfD
                            End If
                            'Boucle fait trois fois pour trouver toutes
                            'les bande inter-carrefour
                        Loop Until unI = 1 And unJ = 1
                            
                    End If 'Fin du dessin de bande verte descendante inter-carrefours
                                                                                                 
                    'Stockage du d�but de vert pr�c�dent
                    If unNbCarf + 1 - i = unIndCarfD Then
                        'Calcul sp�cial pour le carrefour le + haut descendant
                        unDebVertDPred = unCarfRedD1.monCarrefour.monDecModif + unCarfRedD1.DonnerPosRefSens(False)
                        unDebVertDPred = ModuloZeroCycle(unDebVertDPred, .maDur�eDeCycle)
                        If unTD1 < unDebVertDPred - 0.001 Then
                            'D�but de vert > T de d�part onde descendante
                            '==> Recul d'un cycle
                            unDebVertDPred = unDebVertDPred - .maDur�eDeCycle
                        ElseIf unTD1 > unDebVertDPred + unCarfRedD1.DonnerDureeVertSens(False) + 0.001 Then
                            'Fin de vert < T de d�part onde descendante
                            '==> Avanc� d'un cycle
                            unDebVertDPred = unDebVertDPred + .maDur�eDeCycle
                        End If
                        'Initialisation de la dur�e de vert du premier
                        'carrefour descendant
                        uneDur�eVertDPred = unCarfRedD1.DonnerDureeVertSens(False)
                    Else
                        'Affectation du d�but de vert du carf r�duit pr�c�dent
                        unDebVertDPred = unLastDebVertD
                        'Affectation de la dur�e de vert du carf r�duit pr�c�dent
                        '� faire car modifs possibles dans RecalculerDebEtFinVert
                        uneDur�eVertDPred = unFinVertD - unDebVertD
                    End If
                    
                    'Stockage de l'indice de ce carrefour
                    unIndCarfDPred = unNbCarf + 1 - i
                    'Stockage du Y �cran  du point pr�c�dent pour le coup suivant
                    unYDpred = unY
                End If
            End If
        Next i
    End With
    
    'Restauration du style de dessin pr�c�dent cette fonction
    uneZoneDessin.DrawStyle = unDrawStyleSave
End Sub

Public Sub RecalculerDebEtFinVert(unCarfRed As Object, unSensMontant As Boolean, uneDur�eDeCycle As Integer, uneVitesse As Single, unDebVert As Single, unFinVert As Single)
    'Recalcul des d�but et fin de vert du carrefour r�duit pass� en
    'param�tre car la r�duction en onde TC est diff�rente de celle
    '� vitesse variable ou constante
    'Les nouveaux d�but et fin de vert sont retourn�s et modifi�s dans les
    'variables unDebVert et unFinVert. De plus leurs valeurs pass�es en
    'param�tre servent � initialiser pour les recherches de min et de max
    'd�crite ci-dessous.
    
    'On garde le max des d�but de vert projet� sur le feu le plus haut en
    'montant (le plus bas en descendant) et du d�but de vert du carf r�duit
    
    'On garde le min des fin de vert projet� sur le feu le plus haut en
    'montant (le plus bas en descendant) et du fin de vert du carf r�duit
    
    Dim unYCarfRed As Integer, unYFeu As Integer
    Dim unCarf As Carrefour, unNbFeux As Integer
    Dim unDebVertFeu As Single, unFinVertFeu As Single
    Dim unMaxDebVert As Single, unMinFinVert As Single
    Dim unFeu As Feu
    
    'Le Y du feu le plus haut en montant ou le plus bas en descendant est
    'celui du carrefour r�duit
    unYCarfRed = unCarfRed.DonnerYSens(unSensMontant)
    
    'Le d�but de vert et du fin de vert � trouver
    'sont initialis�s avec ceux du carrefour r�duit
    'Ce sont ceux pass� en param�tre
    unMaxDebVert = unDebVert
    unMinFinVert = unFinVert
    
    'Recherche du max d�but de vert et du min fin de vert
    Set unCarf = unCarfRed.monCarrefour
    unNbFeux = unCarf.mesFeux.Count
    For i = 1 To unNbFeux
        Set unFeu = unCarf.mesFeux(i)
        If unFeu.monSensMontant = unSensMontant Then
            'Calcul du d�but de vert ramen� entre 0 et cycle
            unDebVertFeu = ModuloZeroCycle(unCarf.monDecModif + unFeu.maPositionPointRef, uneDur�eDeCycle)
            'On le ram�ne dans le cycle du DebVert (= le min actuel)
            If unDebVert - unDebVertFeu >= 0 Then
                unePrec = 0.001
            Else
                unePrec = -0.001
            End If
            unNbCycle = Fix((unDebVert - unDebVertFeu + unePrec) / uneDur�eDeCycle)
            unDebVertFeu = unDebVertFeu + unNbCycle * uneDur�eDeCycle
            'Calcul du d�but de vert du feu projet� sur le Y du carrefour r�duit
            unDebVertFeu = unDebVertFeu + (unYCarfRed - unFeu.monOrdonn�e) / uneVitesse
            'En sens M , Diff des Y > 0 et Vitesse M > 0 ==> tout > 0 donc OK
            'En sens D , Diff des Y < 0 et Vitesse D < 0 ==> tout > 0 donc OK
            If unDebVertFeu > unMaxDebVert Then unMaxDebVert = unDebVertFeu
            
            'Calcul du fin de vert du feu projet� sur le Y du carrefour r�duit
            unFinVertFeu = unDebVertFeu + unFeu.maDur�eDeVert
            If unFinVertFeu < unMinFinVert Then unMinFinVert = unFinVertFeu
        End If
    Next i
    
    'Retour des nouveaux d�buts et fin de vert
    unDebVert = unMaxDebVert
    unFinVert = unMinFinVert
End Sub
