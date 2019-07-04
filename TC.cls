VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmonTDep As Integer 'copie locale
Private mvarmonCarfDep As Carrefour 'copie locale
Private mvarmonCarfArr As Carrefour 'copie locale
Private mvarmaCouleur As Long 'copie locale
Private mvarmesArrets As ColArretTC 'copie locale
Private mvarmonNom As String 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmesObjGraphics As Collection 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaDistAccFrein As Integer 'copie locale
Private mvarmaDureeAccFrein As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmesPhasesTMOnde As ColPhaseTM 'Tableau de marche cadrant l'onde
'variables locales de stockage des valeurs de propriétés
Private mvarmesPhasesTMProg As ColPhaseTM 'Tableau de marche pour tracer la progression
'variables locales de stockage des valeurs de propriétés
Private mvarmonTempsArretFeux As Integer 'copie locale
Private mvarmonNbArretsFeux As Integer 'copie locale
Public Property Let monNbArretsFeux(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNbArretsFeux = 5
    mvarmonNbArretsFeux = vData
End Property


Public Property Get monNbArretsFeux() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNbArretsFeux
    monNbArretsFeux = mvarmonNbArretsFeux
End Property



Public Property Let monTempsArretFeux(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTempsArretFeux = 5
    mvarmonTempsArretFeux = vData
End Property


Public Property Get monTempsArretFeux() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTempsArretFeux
    monTempsArretFeux = mvarmonTempsArretFeux
End Property




Public Property Set mesPhasesTMProg(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesPhasesTMProg = Form1
    Set mvarmesPhasesTMProg = vData
End Property


Public Property Get mesPhasesTMProg() As ColPhaseTM
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesPhasesTMProg
    Set mesPhasesTMProg = mvarmesPhasesTMProg
End Property

Public Function CalculerDecalCauseProgTC(uneColPhases As ColPhaseTM, unY As Single, unSens As Integer) As Single
    'Fonction retournant le décalage en temps du à une progression TC
    'pour rester cohérent avec le calcul du décalage du aux vitesses des
    'véhicules, les décalages en temps qui correspond aux temps de parcours
    'entre les carrefours sont affectés par ordre croissant par rapport aux
    'Y des carrefours.
    'Les carrefours les plus haut auront les décalages en temps les plus
    'élévés. Ainsi l'algorithme de dessin et cadrage maxi de l'onde verte
    'ne change pas
    Dim unIndPhase As Integer, unNbPhases As Integer
        
    If unSens = 1 Then
        'Cas d'une onde cadrée par un TC montant
        CalculerDecalCauseProgTC = CalculerDateDansTabMarche(uneColPhases, unY, unIndPhase, 1) - uneColPhases(1).monTDeb
    ElseIf unSens = -1 Then
        'Cas d'une onde cadrée par un TC descendant
        'Inversion des signes des Y et décalages en temps translatés
        'pour avoir les plus grands dans les carrefours les plus haut d'où
        'DécTemps = date fin dernière phase TC - date dans phases TC en -unY
        unNbPhases = uneColPhases.Count
        CalculerDecalCauseProgTC = uneColPhases(unNbPhases).monTDeb + uneColPhases(unNbPhases).maDureePhase - CalculerDateDansTabMarche(uneColPhases, -unY, unIndPhase, 1)
    Else
        MsgBox "ERREUR de programmation dans OndeV dans CalculerDecalCauseProgTC", vbCritical
    End If
End Function


Public Function CalculerDateDansTabMarche(uneColPhases As ColPhaseTM, unY As Single, unIndPhase As Integer, unIndPhaseDep As Integer) As Single
    'Fonction retournant une date de passage correspondant
    'à une ordonnée passée en paramètre dans le tableau de
    'marche du TC, liste de phase passées en paramètres
    'La phase contenant le Y est retourné par son indice dans la liste
    'des phases de progression du TC dans unIndPhase
    'La liste des phases est parcouru à partir de l'indice unIndPhaseDep
    'jusqu'à trouver la phase par souci d'optimisation
    Dim unYDeb As Single
    Dim unePhase As PhaseTabMarche
    
    'Recherche de la phase où est situé unY
    'avec une précision de calcul de 0.001 mètres
    'Il y a toujours au moins une phase dans un tableau de marche de TC
    i = unIndPhaseDep - 1
    Do
        i = i + 1
        Set unePhase = uneColPhases(i)
        unIndPhase = i
        If unePhase.monType = Arret Then
            'On passe les phases d'arrêt aux arrêts TC
            i = i + 1
            Set unePhase = uneColPhases(i)
            unIndPhase = i
        End If
        unYDeb = unePhase.monYDeb
    '===> Début correction de la version 1.0.1 par rapport à la 1.0
    'Loop Until (unYDeb - 0.001 < unY) And (unY < unYDeb + unePhase.maLongPhase + 0.001)
    Loop Until ((unYDeb - 0.001 < unY) And (unY < unYDeb + unePhase.maLongPhase + 0.001)) Or i = uneColPhases.Count
    If i = uneColPhases.Count And unYDeb - 0.001 > unY Then
        'Cas où le Y est avant le début du tableau de marche
        '==> On sort de la plage de la collection uneColPhase, d'où plantage
        'On prend la première phase pour calculer la date
        Set unePhase = uneColPhases(1)
    End If
    '===> Fin correction de la version 1.0.1 par rapport à la 1.0
    
    'Calcul de la date de passage à Y = unY dans la phase trouvée
    CalculerDateDansTabMarche = CalculerDateDansPhase(unePhase, unY)
End Function

Public Function DonnerVitesseDansTabMarche(uneColPhases As ColPhaseTM, unY As Single) As Single
    'Fonction retournant la vitesse de la phase contenant
    'l'ordonnée passée en paramètre dans le tableau de
    'marche du TC, liste de phases passées en paramètres
    Dim unYDeb As Single
    Dim unePhase As PhaseTabMarche
    
    'Valeur de retour si unY n'est pas dans une phase du TC
    DonnerVitesseDansTabMarche = 99999
    
    'Recherche de la phase où est situé unY
    'avec une précision de calcul de 0.001 mètres
    'Il y a toujours au moins une phase dans un tableau de marche de TC
    unNbPhases = uneColPhases.Count
    i = 0
    Do
        i = i + 1
        Set unePhase = uneColPhases(i)
        unIndPhase = i
        If unePhase.monType = Arret Then
            'On passe les phases d'arrêt aux arrêts TC
            i = i + 1
            Set unePhase = uneColPhases(i)
            unIndPhase = i
        End If
        unYDeb = unePhase.monYDeb
    Loop Until ((unYDeb - 0.001 < unY) And (unY < unYDeb + unePhase.maLongPhase + 0.001)) Or i = unNbPhases
    
    'Renvoi de la vitesse de la phase trouvée contenant Y = unY
    DonnerVitesseDansTabMarche = unePhase.maVitPhase
End Function

Public Sub CalculerTableauMarcheSansFeux(uneColPhases As ColPhaseTM, unYDeb As Integer, unYFin As Integer, unSensMontant As Boolean)
    'Calcul et alimentation d'un tableau de marche, donc une liste
    'de phase, passé en paramètre sans tenir compte des feux
    'Elle retourne dans unYDeb, unYFin et unSensMontant les Y de départ
    'et de fin du parcours du TC et le sens du TC
    Dim unTabYArret() As Integer, unIndFeu As Integer
    Dim unTabIndArret() As Integer, uneDistAF As Single
    Dim unY As Integer, unArret As ArretTC
    Dim unePhase As PhaseTabMarche, uneVitPhase As Single
    Dim unType As Integer, unTDeb As Single, unY0 As Single
    Dim uneLongPhase As Single, uneDureePhase As Single
    Dim uneAcc As Single, unNbArretUtil As Integer
    Dim unTDep As Single
    
    'On vide la collection avant de la réalimenter
    For i = 1 To uneColPhases.Count
        uneColPhases.Remove 1
    Next i
    
    'Détermination du sens de parcours du TC et des Y de départ
    'et de fin du parcours du TC
    If DonnerYCarrefour(monCarfDep) >= DonnerYCarrefour(monCarfArr) Then
        'Cas d'un TC descendant
        unSensMontant = False
        unSens = -1
        unYFin = DonnerYMinCarfSens(monCarfArr, False, unIndFeu) * unSens
        unYDeb = DonnerYMaxCarfSens(monCarfDep, False, unIndFeu) * unSens
    Else
        'Cas d'un TC montant
        unSensMontant = True
        unSens = 1
        unYFin = DonnerYMaxCarfSens(monCarfArr, True, unIndFeu)
        unYDeb = DonnerYMinCarfSens(monCarfDep, True, unIndFeu)
    End If
    
    'Calcul du début de vert du feu de départ dans le premier cycle
    'donc InsTant de départ (=monTDep) + début de vert (Décalage + posRef)
    'si le carrefour est utilisé dans le calcul d'onde sinon le TC part à la
    'date unTDep en absolu
    unTDep = monTDep
    If monCarfDep.monDecCalcul <> -99 Then unTDep = monTDep + monCarfDep.monDecModif + monCarfDep.mesFeux(unIndFeu).maPositionPointRef
    unTDep = ModuloZeroCycle(unTDep, monSite.maDuréeDeCycle)
   
    'Création d'un tableau contenant les Y des arrêts du TC compris
    'strictement entre YDeb et YFin.
    'l'indice 0 réservé à YDeb et le dernier indice à YFin
    unNbArret = mesArrets.Count
    ReDim unTabYArret(unNbArret + 1)   '+1 au cas où tous les arrêts
    ReDim unTabIndArret(unNbArret + 1) 'sont strictement entre Deb et Fin
    
    unNbArretUtil = 0
    For i = 1 To unNbArret
        'Récupération des Y des arrêts inversant le signe pour un
        'TC descendant ou en gardant le même pour un TC montant
        unY = mesArrets(i).monOrdonnee * unSens
        If unYDeb < unY And unY < unYFin Then
            unNbArretUtil = unNbArretUtil + 1
            unTabYArret(unNbArretUtil) = unY
            'Stockage des indices des arrêts dans un tableau parallèle
            unTabIndArret(unNbArretUtil) = i
        End If
    Next i
    
    'Réorganisation par ordre croissant du tableau des Y arrêts
    'ainsi que du tableau d'indice pour garder la cohérence entre les 2
    For j = 2 To unNbArretUtil
            uneFinBoucle = False
            unYTmp = unTabYArret(j)
            unIndTmp = unTabIndArret(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                If unTabYArret(i) > unYTmp Then
                    unTabYArret(i + 1) = unTabYArret(i)
                    unTabIndArret(i + 1) = unTabIndArret(i)
                    i = i - 1
                Else
                    'Fin de boucle
                    uneFinBoucle = True
                End If
            Loop
            unTabYArret(i + 1) = unYTmp
            unTabIndArret(i + 1) = unIndTmp
    Next j
    
    'Ajout de la fin et du début
    unTabYArret(0) = unYDeb
    unTabYArret(unNbArretUtil + 1) = unYFin
    
    'Stockage de l'accélération du TC
    uneAcc = DonnerAcc
        
    'Création des phases, avec les hypothèses suivantes :
    '   Durée d'accélération = Durée de décélération/freinage
    '   Distance d'accélération = Distance de décélération/freinage
    If unNbArretUtil = 0 Then
        'Création d'une seule phase en prenant par défaut
        'la vitesse du premier arrêt du TC
        uneVitPhase = mesArrets(1).maVitesseMarche / 3.6
        unY0 = unYDeb
        Set unArret = Nothing 'Prochain arrêt aucun
        Set unePhase = uneColPhases.Add(VConst, unY0, unTDep, CSng(unYFin - unYDeb), (unYFin - unYDeb) / uneVitPhase, uneVitPhase, unArret)
    Else
        'Cas où des arrêts sont entre YDeb et YFin
        i = 0
        Do
            If i = 0 Then
                'Cas de l'ordonnée de départ, Y = unYDeb
                'Affectation du prochain arret
                Set unArret = mesArrets(unTabIndArret(1))
                uneVitPhase = mesArrets(unTabIndArret(1)).maVitesseMarche / 3.6
                uneDistAF = DonnerDistAccFrein(uneVitPhase)
                If unYDeb > unTabYArret(1) - uneDistAF - 0.001 Then
                    'Cas où l'écart entre le départ et le premier arrêt
                    'est < à la distance de décélération
                    '==> Création d'une phase de décélération
                    'Calcul de la vitesse avant décélération
                    uneVitPhase = Sqr(2 * uneAcc * (unTabYArret(1) - unYDeb))
                    unY0 = unYDeb
                    Set unePhase = uneColPhases.Add(Decel, unY0, unTDep, DonnerDistAccFrein(uneVitPhase), DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                Else
                    'Cas où la phase de décélération peut exister entiérement
                    '==> Création de deux phases : une à V const et une décel
                    uneLongPhase = unTabYArret(1) - uneDistAF - unYDeb
                    uneDureePhase = uneLongPhase / uneVitPhase
                    unY0 = unYDeb
                    Set unePhase = uneColPhases.Add(VConst, unY0, unTDep, uneLongPhase, uneDureePhase, uneVitPhase, unArret)
                    Set unePhase = uneColPhases.Add(Decel, unYDeb + uneLongPhase, unePhase.monTDeb + uneDureePhase, uneDistAF, DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                End If
                'Ajout de la phase de stop à l'arrêt suivant
                unY0 = unTabYArret(i + 1)
                uneDureePhase = mesArrets(unTabIndArret(i + 1)).monTempsArret
                Set unePhase = uneColPhases.Add(Arret, unY0, unePhase.monTDeb + unePhase.maDureePhase, 0, uneDureePhase, 0, unArret)
            ElseIf i = unNbArretUtil Then
                'Cas de l'ordonnée d'arrivée, Y = unYFin
                'Affectation du prochain arret, nothing car aprés c'est YFin
                Set unArret = Nothing
                uneVitPhase = mesArrets(unTabIndArret(i)).maVitesseMarche / 3.6
                uneDistAF = DonnerDistAccFrein(uneVitPhase)
                If unTabYArret(i) + uneDistAF > unTabYArret(i + 1) - 0.001 Then
                    'Cas où l'écart entre le dernier arrêt et la fin
                    'est > à la distance d'accélération
                    '==> Création d'une phase d'accélération jusqu'à YFin
                    'Calcul de la vitesse obtenue par l'accélération
                    uneVitPhase = Sqr(2 * uneAcc * (unYFin - unTabYArret(i)))
                    unY0 = unTabYArret(i)
                    Set unePhase = uneColPhases.Add(Accel, unY0, unePhase.monTDeb + unePhase.maDureePhase, DonnerDistAccFrein(uneVitPhase), DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                Else
                    'Cas où la phase d'accélération peut exister entiérement
                    '==> Création de deux phases : une accél et une à V const
                    unY0 = unTabYArret(i)
                    Set unePhase = uneColPhases.Add(Accel, unY0, unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                    uneLongPhase = unYFin - unTabYArret(i) - uneDistAF
                    uneDureePhase = uneLongPhase / uneVitPhase
                    Set unePhase = uneColPhases.Add(VConst, unYFin - uneLongPhase, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneDureePhase, uneVitPhase, unArret)
                End If
            Else
                'Cas des arrêts, unYDeb < Y < unYFin
                'Affectation du prochain arret
                Set unArret = mesArrets(unTabIndArret(i + 1))
                uneVitPhase = mesArrets(unTabIndArret(i + 1)).maVitesseMarche / 3.6
                uneDistAF = DonnerDistAccFrein(uneVitPhase)
                If unTabYArret(i) + uneDistAF > unTabYArret(i + 1) - uneDistAF - 0.001 Then
                    'Cas où le prochain arrêt est dans la zone de décélération
                    '==> Création de deux phases : une accél et une décel
                    uneVitPhase = Sqr(uneAcc * (unTabYArret(i + 1) - unTabYArret(i)))
                    uneLongPhase = DonnerDistAccFrein(uneVitPhase)
                    uneDureePhase = DonnerDureeAccFrein(uneVitPhase)
                    unY0 = unTabYArret(i)
                    Set unePhase = uneColPhases.Add(Accel, unY0, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneDureePhase, uneVitPhase, unArret)
                    Set unePhase = uneColPhases.Add(Decel, unY0 + uneLongPhase, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneDureePhase, uneVitPhase, unArret)
                Else
                    'Cas où le prochain arrêt n'est pas dans la zone de décélération
                    '==> Création de 3 phases : une accél, une à V const et une décel
                    unY0 = unTabYArret(i)
                    uneDureePhase = DonnerDureeAccFrein(uneVitPhase)
                    Set unePhase = uneColPhases.Add(Accel, unY0, unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, uneDureePhase, uneVitPhase, unArret)
                    uneLongPhase = unTabYArret(i + 1) - unTabYArret(i) - 2 * uneDistAF
                    Set unePhase = uneColPhases.Add(VConst, unTabYArret(i) + uneDistAF, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneLongPhase / uneVitPhase, uneVitPhase, unArret)
                    Set unePhase = uneColPhases.Add(Decel, unTabYArret(i + 1) - uneDistAF, unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, uneDureePhase, uneVitPhase, unArret)
                End If
                'Ajout de la phase de stop à l'arrêt
                unY0 = unTabYArret(i + 1)
                uneDureePhase = mesArrets(unTabIndArret(i + 1)).monTempsArret
                Set unePhase = uneColPhases.Add(Arret, unY0, unePhase.monTDeb + unePhase.maDureePhase, 0, uneDureePhase, 0, unArret)
            End If
            
            'Incrémentation suivante
            i = i + 1
        Loop While i < unNbArretUtil + 1
    End If
End Sub

Public Function CalculerTableauMarcheProg() As Integer
    'Calcul et alimentation du tableau de marche, donc une liste
    'de phase, servant à tracer la progression du TC
    
    'Le sens du TC est retourné par cette fonction :
    '       1 pour montant
    '       -1 pour descendant
    
    Dim unTabFeu() As Feu, unY As Integer, unYNextArret As Long
    Dim unCarf As Carrefour, unFeu As Feu
    Dim unSens As Integer, unFeuSurArret As Boolean
    Dim unArret As ArretTC, unArretTmp As ArretTC
    Dim unSensMontant As Boolean, unNbFeux As Integer
    Dim unYDeb As Integer, unYFin As Integer
    Dim unePhase As PhaseTabMarche
    Dim unIndPhase As Integer, unIndPhaseDep As Integer
    Dim uneDatePassage As Single, unTNextDebVert  As Single
    Dim uneDistAF As Single, uneDureeAF As Single
    Dim uneVitPhase As Single, uneLongPhase As Single
    Dim uneDureePhase As Single
    
    'Calcul du tableau de marche de progression sans s'occuper des feux et
    'des Y de départ et de fin du parcours du sens du TC
    'Les Y sont de signe inversé si le TC est descendant
    CalculerTableauMarcheSansFeux mesPhasesTMProg, unYDeb, unYFin, unSensMontant

    'Recherche et stockage des feux de même sens que le TC entre les
    'carrefours départ et arrivée dans un tableau dynamique qui
    'contiendra le feu
    unNbFeux = 0
    If unSensMontant Then
        unSens = 1
    Else
        unSens = -1
    End If
    
    'Retour du sens du TC
    CalculerTableauMarcheProg = unSens
    
    For i = 1 To monSite.mesCarrefours.Count
        Set unCarf = monSite.mesCarrefours(i)
        If unCarf.monDecCalcul <> -99 Then
            For j = 1 To unCarf.mesFeux.Count
                Set unFeu = unCarf.mesFeux(j)
                If unFeu.monSensMontant = unSensMontant Then
                    'Inversion du signe des Y si le TC est descendant
                    unY = unFeu.monOrdonnée * unSens
                    If unYDeb <= unY And unY <= unYFin Then
                        unNbFeux = unNbFeux + 1
                        ReDim Preserve unTabFeu(1 To unNbFeux)
                        Set unTabFeu(unNbFeux) = unFeu
                    End If
                End If
            Next j
        End If
    Next i
    
    'Tri par ordre croissant des Y des feux
    For j = 2 To unNbFeux
            uneFinBoucle = False
            Set unFeu = unTabFeu(j)
            i = j - 1
            Do While i > 0 And uneFinBoucle = False
                'Inversion du signe des Y si le TC est descendant
                If unTabFeu(i).monOrdonnée * unSens > unFeu.monOrdonnée * unSens Then
                    Set unTabFeu(i + 1) = unTabFeu(i)
                    i = i - 1
                Else
                    'Fin de boucle
                    uneFinBoucle = True
                End If
            Loop
            Set unTabFeu(i + 1) = unFeu
    Next j
    
    'Initialisation du temps et du nombre d'arrêts aux feux
    'lors de la progression du TC
    monNbArretsFeux = 0
    monTempsArretFeux = 0
    
    'Initialisation de l'indice de départ pour le parcours des phases
    'du TC dans lors de la fonction CalculerDateDansTabMarche
    unIndPhaseDep = 1
    
    'Modifications éventuelles du tableau de marche de progression
    'à cause de la prise en compte des feux rencontrés.
    'Modif si le TC arrive au feu rouge, si vert pas de changement
    For i = 1 To unNbFeux
        Set unFeu = unTabFeu(i)
        'Inversion des signes des Y si TC en sens descendant, sinon rien
        unY = unFeu.monOrdonnée * unSens
        
        'Calcul de la date de passage du TC au feu d'ordonnée unY, ainsi que
        'l'indice de la phase dans la liste des phases de progression TC
        uneDatePassage = CalculerDateDansTabMarche(mesPhasesTMProg, CSng(unY), unIndPhase, unIndPhaseDep)
        
        'Stockage de l'indice de départ pour le parcours des phases du TC
        'dans lors du i suivant qui fera le prochain appel à la fonction
        'CalculerDateDansTabMarche
        unIndPhaseDep = unIndPhase
        
        'Test de l'état du feu, s'il est rouge
        '==> Modif du tableau de marche, sinon rien
        If unFeu.IsVert(uneDatePassage) = False Then
            'Récupération de la phase
            Set unePhase = mesPhasesTMProg(unIndPhase)

            'Calcul de la date de départ du TC après le feu rouge
            'donc du prochain début de vert du feu
            K = Int((uneDatePassage - unFeu.monCarrefour.monDecModif - unFeu.maPositionPointRef) / monSite.maDuréeDeCycle)
            unTNextDebVert = unFeu.monCarrefour.monDecModif + unFeu.maPositionPointRef + (K + 1) * monSite.maDuréeDeCycle
            
            'Suppression de toutes les phases suivantes de la phase
            'contenant le Y du feu qui ont le même prochain arrêt
            'sauf si c'est une phase d'arrêt (Longueur de ces phases = 0)
            Set unArret = unePhase.monProchainArret 'Recup du prochain arrêt
            For j = unIndPhase + 1 To mesPhasesTMProg.Count
                'Récupération du prochain arrêt, on prend toujours
                'la phase d'indice unIndphase + 1 comme on l'explique pour
                'le remove ci-dessous
                Set unArretTmp = mesPhasesTMProg(unIndPhase + 1).monProchainArret
                If unArret Is unArretTmp And mesPhasesTMProg(unIndPhase + 1).maLongPhase > 0 Then
                    'On enléve toujours le même car le remove d'une collection
                    'VB réorganise en décrémentant la collection vers le bas
                    'ex : le j+1 ème devient le j ème après un remove de j
                    mesPhasesTMProg.Remove unIndPhase + 1
                Else
                    Exit For
                End If
            Next j
            
            'Cas où le feu a le même Y qu'un arrêt
            unFeuSurArret = False
            unYNextArret = unYFin
            If Not unePhase.monProchainArret Is Nothing Then
                'Cas où le prochain arrêt n'est pas l'arrêt nul
                'celui qui suit le feu du carrefour d'arrivée ( = unYFin)
                unYNextArret = unePhase.monProchainArret.monOrdonnee * unSens
                'Inversion des signes des Y si TC en sens descendant, sinon rien
                If unY = unYNextArret Then
                    unFeuSurArret = True
                    If unTNextDebVert - uneDatePassage < unePhase.monProchainArret.monTempsArret - 0.001 Then
                        'Cas où le prochain début de vert est < au temps d'arrêt
                        'Le TC ne peut pas redémarrer tout de suite, il doit
                        'attendre uneDatePassage + Temps de l'arrêt si à cette
                        'date le feu est vert, sinon il démarre au prochain début
                        'de vert du cycle suivant (unTNextDebVert + unCycle)
                        If unFeu.IsVert(uneDatePassage + unePhase.monProchainArret.monTempsArret) Then
                            unTNextDebVert = uneDatePassage + unePhase.monProchainArret.monTempsArret
                            'Mise à jour du nombre d'arrêts aux feux, on lui
                            'enlève cet arrêt car c'est l'arêt qui impose le
                            'temps d'arrêt et pas la phase de rouge
                            monNbArretsFeux = monNbArretsFeux - 1
                        Else
                            unTNextDebVert = unTNextDebVert + monSite.maDuréeDeCycle
                        End If
                    End If
                    'Mise à jour du temps d'arrêt au feu, on lui enlève
                    'le temps d'arrêt à l'arrêt TC
                    monTempsArretFeux = monTempsArretFeux - unePhase.monProchainArret.monTempsArret
                End If
            End If
            
            'Modif de la phase contenant le Y du feu
            unePhase.maDureePhase = uneDatePassage - unePhase.monTDeb
            unePhase.maLongPhase = unY - unePhase.monYDeb
            
            'Ajout de la phase d'arrêt au feu rouge si le feu n'a pas le
            'même Y qu'un arrêt, sinon on modifie la phase d'arrêt qui n'a
            'pas été supprimée plus haut
            If unFeuSurArret Then
                mesPhasesTMProg(unIndPhase + 1).maDureePhase = unTNextDebVert - uneDatePassage
                mesPhasesTMProg(unIndPhase + 1).monTDeb = unePhase.monTDeb + unePhase.maDureePhase
            Else
                Set unePhase = mesPhasesTMProg.Add(Arret, CSng(unY), unePhase.monTDeb + unePhase.maDureePhase, 0, unTNextDebVert - uneDatePassage, 0, unePhase.monProchainArret, unIndPhase)
                'Nombre de nouvelles phases ajoutées dans la collection
                unNbNewPhase = 1
            End If
            
            'Mise à jour du temps et du nombre d'arrêts aux feux
            'lors de la progression du TC
            monNbArretsFeux = monNbArretsFeux + 1
            monTempsArretFeux = monTempsArretFeux + unTNextDebVert - uneDatePassage
            
            'Modification du tableau de marche de progression
            If i < unNbFeux And unFeuSurArret = False Then
                'Cas de tous les feux sauf le dernier, celui du feu d'arrivée
                'et ne se trouvant pas sur un arrêt
                
                'Ajout des phases dues à l'existence d'un feu rouge
                uneVitPhase = mesPhasesTMProg(unIndPhase).maVitPhase
                uneDistAF = DonnerDistAccFrein(uneVitPhase)
                Set unArret = mesPhasesTMProg(unIndPhase).monProchainArret
                If unY + uneDistAF > unYNextArret - uneDistAF - 0.001 Then
                    'Cas où le prochain arrêt est dans la zone de décélération
                    If unArret Is Nothing Then
                        'Cas où le prochain arrêt est le dernier feu du
                        'carrefour d'arrivée
                        '==> Création d'une phase d'accélération jusqu'à unYNextArret = Y du dernier feu
                        'Calcul de la vitesse obtenue par l'accélération
                        uneVitPhase = Sqr(2 * DonnerAcc * (unYNextArret - unY))
                        Set unePhase = mesPhasesTMProg.Add(Accel, CSng(unY), unePhase.monTDeb + unePhase.maDureePhase, DonnerDistAccFrein(uneVitPhase), DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                        'Nombre de nouvelles phases ajoutées dans la collection
                        unNbNewPhase = unNbNewPhase + 1
                    Else
                        'Cas où le prochain arrêt en est vraiment un
                        '==> Création de deux phases : une accél et une décel
                        uneVitPhase = Sqr(DonnerAcc * (unYNextArret - unY))
                        uneDistAF = DonnerDistAccFrein(uneVitPhase)
                        uneDureeAF = DonnerDureeAccFrein(uneVitPhase)
                        Set unePhase = mesPhasesTMProg.Add(Accel, CSng(unY), unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, uneDureeAF, uneVitPhase, unArret, unIndPhase + 1)
                        Set unePhase = mesPhasesTMProg.Add(Decel, unY + uneDistAF, unePhase.monTDeb + uneDureeAF, uneDistAF, uneDureeAF, uneVitPhase, unArret, unIndPhase + 2)
                        'Nombre de nouvelles phases ajoutées dans la collection
                        unNbNewPhase = unNbNewPhase + 2
                    End If
                Else
                    'Cas où le prochain arrêt n'est pas dans la zone de décélération
                    If unArret Is Nothing Then
                        'Cas où le prochain arrêt est le dernier feu du
                        'carrefour d'arrivée
                        '==> Création de deux phases : une accél et une à V
                        'const jusqu'à unYNextArret = Y du dernier feu
                        Set unePhase = mesPhasesTMProg.Add(Accel, CSng(unY), unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, DonnerDureeAccFrein(uneVitPhase), uneVitPhase, unArret)
                        uneLongPhase = unYNextArret - unY - uneDistAF
                        uneDureePhase = uneLongPhase / uneVitPhase
                        Set unePhase = mesPhasesTMProg.Add(VConst, unYNextArret - uneLongPhase, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneDureePhase, uneVitPhase, unArret)
                        'Nombre de nouvelles phases ajoutées dans la collection
                        unNbNewPhase = unNbNewPhase + 2
                    Else
                        'Cas où le prochain arrêt en est vraiment un
                        '==> Création de 3 phases : une accél, une à V const et une décel
                        uneDureeAF = DonnerDureeAccFrein(uneVitPhase)
                        Set unePhase = mesPhasesTMProg.Add(Accel, CSng(unY), unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, uneDureeAF, uneVitPhase, unArret, unIndPhase + 1)
                        uneLongPhase = unYNextArret - unY - 2 * uneDistAF
                        Set unePhase = mesPhasesTMProg.Add(VConst, unY + uneDistAF, unePhase.monTDeb + unePhase.maDureePhase, uneLongPhase, uneLongPhase / uneVitPhase, uneVitPhase, unArret, unIndPhase + 2)
                        Set unePhase = mesPhasesTMProg.Add(Decel, unYNextArret - uneDistAF, unePhase.monTDeb + unePhase.maDureePhase, uneDistAF, uneDureeAF, uneVitPhase, unArret, unIndPhase + 3)
                        'Nombre de nouvelles phases ajoutées dans la collection
                        unNbNewPhase = unNbNewPhase + 3
                    End If
                End If
            End If
            
            'Translation des débuts de phases suivant celles du prochain arrêt
            unTransDeb = unePhase.monTDeb + unePhase.maDureePhase
            For j = unIndPhase + unNbNewPhase + 1 To mesPhasesTMProg.Count
                mesPhasesTMProg(j).monTDeb = unTransDeb
                'Modif de unTransDeb pour l'incrémentation suivante
                unTransDeb = unTransDeb + mesPhasesTMProg(j).maDureePhase
            Next j
        End If
    Next i
End Function

Public Sub CalculerTableauMarcheOnde()
    'Calcul et alimentation du tableau de marche, donc une liste
    'de phase, servant à cadrer l'onde verte avec ce TC
    Dim unYDeb As Integer, unYFin As Integer
    Dim unSensMontant As Boolean
    
    'Calcul du tableau de marche de cadrage d'onde, pour celui-ci
    'on ne s'occupe pas des feux,
    'et calcul des Y de départ et de fin du parcours et du sens du TC
    'Les Y sont de signe inversé si le TC est descendant
    CalculerTableauMarcheSansFeux mesPhasesTMOnde, unYDeb, unYFin, unSensMontant
End Sub


Public Function DonnerAcc() As Single
    'Fonction calculant l'accélération = décélération (Hypothèse OndeV)
    DonnerAcc = 2 * maDistAccFrein / maDureeAccFrein / maDureeAccFrein
End Function

Public Function DonnerDureeAccFrein(uneV As Single) As Single
    'Fonction calculant la durée d'accélération = durée de décélération
    'ou freinage de 0 m/s à une vitesse valant uneV m/s
    DonnerDureeAccFrein = uneV * maDureeAccFrein * maDureeAccFrein / 2 / maDistAccFrein
End Function

Public Function DonnerDistAccFrein(uneV As Single) As Single
    'Fonction calculant la distance d'accélération = distance de
    'décélération ou freinage de 0 m/s à une vitesse valant uneV m/s
    DonnerDistAccFrein = uneV * uneV * maDureeAccFrein * maDureeAccFrein / 4 / maDistAccFrein
End Function

Public Property Set mesPhasesTMOnde(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesPhasesTMOnde = Form1
    Set mvarmesPhasesTMOnde = vData
End Property


Public Property Get mesPhasesTMOnde() As ColPhaseTM
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesPhasesTMOnde
    Set mesPhasesTMOnde = mvarmesPhasesTMOnde
End Property



Public Property Let maDureeAccFrein(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDureeAccFrein = 5
    mvarmaDureeAccFrein = vData
End Property


Public Property Get maDureeAccFrein() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDureeAccFrein
    maDureeAccFrein = mvarmaDureeAccFrein
End Property



Public Property Let maDistAccFrein(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDistAccFrein = 5
    mvarmaDistAccFrein = vData
End Property


Public Property Get maDistAccFrein() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDistAccFrein
    maDistAccFrein = mvarmaDistAccFrein
End Property



Public Property Set mesObjGraphics(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesObjGraphics = Form1
    Set mvarmesObjGraphics = vData
End Property


Public Property Get mesObjGraphics() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesObjGraphics
    Set mesObjGraphics = mvarmesObjGraphics
End Property






Public Property Let monNom(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNom = 5
    mvarmonNom = vData
End Property


Public Property Get monNom() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNom
    monNom = mvarmonNom
End Property



Public Property Set mesArrets(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesArrets = Form1
    Set mvarmesArrets = vData
End Property


Public Property Get mesArrets() As ColArretTC
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesArrets
    Set mesArrets = mvarmesArrets
End Property



Public Property Let maCouleur(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCouleur = 5
    mvarmaCouleur = vData
End Property


Public Property Get maCouleur() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCouleur
    maCouleur = mvarmaCouleur
End Property



Public Property Set monCarfArr(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarfArr = Form1
    Set mvarmonCarfArr = vData
End Property


Public Property Get monCarfArr() As Carrefour
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarfArr
    Set monCarfArr = mvarmonCarfArr
End Property



Public Property Set monCarfDep(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarfDep = Form1
    Set mvarmonCarfDep = vData
End Property


Public Property Get monCarfDep() As Carrefour
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarfDep
    Set monCarfDep = mvarmonCarfDep
End Property



Public Property Let monTDep(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTDep = 5
    mvarmonTDep = vData
End Property


Public Property Get monTDep() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTDep
    monTDep = mvarmonTDep
End Property



