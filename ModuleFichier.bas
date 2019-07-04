Attribute VB_Name = "ModuleFichier"

Public Sub EcrireDansFichier(unNomFich As String, uneForm As Form)
    'Ecriture dans le fichier unNomFich du contenu du site uneForm
    Dim unCarf As Carrefour
    Dim unFeu As Feu
    Dim unTC As TC
    Dim unArret As ArretTC
    
    'V�rification de la validit� de la protection
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurEcriture
    
    ' Fermeture du fichier pour d�lock� et ainsi pouvoir �crire dedans.
    If uneForm.monFichId <> 0 Then
        'Cas d'un Site qui n'est pas Sans Nom unNum�ro
        unFichId = uneForm.monFichId
        Close #unFichId
    End If
    
    'Mise � jour des variables donnant l'�tat de modif du fichier
    'Pour ne pas le sauvegarder encore si on a d�j� fait un save
    With uneForm
        If .maModifDataOndeTC Or .maModifDataOnde Or .maModifDataCarf Then
            'Etat incoherent entre donn�es et r�sultats du calcul
            'Etat incoherent permet de relancer un calcul d'onde
            '(cf CalculerOndeVerte)
            uneForm.maCoherenceDataCalc = IncoherenceDonneeCalcul
        End If
        uneForm.InitIndiqModif 'Remise de tous � false
    End With
    
    'Ouvre le fichier en �criture.
    unFichId = FreeFile(0)
    uneForm.monFichId = unFichId
    Open unNomFich For Output As #unFichId
    
    'Remplissage du fichier � partir des donn�es du site (=uneForm)
    '(cf Format de fichier OndeV .tal)
    With uneForm
        'Ecriture des 6 premi�res lignes ==> Donn�s globales du site
        Write #unFichId, "Fichier Talon 3.0"
        Write #unFichId, .monTitreEtude
        Write #unFichId, .maDur�eDeCycle, .monYMinFeu, .monYMaxFeu, .maCoherenceDataCalc
        Write #unFichId, .monTypeOnde, .monPoidsSensM, .monPoidsSensD, .monTCM, .monTCD, .maBandeTCM, .maBandeTCD, .monOndeDoubleTrouve
        Write #unFichId, .monTypeVit, .maVitSensM, .maVitSensD
        Write #unFichId, .maTransDec, .maBandeM, .maBandeD, .maBandeModifM, .maBandeModifD
            
        'Remplissage des donn�es carrefours
        For i = 1 To .mesCarrefours.Count
            Set unCarf = .mesCarrefours(i)
            Write #unFichId, "Carrefour", unCarf.monNom, unCarf.maVitSensM, unCarf.maVitSensD, unCarf.monIsUtil, unCarf.maDemandeM, unCarf.monDebSatM, unCarf.maDemandeD, unCarf.monDebSatD, unCarf.monDecCalcul, unCarf.monDecModif, unCarf.maVitTCSensM, unCarf.maVitTCSensD, unCarf.monDecImp
                        
            'Remplissage des donn�es des feux du carrefour
            For j = 1 To unCarf.mesFeux.Count
                Set unFeu = unCarf.mesFeux(j)
                Write #unFichId, "Feu", unFeu.monSensMontant, unFeu.monOrdonn�e, unFeu.maDur�eDeVert, -unFeu.maPositionPointRef
                '- pour la position de r�f�rence car en interne elle est invers�e par rapport � la saisie
            Next j
        Next i
            
        'Remplisage des donn�es TC
        For i = 1 To .mesTC.Count
            Set unTC = .mesTC(i)
            Write #unFichId, "TC", unTC.monNom, unTC.monTDep, unTC.maDistAccFrein, unTC.maDureeAccFrein, unTC.monCarfDep.maPosition, unTC.monCarfArr.maPosition, unTC.maCouleur
                        
            'Remplissage des donn�es des arr�ts du TC
            For j = 1 To unTC.mesArrets.Count
                Set unArret = unTC.mesArrets(j)
                Write #unFichId, "Arret", unArret.monOrdonnee, unArret.monTempsArret, unArret.maVitesseMarche, unArret.monLibelle
            Next j
        Next i
    End With
    
    'Mise � jour du titre de la fenetre site courante
    uneForm.Caption = "Site : " + unNomFich
    
    'Fermeture du fichier.
    Close #unFichId
        
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    ' Quitte pour �viter le gestionnaire d'erreur.
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurEcriture:
    
    Select Case Err.Number
        Case 55 'Erreur "Ce fichier est d�j� ouvert".
            MsgBox "Le fichier " + unFich + " est d�j� ouvert", vbCritical
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'fermeture et Sortie du menu Ouvrir
    Close #unFichId
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    Exit Sub
End Sub
