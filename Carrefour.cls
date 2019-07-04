VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Carrefour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe représentant un carrefour Talon"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmonNom As String 'copie locale
Private mvarmesFeux As ColFeu 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonCarfGraphic As Object 'copie locale
Private mvarmesFeuxGraphics As Collection 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaPosition As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaVitSensM As Integer 'copie locale
Private mvarmaVitSensD As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonIsUtil As Boolean 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecModif As Single 'copie locale
Private mvarmonDecCalcul As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecVitSensM As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecVitSensD As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaVitTCSensM As Single 'copie locale
Private mvarmaVitTCSensD As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaDemandeM As Long 'copie locale
Private mvarmaDemandeD As Long 'copie locale
Private mvarmonDebSatM As Long 'copie locale
Private mvarmonDebSatD As Long 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonCarfRed As Object 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecImp As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecCalculSave As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDecModifSave As Single 'copie locale
Public Property Let monDecModifSave(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecModifSave = 5
    mvarmonDecModifSave = vData
End Property


Public Property Get monDecModifSave() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecModifSave
    monDecModifSave = mvarmonDecModifSave
End Property



Public Property Let monDecCalculSave(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecCalculSave = 5
    mvarmonDecCalculSave = vData
End Property


Public Property Get monDecCalculSave() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecCalculSave
    monDecCalculSave = mvarmonDecCalculSave
End Property



Public Property Let monDecImp(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecImp = 5
    mvarmonDecImp = vData
End Property


Public Property Get monDecImp() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecImp
    monDecImp = mvarmonDecImp
End Property




Public Property Set monCarfRed(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarfRed = Form1
    Set mvarmonCarfRed = vData
End Property


Public Property Get monCarfRed() As Object
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarfRed
    Set monCarfRed = mvarmonCarfRed
End Property



Public Property Let monDebSatD(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebSatD = 5
    mvarmonDebSatD = vData
End Property


Public Property Get monDebSatD() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebSatD
    monDebSatD = mvarmonDebSatD
End Property



Public Property Let monDebSatM(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebSatM = 5
    mvarmonDebSatM = vData
End Property


Public Property Get monDebSatM() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebSatM
    monDebSatM = mvarmonDebSatM
End Property



Public Property Let maDemandeD(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDemandeD = 5
    mvarmaDemandeD = vData
End Property


Public Property Get maDemandeD() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDemandeD
    maDemandeD = mvarmaDemandeD
End Property



Public Property Let maDemandeM(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDemandeM = 5
    mvarmaDemandeM = vData
End Property


Public Property Get maDemandeM() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDemandeM
    maDemandeM = mvarmaDemandeM
End Property



Public Sub SetDemDeb(uneDemM As Long, unDebSatM As Long, uneDemD As Long, unDebSatD As Long)
    'Affectation des demandes montantes et descendantes du carrefour
    maDemandeM = uneDemM
    maDemandeD = uneDemD
    'Affectation des débits de saturation montant et descendant du carrefour
    monDebSatM = unDebSatM
    monDebSatD = unDebSatD
End Sub

Public Function DonnerVitSens(unSensMontant As Boolean) As Single
    'Calcul de la vitesse en m/s (d'où les divisions par 3.6)
    'du carrefour dans le sens montant si unSensMontant est true,
    'dans le sens descendant sinon, en tenant compte du type d'onde
    'verte, du type de vitesse choisi (cte ou variable) et des TC
    'cadrant l'onde verte en sens montant et/ou descendant
    Dim unTC As TC
    
    If unSensMontant Then
        'Calcul de la vitesse en sens montant
        If monSite.monTCM <> 0 And monSite.monTypeOnde = OndeTC Then
            'Cas d'une onde verte à cadrer par un TC montant
            Set unTC = monSite.mesTC(monSite.monTCM)
            maVitTCSensM = unTC.DonnerVitesseDansTabMarche(unTC.mesPhasesTMOnde, DonnerYCarrefour(Me))
            DonnerVitSens = maVitTCSensM
        ElseIf monSite.monTypeVit = VitVar Then
            'Cas d'une onde verte à cadrer sans TC montant
            'avec des vitesses variables en sens montant
            DonnerVitSens = maVitSensM / 3.6
        ElseIf monSite.monTypeVit = VitConst Then
            'Cas d'une onde verte à cadrer sans TC montant
            'avec une vitesse constante en sens montant
            DonnerVitSens = monSite.maVitSensM / 3.6
        Else
            MsgBox "Erreur de programmation dans OndeV : Méthode DonnerVitSens Classe Carrefour"
        End If
    Else
        'Calcul de la vitesse en sens descendant
        If monSite.monTCD <> 0 And monSite.monTypeOnde = OndeTC Then
            'Cas d'une onde verte à cadrer par un TC descendant
            Set unTC = monSite.mesTC(monSite.monTCD)
            maVitTCSensD = unTC.DonnerVitesseDansTabMarche(unTC.mesPhasesTMOnde, -DonnerYCarrefour(Me))
            DonnerVitSens = maVitTCSensD
        ElseIf monSite.monTypeVit = VitVar Then
            'Cas d'une onde verte à cadrer sans TC descendant
            'avec des vitesses variables en sens descendant
            DonnerVitSens = maVitSensD / 3.6
        ElseIf monSite.monTypeVit = VitConst Then
            'Cas d'une onde verte à cadrer sans TC descendant
            'avec une vitesse constante en sens descendant
            DonnerVitSens = monSite.maVitSensD / 3.6
        Else
            MsgBox "Erreur de programmation dans OndeV : Méthode DonnerVitSens Classe Carrefour"
        End If
        'Transformation en négatif d'une vitesse descendante
        'saisie ou calculée > 0 pour le cas d'un TC montant ou descendant
        DonnerVitSens = -DonnerVitSens
    End If
End Function

Public Property Let maVitTCSensD(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVitTCSensD = 5
    mvarmaVitTCSensD = vData
End Property


Public Property Get maVitTCSensD() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVitTCSensD
    maVitTCSensD = mvarmaVitTCSensD
End Property



Public Property Let maVitTCSensM(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVitTCSensM = 5
    mvarmaVitTCSensM = vData
End Property


Public Property Get maVitTCSensM() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVitTCSensM
    maVitTCSensM = mvarmaVitTCSensM
End Property



Public Property Let monDecVitSensD(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecVitSensD = 5
    mvarmonDecVitSensD = vData
End Property


Public Property Get monDecVitSensD() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecVitSensD
    monDecVitSensD = mvarmonDecVitSensD
End Property



Public Property Let monDecVitSensM(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecVitSensM = 5
    mvarmonDecVitSensM = vData
End Property


Public Property Get monDecVitSensM() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecVitSensM
    monDecVitSensM = mvarmonDecVitSensM
End Property



Public Property Let monDecCalcul(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecCalcul = 5
    mvarmonDecCalcul = vData
End Property


Public Property Get monDecCalcul() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecCalcul
    monDecCalcul = mvarmonDecCalcul
End Property



Public Property Let monDecModif(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDecModif = 5
    mvarmonDecModif = vData
End Property


Public Property Get monDecModif() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDecModif
    monDecModif = mvarmonDecModif
End Property



Public Property Let monIsUtil(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monIsUtil = 5
    mvarmonIsUtil = vData
End Property


Public Property Get monIsUtil() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monIsUtil
    monIsUtil = mvarmonIsUtil
End Property



Public Property Let maVitSensD(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVitSensD = 5
    mvarmaVitSensD = vData
End Property


Public Property Get maVitSensD() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVitSensD
    maVitSensD = mvarmaVitSensD
End Property



Public Property Let maVitSensM(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVitSensM = 5
    mvarmaVitSensM = vData
End Property


Public Property Get maVitSensM() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVitSensM
    maVitSensM = mvarmaVitSensM
End Property



Public Property Let maPosition(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maPosition = 5
    mvarmaPosition = vData
End Property


Public Property Get maPosition() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maPosition
    maPosition = mvarmaPosition
End Property



Public Property Set mesFeuxGraphics(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesFeuxGraphics = Form1
    Set mvarmesFeuxGraphics = vData
End Property


Public Property Get mesFeuxGraphics() As Collection
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesFeuxGraphics
    Set mesFeuxGraphics = mvarmesFeuxGraphics
End Property



Public Property Set monCarfGraphic(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarfGraphic = Form1
    Set mvarmonCarfGraphic = vData
End Property


Public Property Get monCarfGraphic() As Object
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarfGraphic
    Set monCarfGraphic = mvarmonCarfGraphic
End Property




Public Property Set mesFeux(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.mesFeux = Form1
    Set mvarmesFeux = vData
End Property


Public Property Get mesFeux() As ColFeu
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.mesFeux
    Set mesFeux = mvarmesFeux
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


Public Sub DonnerNbFeuxMetD(unNbFeuxM As Integer, unNbFeuxD As Integer)
    'Méthode retournant le nombre de feux montants
    'et descendants du carrefour
    unNbFeuxM = 0
    unNbFeuxD = 0
    For i = 1 To mesFeux.Count
        If mesFeux(i).monSensMontant Then
            unNbFeuxM = unNbFeuxM + 1
        Else
            unNbFeuxD = unNbFeuxD + 1
        End If
    Next i
End Sub



Public Function DonnerVitCarfSens(unSensMontant As Boolean) As Single
    'Calcul de la vitesse en m/s (d'où les divisions par 3.6)
    'du carrefour dans le sens montant si unSensMontant est true,
    'dans le sens descendant sinon, en tenant compte uniquement
    'du type de vitesse choisi (cte ou variable) pour retrouver
    'la vitesse propre du carrefour
    
    'Utiliser pour dessiner les bandes inter-carrefours voitures
    'dans le cas d'une onde TC
    Dim unTC As TC
    
    If unSensMontant Then
        'Calcul de la vitesse en sens montant
        If monSite.monTypeVit = VitVar Then
            'Cas des vitesses variables en sens montant
            DonnerVitCarfSens = maVitSensM / 3.6
        ElseIf monSite.monTypeVit = VitConst Then
            'Cas d'une vitesse constante en sens montant
            DonnerVitCarfSens = monSite.maVitSensM / 3.6
        Else
            MsgBox "Erreur de programmation dans OndeV : TypeVitesse inconnu dans Méthode Carrefour:DonnerVitCarfSens."
        End If
    Else
        'Calcul de la vitesse en sens descendant
        If monSite.monTypeVit = VitVar Then
            'Cas des vitesses variables en sens descendant
            DonnerVitCarfSens = maVitSensD / 3.6
        ElseIf monSite.monTypeVit = VitConst Then
            'Cas d'une vitesse constante en sens descendant
            DonnerVitCarfSens = monSite.maVitSensD / 3.6
        Else
            MsgBox "Erreur de programmation dans OndeV : TypeVitesse inconnu dans Méthode Carrefour:DonnerVitCarfSens."
        End If
        'Transformation en négatif d'une vitesse descendante
        DonnerVitCarfSens = -DonnerVitCarfSens
    End If
End Function
