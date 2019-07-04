VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarfReduitSensDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmonCarrefour As Carrefour 'copie locale
Private mvarmonEcart As Single 'copie locale
Private mvarmaDureeVertM As Single 'copie locale
Private mvarmaDureeVertD As Single 'copie locale
Private mvarmonOrdonneeM As Integer 'copie locale
Private mvarmonOrdonneeD As Integer 'copie locale
Private mvarmaPosRefM As Single 'copie locale
Private mvarmaPosRefD As Single 'copie locale
Private mvarmonDebOndeM As Single 'copie locale
Private mvarmonDebOndeD As Single 'copie locale
Public Property Let monDebOndeD(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebOndeD = 5
    mvarmonDebOndeD = vData
End Property


Public Property Get monDebOndeD() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebOndeD
    monDebOndeD = mvarmonDebOndeD
End Property



Public Property Let monDebOndeM(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebOndeM = 5
    mvarmonDebOndeM = vData
End Property


Public Property Get monDebOndeM() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebOndeM
    monDebOndeM = mvarmonDebOndeM
End Property



Public Sub AffecterDebOndeSens(unDebOnde As Single, unSensMontant As Boolean)
    'Elle affecte le point début de l'onde verte
    'du carrefour réduit correspondant
    'au feu équivalent dans le sens choisi.
    If unSensMontant Then
        monDebOndeM = unDebOnde
    Else
        monDebOndeD = unDebOnde
    End If
End Sub

Public Function DonnerDebOndeSens(unSensMontant As Boolean) As Single
    'Elle retourne le point début de l'onde verte
    'du carrefour réduit correspondant
    'au feu équivalent dans le sens choisi.
    If unSensMontant Then
        DonnerDebOndeSens = monDebOndeM
    Else
        DonnerDebOndeSens = monDebOndeD
    End If
End Function


Public Function DonnerVitSens(unSensMontant As Boolean) As Single
    'Méthode utilisée uniquement par la procédure globale CalculerDecVitesse
    'Elle calcule la vitesse du carrefour réduit correspondant
    'à la vitesse variable, constante ou contrainte par un TC du carrefour.
    
    'Retour : Vitesse sens montant si unSensMontant est Vrai
    '         vitesse sens descendant sinon
    
    DonnerVitSens = monCarrefour.DonnerVitSens(unSensMontant)
End Function

Public Function DonnerYSens(unSensMontant As Boolean) As Integer
    'Méthode à utiliser uniquement dans le cas de vitesse variable
    'Appelé par la procédure globale CalculerDecVitesse
    'Elle calcule l'ordonnée du carrefour réduit correspondant
    'au feu équivalent dans le sens choisi.
    'Retour :
    '   ordonnée du feu équivalent montant si unSensMontant est Vrai
    '   ordonnée du feu équivalent descendant sinon
    If unSensMontant Then
        DonnerYSens = monOrdonneeM
    Else
        DonnerYSens = monOrdonneeD
    End If
End Function

Public Function DonnerPosRefSens(unSensMontant As Boolean) As Single
    'Elle calcule la position du point de référence du carrefour réduit
    'correspondant au feu équivalent dans le sens choisi.
    '
    'Retour :
    '   position du point de référence du feu équivalent montant
    '   si unSensMontant est Vrai
    '   position du point de référence du feu équivalent descendant sinon
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    
    If unSensMontant Then
        DonnerPosRefSens = maPosRefM
    Else
        DonnerPosRefSens = maPosRefD
    End If
End Function

Public Function DonnerDureeVertSens(unSensMontant As Boolean) As Single
    'Elle calcule la durée de vert du carrefour réduit
    'correspondant au feu équivalent dans le sens choisi.
    '
    'Retour :
    '   durée de vert du feu équivalent montant
    '   si unSensMontant est Vrai
    '   durée de vert du feu équivalent descendant sinon
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    
    If unSensMontant Then
        DonnerDureeVertSens = maDureeVertM
    Else
        DonnerDureeVertSens = maDureeVertD
    End If
End Function

Public Function HasFeuDescendant() As Boolean
    'Elle indique si le carrefour réduit
    'a un feu équivalent dans le sens descendant.
    '
    'Retour :
    '   toujours VRAI car le carrefour réduit à double sens ont
    '   toujours au moins un feu équivalent dans le sens descendant.
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    HasFeuDescendant = True
End Function

Public Function HasFeuMontant() As Boolean
    'Elle indique si le carrefour réduit
    'a un feu équivalent dans le sens montant.
    '
    'Retour :
    '   toujours VRAI car le carrefour réduit à double sens ont
    '   toujours au moins un feu équivalent dans le sens montant.
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    HasFeuMontant = True
End Function

Public Property Let monEcart(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monEcart = 5
    mvarmonEcart = vData
End Property


Public Property Get monEcart() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monEcart
    monEcart = mvarmonEcart
End Property



Public Sub SetPropsSensM(uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer)
    mvarmaDureeVertM = uneDureeVert
    mvarmonOrdonneeM = uneOrdonnee
    mvarmaPosRefM = unePosRef
End Sub

Public Sub SetPropsSensD(uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer)
    mvarmaDureeVertD = uneDureeVert
    mvarmonOrdonneeD = uneOrdonnee
    mvarmaPosRefD = unePosRef
End Sub




Public Property Let maPosRefD(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maPosRefD = 5
    mvarmaPosRefD = vData
End Property


Public Property Get maPosRefD() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maPosRefD
    maPosRefD = mvarmaPosRefD
End Property



Public Property Let maPosRefM(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maPosRefM = 5
    mvarmaPosRefM = vData
End Property


Public Property Get maPosRefM() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maPosRefM
    maPosRefM = mvarmaPosRefM
End Property



Public Property Let monOrdonneeD(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monOrdonneeD = 5
    mvarmonOrdonneeD = vData
End Property


Public Property Get monOrdonneeD() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monOrdonneeD
    monOrdonneeD = mvarmonOrdonneeD
End Property



Public Property Let monOrdonneeM(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monOrdonneeM = 5
    mvarmonOrdonneeM = vData
End Property


Public Property Get monOrdonneeM() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monOrdonneeM
    monOrdonneeM = mvarmonOrdonneeM
End Property



Public Property Let maDureeVertD(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDureeVertD = 5
    mvarmaDureeVertD = vData
End Property


Public Property Get maDureeVertD() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDureeVertD
    maDureeVertD = mvarmaDureeVertD
End Property



Public Property Let maDureeVertM(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDureeVertM = 5
    mvarmaDureeVertM = vData
End Property


Public Property Get maDureeVertM() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDureeVertM
    maDureeVertM = mvarmaDureeVertM
End Property



Public Property Set monCarrefour(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarrefour = Form1
    Set mvarmonCarrefour = vData
End Property


Public Property Get monCarrefour() As Carrefour
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarrefour
    Set monCarrefour = mvarmonCarrefour
End Property

