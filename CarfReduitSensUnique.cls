VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarfReduitSensUnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonOrdonnee As Integer 'copie locale
Private mvarmaDureeVert As Single 'copie locale
Private mvarmonCarrefour As Carrefour 'copie locale
Private mvarmaPosRef As Single 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonSensMontant As Boolean 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonDebOnde As Single 'copie locale
Public Sub AffecterDebOndeSens(unDebOnde As Single, unSensMontant As Boolean)
    'Elle affecte le point début de l'onde verte
    'du carrefour réduit correspondant
    
    'Le paramètre unSensMontant sert uniquement à avoir la même signature
    'que la méthode AffecterDebOndeSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    monDebOnde = unDebOnde
End Sub

Public Function DonnerDebOndeSens(unSensMontant As Boolean) As Single
    'Elle retourne le point début de l'onde verte
    'du carrefour réduit correspondant
    
    'Le paramètre unSensMontant sert uniquement à avoir la même signature
    'que la méthode AffecterDebOndeSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    DonnerDebOndeSens = monDebOnde
End Function

Public Property Let monDebOnde(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebOnde = 5
    mvarmonDebOnde = vData
End Property


Public Property Get monDebOnde() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebOnde
    monDebOnde = mvarmonDebOnde
End Property



Public Function DonnerVitSens(unSensMontant As Boolean) As Single
    'Méthode à utiliser uniquement dans le cas de vitesse variable
    'Appelé par la procédure globale CalculerDecVitesse
    'Elle calcule la vitesse du carrefour réduit correspondant
    'à la vitesse variable du carrefour.
    '
    'Retour : Vitesse sens montant si unSensMontant est Vrai (cas d'un carrefour unique sens montant)
    '         vitesse sens descendant sinon (cas d'un carrefour unique sens descendant)
    
    DonnerVitSens = monCarrefour.DonnerVitSens(unSensMontant)
End Function

Public Function DonnerYSens(unSensMontant As Boolean) As Integer
    'Méthode à utiliser uniquement dans le cas de vitesse variable
    'Appelé par la procédure globale CalculerDecVitesse
    'Elle calcule l'ordonnée du carrefour réduit correspondant
    'au feu équivalent dans le sens choisi.
    '
    'Retour :
    '   ordonnée du feu équivalent
    '
    'Le paramètre unSensMontant sert uniquement à avoir la même signature
    'que la méthode DonnerYSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    '(voir la procédure globale CalculerDecVitesse)
    DonnerYSens = monOrdonnee
End Function

Public Function DonnerPosRefSens(unSensMontant As Boolean) As Single
    'Elle calcule la position du point de référence du carrefour réduit
    'correspondant au feu équivalent dans le sens choisi.
    '
    'Retour :
    '   position du point de référence du feu équivalent
    '
    'Le paramètre unSensMontant sert uniquement à avoir la même signature
    'que la méthode DonnerPosRefSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    DonnerPosRefSens = maPosRef
End Function

Public Function DonnerDureeVertSens(unSensMontant As Boolean) As Single
    'Elle calcule la durée de vert du carrefour réduit
    'correspondant au feu équivalent dans le sens choisi.
    '
    'Retour :
    '   durée de vert du feu équivalent
    '
    'Le paramètre unSensMontant sert uniquement à avoir la même signature
    'que la méthode DonnerDureeVertSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    DonnerDureeVertSens = maDureeVert
End Function

Public Function HasFeuDescendant() As Boolean
    'Elle indique si le carrefour réduit
    'a un feu équivalent dans le sens descendant,
    'donc s'il est à sens unique descendant
    '
    'Retour :
    '   VRAI si le carrefour réduit est à sens unique descendant
    '   FAUX si le carrefour réduit est à sens unique montant
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    
    HasFeuDescendant = Not monSensMontant
End Function

Public Function HasFeuMontant() As Boolean
    'Elle indique si le carrefour réduit
    'a un feu équivalent dans le sens montant,
    'donc s'il est à sens unique montant
    '
    'Retour :
    '   FAUX si le carrefour réduit est à sens unique descendant
    '   VRAI si le carrefour réduit est à sens unique montant
    '
    '==> Utilisation du polymorphisme grâce à une variable de type Object
    
    HasFeuMontant = monSensMontant
End Function

Public Property Let monSensMontant(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monSensMontant = 5
    mvarmonSensMontant = vData
End Property


Public Property Get monSensMontant() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monSensMontant
    monSensMontant = mvarmonSensMontant
End Property




Public Property Let maPosRef(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maPosRef = 5
    mvarmaPosRef = vData
End Property


Public Property Get maPosRef() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maPosRef
    maPosRef = mvarmaPosRef
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



Public Property Let maDureeVert(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDureeVert = 5
    mvarmaDureeVert = vData
End Property


Public Property Get maDureeVert() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDureeVert
    maDureeVert = mvarmaDureeVert
End Property



Public Property Let monOrdonnee(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monOrdonnee = 5
    mvarmonOrdonnee = vData
End Property


Public Property Get monOrdonnee() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monOrdonnee
    monOrdonnee = mvarmonOrdonnee
End Property



