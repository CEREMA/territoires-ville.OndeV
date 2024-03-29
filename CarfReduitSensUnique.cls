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
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonSensMontant As Boolean 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonDebOnde As Single 'copie locale
Public Sub AffecterDebOndeSens(unDebOnde As Single, unSensMontant As Boolean)
    'Elle affecte le point d�but de l'onde verte
    'du carrefour r�duit correspondant
    
    'Le param�tre unSensMontant sert uniquement � avoir la m�me signature
    'que la m�thode AffecterDebOndeSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    monDebOnde = unDebOnde
End Sub

Public Function DonnerDebOndeSens(unSensMontant As Boolean) As Single
    'Elle retourne le point d�but de l'onde verte
    'du carrefour r�duit correspondant
    
    'Le param�tre unSensMontant sert uniquement � avoir la m�me signature
    'que la m�thode AffecterDebOndeSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    DonnerDebOndeSens = monDebOnde
End Function

Public Property Let monDebOnde(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDebOnde = 5
    mvarmonDebOnde = vData
End Property


Public Property Get monDebOnde() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDebOnde
    monDebOnde = mvarmonDebOnde
End Property



Public Function DonnerVitSens(unSensMontant As Boolean) As Single
    'M�thode � utiliser uniquement dans le cas de vitesse variable
    'Appel� par la proc�dure globale CalculerDecVitesse
    'Elle calcule la vitesse du carrefour r�duit correspondant
    '� la vitesse variable du carrefour.
    '
    'Retour : Vitesse sens montant si unSensMontant est Vrai (cas d'un carrefour unique sens montant)
    '         vitesse sens descendant sinon (cas d'un carrefour unique sens descendant)
    
    DonnerVitSens = monCarrefour.DonnerVitSens(unSensMontant)
End Function

Public Function DonnerYSens(unSensMontant As Boolean) As Integer
    'M�thode � utiliser uniquement dans le cas de vitesse variable
    'Appel� par la proc�dure globale CalculerDecVitesse
    'Elle calcule l'ordonn�e du carrefour r�duit correspondant
    'au feu �quivalent dans le sens choisi.
    '
    'Retour :
    '   ordonn�e du feu �quivalent
    '
    'Le param�tre unSensMontant sert uniquement � avoir la m�me signature
    'que la m�thode DonnerYSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    '(voir la proc�dure globale CalculerDecVitesse)
    DonnerYSens = monOrdonnee
End Function

Public Function DonnerPosRefSens(unSensMontant As Boolean) As Single
    'Elle calcule la position du point de r�f�rence du carrefour r�duit
    'correspondant au feu �quivalent dans le sens choisi.
    '
    'Retour :
    '   position du point de r�f�rence du feu �quivalent
    '
    'Le param�tre unSensMontant sert uniquement � avoir la m�me signature
    'que la m�thode DonnerPosRefSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    DonnerPosRefSens = maPosRef
End Function

Public Function DonnerDureeVertSens(unSensMontant As Boolean) As Single
    'Elle calcule la dur�e de vert du carrefour r�duit
    'correspondant au feu �quivalent dans le sens choisi.
    '
    'Retour :
    '   dur�e de vert du feu �quivalent
    '
    'Le param�tre unSensMontant sert uniquement � avoir la m�me signature
    'que la m�thode DonnerDureeVertSens des CarfReduitSensDouble
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    DonnerDureeVertSens = maDureeVert
End Function

Public Function HasFeuDescendant() As Boolean
    'Elle indique si le carrefour r�duit
    'a un feu �quivalent dans le sens descendant,
    'donc s'il est � sens unique descendant
    '
    'Retour :
    '   VRAI si le carrefour r�duit est � sens unique descendant
    '   FAUX si le carrefour r�duit est � sens unique montant
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    
    HasFeuDescendant = Not monSensMontant
End Function

Public Function HasFeuMontant() As Boolean
    'Elle indique si le carrefour r�duit
    'a un feu �quivalent dans le sens montant,
    'donc s'il est � sens unique montant
    '
    'Retour :
    '   FAUX si le carrefour r�duit est � sens unique descendant
    '   VRAI si le carrefour r�duit est � sens unique montant
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    
    HasFeuMontant = monSensMontant
End Function

Public Property Let monSensMontant(ByVal vData As Boolean)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monSensMontant = 5
    mvarmonSensMontant = vData
End Property


Public Property Get monSensMontant() As Boolean
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monSensMontant
    monSensMontant = mvarmonSensMontant
End Property




Public Property Let maPosRef(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maPosRef = 5
    mvarmaPosRef = vData
End Property


Public Property Get maPosRef() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maPosRef
    maPosRef = mvarmaPosRef
End Property



Public Property Set monCarrefour(ByVal vData As Object)
'utilis� lors de l'affectation d'un Object � la propri�t�, du cot� gauche de l'instruction Set
'Syntax: Set x.monCarrefour = Form1
    Set mvarmonCarrefour = vData
End Property


Public Property Get monCarrefour() As Carrefour
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monCarrefour
    Set monCarrefour = mvarmonCarrefour
End Property



Public Property Let maDureeVert(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDureeVert = 5
    mvarmaDureeVert = vData
End Property


Public Property Get maDureeVert() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDureeVert
    maDureeVert = mvarmaDureeVert
End Property



Public Property Let monOrdonnee(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monOrdonnee = 5
    mvarmonOrdonnee = vData
End Property


Public Property Get monOrdonnee() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monOrdonnee
    monOrdonnee = mvarmonOrdonnee
End Property



