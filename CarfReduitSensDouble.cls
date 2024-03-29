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
'variables locales de stockage des valeurs de propri�t�s
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
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDebOndeD = 5
    mvarmonDebOndeD = vData
End Property


Public Property Get monDebOndeD() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDebOndeD
    monDebOndeD = mvarmonDebOndeD
End Property



Public Property Let monDebOndeM(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDebOndeM = 5
    mvarmonDebOndeM = vData
End Property


Public Property Get monDebOndeM() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDebOndeM
    monDebOndeM = mvarmonDebOndeM
End Property



Public Sub AffecterDebOndeSens(unDebOnde As Single, unSensMontant As Boolean)
    'Elle affecte le point d�but de l'onde verte
    'du carrefour r�duit correspondant
    'au feu �quivalent dans le sens choisi.
    If unSensMontant Then
        monDebOndeM = unDebOnde
    Else
        monDebOndeD = unDebOnde
    End If
End Sub

Public Function DonnerDebOndeSens(unSensMontant As Boolean) As Single
    'Elle retourne le point d�but de l'onde verte
    'du carrefour r�duit correspondant
    'au feu �quivalent dans le sens choisi.
    If unSensMontant Then
        DonnerDebOndeSens = monDebOndeM
    Else
        DonnerDebOndeSens = monDebOndeD
    End If
End Function


Public Function DonnerVitSens(unSensMontant As Boolean) As Single
    'M�thode utilis�e uniquement par la proc�dure globale CalculerDecVitesse
    'Elle calcule la vitesse du carrefour r�duit correspondant
    '� la vitesse variable, constante ou contrainte par un TC du carrefour.
    
    'Retour : Vitesse sens montant si unSensMontant est Vrai
    '         vitesse sens descendant sinon
    
    DonnerVitSens = monCarrefour.DonnerVitSens(unSensMontant)
End Function

Public Function DonnerYSens(unSensMontant As Boolean) As Integer
    'M�thode � utiliser uniquement dans le cas de vitesse variable
    'Appel� par la proc�dure globale CalculerDecVitesse
    'Elle calcule l'ordonn�e du carrefour r�duit correspondant
    'au feu �quivalent dans le sens choisi.
    'Retour :
    '   ordonn�e du feu �quivalent montant si unSensMontant est Vrai
    '   ordonn�e du feu �quivalent descendant sinon
    If unSensMontant Then
        DonnerYSens = monOrdonneeM
    Else
        DonnerYSens = monOrdonneeD
    End If
End Function

Public Function DonnerPosRefSens(unSensMontant As Boolean) As Single
    'Elle calcule la position du point de r�f�rence du carrefour r�duit
    'correspondant au feu �quivalent dans le sens choisi.
    '
    'Retour :
    '   position du point de r�f�rence du feu �quivalent montant
    '   si unSensMontant est Vrai
    '   position du point de r�f�rence du feu �quivalent descendant sinon
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    
    If unSensMontant Then
        DonnerPosRefSens = maPosRefM
    Else
        DonnerPosRefSens = maPosRefD
    End If
End Function

Public Function DonnerDureeVertSens(unSensMontant As Boolean) As Single
    'Elle calcule la dur�e de vert du carrefour r�duit
    'correspondant au feu �quivalent dans le sens choisi.
    '
    'Retour :
    '   dur�e de vert du feu �quivalent montant
    '   si unSensMontant est Vrai
    '   dur�e de vert du feu �quivalent descendant sinon
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    
    If unSensMontant Then
        DonnerDureeVertSens = maDureeVertM
    Else
        DonnerDureeVertSens = maDureeVertD
    End If
End Function

Public Function HasFeuDescendant() As Boolean
    'Elle indique si le carrefour r�duit
    'a un feu �quivalent dans le sens descendant.
    '
    'Retour :
    '   toujours VRAI car le carrefour r�duit � double sens ont
    '   toujours au moins un feu �quivalent dans le sens descendant.
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    HasFeuDescendant = True
End Function

Public Function HasFeuMontant() As Boolean
    'Elle indique si le carrefour r�duit
    'a un feu �quivalent dans le sens montant.
    '
    'Retour :
    '   toujours VRAI car le carrefour r�duit � double sens ont
    '   toujours au moins un feu �quivalent dans le sens montant.
    '
    '==> Utilisation du polymorphisme gr�ce � une variable de type Object
    HasFeuMontant = True
End Function

Public Property Let monEcart(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monEcart = 5
    mvarmonEcart = vData
End Property


Public Property Get monEcart() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
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
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maPosRefD = 5
    mvarmaPosRefD = vData
End Property


Public Property Get maPosRefD() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maPosRefD
    maPosRefD = mvarmaPosRefD
End Property



Public Property Let maPosRefM(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maPosRefM = 5
    mvarmaPosRefM = vData
End Property


Public Property Get maPosRefM() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maPosRefM
    maPosRefM = mvarmaPosRefM
End Property



Public Property Let monOrdonneeD(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monOrdonneeD = 5
    mvarmonOrdonneeD = vData
End Property


Public Property Get monOrdonneeD() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monOrdonneeD
    monOrdonneeD = mvarmonOrdonneeD
End Property



Public Property Let monOrdonneeM(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monOrdonneeM = 5
    mvarmonOrdonneeM = vData
End Property


Public Property Get monOrdonneeM() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monOrdonneeM
    monOrdonneeM = mvarmonOrdonneeM
End Property



Public Property Let maDureeVertD(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDureeVertD = 5
    mvarmaDureeVertD = vData
End Property


Public Property Get maDureeVertD() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDureeVertD
    maDureeVertD = mvarmaDureeVertD
End Property



Public Property Let maDureeVertM(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDureeVertM = 5
    mvarmaDureeVertM = vData
End Property


Public Property Get maDureeVertM() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDureeVertM
    maDureeVertM = mvarmaDureeVertM
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

