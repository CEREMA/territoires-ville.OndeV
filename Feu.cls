VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe repr�sentant un feu d'un carrefour contenant juste sa g�om�trie, son sens et son ordonn�e."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonSensMontant As Boolean 'copie locale
Private mvarmonOrdonn�e As Integer 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmaDur�eDeVert As Single 'copie locale
Private mvarmaPositionPointRef As Single 'copie locale
Private mvarmonCarrefour As Carrefour 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmaPosition As Integer 'copie locale

Public Function IsVert(unTemps As Single) As Boolean
    'Retourne Vrai si le feu est vert � l'instant unTemps
    'Faux sinon
    Dim unDebVert As Single, unFinVert As Single
    
    'Calcul du d�but de vert
    unDebVert = monCarrefour.monDecModif + maPositionPointRef
    If unTemps < unDebVert - 0.001 Then
        'Recul d'un cycle pourles plages de vert commen�ant
        'sur le graphique en n�gatif
        unDebVert = unDebVert - monSite.maDur�eDeCycle
    End If
    
    'Calcul de la fin du vert
    unFinVert = unDebVert + maDur�eDeVert
    
    'Recherche si l'instant unTemps est dans les plages de vert
    'du feu des diff�rents cycles
    'Pr�cision de calcul = 0.001
    IsVert = False
    Do
        If unDebVert - 0.001 < unTemps And unTemps < unFinVert + 0.001 Then
            IsVert = True
        Else
            'Incr�mentation suivante, donc cycle suivant
            unDebVert = unDebVert + monSite.maDur�eDeCycle
            unFinVert = unFinVert + monSite.maDur�eDeCycle
        End If
    Loop Until unTemps < unDebVert - 0.001 Or IsVert
End Function

Public Property Let maPosition(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maPosition = 5
    mvarmaPosition = vData
End Property


Public Property Get maPosition() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maPosition
    maPosition = mvarmaPosition
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



Public Property Let maPositionPointRef(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maPositionPointRef = 5
    mvarmaPositionPointRef = vData
End Property


Public Property Get maPositionPointRef() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maPositionPointRef
    maPositionPointRef = mvarmaPositionPointRef
End Property



Public Property Let maDur�eDeVert(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDur�eDeVert = 5
    mvarmaDur�eDeVert = vData
End Property


Public Property Get maDur�eDeVert() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDur�eDeVert
    maDur�eDeVert = mvarmaDur�eDeVert
End Property



Public Property Let monOrdonn�e(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monOrdonn�e = 5
    mvarmonOrdonn�e = vData
End Property


Public Property Get monOrdonn�e() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monOrdonn�e
    monOrdonn�e = mvarmonOrdonn�e
End Property



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



