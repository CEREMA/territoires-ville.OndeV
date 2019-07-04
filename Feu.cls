VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe représentant un feu d'un carrefour contenant juste sa géométrie, son sens et son ordonnée."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonSensMontant As Boolean 'copie locale
Private mvarmonOrdonnée As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaDuréeDeVert As Single 'copie locale
Private mvarmaPositionPointRef As Single 'copie locale
Private mvarmonCarrefour As Carrefour 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaPosition As Integer 'copie locale

Public Function IsVert(unTemps As Single) As Boolean
    'Retourne Vrai si le feu est vert à l'instant unTemps
    'Faux sinon
    Dim unDebVert As Single, unFinVert As Single
    
    'Calcul du début de vert
    unDebVert = monCarrefour.monDecModif + maPositionPointRef
    If unTemps < unDebVert - 0.001 Then
        'Recul d'un cycle pourles plages de vert commençant
        'sur le graphique en négatif
        unDebVert = unDebVert - monSite.maDuréeDeCycle
    End If
    
    'Calcul de la fin du vert
    unFinVert = unDebVert + maDuréeDeVert
    
    'Recherche si l'instant unTemps est dans les plages de vert
    'du feu des différents cycles
    'Précision de calcul = 0.001
    IsVert = False
    Do
        If unDebVert - 0.001 < unTemps And unTemps < unFinVert + 0.001 Then
            IsVert = True
        Else
            'Incrémentation suivante, donc cycle suivant
            unDebVert = unDebVert + monSite.maDuréeDeCycle
            unFinVert = unFinVert + monSite.maDuréeDeCycle
        End If
    Loop Until unTemps < unDebVert - 0.001 Or IsVert
End Function

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



Public Property Let maPositionPointRef(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maPositionPointRef = 5
    mvarmaPositionPointRef = vData
End Property


Public Property Get maPositionPointRef() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maPositionPointRef
    maPositionPointRef = mvarmaPositionPointRef
End Property



Public Property Let maDuréeDeVert(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDuréeDeVert = 5
    mvarmaDuréeDeVert = vData
End Property


Public Property Get maDuréeDeVert() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDuréeDeVert
    maDuréeDeVert = mvarmaDuréeDeVert
End Property



Public Property Let monOrdonnée(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monOrdonnée = 5
    mvarmonOrdonnée = vData
End Property


Public Property Get monOrdonnée() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monOrdonnée
    monOrdonnée = mvarmonOrdonnée
End Property



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



