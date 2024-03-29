VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PeriodeVert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*************************************************************************
'Classe permettant de stocker les p�riodes de vert
'trouv�e lors du calcul du feu �quivalent d'un carrefour
'
'maDuree As Single                  Dur�e de la p�riode de vert trouv�e
'monIsTousFeuxVert As Boolean       Vrai si tous les feux du carrefour
'                                   sont vert dans cette p�riode
'monDernierFeuDevenuVert As Integer Indice du dernier feu pass� au vert
'                                   dans cette p�riode
'*************************************************************************

'variables locales de stockage des valeurs de propri�t�s
Private mvarmaDuree As Single 'copie locale
Private mvarmonDebutVert As Single 'copie locale
Public Property Let monDebutVert(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDebutVert = 5
    mvarmonDebutVert = vData
End Property


Public Property Get monDebutVert() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDebutVert
    monDebutVert = mvarmonDebutVert
End Property


Public Property Let maDuree(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDuree = 5
    mvarmaDuree = vData
End Property


Public Property Get maDuree() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDuree
    maDuree = mvarmaDuree
End Property



