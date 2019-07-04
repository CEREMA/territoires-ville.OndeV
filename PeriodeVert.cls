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
'Classe permettant de stocker les périodes de vert
'trouvée lors du calcul du feu équivalent d'un carrefour
'
'maDuree As Single                  Durée de la période de vert trouvée
'monIsTousFeuxVert As Boolean       Vrai si tous les feux du carrefour
'                                   sont vert dans cette période
'monDernierFeuDevenuVert As Integer Indice du dernier feu passé au vert
'                                   dans cette période
'*************************************************************************

'variables locales de stockage des valeurs de propriétés
Private mvarmaDuree As Single 'copie locale
Private mvarmonDebutVert As Single 'copie locale
Public Property Let monDebutVert(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebutVert = 5
    mvarmonDebutVert = vData
End Property


Public Property Get monDebutVert() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebutVert
    monDebutVert = mvarmonDebutVert
End Property


Public Property Let maDuree(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maDuree = 5
    mvarmaDuree = vData
End Property


Public Property Get maDuree() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maDuree
    maDuree = mvarmaDuree
End Property



