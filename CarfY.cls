VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarfY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************************
'* Classe contenant un carrefour réduit avec son Y qui ont  *
'* été calculée par la moyenne de Y de ces feux équivalents *
'************************************************************

'variables locales de stockage des valeurs de propriétés
Private mvarmonCarfReduit As Object 'copie locale
Private mvarmonY As Single 'copie locale
Public Property Let monY(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monY = 5
    mvarmonY = vData
End Property


Public Property Get monY() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monY
    monY = mvarmonY
End Property



Public Property Set monCarfReduit(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monCarfReduit = Form1
    Set mvarmonCarfReduit = vData
End Property


Public Property Get monCarfReduit() As Object
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monCarfReduit
    Set monCarfReduit = mvarmonCarfReduit
End Property



