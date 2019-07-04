VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'variables locales de stockage des valeurs de propriétés
Private mvarmonPtRef As Long 'copie locale
Private mvarmonDebVert As Long 'copie locale
Private mvarmonFinVert As Long 'copie locale
Private mvarmonIndCarf As Integer 'copie locale
Private mvarmonIndFeu As Integer 'copie locale
Public Property Let monIndFeu(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monIndFeu = 5
    mvarmonIndFeu = vData
End Property


Public Property Get monIndFeu() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monIndFeu
    monIndFeu = mvarmonIndFeu
End Property



Public Property Let monIndCarf(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monIndCarf = 5
    mvarmonIndCarf = vData
End Property


Public Property Get monIndCarf() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monIndCarf
    monIndCarf = mvarmonIndCarf
End Property



Public Property Let monFinVert(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monFinVert = 5
    mvarmonFinVert = vData
End Property


Public Property Get monFinVert() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monFinVert
    monFinVert = mvarmonFinVert
End Property



Public Property Let monDebVert(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monDebVert = 5
    mvarmonDebVert = vData
End Property


Public Property Get monDebVert() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monDebVert
    monDebVert = mvarmonDebVert
End Property



Public Property Let monPtRef(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monPtRef = 5
    mvarmonPtRef = vData
End Property


Public Property Get monPtRef() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monPtRef
    monPtRef = mvarmonPtRef
End Property



