VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlageGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Classe symbolisant la plage de vert contenant
'l'onde verte montante ou descendante apr�s le dessin d'une
'onde, cette ligne sera s�lectionnable interactivement

'variables locales de stockage des valeurs de propri�t�s
Private mvarmonDebVert As Long 'copie locale
Private mvarmonFinVert As Long 'copie locale
Private mvarmonIndCarf As Integer 'copie locale
Private mvarmonIndFeu As Integer 'copie locale
Public Sub AffecterAttributs(unDebVert As Long, unFinVert As Long, unIndCarf As Integer, unIndFeu As Integer)
    monDebVert = unDebVert
    monFinVert = unFinVert
    monIndCarf = unIndCarf
    monIndFeu = unIndFeu
End Sub

Public Property Let monIndFeu(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monIndFeu = 5
    mvarmonIndFeu = vData
End Property


Public Property Get monIndFeu() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monIndFeu
    monIndFeu = mvarmonIndFeu
End Property



Public Property Let monIndCarf(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monIndCarf = 5
    mvarmonIndCarf = vData
End Property


Public Property Get monIndCarf() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monIndCarf
    monIndCarf = mvarmonIndCarf
End Property



Public Property Let monFinVert(ByVal vData As Long)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monFinVert = 5
    mvarmonFinVert = vData
End Property


Public Property Get monFinVert() As Long
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monFinVert
    monFinVert = mvarmonFinVert
End Property



Public Property Let monDebVert(ByVal vData As Long)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDebVert = 5
    mvarmonDebVert = vData
End Property


Public Property Get monDebVert() As Long
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDebVert
    monDebVert = mvarmonDebVert
End Property






