VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Classe symbolisant le point de r�ference graphique de
'l'onde verte montante ou descendante apr�s le dessin d'une
'onde, ce point sera s�lectionnable interactivement

'variables locales de stockage des valeurs de propri�t�s
Private mvarmonDecal As Long 'copie locale
Private mvarmonIndCarf As Integer 'copie locale
Public Sub AffecterAttributs(unDecal As Long, unIndCarf As Integer)
    monDecal = unDecal
    monIndCarf = unIndCarf
End Sub

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



Public Property Let monDecal(ByVal vData As Long)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monDecal = 5
    mvarmonDecal = vData
End Property


Public Property Get monDecal() As Long
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monDecal
    monDecal = mvarmonDecal
End Property



