VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PhaseTabMarche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe repr�sentant une phase d'un tableau de marche de TC. Ce sont les phases d'acc�l�ration, de d�c�l�ration ou de vitesses constante."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonType As Integer 'copie locale
Private mvarmonYDeb As Single 'copie locale
Private mvarmonTDeb As Single 'copie locale
Private mvarmaLongPhase As Single 'copie locale
Private mvarmaDureePhase As Single 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmaVitPhase As Single 'copie locale
'variables locales de stockage des valeurs de propri�t�s
Private mvarmonProchainArret As ArretTC 'copie locale
Public Property Set monProchainArret(ByVal vData As Object)
'utilis� lors de l'affectation d'un Object � la propri�t�, du cot� gauche de l'instruction Set
'Syntax: Set x.monProchainArret = Form1
    Set mvarmonProchainArret = vData
End Property


Public Property Get monProchainArret() As ArretTC
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monProchainArret
    Set monProchainArret = mvarmonProchainArret
End Property



Public Property Let maVitPhase(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maVitPhase = 5
    mvarmaVitPhase = vData
End Property


Public Property Get maVitPhase() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maVitPhase
    maVitPhase = mvarmaVitPhase
End Property



Public Sub CalculerDateDansPhase(unY As Single)
End Sub

Public Function EstDansPhase(unY As Single) As Boolean
End Function

Public Property Let maDureePhase(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maDureePhase = 5
    mvarmaDureePhase = vData
End Property


Public Property Get maDureePhase() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maDureePhase
    maDureePhase = mvarmaDureePhase
End Property



Public Property Let maLongPhase(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maLongPhase = 5
    mvarmaLongPhase = vData
End Property


Public Property Get maLongPhase() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maLongPhase
    maLongPhase = mvarmaLongPhase
End Property



Public Property Let monTDeb(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monTDeb = 5
    mvarmonTDeb = vData
End Property


Public Property Get monTDeb() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monTDeb
    monTDeb = mvarmonTDeb
End Property



Public Property Let monYDeb(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monYDeb = 5
    mvarmonYDeb = vData
End Property


Public Property Get monYDeb() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monYDeb
    monYDeb = mvarmonYDeb
End Property



Public Property Let monType(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.monType = 5
    mvarmonType = vData
End Property


Public Property Get monType() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.monType
    monType = mvarmonType
End Property



