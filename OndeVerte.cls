VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OndeVerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe repr�sentant les r�sultats du calcul d'un onde verte avec ses donn�es de d�part : vitesses et largeur de bande passante montante et descendante, liste des carrefours pris en compte dans le calcul."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmaVitesseSensMont As Integer 'copie locale
Private mvarmaVitesseSensDesc As Integer 'copie locale
Private mvarmaLargeurBandeMont As Single 'copie locale
Private mvarmalargeurBandeDesc As Single 'copie locale
Private mvarmesD�calages As Collection 'copie locale
Private mvarmesCarrefoursUtiles As ColCarrefour 'copie locale
Public Property Set mesCarrefoursUtiles(ByVal vData As Object)
'utilis� lors de l'affectation d'un Object � la propri�t�, du cot� gauche de l'instruction Set
'Syntax: Set x.mesCarrefoursUtiles = Form1
    Set mvarmesCarrefoursUtiles = vData
End Property


Public Property Get mesCarrefoursUtiles() As ColCarrefour
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.mesCarrefoursUtiles
    Set mesCarrefoursUtiles = mvarmesCarrefoursUtiles
End Property



Public Property Set mesD�calages(ByVal vData As Object)
'utilis� lors de l'affectation d'un Object � la propri�t�, du cot� gauche de l'instruction Set
'Syntax: Set x.mesD�calages = Form1
    Set mvarmesD�calages = vData
End Property


Public Property Get mesD�calages() As Collection
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.mesD�calages
    Set mesD�calages = mvarmesD�calages
End Property



Public Property Let malargeurBandeDesc(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.malargeurBandeDesc = 5
    mvarmalargeurBandeDesc = vData
End Property


Public Property Get malargeurBandeDesc() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.malargeurBandeDesc
    malargeurBandeDesc = mvarmalargeurBandeDesc
End Property



Public Property Let maLargeurBandeMont(ByVal vData As Single)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maLargeurBandeMont = 5
    mvarmaLargeurBandeMont = vData
End Property


Public Property Get maLargeurBandeMont() As Single
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maLargeurBandeMont
    maLargeurBandeMont = mvarmaLargeurBandeMont
End Property



Public Property Let maVitesseSensDesc(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maVitesseSensDesc = 5
    mvarmaVitesseSensDesc = vData
End Property


Public Property Get maVitesseSensDesc() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maVitesseSensDesc
    maVitesseSensDesc = mvarmaVitesseSensDesc
End Property



Public Property Let maVitesseSensMont(ByVal vData As Integer)
'utilis� lors de l'affectation d'une valeur � la propri�t�, du cot� gauche de l'affectation.
'Syntax: X.maVitesseSensMont = 5
    mvarmaVitesseSensMont = vData
End Property


Public Property Get maVitesseSensMont() As Integer
'utilis� lors de la lecture de la valeur de la propri�t�, du cot� droit de l'instruction.
'Syntax: Debug.Print X.maVitesseSensMont
    maVitesseSensMont = mvarmaVitesseSensMont
End Property



