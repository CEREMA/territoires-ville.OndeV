VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArretTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarmonOrdonnee As Integer 'copie locale
Private mvarmonTempsArret As Integer 'copie locale
Private mvarmaVitesseMarche As Integer 'copie locale
Private mvarmonLibelle As String 'copie locale
Public Property Let monLibelle(ByVal vData As String)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monLibelle = 5
    mvarmonLibelle = vData
End Property


Public Property Get monLibelle() As String
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monLibelle
    monLibelle = mvarmonLibelle
End Property



Public Property Let maVitesseMarche(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVitesseMarche = 5
    mvarmaVitesseMarche = vData
End Property


Public Property Get maVitesseMarche() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVitesseMarche
    maVitesseMarche = mvarmaVitesseMarche
End Property



Public Property Let monTempsArret(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monTempsArret = 5
    mvarmonTempsArret = vData
End Property


Public Property Get monTempsArret() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monTempsArret
    monTempsArret = mvarmonTempsArret
End Property



Public Property Let monOrdonnee(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monOrdonnee = 5
    mvarmonOrdonnee = vData
End Property


Public Property Get monOrdonnee() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monOrdonnee
    monOrdonnee = mvarmonOrdonnee
End Property



