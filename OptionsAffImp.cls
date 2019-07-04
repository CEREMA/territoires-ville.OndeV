VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionsAffImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales de stockage des valeurs de propriétés
Private mvarmaVisuBandComM As Boolean 'copie locale
Private mvarmaVisuBandInterCarfM As Boolean 'copie locale
Private mvarmaCoulBandComM As Long 'copie locale
Private mvarmaCoulBandInterCarfM As Long 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaVisuBandComD As Boolean 'copie locale
Private mvarmaVisuBandInterCarfD As Boolean 'copie locale
Private mvarmaCoulBandComD As Long 'copie locale
Private mvarmaCoulBandInterCarfD As Long 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmaCoulPtRef As Long 'copie locale
Private mvarmaCoulNomCarf As Long 'copie locale
Private mvarmaCoulNomArret As Long 'copie locale
Private mvarmaCoulTitreEch As Long 'copie locale
Private mvarmaCoulLigne As Long 'copie locale
Private mvarmaVisuLigne As Boolean 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonEpaisseurLigne As Integer 'copie locale
'variables locales de stockage des valeurs de propriétés
Private mvarmonNbSecondesRappel As Integer 'copie locale
Public Property Let monNbSecondesRappel(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNbSecondesRappel = 5
    mvarmonNbSecondesRappel = vData
End Property


Public Property Get monNbSecondesRappel() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNbSecondesRappel
    monNbSecondesRappel = mvarmonNbSecondesRappel
End Property



Public Property Let monEpaisseurLigne(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monEpaisseurLigne = 5
    mvarmonEpaisseurLigne = vData
End Property


Public Property Get monEpaisseurLigne() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monEpaisseurLigne
    monEpaisseurLigne = mvarmonEpaisseurLigne
End Property



Public Property Let maVisuLigne(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVisuLigne = 5
    mvarmaVisuLigne = vData
End Property


Public Property Get maVisuLigne() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVisuLigne
    maVisuLigne = mvarmaVisuLigne
End Property



Public Property Let maCoulLigne(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulLigne = 5
    mvarmaCoulLigne = vData
End Property


Public Property Get maCoulLigne() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulLigne
    maCoulLigne = mvarmaCoulLigne
End Property



Public Property Let maCoulTitreEch(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulTitreEch = 5
    mvarmaCoulTitreEch = vData
End Property


Public Property Get maCoulTitreEch() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulTitreEch
    maCoulTitreEch = mvarmaCoulTitreEch
End Property



Public Property Let maCoulNomArret(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulNomArret = 5
    mvarmaCoulNomArret = vData
End Property


Public Property Get maCoulNomArret() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulNomArret
    maCoulNomArret = mvarmaCoulNomArret
End Property



Public Property Let maCoulNomCarf(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulNomCarf = 5
    mvarmaCoulNomCarf = vData
End Property


Public Property Get maCoulNomCarf() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulNomCarf
    maCoulNomCarf = mvarmaCoulNomCarf
End Property



Public Property Let maCoulPtRef(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulPtRef = 5
    mvarmaCoulPtRef = vData
End Property


Public Property Get maCoulPtRef() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulPtRef
    maCoulPtRef = mvarmaCoulPtRef
End Property



Public Property Let maCoulBandInterCarfD(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulBandInterCarfD = 5
    mvarmaCoulBandInterCarfD = vData
End Property


Public Property Get maCoulBandInterCarfD() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulBandInterCarfD
    maCoulBandInterCarfD = mvarmaCoulBandInterCarfD
End Property



Public Property Let maCoulBandComD(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulBandComD = 5
    mvarmaCoulBandComD = vData
End Property


Public Property Get maCoulBandComD() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulBandComD
    maCoulBandComD = mvarmaCoulBandComD
End Property



Public Property Let maVisuBandInterCarfD(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVisuBandInterCarfD = 5
    mvarmaVisuBandInterCarfD = vData
End Property


Public Property Get maVisuBandInterCarfD() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVisuBandInterCarfD
    maVisuBandInterCarfD = mvarmaVisuBandInterCarfD
End Property



Public Property Let maVisuBandComD(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVisuBandComD = 5
    mvarmaVisuBandComD = vData
End Property


Public Property Get maVisuBandComD() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVisuBandComD
    maVisuBandComD = mvarmaVisuBandComD
End Property



Public Property Let maCoulBandInterCarfM(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulBandInterCarfM = 5
    mvarmaCoulBandInterCarfM = vData
End Property


Public Property Get maCoulBandInterCarfM() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulBandInterCarfM
    maCoulBandInterCarfM = mvarmaCoulBandInterCarfM
End Property



Public Property Let maCoulBandComM(ByVal vData As Long)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maCoulBandComM = 5
    mvarmaCoulBandComM = vData
End Property


Public Property Get maCoulBandComM() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maCoulBandComM
    maCoulBandComM = mvarmaCoulBandComM
End Property



Public Property Let maVisuBandInterCarfM(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVisuBandInterCarfM = 5
    mvarmaVisuBandInterCarfM = vData
End Property


Public Property Get maVisuBandInterCarfM() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVisuBandInterCarfM
    maVisuBandInterCarfM = mvarmaVisuBandInterCarfM
End Property



Public Property Let maVisuBandComM(ByVal vData As Boolean)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maVisuBandComM = 5
    mvarmaVisuBandComM = vData
End Property


Public Property Get maVisuBandComM() As Boolean
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maVisuBandComM
    maVisuBandComM = mvarmaVisuBandComM
End Property



