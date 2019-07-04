VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeuSens1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe représentant un feu dans un sens choisi\r\nmontant ou descendant\r\n   Attribut monFeu est le feu utilisé\r\n   Attribut maBorneVert1 et maBorneVert2 sont les bornes\r\n            inf et sup des plages de vert du feu entre 0\r\n            et la duréee du cycle\r\n   Attribut monNbPlageVert est le nombre de plages de durée\r\n            de vert entre 0 et la durée du cycle\r\n       Si 2 plages ==> phases de vert = [0,uneBorneVert1[                               union [uneBorneVert2, durée du cycle[\r\n       Sinon 1 plage ==> phase de vert =                                          [uneBorneVert1,uneBorneVert2[\r\n"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================
' Commentaires sur cette classe :
'=================================
'
'Classe représentant un feu dans un sens choisi
'montant ou descendant
'   Attribut monFeu est le feu utilisé
'   Attribut maBorneVert1 et maBorneVert2 sont les bornes
'            inf et sup des plages de vert du feu entre 0
'            et la duréee du cycle
'   Attribut monNbPlageVert est le nombre de plages de durée
'            de vert entre 0 et la durée du cycle
'       Si 2 plages ==> phases de vert = [0,uneBorneVert1[ union [uneBorneVert2, durée du cycle[
'       Sinon 1 plage ==> phase de vert = [uneBorneVert1,uneBorneVert2[

'variables locales de stockage des valeurs de propriétés
Private mvarmonFeu As Feu 'copie locale
Private mvarmaBorneVert1 As Single 'copie locale
Private mvarmaBorneVert2 As Single 'copie locale
Private mvarmonNbPlageVert As Integer 'copie locale
Public Function IsVert(unInstantT As Single) As Boolean
    'Retourne Vrai si le feu de sens 1 (le sens étudié montant
    'ou descendant) est vert à T = unInstantT secondes
    'Faux sinon
    If monNbPlageVert = 1 Then
        'Cas d'un feu étant vert dans
        'l'intervalle [maBorneVert1, maBorneVert2[
        If unInstantT >= maBorneVert1 And unInstantT < maBorneVert2 Then
            IsVert = True
        Else
            IsVert = False
        End If
    Else
        'Cas d'un feu étant vert dans l'union des 2 intervalles
        '[0, maBorneVert1[ et [maBorneVert2, duréee du cycle[
        If unInstantT < maBorneVert1 Or unInstantT >= maBorneVert2 Then
            'Test seulement par rapport aux bornes de vert car
            'unInstanT appartient à l'intervalle [0, durée du cycle[
            IsVert = True
        Else
            IsVert = False
        End If
    End If
End Function


Public Property Let monNbPlageVert(ByVal vData As Integer)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.monNbPlageVert = 5
    mvarmonNbPlageVert = vData
End Property


Public Property Get monNbPlageVert() As Integer
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monNbPlageVert
    monNbPlageVert = mvarmonNbPlageVert
End Property



Public Property Let maBorneVert2(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maBorneVert2 = 5
    mvarmaBorneVert2 = vData
End Property


Public Property Get maBorneVert2() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maBorneVert2
    maBorneVert2 = mvarmaBorneVert2
End Property



Public Property Let maBorneVert1(ByVal vData As Single)
'utilisé lors de l'affectation d'une valeur à la propriété, du coté gauche de l'affectation.
'Syntax: X.maBorneVert1 = 5
    mvarmaBorneVert1 = vData
End Property


Public Property Get maBorneVert1() As Single
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.maBorneVert1
    maBorneVert1 = mvarmaBorneVert1
End Property



Public Property Set monFeu(ByVal vData As Object)
'utilisé lors de l'affectation d'un Object à la propriété, du coté gauche de l'instruction Set
'Syntax: Set x.monFeu = Form1
    Set mvarmonFeu = vData
End Property


Public Property Get monFeu() As Feu
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print X.monFeu
    Set monFeu = mvarmonFeu
End Property



