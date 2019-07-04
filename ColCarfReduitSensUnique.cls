VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColCarfReduitSensUnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CarfReduitSensUnique"
Attribute VB_Ext_KEY = "Member0" ,"CarfReduitSensUnique"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"CarfReduitSensUnique"
'variable locale pour la collection
Private mCol As Collection
Public Sub Vider()
    'Procédure vidant la collection.
    'Si les éléments de la collection ne sont plus référencés
    '==> ils sont détruits en mémoire
    
    'Algo : Puisque les collections sont réindexées
    '       automatiquement, en supprimant le premier
    '       membre à chaque itération, on supprime tout.
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next i
End Sub

Public Function Add(unCarf As Carrefour, unSensMontant As Boolean, uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer, Optional sKey As String) As CarfReduitSensUnique
    'créer un nouvel objet
    Dim objNewMember As CarfReduitSensUnique
    Set objNewMember = New CarfReduitSensUnique

    'définir les propriétés passées dans la méthode
    With objNewMember
        Set .monCarrefour = unCarf
        .monSensMontant = unSensMontant
        .maDureeVert = uneDureeVert
        .monOrdonnee = uneOrdonnee
        .maPosRef = unePosRef
    End With
    
    'définir les propriétés passées dans la méthode
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'renvoyé l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CarfReduitSensUnique
Attribute Item.VB_UserMemId = 0
    'utilisée pour faire référence à un élément de la collection
    'vntIndexKey contient l'index ou la clé, elle est donc
    'déclarée comme un Variant
    'Syntaxe: Set toto = x.Item(xyz) oo Set toto = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'utilisée lors de l'extraction du nombre d'éléments dans
    'la collection. Syntaxe: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'utilisée pour la suppression d'un élément de la collection
    'vntIndexKey contient l'index ou la clé, elle est donc
    'déclarée comme un Variant
    'Syntaxe: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'cette propriété vous permet d'énumérer cette
    'collection avec la syntaxe For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crée la collection lors de la création de cette classe
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'détruit la collection lorsque cette classe est terminée
    Set mCol = Nothing
End Sub


