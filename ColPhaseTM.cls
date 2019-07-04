VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColPhaseTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PhaseTabMarche"
Attribute VB_Ext_KEY = "Member0" ,"PhaseTabMarche"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variable locale pour la collection
Private mCol As Collection
Public Function Add(monType As Integer, monYDeb As Single, monTDeb As Single, maLongPhase As Single, maDureePhase As Single, maVitPhase As Single, monProchainArret As ArretTC, Optional unIndAfter As Integer = 0) As PhaseTabMarche
    'créer un nouvel objet
    Dim objNewMember As PhaseTabMarche
    Set objNewMember = New PhaseTabMarche


    'définir les propriétés passées dans la méthode
    objNewMember.monType = monType
    objNewMember.monYDeb = monYDeb
    objNewMember.monTDeb = monTDeb
    objNewMember.maLongPhase = maLongPhase
    objNewMember.maDureePhase = maDureePhase
    objNewMember.maVitPhase = maVitPhase
    Set objNewMember.monProchainArret = monProchainArret

    If unIndAfter = 0 Then
        'Cas d'ajout en fin de collection
        mCol.Add objNewMember
    Else
        'Cas d'ajout aprés la position unIndAfter
        mCol.Add objNewMember, , , unIndAfter
    End If


    'renvoyé l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As PhaseTabMarche
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

