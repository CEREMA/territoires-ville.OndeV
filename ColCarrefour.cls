VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColCarrefour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Carrefour"
Attribute VB_Ext_KEY = "Member0" ,"Carrefour"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Carrefour"
'variable locale pour la collection
Private mCol As Collection

Public Function Add(unNom As String, uneVitSensM As Integer, uneVitSensD As Integer, Optional unIsUtil As Boolean = True, Optional unDecImp As Integer = 0, Optional sKey As String) As Carrefour
    'créer un nouvel objet
    Dim objNewMember As Carrefour
    Set objNewMember = New Carrefour

    'définir les propriétés passées dans la méthode
    objNewMember.monNom = unNom
    objNewMember.monIsUtil = unIsUtil
    objNewMember.monDecImp = unDecImp
    objNewMember.maVitSensM = uneVitSensM
    objNewMember.maVitSensD = uneVitSensD
    objNewMember.maVitTCSensM = uneVitSensM
    objNewMember.maVitTCSensD = uneVitSensD
    Set objNewMember.mesFeux = New ColFeu
    Set objNewMember.mesFeuxGraphics = New Collection

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    'définir les propriétés passées dans la méthode
    objNewMember.maPosition = mCol.Count

    'renvoyé l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Carrefour
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


Public Function HasCarfRed(unCarfRed As Object) As Boolean
    'Retourne true si le carrefour réduit a son
    'carrefour dans cette collection
    HasCarfRed = False
    For i = 1 To mCol.Count
        If unCarfRed.monCarrefour Is mCol(i) Then
            HasCarfRed = True
            Exit For
        End If
    Next i
End Function

Public Sub AjouterCarf(unCarf As Carrefour)
    'Ajout d'un carrefour déjà créé dans la collection
    mCol.Add unCarf
End Sub

