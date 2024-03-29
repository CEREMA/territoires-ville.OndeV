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
    'Proc�dure vidant la collection.
    'Si les �l�ments de la collection ne sont plus r�f�renc�s
    '==> ils sont d�truits en m�moire
    
    'Algo : Puisque les collections sont r�index�es
    '       automatiquement, en supprimant le premier
    '       membre � chaque it�ration, on supprime tout.
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next i
End Sub

Public Function Add(unCarf As Carrefour, unSensMontant As Boolean, uneDureeVert As Single, unePosRef As Single, uneOrdonnee As Integer, Optional sKey As String) As CarfReduitSensUnique
    'cr�er un nouvel objet
    Dim objNewMember As CarfReduitSensUnique
    Set objNewMember = New CarfReduitSensUnique

    'd�finir les propri�t�s pass�es dans la m�thode
    With objNewMember
        Set .monCarrefour = unCarf
        .monSensMontant = unSensMontant
        .maDureeVert = uneDureeVert
        .monOrdonnee = uneOrdonnee
        .maPosRef = unePosRef
    End With
    
    'd�finir les propri�t�s pass�es dans la m�thode
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'renvoy� l'objet cr��
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CarfReduitSensUnique
Attribute Item.VB_UserMemId = 0
    'utilis�e pour faire r�f�rence � un �l�ment de la collection
    'vntIndexKey contient l'index ou la cl�, elle est donc
    'd�clar�e comme un Variant
    'Syntaxe: Set toto = x.Item(xyz) oo Set toto = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'utilis�e lors de l'extraction du nombre d'�l�ments dans
    'la collection. Syntaxe: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'utilis�e pour la suppression d'un �l�ment de la collection
    'vntIndexKey contient l'index ou la cl�, elle est donc
    'd�clar�e comme un Variant
    'Syntaxe: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'cette propri�t� vous permet d'�num�rer cette
    'collection avec la syntaxe For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'cr�e la collection lors de la cr�ation de cette classe
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'd�truit la collection lorsque cette classe est termin�e
    Set mCol = Nothing
End Sub


