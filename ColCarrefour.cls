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
    'cr�er un nouvel objet
    Dim objNewMember As Carrefour
    Set objNewMember = New Carrefour

    'd�finir les propri�t�s pass�es dans la m�thode
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
    
    'd�finir les propri�t�s pass�es dans la m�thode
    objNewMember.maPosition = mCol.Count

    'renvoy� l'objet cr��
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Carrefour
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


Public Function HasCarfRed(unCarfRed As Object) As Boolean
    'Retourne true si le carrefour r�duit a son
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
    'Ajout d'un carrefour d�j� cr�� dans la collection
    mCol.Add unCarf
End Sub

