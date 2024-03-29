VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TC"
Attribute VB_Ext_KEY = "Member0" ,"TC"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"TC"
'variable locale pour la collection
Private mCol As Collection

Public Function Add(unNom As String, unTDep As Integer, uneDistAccFrein As Integer, uneDureeAccFrein As Integer, unCarfDep As Carrefour, unCarfArr As Carrefour, uneCouleur As Long, Optional sKey As String) As TC
    'cr�er un nouvel objet
    Dim objNewMember As TC
    Set objNewMember = New TC


    'd�finir les propri�t�s pass�es dans la m�thode
    objNewMember.monNom = unNom
    objNewMember.monTDep = unTDep
    objNewMember.maCouleur = uneCouleur
    objNewMember.maDistAccFrein = uneDistAccFrein
    objNewMember.maDureeAccFrein = uneDureeAccFrein
    Set objNewMember.monCarfDep = unCarfDep
    Set objNewMember.monCarfArr = unCarfArr
    Set objNewMember.mesArrets = New ColArretTC
    Set objNewMember.mesObjGraphics = New Collection
    'Cr�ation du tableau de marche pour cadrer une onde par ce TC
    Set objNewMember.mesPhasesTMOnde = New ColPhaseTM
    'Cr�ation du tableau de marche pour tracer la progression du TC
    Set objNewMember.mesPhasesTMProg = New ColPhaseTM
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'renvoy� l'objet cr��
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As TC
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

