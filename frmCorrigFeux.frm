VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmCorrigFeux 
   Caption         =   "Correction des durées de vert des feux"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   Icon            =   "frmCorrigFeux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread TabFeuxCorrect 
      Height          =   3015
      Left            =   480
      OleObjectBlob   =   "frmCorrigFeux.frx":0442
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton OKBouton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton CancelBouton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label FondLock 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Fond pour les cellules lockées"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2130
   End
End
Attribute VB_Name = "frmCorrigFeux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBouton_Click()
    'Pas de validation des corrections
    frmModifCycle.IsCorrigé = False
    'Fermeture fenêtre
    Unload Me
End Sub

Private Sub Form_Load()
    'Affichage de tous les feux ayant des durées de vert
    'inférieure à la durée de cycle précédemment choisi
    Dim unFeu As Feu
    Dim unNomCarfPred As String
    
    unNomCarfPred = ""
    'Affectation d'une couleur pour les cellules lockées
    TabFeuxCorrect.LockBackColor = FondLock.BackColor
    'Remplissage du tableau des feux à corriger
    TabFeuxCorrect.MaxRows = frmModifCycle.maColFeux.Count
    For i = 1 To TabFeuxCorrect.MaxRows
        Set unFeu = frmModifCycle.maColFeux(i)
        TabFeuxCorrect.Row = i
        'Affichage du nom du carrefour du feu
        TabFeuxCorrect.Col = 1
        If unFeu.monCarrefour.monNom = unNomCarfPred Then
            'Cas où le nom du carrefour est le même que le précédent
            '==> on n'écrit rien
            TabFeuxCorrect.Text = "        ''"
        Else
            'Cas où le nom du carrefour est différent du précédent
            '==> on l'écrit
            TabFeuxCorrect.Text = unFeu.monCarrefour.monNom
            unNomCarfPred = unFeu.monCarrefour.monNom
        End If
        'Affichage du numéro du feu
        TabFeuxCorrect.Col = 2
        TabFeuxCorrect.Text = Format(unFeu.maPosition)
        'Affichage de la durée de vert actuelle du feu
        TabFeuxCorrect.Col = 3
        TabFeuxCorrect.Text = Format(unFeu.maDuréeDeVert)
        'Affichage de la durée de vert proposée pour corriger ce feu
        TabFeuxCorrect.Col = 4
        TabFeuxCorrect.Text = Format(Val(frmModifCycle.DuréeCycle.Text) / 2)
    Next i
End Sub

Private Sub OKBouton_Click()
    'Validation des corrections
    frmModifCycle.IsCorrigé = True
    'Modification dans les objets feux
    TabFeuxCorrect.Col = 4
    For i = 1 To frmModifCycle.maColFeux.Count
        TabFeuxCorrect.Row = i
        frmModifCycle.maColFeux(i).maDuréeDeVert = Val(TabFeuxCorrect.Text)
    Next i
    'Modification dans l'onglet Carrefour si visible
    If monSite.TabFeux.Tab = 0 Then
        monSite.TabPropCarf.Col = 3
        For i = 1 To monSite.TabPropCarf.MaxRows
            monSite.TabPropCarf.Row = i
            monSite.TabPropCarf.Text = Format(monSite.monCarrefourCourant.mesFeux(i).maDuréeDeVert)
        Next i
    End If
    'Fermeture fenêtre
    Unload Me
End Sub

Private Sub VerifMinMaxDuréeVert()
    'stockage de la cellule active
    uneRow = TabFeuxCorrect.ActiveRow
    uneCol = TabFeuxCorrect.ActiveCol
    'Positionnement sur la cellule active
    TabFeuxCorrect.Col = uneCol
    TabFeuxCorrect.Row = uneRow
    If Val(TabFeuxCorrect.Text) < 1 Or Val(TabFeuxCorrect.Text) >= Val(frmModifCycle.DuréeCycle.Text) Then
        'Test que la valeur de la durée de vert (colonne 4) est
        'entre 1 et la durée du cycle
        unMsg = "La durée de vert doit être >= 1 et < Durée du cycle, qui vaut " + frmModifCycle.DuréeCycle.Text
        unMsg = unMsg + Chr(13) + Chr(13) + "OndeV lui donnera comme valeur la moitié de la durée du cycle"
        MsgBox unMsg, vbCritical, "Message d'erreur de OndeV"
        'Positionnement sur la cellule initialement active
        TabFeuxCorrect.Col = uneCol
        TabFeuxCorrect.Row = uneRow
        TabFeuxCorrect.Text = Format(Val(frmModifCycle.DuréeCycle.Text) / 2)
        'Positionnement sur la cellule initialement active
        TabFeuxCorrect.Col = uneCol
        TabFeuxCorrect.Row = uneRow
        TabFeuxCorrect.Action = SS_ACTION_ACTIVE_CELL
    End If
End Sub

Private Sub TabFeuxCorrect_KeyUp(KeyCode As Integer, Shift As Integer)
    If TabFeuxCorrect.ActiveCol = 4 Then
        'Cas d'une saisie d'une durée de vert d'un feu
        Call VerifMinMaxDuréeVert
    End If
End Sub
