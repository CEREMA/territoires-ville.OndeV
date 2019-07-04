VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmOptiVit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recherche des vitesses optimisant les bandes passantes"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   Icon            =   "frmOptiVit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread TabCouple 
      Height          =   2295
      Left            =   1080
      OleObjectBlob   =   "frmOptiVit.frx":0442
      TabIndex        =   5
      Top             =   3360
      Width           =   8535
   End
   Begin FPSpread.vaSpread TabOptiVit 
      Height          =   1455
      Left            =   120
      OleObjectBlob   =   "frmOptiVit.frx":0778
      TabIndex        =   3
      Top             =   720
      Width           =   6615
   End
   Begin FPSpread.vaSpread TabResult 
      Height          =   975
      Left            =   1080
      OleObjectBlob   =   "frmOptiVit.frx":0C80
      TabIndex        =   10
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Frame FrameSeparateur 
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   9615
   End
   Begin VB.CommandButton BoutonChangerVitesse 
      Caption         =   "Utiliser les vitesses optimales"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton BoutonFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton BoutonCalculerVitesse 
      Caption         =   "Calculer les vitesses optimales"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "en km/h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Proposition de vitesses et de bandes passantes optimales :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   5040
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   5160
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "En rouge, les bandes maximales trouvées"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   2640
      Width           =   3525
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vitesses descendantes en km/h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   3000
      Width           =   2730
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Couples de bandes passantes possibles en secondes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   4650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "montantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vitesses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   720
   End
   Begin VB.Label LabelDureeCycle 
      AutoSize        =   -1  'True
      Caption         =   "Durée du cycle : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1485
   End
End
Attribute VB_Name = "frmOptiVit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant des paramètres du site en cours
'qui vont modifiés lors du calcul d'optimun
Private maSaveTypeOnde As Integer
Private maSavePoidsM As Integer
Private maSavePoidsD As Integer
Private maSaveTypeVit As Integer
Private maSaveVitSensM As Integer
Private maSaveVitSensD As Integer
'Variables stockant les carrefours réduits en sens montants et descendant
Private unTabCarfM() As CarfRed
Private unTabCarfD() As CarfRed
'Tableau stockant par carrefour son type de sens (unique ou double)
Private unTabSensDouble() As Boolean


Private Sub BoutonCalculerVitesse_Click()
    'Changement du pointeur souris en sablier
    MousePointer = vbHourglass
    'Lancement du calcul des vitesses optimales
    If RechercherOptimun Then
        'Affichage des controls se trouvant sous FrameSeparateur
        Height = Height - ScaleHeight + TabResult.Top + TabResult.Height
        'Désinhibition du bouton Utiliser les vitesses optimales
        BoutonChangerVitesse.Enabled = True
    End If
    'Restauration du pointeur souris
    MousePointer = vbDefault
End Sub

Private Sub BoutonChangerVitesse_Click()
    unMsg = "ATTENTION, les vitesses constantes et les poids dans les deux sens vont être changés."
    If MsgBox(unMsg, vbCritical + vbYesNo) = vbYes Then
        'Affichage des nouvelles valeurs des paramètres du site en cours
        'qui ont été modifiés lors du calcul
        monSite.OptionOndeDouble.Value = True
        monSite.TextPoidsM.Text = Format(monSite.monPoidsSensM)
        monSite.TextPoidsD.Text = Format(monSite.monPoidsSensD)
        monSite.OptionVitConst.Value = True
        monSite.TextVitM.Text = Format(monSite.maVitSensM)
        monSite.TextVitD.Text = Format(monSite.maVitSensD)
        'Stockage d'une modif dans les données du calcul d'onde
        monSite.maModifDataOnde = True
        'Fermeture
        Unload Me
    End If
End Sub

Private Sub BoutonFermer_Click()
    'Restauration des paramètres du site en cours qui ont été
    'modifiés lors du calcul
    monSite.monTypeOnde = maSaveTypeOnde
    monSite.monPoidsSensM = maSavePoidsM
    monSite.monPoidsSensD = maSavePoidsD
    monSite.monTypeVit = maSaveTypeVit
    monSite.maVitSensM = maSaveVitSensM
    monSite.maVitSensD = maSaveVitSensD
    'Fermeture
    Unload Me
End Sub


Private Sub Form_Load()
    'Index pour l'aide
    HelpContextID = IDhlp_WinOptiVit
    
    'Centrage de la fenetre
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'Masquage des controls se trouvant sous FrameSeparateur
    'Il apparaitront lors du click sur le bouton calculer
    Height = Height - ScaleHeight + FrameSeparateur.Top
    
    'Affectation d'une couleur pour les cellules lockées
    'grâce à la couleur des Info bulles stockées dans LabelTrait
    'du site actif
    TabCouple.LockBackColor = monSite.LabelTrait.BackColor
    TabResult.LockBackColor = monSite.LabelTrait.BackColor
    
    'Initialisation des poids avec ceux du site courant
    TabOptiVit.Row = 3
    TabOptiVit.Col = 1
    TabOptiVit.Text = Format(monSite.monPoidsSensM)
    TabOptiVit.Col = 2
    TabOptiVit.Text = Format(monSite.monPoidsSensD)
    
    'Affichage de la durée du cycle
    LabelDureeCycle.Caption = "Durée du cycle : " + Format(monSite.maDuréeDeCycle) + " secondes"
    
    'Stockage des paramètres du site en cours qui vont modifiés
    'lors du calcul
    maSaveTypeOnde = monSite.monTypeOnde
    maSavePoidsM = monSite.monPoidsSensM
    maSavePoidsD = monSite.monPoidsSensD
    maSaveTypeVit = monSite.monTypeVit
    maSaveVitSensM = monSite.maVitSensM
    maSaveVitSensD = monSite.maVitSensD
End Sub

Public Function RechercherOptimun() As Boolean
    'Recherche des vitesses optimuns avec les paramètres saisis
    'dans la fenêtre et affichage des résultats dans celle-ci
    
    Dim unPasSensM As Integer, unPasSensD As Integer
    Dim unB1 As Single 'valeur de la bande passante de vert du sens montant
    Dim unB2 As Single 'valeur de la bande passante de vert du sens descendant
    Dim unH As Single  'Temps écoulé entre les événements
                       '"Passage au vert montant" et "Fin de vert descendant"
    'Variable stockant les bandes passantes optimales
    Dim unB1opt As Single
    Dim unB2opt As Single
    Dim unCalculBande As Integer
    'Variable stockant les vitesses optimisant les bandes
    Dim uneVitOptM As Integer
    Dim uneVitOptD As Integer
    
    Dim uneVitInitM As Integer
    Dim uneVitInitD As Integer
    Dim unOpt As String
    
    Dim uneDureeVert As Single, unePosRef As Single
    Dim uneOrdonnee As Integer, unCarf As Carrefour
    Dim unNbCarf As Integer
    
    'Tableau stockant le minimun des durées de vert montante (resp descendante)
    'pour chaque vitesse montante (resp descendante)
    Dim unTabMinVertM(0 To 7) As Single
    Dim unTabMinVertD(0 To 7) As Single
    'Variable stockant les carrefours à double sens
    Dim uneColCarfDouble As New Collection
    Dim i As Integer, j As Integer
    Dim unCarfMExist As Boolean, unCarfDExist As Boolean
    
    'Tableau stockant par carrefour son type de sens (unique ou double)
    ReDim unTabSensDouble(1 To monSite.mesCarrefours.Count)
    
    'Tableau des carrefours réduits montants des carrefours à double sens
    'une valeur pour chaque vitesse montante, elles sont huit
    ReDim unTabCarfM(0 To 7, 1 To monSite.mesCarrefours.Count)
    
    'Tableau des carrefours réduits descendants des carrefours à double sens
    'une valeur pour chaque vitesse descendante, elles sont huit
    ReDim unTabCarfD(0 To 7, 1 To monSite.mesCarrefours.Count)

    'Initialisation
    unB1opt = 0
    unB2opt = 0
    unNbCarfDouble = 0
    unCarfMExist = False
    unCarfDExist = False
    
    'On fait un calcul d'onde à double sens à vitesse constante
    monSite.monTypeOnde = OndeDouble
    monSite.monTypeVit = VitConst
    
    'Initialisation de la vitesse montant avec la vitesse
    'montante de début de recherche
    TabOptiVit.Row = 1
    TabOptiVit.Col = 1
    monSite.maVitSensM = Format(TabOptiVit.Text)
    
    'Initialisation de la vitesse descendante avec la vitesse
    'descendante de début de recherche
    TabOptiVit.Col = 2
    monSite.maVitSensD = Format(TabOptiVit.Text)
    
    'Utilisation des pas de recherche de la fenetre Optimisation
    TabOptiVit.Row = 2
    TabOptiVit.Col = 1
    unPasSensM = Format(TabOptiVit.Text)
    TabOptiVit.Col = 2
    unPasSensD = Format(TabOptiVit.Text)
    
    'Utilisation des poids de la fenetre Optimisation
    TabOptiVit.Row = 3
    TabOptiVit.Col = 1
    monSite.monPoidsSensM = Format(TabOptiVit.Text)
    TabOptiVit.Col = 2
    monSite.monPoidsSensD = Format(TabOptiVit.Text)
   
    
    'Lancement du calcul
    uneVitInitM = monSite.maVitSensM
    uneVitInitD = monSite.maVitSensD
    
    'Test de non nullité des pas et vitesses montantes et descendantes
    If monSite.monPoidsSensM = 0 Or monSite.monPoidsSensD = 0 Or monSite.maVitSensM = 0 Or monSite.maVitSensD = 0 Or unPasSensM = 0 Or unPasSensD = 0 Then
        MsgBox "Les vitesses de début de recherche, les pas de recherche et les poids doivent être non nulles dans les deux sens", vbCritical
        RechercherOptimun = False
        Exit Function
    Else
        RechercherOptimun = True
    End If
    
    'Parcours de tous les carrefours passés en paramètres pour faire
    'les réductions dans le sens montant et pour trouver le minimun des
    'durées de vert montantes. Pour cela on fait toutes les vitesses sens M
    For K = 0 To 7
        'Initialisation du min des durées de vert montant
        unTabMinVertM(K) = monSite.maDuréeDeCycle
        'Affectation de la vitesse montant
        monSite.maVitSensM = uneVitInitM + K * unPasSensM
        'Affichage de la vitesse montante en ligne k+1, colonne 0
        TabCouple.Row = K + 1
        TabCouple.Col = 0
        TabCouple.Text = Format(monSite.maVitSensM)
        'Parcours de tous les carrefours
        unNbCarf = monSite.mesCarrefours.Count
        For i = 1 To unNbCarf
            Set unCarf = monSite.mesCarrefours(i)
            'On ne travaille que sur les carrefours choisis par l'utilisateur
            If unCarf.monIsUtil Then
                'Parcours de tous les feux du carrefour pour voir s'ils sont
                'tous dans le même sens ou dans deux sens différents
                'Juste à la première réduction car les carrefours ne changent
                'pas de sens lors d'un changement de vitesse M ou D
                If K = 0 Then
                    unTabSensDouble(i) = False
                    j = 2
                    'Test des sens de deux feux consécutifs
                    'Sortie si les sens sont différents ==> Carrefour à double sens
                    'Sinon Carrefour à sens unique celui du feu 1 par exemple
                    'Si un seul feu dans le carrefour, on ne rentre pas dans la boucle
                    '==> Carrefour à sens unique, celui du seul feu
                    Do While j <= unCarf.mesFeux.Count And unTabSensDouble(i) = False
                        If unCarf.mesFeux(j - 1).monSensMontant <> unCarf.mesFeux(j).monSensMontant Then
                            'Cas d'un carrefour ayant des feux dans les deux sens
                            unTabSensDouble(i) = True
                            'Stockage du carrefour double sens par son indice
                            uneColCarfDouble.Add i
                        End If
                        j = j + 1
                    Loop
                End If
                'Alimentation du Carrefour réduit i pour la Kème vitesse M
                If unTabSensDouble(i) Or (Not unTabSensDouble(i) And unCarf.mesFeux(1).monSensMontant) Then
                    'Cas d'un carrefour ayant des feux dans les deux sens ou
                    'tous ses feux dans le sens montant
                    'Calcul du feu équivalent dans le sens montant
                    unFeuEquivExist = CalculerFeuEquivalent(unCarf, True, uneDureeVert, unePosRef, uneOrdonnee)
                    If unFeuEquivExist = False Then
                        'Cas d'erreur
                        BoutonChangerVitesse.Enabled = False
                        'Masquage des controls se trouvant sous FrameSeparateur
                        Height = Height - ScaleHeight + FrameSeparateur.Top
                        RechercherOptimun = False
                        Exit Function
                    End If
                    If unTabSensDouble(i) Then
                        'Cas du carrefour à sens double,
                        'Il faut calculer la partie montante de l'écart
                        'Stockage des carrefours réduits pour la Kème vitesse M
                        unTabCarfM(K, i).maDureeVert = uneDureeVert
                        unTabCarfM(K, i).monEcart = unePosRef - uneOrdonnee / monSite.maVitSensM * 3.6
                    End If
                    'Stockage de la durée de vert minimun montante (1ère condition onde verte)
                    If uneDureeVert < unTabMinVertM(K) Then
                        unTabMinVertM(K) = uneDureeVert
                    End If
                Else
                    'Cas des carrefours ayant des feux uniquement descendant
                    unCarfDExist = True
                End If
            End If
        Next i
    Next K
    
    'Parcours de tous les carrefours passés en paramètres pour faire
    'les réductions dans le sens descendant et pour trouver le minimun des
    'durées de vert descendantes. Pour cela on fait toutes les vitesses sens D
    For K = 0 To 7
        'Initialisation du min des durées de vert descendante
        unTabMinVertD(K) = monSite.maDuréeDeCycle
        'Affectation de la vitesse descendante
        monSite.maVitSensD = uneVitInitD + K * unPasSensD
        'Affichage de la vitesse descendante en ligne 0, colonne k+1
        TabCouple.Row = 0
        TabCouple.Col = K + 1
        TabCouple.Text = Format(monSite.maVitSensD)
        'Parcours de tous les carrefours
        unNbCarf = monSite.mesCarrefours.Count
        For i = 1 To unNbCarf
            Set unCarf = monSite.mesCarrefours(i)
            'On ne travaille que sur les carrefours choisis par l'utilisateur
            If unCarf.monIsUtil Then
                'Alimentation des Carrefour réduit
                If unTabSensDouble(i) Or (Not unTabSensDouble(i) And unCarf.mesFeux(1).monSensMontant = False) Then
                    'Cas d'un carrefour ayant des feux dans les deux sens ou
                    'tous ses feux dans le sens descendant
                    'Calcul du feu équivalent dans le sens descendant
                    unFeuEquivExist = CalculerFeuEquivalent(unCarf, False, uneDureeVert, unePosRef, uneOrdonnee)
                    If unFeuEquivExist = False Then
                        'Cas d'erreur
                        BoutonChangerVitesse.Enabled = False
                        'Masquage des controls se trouvant sous FrameSeparateur
                        Height = Height - ScaleHeight + FrameSeparateur.Top
                        RechercherOptimun = False
                        Exit Function
                    End If
                    If unTabSensDouble(i) Then
                        'Cas du carrefour à sens double,
                        'Il faut calculer la partie descendante de l'écart
                        'Stockage des carrefours réduits pour la Kème vitesse D
                        unTabCarfD(K, i).maDureeVert = uneDureeVert
                        'On a + ordonnee/vitesse descendante car la formule est
                        '-D/V avec V algébrique, or ici vitesse descendante > 0 or
                        'algébriquement elle est < 0 d'où le - par - qui donne +
                        unTabCarfD(K, i).monEcart = unePosRef + uneOrdonnee / monSite.maVitSensD * 3.6 + uneDureeVert
                    End If
                    'Stockage de la durée de vert minimun descendante (1ère condition onde verte)
                    If uneDureeVert < unTabMinVertD(K) Then
                        unTabMinVertD(K) = uneDureeVert
                    End If
                Else
                    'Cas des carrefours ayant des feux uniquement montant
                    unCarfMExist = True
                End If
            End If
        Next i
    Next K
        
    'Calcul pour toutes les vitesses sans refaire les réductions
    For i = 0 To 7
        'Affectation de la vitesse montant
        monSite.maVitSensM = uneVitInitM + i * unPasSensM
        For j = 0 To 7
            'Affectation de la vitesse descendante
            monSite.maVitSensD = uneVitInitD + j * unPasSensD
        
            'Calcul des bandes passantes maximales
            unCalculBande = OptimiserCalculBandes(monSite, unB1, unB2, unH, unTabMinVertM, i, unTabMinVertD, j, uneColCarfDouble, unCarfMExist, unCarfDExist)
            If unCalculBande <> AucuneSolution Then
                'Cas où les bandes passantes existent
                'avec une solution à double sens
                
                'test si les bandes passantes calculées sont les optimuns
                If unB1 >= unB1opt And unB2 >= unB2opt Then
                    unB1opt = unB1
                    unB2opt = unB2
                    'Stockage des vitesses optimales
                    uneVitOptM = monSite.maVitSensM
                    uneVitOptD = monSite.maVitSensD
                End If
            End If
            'affichage dans le tableau des couples
            TabCouple.Row = i + 1
            TabCouple.Col = j + 1
            TabCouple.Text = Format(Int(unB1)) + " / " + Format(Int(unB2))
        Next j 'vitesse descendante suivante
    Next i 'vitesse montante suivante
    
    'Alimentation du tableau de résultat
    RemplirTabResult uneVitOptM, uneVitOptD, Int(unB1opt), Int(unB2opt)
    
    'Mise en rouge des cellules correspondant à l'optimun
    unOpt = Format(Int(unB1opt)) + " / " + Format(Int(unB2opt))
    For i = 1 To 8
        TabCouple.Row = i
        For j = 1 To 8
            TabCouple.Col = j
            If TabCouple.Text = unOpt Then
                TabCouple.ForeColor = RGB(255, 0, 0)
            Else
                TabCouple.ForeColor = monSite.ForeColor
            End If
        Next j
    Next i
    'Destruction de la collection des indices des carrefours vides
    Set uneColCarfDouble = Nothing
End Function


Private Sub TabCouple_Click(ByVal Col As Long, ByVal Row As Long)
    'Alimentation du tableau de résultat avec les
    'valeurs de la cellule cliquée
    Dim uneVM As Integer, uneVD As Integer
    Dim uneBM As Integer, uneBD As Integer
    
    TabCouple.Col = 0
    TabCouple.Row = Row
    uneVM = Val(TabCouple.Text)
    TabCouple.Col = Col
    TabCouple.Row = 0
    uneVD = Val(TabCouple.Text)
    
    TabCouple.Col = Col
    TabCouple.Row = Row
    unePosSlash = InStr(1, TabCouple.Text, "/")
    uneBM = Val(Mid$(TabCouple.Text, 1, unePosSlash))
    uneBD = Val(Mid$(TabCouple.Text, unePosSlash + 1))
    
    'Alimentation du tableau de résultat
    RemplirTabResult uneVM, uneVD, uneBM, uneBD
End Sub

Public Sub RemplirTabResult(uneVM As Integer, uneVD As Integer, uneBM As Integer, uneBD As Integer)
    TabResult.Row = 1
    TabResult.Col = 1
    TabResult.Text = Format(uneVM)
    TabResult.Col = 2
    TabResult.Text = Format(uneVD)
    TabResult.Row = 2
    TabResult.Col = 1
    TabResult.Text = Format(uneBM)
    TabResult.Col = 2
    TabResult.Text = Format(uneBD)
    'Stockage dans les vitesses constantes du site
    'si on click sur le bouton utiliser les vitesses optimales
    monSite.maVitSensM = uneVM
    monSite.maVitSensD = uneVD
End Sub


Private Sub TabOptiVit_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    'Cas d'une saisie
    BoutonChangerVitesse.Enabled = False
    'Masquage des controls se trouvant sous FrameSeparateur
    Height = Height - ScaleHeight + FrameSeparateur.Top
End Sub

Public Function OptimiserCalculBandes(uneForm As Form, unB1 As Single, unB2 As Single, unH As Single, unTabMinVertM() As Single, unI As Integer, unTabMinVertD() As Single, unJ As Integer, uneColCarfDouble As Collection, unCarfMExist As Boolean, unCarfDExist As Boolean) As Integer
    'Fonction cherchant les bandes passantes maximales
    'dans les deux sens
    Dim unS As Single 'Correspond au S des spécifs
    Dim unMin As Single, unMinLoc As Single
    Dim unPMsurPD As Single
    Dim unMinVertSensM As Single
    Dim unMinVertSensD As Single
    Dim unEcart_I As Single, unEcart_J As Single
    Dim unNbCarfDouble As Integer
    
    'Initialisation des durées de vert minimales
    unMinVertSensM = unTabMinVertM(unI)
    unMinVertSensD = unTabMinVertD(unJ)
    
    'Initialisation du unS qui est un maximun, pour le trouver dans
    'le code plus bas explicant la 2ème condition
    unS = -3 * uneForm.maDuréeDeCycle 'Car toutes les valeurs sont dans [0,durée du cycle[
    unNbCarfDouble = uneColCarfDouble.Count
    For ii = 1 To unNbCarfDouble
        '2ème condition sur l'onde verte s'il y a
        'des carrefours réduits à double sens
                
        'Utilisation des carrefours doubles par leur indice
        'stockée dans la collection uneColCarfDouble
        i = uneColCarfDouble(ii)
        'Calcul de l'écart du carrefour i
        unEcart_I = unTabCarfD(unJ, i).monEcart - unTabCarfM(unI, i).monEcart
        'On ramène l'écart modulo entre [0, duréee du cycle[
        unEcart_I = ModuloZeroCycle(unEcart_I, uneForm.maDuréeDeCycle)
        'Calcul sur tous les carrefours réduits double sens de la fonction :
        'Min(Z) = Minimun(DureeVertSensM_Carf_i + DureeVertSensD_Carf_i - Ecart_Carf_i(Z) + Z)
        'pour tout i variant de 1 à nombre de carrefours réduits double sens
        'et Z variant de monEcart du premier carrefour réduit double sens à
        'monEcart du dernier carrefour réduit double sens
        'En même temps on cherche unS = Max de ces min et on stocke dans unH
        'le Z correspondant au Max
        unMin = 3 * uneForm.maDuréeDeCycle 'Car toutes les valeurs sont dans [0,durée du cycle[
        For jj = 1 To unNbCarfDouble
            'Utilisation des carrefours doubles par leur indice
            'stockée dans la collection uneColCarfDouble
            j = uneColCarfDouble(jj)
            'Calcul de l'écart du carrefour j
            unEcart_J = unTabCarfD(unJ, j).monEcart - unTabCarfM(unI, j).monEcart
            'On ramène l'écart modulo entre [0, duréee du cycle[
            unEcart_J = ModuloZeroCycle(unEcart_J, uneForm.maDuréeDeCycle)
            unMinLoc = unTabCarfM(unI, j).maDureeVert + unTabCarfD(unJ, j).maDureeVert - Ecart(unEcart_J, unEcart_I, uneForm.maDuréeDeCycle) + unEcart_I
            If unMinLoc < unMin Then
                'Stockage du minimun
                unMin = unMinLoc
            End If
        Next jj
        
        'Stockage du maximun des minimuns sur tous les
        'carrefours réduits et l'écart réalisant ce max
        If unMin > unS Then
            unS = unMin
            unH = unEcart_I
        End If
    Next ii
    
    'Détermination des bandes passantes maximales
    If unNbCarfDouble = 0 Then
        'Cas où tous les carrefours sont à sens unique
        OptimiserCalculBandes = DoubleSensPossible
        If unCarfMExist = False Then
            'Cas où tous les carrefours à sens unique descendant
            unB1 = 0
            unB2 = unMinVertSensD
        ElseIf unCarfDExist = False Then
            'Cas où tous les carrefours à sens unique montant
            unB1 = unMinVertSensM
            unB2 = 0
        Else
            unB1 = unMinVertSensM
            unB2 = unMinVertSensD
            unH = 0 'Tous les H dans [0, Durée du cycle[ sont possibles
        End If
    Else
        'Cas où il y a des carrefours réduits à double sens
        'Dans l'optimisation on calcule des ondes double sens uniquement
        OptimiserCalculBandes = DoubleSensPossible
        'La valeur ci-dessus sera modifiée uniquement
        'si aucune solution trouvée
        unPMsurPD = uneForm.monPoidsSensM / uneForm.monPoidsSensD
        If unS <= 0 Then
            'Cas sans solution
            OptimiserCalculBandes = AucuneSolution
        ElseIf unS >= unMinVertSensM + unMinVertSensD Then
            unB1 = unMinVertSensM
            unB2 = unMinVertSensD
        ElseIf unS / (1 + 1 / unPMsurPD) <= unMinVertSensM And unS / (1 + unPMsurPD) <= unMinVertSensD Then
            unB1 = unS / (1 + 1 / unPMsurPD)
            unB2 = unS / (1 + unPMsurPD)
        ElseIf unS / (1 + 1 / unPMsurPD) > unMinVertSensM Then
            unB1 = unMinVertSensM
            unB2 = unS - unMinVertSensM
        ElseIf unS / (1 + unPMsurPD) > unMinVertSensD Then
            unB1 = unS - unMinVertSensD
            unB2 = unMinVertSensD
        End If
    End If
End Function


