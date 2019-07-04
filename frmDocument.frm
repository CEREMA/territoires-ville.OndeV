VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDocument"
   ClientHeight    =   7620
   ClientLeft      =   840
   ClientTop       =   660
   ClientWidth     =   10965
   Icon            =   "frmDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   10965
   Begin VB.TextBox TitreEtude 
      Height          =   285
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   1
      Top             =   120
      Width           =   6090
   End
   Begin VB.TextBox DuréeCycle 
      Height          =   285
      Left            =   9315
      MaxLength       =   3
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame FrameVisuCarf 
      Height          =   7095
      HelpContextID   =   61
      Left            =   60
      TabIndex        =   76
      Top             =   480
      Width           =   4095
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Axe des Ordonnées"
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
         Left            =   2280
         TabIndex        =   83
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Carrefour et Arrêt TC"
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
         TabIndex        =   82
         Top             =   150
         Width           =   1785
      End
      Begin VB.Label LabelTrait 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "LabelTrait"
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   1680
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image IconeArret 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmDocument.frx":0442
         Top             =   2460
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label NomArret 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10 nom tc num_____________"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   80
         Top             =   2520
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Origine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0_"
         Height          =   195
         Left            =   2940
         TabIndex        =   79
         Top             =   3000
         Width           =   180
      End
      Begin VB.Image IconeFeu 
         DragIcon        =   "frmDocument.frx":0784
         Height          =   240
         Index           =   0
         Left            =   2400
         Picture         =   "frmDocument.frx":0BC6
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label NomCarf 
         BackStyle       =   0  'Transparent
         Caption         =   "NOM CARREFOUR 20 CAR"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label NumFeu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10___"
         DragIcon        =   "frmDocument.frx":0E48
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   77
         Top             =   1080
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Line AxeOrdonnée 
         X1              =   3120
         X2              =   3120
         Y1              =   720
         Y2              =   6740
      End
   End
   Begin TabDlg.SSTab TabFeux 
      Height          =   6975
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Carrefours"
      TabPicture(0)   =   "frmDocument.frx":128A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LabelCarf"
      Tab(0).Control(1)=   "ComboNomCarf"
      Tab(0).Control(2)=   "RenameCarf"
      Tab(0).Control(3)=   "AjoutCarf"
      Tab(0).Control(4)=   "SupprCarf"
      Tab(0).Control(5)=   "AjoutFeu"
      Tab(0).Control(6)=   "SupprFeu"
      Tab(0).Control(7)=   "CarfSuiv"
      Tab(0).Control(8)=   "CarfPred"
      Tab(0).Control(9)=   "TabPropCarf"
      Tab(0).Control(10)=   "TabTrafCarf"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Transports collectifs"
      TabPicture(1)   =   "frmDocument.frx":12A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelTC"
      Tab(1).Control(1)=   "FrameTC"
      Tab(1).Control(2)=   "RenameTC"
      Tab(1).Control(3)=   "ComboTC"
      Tab(1).Control(4)=   "NewTC"
      Tab(1).Control(5)=   "DelTC"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Cadrage onde verte"
      TabPicture(2)   =   "frmDocument.frx":12C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabInfoCalc"
      Tab(2).Control(1)=   "FrameOndeTC"
      Tab(2).Control(2)=   "FrameTypeOnde"
      Tab(2).Control(3)=   "FramePoids"
      Tab(2).Control(4)=   "FrameVitesse"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Résultat décalages"
      TabPicture(3)   =   "frmDocument.frx":12DE
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "FrameTransDec"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "TabDecal"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TabBande"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Dessin onde verte"
      TabPicture(4)   =   "frmDocument.frx":12FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ZoneDessin"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Fiche Résultats"
      TabPicture(5)   =   "frmDocument.frx":1316
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LabelTypeOnde"
      Tab(5).Control(1)=   "LabelResCarf"
      Tab(5).Control(2)=   "LabelResTC"
      Tab(5).Control(3)=   "TabFicheCarf"
      Tab(5).Control(4)=   "TabFicheTC"
      Tab(5).Control(5)=   "TabFicheOnde"
      Tab(5).ControlCount=   6
      Begin FPSpread.vaSpread TabFicheOnde 
         Height          =   1035
         Left            =   -74880
         TabIndex        =   71
         Top             =   840
         Width           =   6000
         _Version        =   131077
         _ExtentX        =   10583
         _ExtentY        =   1826
         _StockProps     =   64
         Enabled         =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   2
         OperationMode   =   1
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":1332
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   5
         VisibleRows     =   2
      End
      Begin FPSpread.vaSpread TabFicheTC 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   75
         Top             =   4125
         Width           =   6255
         _Version        =   131077
         _ExtentX        =   11033
         _ExtentY        =   1931
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   10
         OperationMode   =   1
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":172D
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread TabTrafCarf 
         Height          =   795
         Left            =   -74880
         TabIndex        =   11
         Top             =   1050
         Width           =   6015
         _Version        =   131077
         _ExtentX        =   10610
         _ExtentY        =   1402
         _StockProps     =   64
         BorderStyle     =   0
         EditEnterAction =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   2
         RowHeaderDisplay=   2
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":1B8E
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   2
         VisibleRows     =   2
      End
      Begin FPSpread.vaSpread TabBande 
         Height          =   1050
         Left            =   120
         TabIndex        =   67
         Top             =   1800
         Width           =   5895
         _Version        =   131077
         _ExtentX        =   10398
         _ExtentY        =   1852
         _StockProps     =   64
         BackColorStyle  =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   2
         RowHeaderDisplay=   2
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":1EFC
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread TabInfoCalc 
         Height          =   2160
         Left            =   -74880
         TabIndex        =   62
         Top             =   2880
         Width           =   6015
         _Version        =   131077
         _ExtentX        =   10610
         _ExtentY        =   3810
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   10
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":228A
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread TabDecal 
         Height          =   3015
         Left            =   120
         TabIndex        =   69
         Top             =   3240
         Width           =   5895
         _Version        =   131077
         _ExtentX        =   10398
         _ExtentY        =   5318
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   10
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":260A
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread TabFicheCarf 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   73
         Top             =   2280
         Width           =   6255
         _Version        =   131077
         _ExtentX        =   11033
         _ExtentY        =   2778
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   10
         OperationMode   =   1
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":2969
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread TabPropCarf 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   98
         Top             =   1995
         Width           =   5955
         _Version        =   131077
         _ExtentX        =   10504
         _ExtentY        =   5953
         _StockProps     =   64
         EditEnterAction =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   5
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmDocument.frx":2DBB
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin VB.Frame FrameOndeTC 
         Caption         =   "Sens ---------------- TC --------------Bande"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   51
         Top             =   5160
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox ComboTCM 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox ComboTCD 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   750
            Width           =   1215
         End
         Begin VB.TextBox TextBTCD 
            Height          =   315
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   55
            Text            =   "15"
            Top             =   750
            Width           =   375
         End
         Begin VB.TextBox TextBTCM 
            Height          =   315
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   53
            Text            =   "15"
            Top             =   300
            Width           =   375
         End
         Begin VB.Label LabelSecD 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Left            =   3120
            TabIndex        =   93
            Top             =   800
            Width           =   255
         End
         Begin VB.Label LabelSecM 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Left            =   3120
            TabIndex        =   92
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descendant :"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   800
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Montant :"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.CommandButton CarfPred 
         Caption         =   "Carrefour précédent"
         Height          =   375
         Left            =   -72000
         TabIndex        =   14
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton CarfSuiv 
         Caption         =   "Carrefour suivant"
         Height          =   375
         Left            =   -70200
         TabIndex        =   15
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton SupprFeu 
         Caption         =   "Supprimer le feu"
         Height          =   375
         Left            =   -73440
         TabIndex        =   13
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton AjoutFeu 
         Caption         =   "Nouveau feu"
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton SupprCarf 
         Caption         =   "Supprimer"
         Height          =   315
         Left            =   -69600
         TabIndex        =   10
         Top             =   650
         Width           =   975
      End
      Begin VB.CommandButton AjoutCarf 
         Caption         =   "Nouveau"
         Height          =   315
         Left            =   -71695
         TabIndex        =   8
         Top             =   650
         Width           =   975
      End
      Begin VB.CommandButton DelTC 
         Caption         =   "Supprimer"
         Height          =   315
         Left            =   -70560
         TabIndex        =   20
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton NewTC 
         Caption         =   "Nouveau..."
         Height          =   315
         Left            =   -72840
         TabIndex        =   18
         Top             =   600
         Width           =   1050
      End
      Begin VB.ComboBox ComboTC 
         Height          =   315
         Left            =   -74280
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton RenameTC 
         Caption         =   "Renommer..."
         Height          =   315
         Left            =   -71700
         TabIndex        =   19
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton RenameCarf 
         Caption         =   "Renommer..."
         Height          =   315
         Left            =   -70700
         TabIndex        =   9
         Top             =   650
         Width           =   1050
      End
      Begin VB.ComboBox ComboNomCarf 
         Height          =   315
         Left            =   -74280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   650
         Width           =   2535
      End
      Begin VB.Frame FrameTypeOnde 
         Caption         =   "Type d'onde verte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1250
         Left            =   -74880
         TabIndex        =   43
         Top             =   600
         Width           =   2415
         Begin VB.OptionButton OptionTC 
            Caption         =   "Prise en compte des TC"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   2200
         End
         Begin VB.OptionButton OptionOndeDouble 
            Caption         =   "Double sens"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OptionSensM 
            Caption         =   "Sens montant privilégié"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton OptionSensD 
            Caption         =   "Sens descendant privilégié"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   2200
         End
      End
      Begin VB.Frame FramePoids 
         Caption         =   "Poids des sens"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   -72360
         TabIndex        =   48
         Top             =   600
         Width           =   1935
         Begin VB.TextBox TextPoidsM 
            Height          =   300
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   49
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TextPoidsD 
            Height          =   300
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   50
            Text            =   "1"
            Top             =   650
            Width           =   375
         End
         Begin ComCtl2.UpDown UpDownSensM 
            Height          =   360
            Left            =   1560
            TabIndex        =   86
            Top             =   240
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   635
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "TextPoidsM"
            BuddyDispid     =   196649
            OrigLeft        =   2040
            OrigTop         =   240
            OrigRight       =   2280
            OrigBottom      =   600
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDownSensD 
            Height          =   360
            Left            =   1560
            TabIndex        =   87
            Top             =   645
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   635
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "TextPoidsD"
            BuddyDispid     =   196650
            OrigLeft        =   2040
            OrigTop         =   650
            OrigRight       =   2280
            OrigBottom      =   1010
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Montant :"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Descendant :"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   645
            Width           =   960
         End
      End
      Begin VB.Frame FrameVitesse 
         Caption         =   "Vitesse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   56
         Top             =   1920
         Width           =   6015
         Begin VB.OptionButton OptionVitConst 
            Caption         =   "Vitesse constante"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptionVitVar 
            Caption         =   "Vitesse variable"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TextVitM 
            Height          =   285
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   59
            Text            =   "35"
            Top             =   160
            Width           =   300
         End
         Begin VB.TextBox TextVitD 
            Height          =   285
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   60
            Text            =   "35"
            Top             =   480
            Width           =   300
         End
         Begin VB.CommandButton BoutonOptimun 
            Caption         =   "Optimisation des vitesses ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4400
            TabIndex        =   61
            Top             =   160
            Width           =   1455
         End
         Begin VB.Label LabelVitSensM 
            AutoSize        =   -1  'True
            Caption         =   "Sens montant en km/h :"
            Height          =   195
            Left            =   1920
            TabIndex        =   85
            Top             =   240
            Width           =   1710
         End
         Begin VB.Label LabelVitSensD 
            AutoSize        =   -1  'True
            Caption         =   "Sens descendant en km/h :"
            Height          =   195
            Left            =   1920
            TabIndex        =   84
            Top             =   480
            Width           =   1980
         End
      End
      Begin VB.Frame FrameTransDec 
         Caption         =   "Résultats du calcul d'onde verte  à double sens"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   5895
         Begin VB.TextBox TextTransDec 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   65
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton BoutonTrans 
            Caption         =   "Translater les décalages modifiables de :"
            Height          =   435
            Left            =   120
            TabIndex        =   64
            Top             =   290
            Width           =   3375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "secondes"
            Height          =   195
            Left            =   4200
            TabIndex        =   97
            Top             =   380
            Width           =   690
         End
      End
      Begin VB.Frame FrameTC 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   21
         Top             =   860
         Width           =   6015
         Begin FPSpread.vaSpread TabYArret 
            Height          =   1245
            Left            =   120
            TabIndex        =   38
            Top             =   2400
            Width           =   4065
            _Version        =   131077
            _ExtentX        =   7170
            _ExtentY        =   2196
            _StockProps     =   64
            ColHeaderDisplay=   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   10
            MaxRows         =   3
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   1
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmDocument.frx":31C3
            UnitType        =   2
            UserResize      =   0
            VisibleCols     =   500
            VisibleRows     =   500
         End
         Begin VB.TextBox TextDistAF_TC 
            Height          =   285
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   26
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox TextDuréeAF_TC 
            Height          =   285
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   29
            Top             =   900
            Width           =   735
         End
         Begin VB.CommandButton NewArret 
            Caption         =   "Nouvel arrêt"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   4140
            Width           =   1335
         End
         Begin VB.CommandButton DelArret 
            Caption         =   "Supprimer l'arrêt"
            Height          =   315
            Left            =   1560
            TabIndex        =   42
            Top             =   4140
            Width           =   1335
         End
         Begin VB.TextBox TextTDep 
            Height          =   285
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   23
            Top             =   180
            Width           =   495
         End
         Begin VB.CommandButton BoutonInverser 
            Caption         =   "Inverser"
            Height          =   495
            Left            =   4920
            Picture         =   "frmDocument.frx":34F2
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1395
            Width           =   735
         End
         Begin VB.TextBox TextArret 
            Height          =   285
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   40
            Top             =   3765
            Width           =   3015
         End
         Begin VB.PictureBox ColorTC 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   37
            ToolTipText     =   "Cliquer pour changer la couleur"
            Top             =   2055
            Width           =   255
         End
         Begin VB.ComboBox ComboCarfDep 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1260
            Width           =   3015
         End
         Begin VB.ComboBox ComboCarfArr 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1620
            Width           =   3015
         End
         Begin VB.Label LabelDistAF 
            AutoSize        =   -1  'True
            Caption         =   "Distance d'accélération et de freinage :"
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
            TabIndex        =   25
            Top             =   540
            Width           =   3390
         End
         Begin VB.Label LabelMetre 
            AutoSize        =   -1  'True
            Caption         =   "mètres"
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
            Left            =   4440
            TabIndex        =   27
            Top             =   540
            Width           =   570
         End
         Begin VB.Label LabelDuréeAF 
            AutoSize        =   -1  'True
            Caption         =   "Durée d'accélération et de freinage :"
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
            TabIndex        =   28
            Top             =   900
            Width           =   3150
         End
         Begin VB.Label LabelSec2 
            AutoSize        =   -1  'True
            Caption         =   "secondes"
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
            Left            =   4440
            TabIndex        =   30
            Top             =   900
            Width           =   825
         End
         Begin VB.Label LabelCarfArr 
            AutoSize        =   -1  'True
            Caption         =   "Carrefour arrivée : "
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
            TabIndex        =   33
            Top             =   1620
            Width           =   1620
         End
         Begin VB.Label LabelCarfDep 
            AutoSize        =   -1  'True
            Caption         =   "Carrefour départ : "
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
            TabIndex        =   31
            Top             =   1260
            Width           =   1575
         End
         Begin VB.Label LabelSec1 
            AutoSize        =   -1  'True
            Caption         =   "secondes"
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
            Left            =   2400
            TabIndex        =   24
            Top             =   180
            Width           =   825
         End
         Begin VB.Label LabelTDep 
            AutoSize        =   -1  'True
            Caption         =   "Instant de départ :"
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
            TabIndex        =   22
            Top             =   180
            Width           =   1590
         End
         Begin VB.Label LabelColTC 
            AutoSize        =   -1  'True
            Caption         =   "Couleur de représentation :"
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
            TabIndex        =   36
            Top             =   2055
            Width           =   2325
         End
         Begin VB.Label LabelArret 
            AutoSize        =   -1  'True
            Caption         =   "Libellé de l'arrêt : "
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
            TabIndex        =   39
            Top             =   3780
            Width           =   1560
         End
      End
      Begin VB.PictureBox ZoneDessin 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   -74940
         ScaleHeight     =   2145
         ScaleWidth      =   5865
         TabIndex        =   94
         Top             =   560
         Width           =   5895
         Begin VB.Label InfoModif 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "InfoModif"
            Height          =   195
            Left            =   2640
            TabIndex        =   96
            Top             =   840
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Image PtRef 
            Height          =   180
            Left            =   840
            Picture         =   "frmDocument.frx":3A24
            Stretch         =   -1  'True
            Top             =   1320
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Line PlageVert 
            BorderColor     =   &H80000002&
            BorderWidth     =   2
            Index           =   0
            Visible         =   0   'False
            X1              =   840
            X2              =   1800
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Shape PoigneeGauche 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000002&
            FillColor       =   &H80000002&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   600
            Top             =   840
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Shape PoigneeDroite 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000002&
            FillColor       =   &H80000002&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   1560
            Top             =   840
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label LabelFleche 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ">"
            Height          =   195
            Left            =   3120
            TabIndex        =   95
            Top             =   1920
            Width           =   90
         End
         Begin VB.Line AxeTemps 
            X1              =   120
            X2              =   3120
            Y1              =   2040
            Y2              =   2040
         End
      End
      Begin VB.Label LabelResTC 
         AutoSize        =   -1  'True
         Caption         =   "Résultat par Transport Collectif"
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
         Left            =   -74880
         TabIndex        =   74
         Top             =   3900
         Width           =   2670
      End
      Begin VB.Label LabelResCarf 
         AutoSize        =   -1  'True
         Caption         =   "Résultat par carrefour"
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
         Left            =   -74880
         TabIndex        =   72
         Top             =   1980
         Width           =   1875
      End
      Begin VB.Label LabelTypeOnde 
         AutoSize        =   -1  'True
         Caption         =   "Résultat du calcul d'onde verte"
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
         Left            =   -74880
         TabIndex        =   70
         Top             =   600
         Width           =   2685
      End
      Begin VB.Label LabelCarf 
         AutoSize        =   -1  'True
         Caption         =   "Nom :"
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
         Left            =   -74880
         TabIndex        =   6
         Top             =   645
         Width           =   510
      End
      Begin VB.Label LabelTC 
         AutoSize        =   -1  'True
         Caption         =   "Nom :"
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
         Left            =   -74820
         TabIndex        =   16
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Modification des décalages en secondes par carrefour :"
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
         TabIndex        =   68
         Top             =   3000
         Width           =   4755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Largeur maximun des bandes passantes en secondes :"
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
         TabIndex        =   66
         Top             =   1560
         Width           =   4650
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "secondes"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   165
      Width           =   825
   End
   Begin VB.Label Label3 
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
      Left            =   7860
      TabIndex        =   2
      Top             =   165
      Width           =   1485
   End
   Begin VB.Label LabelTitre 
      AutoSize        =   -1  'True
      Caption         =   "Titre de l'étude :"
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
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attributs du site et donc de la fenêtre fille
Public monFichId As Integer
Public mesCarrefours As New ColCarrefour
Public monCarrefourCourant As Carrefour
Public monTitreEtude As String
Public maDuréeDeCycle As Integer
Public maLongueurAxeY As Long
Public mesTC As New ColTC
Public monTypeOnde As Integer 'Type onde verte : Double sens, sens Montant ou Descendant
Public monPoidsSensM As Integer
Public monPoidsSensD As Integer
Public monTypeVit As Integer 'Type de vitesse par carrefour : Constant ou variable
Public maVitSensM As Integer 'Vitesse constant sens montant
Public maVitSensD As Integer 'Vitesse constant sens descendant
Public maVitMaxM As String   'Vitesse maximale possible sens montant
Public maVitMaxD As String   'Vitesse maximale possible sens descendant
Public maTransDec As Integer 'Valeur de la translation de tous les décalages
'Variable indiquant la cohérence entre les données
'et les résultats du calcul d'onde
'valeur OK ou CalculImpossible ou IncoherenceDonneeCalcul
Public maCoherenceDataCalc As Integer
'Collection des TC utilisés (Ceux pris en compte dans l'onde verte
'+ ceux dont on détermine la progression)
Public mesTCutil As New Collection
'Nombre d'objets graphiques total représentant un carrefour dans la fenetre
Public monNbObjGraphicCarf As Integer
'Nombre d'objets graphiques total représentant un feu dans la fenetre
Public monNbObjGraphicFeu As Integer
'Nombre d'objets graphiques total représentant un TC dans la fenetre
Public monNbObjGraphicTC As Integer
'Variables donnant le type d'objet sélectionné graphiquement et son index
Public monObjSel As Integer
Public monIndSel As Integer
'Variable stockant les Y minimun et maximun des feux de tous les carrefours,
'utilisés ou non dans le calcul de l'onde, pour faire un zoom maximun
'Cela correspond à l'englobant maximun
Public monYMinFeu As Integer
Public monYMaxFeu As Integer
'Variable stockant les Y minimun et maximun des feux des carrefours utilisés
'dans le calcul de l'onde verte, pour trouver le niveau de zoom maximun lors
'd'une impression ou d'un affichage du graphique de l'onde en pleine écran
Public monYMaxFeuUtil As Long
Public monYMinFeuUtil As Long
'Variable contenant les largeurs de bandes passantes calculées
Public maBandeM As Single
Public maBandeD As Single
'Variable contenant les largeurs de bandes passantes modifiées par l'utilisateur
Public maBandeModifM As Single
Public maBandeModifD As Single
'Variable contenant les largeurs de bandes passantes imposées pour les TC
Public maBandeTCM As Single
Public maBandeTCD As Single
'Variable contenant les positions, dans la liste des TC du site,
'des TC pris en compte pour l'onde verte
Public monTCM As Integer
Public monTCD As Integer
'Collection contenant les carrefours réduits pour le calcul de l'onde verte
'Carrefours à sens unique montant
Public mesCarfReduitsSensM As New ColCarfReduitSensUnique
'Carrefours à sens unique montant
Public mesCarfReduitsSensD As New ColCarfReduitSensUnique
'Carrefours à sens unique montant
Public mesCarfReduitsSens2 As New ColCarfReduitSensDouble
'Variable indiquant une modification de valeurs
'du site (titre étude, renommer Carrefour, TC et/ou Arrêts TC)
Public maModifDataSite  As Boolean
'Variable indiquant une modification de valeurs dans les carrefours
Public maModifDataCarf  As Boolean
'Variable indiquant une modification de valeurs dans les TC
'ne cadrant pas l'onde verte
Public maModifDataTC  As Boolean
'Variable indiquant une modification de valeurs dans les TC
'cadrant l'onde verte
Public maModifDataOndeTC  As Boolean
'Variable indiquant une modification de valeurs dans les calculs d'onde
Public maModifDataOnde  As Boolean
'Variable indiquant une modification de valeurs dans les décalages
Public maModifDataDec As Boolean
'Variable indiquant une modification de valeurs dans la visu graphique
Public maModifDataDes As Boolean
'Variable stockant les options d'affichage et d'impression
Public mesOptionsAffImp As New OptionsAffImp
'Variable stockant si on a une onde verte dans les deux sens
'Elle ne sert que dans le cas d'un sens privilégié
'==> Il faudra les dessiner toutes les deux si vrai,
'sinon dessin uniquement de celle du sens privilégié
Public monOndeDoubleTrouve As Boolean

'Variable stockant les englobants réelles en Temps et en Y
'Calculées dans DessinerOndeVerte et utilisées dans TracerProgressionTC
Public monDYTotal As Long
Public monTmpTotal As Long

'Variable stockant le minimun en Y et en T en coordonnées réelles
'Calculées dans DessinerOndeVerte et utilisées dans TracerProgressionTC
Public monYMin As Long
Public monTMin As Long
Public monOrigX As Long

'Variables stockant l'englobant en temps pour les progressions des TC
Public monTDepTCMin As Long
Public monTFinTCMax As Long

'Collection stockant les plages de vert graphiques des ondes vertes montantes
'et descendantes pour les sélections et les modifications interactives
'Ces collections contiennent des instances de la classe PlageGraphic
'Remis à jour à chaque dessin d'onde verte (Procédure DessinerOndeVerte)
Public maColPlageGraphicD As New Collection 'Sens descendant
Public maColPlageGraphicM As New Collection 'Sens montant

'Collection stockant les points de référence graphiques des ondes vertes montantes
'et descendantes pour les sélections et les modifications interactives
'Ces collections contiennent des instances de la classe refGraphic
'Remis à jour à chaque dessin d'onde verte (Procédure DessinerOndeVerte)
Public maColRefGraphicD As New Collection 'Sens descendant
Public maColRefGraphicM As New Collection 'Sens montant

'Collection stockant les valeurs de décalages calculés et modifiés
'avant une modif de date imposée ou de décalage
Private maColDecSave As New Collection

'Variable stockant le nombre de feux pris pour le calcul du feu équivalent
'dans le sens montant et dans le sens descendant
Public monNbFeuxMpris As Integer
Public monNbFeuxDpris As Integer

'Variable stockant si on dessine les bandes inter-carrefours voitures
'quand on est en onde verte cadrée par un TC montant et/ou un TC descendant
Public monDessinInterCarfVP As Boolean

Private Sub AjoutCarf_Click()
    CreerCarrefour Me
End Sub

Private Sub AjoutFeu_Click()
    CreerFeu Me
End Sub

Private Sub BoutonInverser_Click()
    Dim unCarfTmp As Carrefour
    Dim unePosTC As Integer
    
    'Récupération de la position du carrefour
    unePosTC = ComboTC.ListIndex + 1
    'Inversion des carrefours de départ et d'arrivée
    'si les paramètres pour le calcul d'onde verte TC l'autorise
    If ChangerParamOndeTC(Me, unePosTC, mesTC(unePosTC).monCarfArr, mesTC(unePosTC).monCarfDep) Then
        'Inversion des carrefours de départ et d'arrivée
        Set unCarfTmp = mesTC(unePosTC).monCarfArr
        Set mesTC(unePosTC).monCarfArr = mesTC(unePosTC).monCarfDep
        Set mesTC(unePosTC).monCarfDep = unCarfTmp
        'Mise à vide de l'affichage de la combobox pour ne pas déclencher
        'd'erreur lors de la vérification de la différence entre les carrefours
        'de départ et d'arrivée lors du click event de ComboCarDep
        ComboCarfDep.ListIndex = -1
        'Mise à jour des combobox, d'abord le carrefour d'arrivée pour que
        'les indices de deux combobox ne soient pas les mêmes
        '==> Pas de message d'erreur dans le click des combobox
        ComboCarfArr.ListIndex = mesTC(unePosTC).monCarfArr.maPosition - 1
        ComboCarfDep.ListIndex = mesTC(unePosTC).monCarfDep.maPosition - 1
        'Indication d'un changement de données TC
        IndiquerModifTC
    End If
End Sub

Private Sub BoutonOptimun_Click()
    'Affichage de la fenetre de recherche d'un couple de vitesse
    'optimisant les bandes passantes
    frmOptiVit.Show vbModal
End Sub

Private Sub BoutonTrans_Click()
    TranslaterDecalages
End Sub

Public Sub TranslaterDecalages()
    'Stockage de la translation effectuée
    maTransDec = Val(TextTransDec.Text)
    
    'Stockage d'une modification de valeurs dans les décalages
    'Ceci permettra aussi de demander une sauvegarde à la fermeture
    maModifDataDec = True
    
    'Translation de tous les décalages modifiables des carrefours
    For i = 1 To mesCarrefours.Count
        If mesCarrefours(i).monDecCalcul <> -99 Then
            'Stockage dans les instances de carrefours
            mesCarrefours(i).monDecModif = mesCarrefours(i).monDecModif + maTransDec
            mesCarrefours(i).monDecModif = ModuloZeroCycle(mesCarrefours(i).monDecModif, maDuréeDeCycle)
            'Affichage dans l'onglet Tableau de résultat en arrondissant à l'entier
            'le plus proche, d'où l'utilisation de la fonction VB5 CInt
            TabDecal.Row = i
            TabDecal.Col = 3
            If CIntCorrigé(mesCarrefours(i).monDecModif) = maDuréeDeCycle Then
                'Une valeur valant durée du cycle s'affiche 0
                TabDecal.Text = "0"
            Else
                TabDecal.Text = CIntCorrigé(mesCarrefours(i).monDecModif)
            End If
        End If
    Next i
End Sub

Private Sub CarfPred_Click()
    'Récupération de la position du carrefour courant
    'dans la collection des carrefours du site courant
    unInd = monCarrefourCourant.maPosition
    If unInd = 1 Then
        MsgBox "Ce carrefour est le premier. Aucun carrefour ne le précéde", vbCritical
    Else
        'Mise à jour sélection graphique et l'onglet Carrefour
        Call MiseAJourSelectionEtOngletCarrefour(Me, CarfSel, unInd - 1, 1)
    End If
End Sub

Private Sub CarfSuiv_Click()
    'Récupération de la position du carrefour courant
    'dans la collection des carrefours du site courant
    unInd = monCarrefourCourant.maPosition
    If unInd = mesCarrefours.Count Then
        MsgBox "Ce carrefour est le dernier. Aucun carrefour ne le suit", vbCritical
    Else
        'Mise à jour sélection graphique et l'onglet Carrefour
        Call MiseAJourSelectionEtOngletCarrefour(Me, CarfSel, unInd + 1, 1)
    End If
End Sub

Private Sub ColorTC_Click()
    frmMain.ChangerCouleurPicBox ColorTC
    'Changement de la couleur des noms de tous les arrêts du TC courant
    Call ModifierObjGraphicTC(ModifColTC)
    'Indication d'un changement de données TC
    maModifDataTC = True
End Sub



Private Sub ComboCarfDep_Click()
    If ComboCarfDep.ListIndex <> -1 Then
        'Cas où ListIndex n'a pas été mis à -1. Il vaut -1
        'pour éviter les erreurs dues aux vérifications de différences
        'entre les carrefours de départ et d'arrivée, avant chaque affectation
        'en même temps de ComboCarfDep et ComboCarfArr
        
        'Test de différence entre les carrefours de départ et d'arrivée
        'Les listes de carrefours sont identiques et ordonnées de la même façon
        'dans les deux combobox ComboCarfDep et ComboCarfArr
        If ComboCarfDep.ListIndex = ComboCarfArr.ListIndex Then
            MsgBox "Erreur : Les carrefours de départ et d'arrivée sont identiques", vbCritical
            'On restaure le carrefour de départ précédent
            ComboCarfDep.ListIndex = mesTC(ComboTC.ListIndex + 1).monCarfDep.maPosition - 1
        Else
            'Changement du carrefour de départ
            If ChangerParamOndeTC(Me, ComboTC.ListIndex + 1, mesCarrefours(ComboCarfDep.ListIndex + 1), mesTC(ComboTC.ListIndex + 1).monCarfArr) Then
                Set mesTC(ComboTC.ListIndex + 1).monCarfDep = mesCarrefours(ComboCarfDep.ListIndex + 1)
                IndiquerModifTC
            Else
                'Restauration du carrefour de départ précédent
                ComboCarfDep.ListIndex = mesTC(ComboTC.ListIndex + 1).monCarfDep.maPosition - 1
            End If
        End If
    End If
End Sub

Private Sub ComboCarfArr_Click()
    If ComboCarfArr.ListIndex <> -1 Then
        'Cas où ListIndex n'a pas été mis à -1. Il vaut -1
        'pour éviter les erreurs dues aux vérifications de différences
        'entre les carrefours de départ et d'arrivée, avant chaque affectation
        'en même temps de ComboCarfDep et ComboCarfArr
        
        'Test de différence entre les carrefours de départ et d'arrivée
        'Les listes de carrefours sont identiques et ordonnées de la même façon
        'dans les deux combobox ComboCarfDep et ComboCarfArr
        If ComboCarfDep.ListIndex = ComboCarfArr.ListIndex Then
            MsgBox "Erreur : Les carrefours de départ et d'arrivée sont identiques", vbCritical
            'On restaure le carrefour d'arrivée précédent
            ComboCarfArr.ListIndex = mesTC(ComboTC.ListIndex + 1).monCarfArr.maPosition - 1
        Else
            'On change le carrefour d'arrivée
            If ChangerParamOndeTC(Me, ComboTC.ListIndex + 1, mesTC(ComboTC.ListIndex + 1).monCarfDep, mesCarrefours(ComboCarfArr.ListIndex + 1)) Then
                Set mesTC(ComboTC.ListIndex + 1).monCarfArr = mesCarrefours(ComboCarfArr.ListIndex + 1)
                IndiquerModifTC
            Else
                'Restauration du carrefour de départ précédent
                ComboCarfArr.ListIndex = mesTC(ComboTC.ListIndex + 1).monCarfArr.maPosition - 1
            End If
        End If
    End If
End Sub

Private Sub ComboNomCarf_Click()
    Dim unInd As Long
    
    If ComboNomCarf.Tag = "Déroulé par Click souris" Then
        'Cas où la combobox a été activé par un click dans sa flèche
        'sinon la combobox a été déroulé par programme lors d'une affectation
        'du genre combobox.listindex = un entier ==> On ne fait rien
        
        'Remise à vide du tag qui a été mis dans l'event ComboNomCarf_DropDown
        ComboNomCarf.Tag = ""
        'Récupération de la position du carrefour choisi
        'dans les listes des noms carrefours de la combobox ComboNomCarf
        unInd = ComboNomCarf.ListIndex + 1 'car les combobox vont de 0 à n-1
        'Mise à jour sélection graphique et l'onglet Carrefour
        Call MiseAJourSelectionEtOngletCarrefour(Me, CarfSel, unInd, 1)
    End If
    
    'Affichage des demandes et de débit de saturation dans TabTrafCarf
    TabTrafCarf.Row = 1
    TabTrafCarf.Col = 1
    TabTrafCarf.Text = monCarrefourCourant.maDemandeM
    TabTrafCarf.Col = 2
    TabTrafCarf.Text = monCarrefourCourant.monDebSatM
    TabTrafCarf.Row = 2
    TabTrafCarf.Col = 1
    TabTrafCarf.Text = monCarrefourCourant.maDemandeD
    TabTrafCarf.Col = 2
    TabTrafCarf.Text = monCarrefourCourant.monDebSatD
End Sub

Private Sub ComboNomCarf_DropDown()
    ComboNomCarf.Tag = "Déroulé par Click souris"
End Sub


Private Sub ComboTC_Click()
    'Remise à jour des zones de l'onglet Transport collectif
    'avec le TC choisi
    Dim unInd As Long
    
    'Calcul de l'index dans la collection des TC à partir de
    'la combobox comboTC
    unInd = ComboTC.ListIndex + 1
    'Remplissage de FrameTC avec les valeurs du TC numéro unInd
    RemplirFrameTC Me, unInd
    'Sélection graphique de l'arrêt TC correspondant
    'à la cellule active du spread TabYArret
    MiseAJourSelectionParCellule Me, ArretSel, unInd, TabYArret.ActiveCol
End Sub

Private Sub ComboTCD_Click()
    'Changement du TC cadrant l'onde verte descendante
    unTCD = TrouverTCParNom(Me, ComboTCD.Text)
    'Test si on choisi un autre TC que celui précédemment choisi
    If unTCD <> monTCD Then
        monTCD = unTCD
        'Mise à jour du TC cadrant l'onde verte en sens descendant
        If monTCD = 0 Then
            'Masquage de la bande passante descendante
            TextBTCD.Visible = False
            LabelSecD.Visible = False
        Else
            'Afichage de la bande passante descendante
            TextBTCD.Visible = True
            LabelSecD.Visible = True
        End If
        'Indication d'un modification des données de calcul de l'onde verte
        maModifDataOnde = True
    End If
End Sub

Private Sub ComboTCM_Click()
    'Changement du TC cadrant l'onde verte montante
    unTCM = TrouverTCParNom(Me, ComboTCM.Text)
    'Test si on choisi un autre TC que celui précédemment choisi
    If unTCM <> monTCM Then
        monTCM = unTCM
        'Mise à jour du TC cadrant l'onde verte en sens montant
        If monTCM = 0 Then
            'Masquage de la bande passante montante
            TextBTCM.Visible = False
            LabelSecM.Visible = False
        Else
            'Afichage de la bande passante montante
            TextBTCM.Visible = True
            LabelSecM.Visible = True
        End If
        'Indication d'un modification des données de calcul de l'onde verte
        maModifDataOnde = True
    End If
End Sub

Private Sub DelArret_Click()
    SupprimerArretTC Me
End Sub

Private Sub DelTC_Click()
    Dim uneListeY As New Collection
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unNbArretsConfondus As Integer
    Dim unControl As Control
        
    unMsg = "Etes-vous sûr de vouloir supprimer le transport collectif " + UCase(ComboTC.Text) + " ?"
    If MsgBox(unMsg, vbYesNo + vbQuestion) = vbYes Then
        'Récupération de la position du TC dans la liste des TC
        unePos = ComboTC.ListIndex
        'Vérification de l'utilisation de ce TC dans une onde verte TC
        If unePos + 1 = monTCM Or unePos + 1 = monTCD Then
            unMsg = "Impossible de supprimer " + mesTC(unePos + 1).monNom
            unMsg = unMsg + " car il est utilisé dans le calcul d'onde verte prenant en compte des TC"
            MsgBox unMsg, vbCritical
            Exit Sub
        Else
            If DonnerYCarrefour(mesTC(unePos + 1).monCarfDep) < DonnerYCarrefour(mesTC(unePos + 1).monCarfArr) Then
                'Suppression dans la liste des TC montant
                i = -1
                Do
                    i = i + 1
                Loop Until mesTC(unePos + 1).monNom = ComboTCM.List(i)
                ComboTCM.RemoveItem i
            Else
                'Suppression dans la liste des TC descendant
                i = -1
                Do
                    i = i + 1
                Loop Until mesTC(unePos + 1).monNom = ComboTCD.List(i)
                ComboTCD.RemoveItem i
            End If
            
            'Suppression dans la liste des TC utilisés
            '(ceux cadrant les ondes vertes M et/ou D et
            'ceux dont on veut afficher la progression)
            unTCtrouv = False
            i = 1
            Do While unTCtrouv = False And i <= mesTCutil.Count
                If mesTC(unePos + 1).monNom = mesTCutil(i).monNom Then
                    unTCtrouv = True
                    mesTCutil.Remove i
                End If
                i = i + 1
            Loop
        End If
        
        'Désélection de la sélection graphique précédente
        Call Deselectionner(Me)
        'Suppression des objets graphiques du TC
        Call ModifierObjGraphicTC(SupprTC)
        'Récupération des Y des arrêts du TC avant sa suppression
        For i = 1 To mesTC(unePos + 1).mesArrets.Count
            uneListeY.Add mesTC(unePos + 1).mesArrets(i).monOrdonnee
        Next i
        'Suppression dans la collection des TC du site
        mesTC.Remove (unePos + 1)
        'Suppression dans la liste des TC de la comboBox comboTC
        ComboTC.RemoveItem (unePos)
        'Mise à jour des noms d'arrêts qui étaient confondus avec
        'ceux du TC détruit
        For i = 1 To uneListeY.Count
            'Recherche des arrêts confondus en un Y pour alimenter
            'les listes d'arrêts et de TC trouvés
            unNb = RechercherArretConfondu(uneListeY(i), uneListeIndexTC, uneListeIndexArret)
            'Mise à jour des décalages des labels NomArrêt
            Call MiseAJourNomArret(Me, uneListeIndexTC, uneListeIndexArret)
            'On vide les listes pour le i suivant
            ViderCollection uneListeIndexTC
            ViderCollection uneListeIndexArret
        Next i
        'affichage dans ComboTC de l'élément précédent celui supprimé
        If ComboTC.ListCount > 0 Then
            If ComboTC.ListCount = unePos Then
                'Cas du dernier élément détruit
                ComboTC.ListIndex = unePos - 1
            Else
                ComboTC.ListIndex = unePos
            End If
            'Remise à jour des tags des NomArret donnés par mesTC.mesObjGraphics
            'et IconeArret des TC qui suivait celui supprimé
            For i = unePos + 1 To mesTC.Count
                unNbArret = mesTC(i).mesArrets.Count
                For j = 1 To unNbArret
                    Set unControl = mesTC(i).mesObjGraphics(j)
                    unControl.Tag = Format(i) + "-" + Format(j)
                    IconeArret(unControl.Index).Tag = unControl.Tag
                Next j
            Next i
        Else
            'Cas où il n'y a plus de TC
            FrameTC.Visible = False
            'Inhibition des boutons de TC
            RenameTC.Enabled = False
            DelTC.Enabled = False
        End If
        'Indication d'une modification dans les données TC
        maModifDataTC = True
    End If
End Sub



Private Sub DuréeCycle_GotFocus()
    'Focus donné aux onglets de TabFeux pour éviter un bouclage
    TabFeux.SetFocus
    'Cas d'un nombre de fois pair où on est entré
    frmModifCycle.Show vbModal
End Sub

Private Sub Form_Activate()
    'Remise à jour des CarfY pour le site courant
    'Ainsi les calculs et les dessins seront juste pour ce site
    'Recalcul du tableau monTabCarfY car le stockage par pointeur
    'de type variant ne marche pas
    If Not (monSite Is Me) Then
        'Initialisation à faux du dessin des bandes
        'inter-carrefours voitures en onde TC
        frmMain.mnuGraphicOndeInterCarfVP.Checked = monDessinInterCarfVP
        
        'Mise à vide de l'objet sélectionné graphiquement
        ViderObjPick
        'Masquage des poignées
        PoigneeDroite.Visible = False
        PoigneeGauche.Visible = False
        
        'Stockage de la fenetre du site courant
        Set monSite = Me
        
        'Réduction des carrefours pour lier le carrefour
        'et son carrefour réduit
        Call ReduireCarrefourSite(Me, mesCarrefours, monTypeOnde)
        
        'Calcul des temps de parcours dans chaque sens à
        'chaque carrefour. Ces temps servent dans le recalcul
        'des bandes passantes lors d'une modif d'un décalage
        CalculerTempsParcours Me
        
        'Mise en grisée ou non du menu Annuler la dernière modification
        'sur le graphique d'onde verte si on change de site courant
        frmMain.mnuGraphicOndeAnnul.Enabled = False
    End If

    'Mise à jour des contextes d'aide
    ChangerHelpID monSite.TabFeux.Tab
End Sub

Private Sub Form_Load()
    If maDemoVersion Then
        'Vérrouillage en modif des ordonnées des feux
        'colonne 2 du spread TabPropCarf
        TabPropCarf.Row = -1
        TabPropCarf.Col = 2
        TabPropCarf.Lock = True
    End If
    
    'Initialisation à faux du dessin des bandes
    'inter-carrefours voitures en onde TC
    monDessinInterCarfVP = False
    frmMain.mnuGraphicOndeInterCarfVP.Checked = False
    
    'Mise à vide de l'objet sélectionné graphiquement
    ViderObjPick
    
    'Augmentation du nombres de fenêtres filles
    monNbFenFilles = monNbFenFilles + 1
    
    If monNbFenFilles = 1 Then
        'On affiche les boutons dans la toolbar permettant l'impression
        'et la sauvegarde car on a une fenêtre fille d'ouverte
        '==> Impression et sauvegarde possible
        frmMain.tbToolBar.Buttons("Print").Visible = True
        frmMain.tbToolBar.Buttons("Save").Visible = True
    End If
    
    'Chargement des options d'affichage et d'impression
    ChargerOptionsAffImp Me
    'Initialisation des combobox des TC calculant l'onde TC
    ComboTCM.AddItem "Aucun"
    ComboTCD.AddItem "Aucun"
    'Retaillage de la fenetre Site
    Height = Screen.Height * 0.7
    Width = Screen.Width * 0.905
    'Retaillage de la frame de visu des carrefours
    FrameVisuCarf.Height = ScaleHeight - FrameVisuCarf.Top - TitreEtude.Top / 4
    AxeOrdonnée.Y2 = FrameVisuCarf.Height - AxeOrdonnée.Y1 / 4
    If AxeOrdonnée.Y2 Mod 2 = 1 Then
        'Pour avoir une longueur d'axe des Y en twips pairs
        '==> le milieu de l'axe des Y sera paire en twips
        AxeOrdonnée.Y2 = AxeOrdonnée.Y2 + 1
    End If
    'Retaillage de l'onglet
    TabFeux.Left = FrameVisuCarf.Width + FrameVisuCarf.Left + LabelTitre.Left
    TabFeux.Width = ScaleWidth - TabFeux.Left - LabelTitre.Left
    TabFeux.Height = ScaleHeight - TabFeux.Top - TitreEtude.Top / 4
    'Centrage du tableau de propriétés de carrefour
    TabPropCarf.Left = LabelCarf.Left
    
    'Positionnement des boutons de l'onglet Carrefours
    unEspacement = 40
    unDecBord = 70
    uneLargeurBouton = (TabFeux.Width - unDecBord * 2 - unEspacement * 3) / 4
    
    AjoutFeu.Top = TabFeux.Height - AjoutFeu.Height - 120
    AjoutFeu.Left = unDecBord
    AjoutFeu.Width = uneLargeurBouton
    
    SupprFeu.Top = AjoutFeu.Top
    SupprFeu.Left = AjoutFeu.Left + uneLargeurBouton + unEspacement
    SupprFeu.Width = uneLargeurBouton
    
    CarfPred.Top = SupprFeu.Top
    CarfPred.Left = SupprFeu.Left + uneLargeurBouton + unEspacement
    CarfPred.Width = uneLargeurBouton
    
    CarfSuiv.Top = SupprFeu.Top
    CarfSuiv.Left = CarfPred.Left + uneLargeurBouton + unEspacement
    CarfSuiv.Width = uneLargeurBouton
    
    'Choix d'une valeur d'espacement
    unEspacement = 100
    'Centrage vertical du labelTC avec le nom du TC
    LabelTC.Top = LabelTC.Top + (ComboTC.Height - LabelTC.Height) / 2
    'Retaillage de la frame FrameTC et du spread TabYArret de l'onglet TC
    FrameTC.Height = TabFeux.Height - FrameTC.Top - unEspacement / 2
    TabYArret.Width = FrameTC.Width - 2 * LabelTDep.Left
    
    'Calcul de la place restant entre le bas de FrameTC et les boutons Nouvel et Suppr Arrets
    'Décalage des controls de FrameTC pour meilleure répartition
    'Division par 10 car on a 9 lignes de controls à centrer verticalement
    unDecal = FrameTC.Height - 120 - TextTDep.Height * 4 - ComboCarfDep.Height * 2 - TabYArret.Height - DelArret.Height - ColorTC.Height
    unDecal = unDecal / 10
    
    unDecaIni = 120 'décalage initiale en twips
    LabelTDep.Top = unDecaIni + unDecal
    TextTDep.Top = unDecaIni + unDecal
    LabelSec1.Top = unDecaIni + unDecal
        
    LabelDistAF.Top = TextTDep.Top + TextTDep.Height + unDecal
    TextDistAF_TC.Top = TextTDep.Top + TextTDep.Height + unDecal
    LabelMetre.Top = TextTDep.Top + TextTDep.Height + unDecal
    
    LabelDuréeAF.Top = TextDistAF_TC.Top + TextDistAF_TC.Height + unDecal
    TextDuréeAF_TC.Top = TextDistAF_TC.Top + TextDistAF_TC.Height + unDecal
    LabelSec2.Top = TextDistAF_TC.Top + TextDistAF_TC.Height + unDecal
    
    LabelCarfDep.Top = TextDuréeAF_TC.Top + TextDuréeAF_TC.Height + unDecal
    ComboCarfDep.Top = TextDuréeAF_TC.Top + TextDuréeAF_TC.Height + unDecal
    
    LabelCarfArr.Top = ComboCarfDep.Top + ComboCarfDep.Height + unDecal
    ComboCarfArr.Top = ComboCarfDep.Top + ComboCarfDep.Height + unDecal
    
    LabelColTC.Top = ComboCarfArr.Top + ComboCarfArr.Height + unDecal
    ColorTC.Top = ComboCarfArr.Top + ComboCarfArr.Height + unDecal
    
    TabYArret.Top = ColorTC.Top + ColorTC.Height + unDecal
    LabelArret.Top = TabYArret.Top + TabYArret.Height + unDecal
    TextArret.Top = TabYArret.Top + TabYArret.Height + unDecal
    
    NewArret.Top = TextArret.Top + TextArret.Height + unDecal
    DelArret.Top = TextArret.Top + TextArret.Height + unDecal
    BoutonInverser.Top = (ComboCarfArr.Top + ComboCarfDep.Top + ComboCarfDep.Height - BoutonInverser.Height) / 2
    
    'Retaillage du spread TabPropCarf
    TabPropCarf.Left = (TabFeux.Width - TabPropCarf.Width) / 2
    TabPropCarf.Height = AjoutFeu.Top - unEspacement - TabPropCarf.Top
    'Retaillage du spread TabTrafCarf
    TabTrafCarf.Left = TabPropCarf.Left
    TabTrafCarf.Width = TabPropCarf.Width
    'Retaillage du tableau TabInfoCalc
    TabInfoCalc.Height = TabFeux.Height - TabInfoCalc.Top - (FrameTypeOnde.Top - TabFeux.TabHeight)
    'Retaillage du tableau TabDecal
    TabDecal.Height = TabFeux.Height - TabDecal.Top - (FrameTransDec.Top - TabFeux.TabHeight)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim unNomFich As String
    
    'Pas de suavegarde en version démo
    If maDemoVersion Then
        MsgBox "LA SAUVEGARDE n'est pas disponible en version DEMO", vbInformation
        Exit Sub
    End If
    
    'Recherche de modif non sauvegardée
    uneModif1 = maModifDataSite Or maModifDataTC Or maModifDataDec Or maModifDataDes
    uneModif2 = maModifDataOndeTC Or maModifDataOnde Or maModifDataCarf
    If uneModif1 Or uneModif2 Or Mid(Caption, 1, 15) = "Site : Sans Nom" Then
        'Cas où la fenêtre a été modifiée depuis son chargement
        'ou sa dernière sauvegarde ou si c'est un site sans nom
        'qui n'a jamais été sauvegardé
        unMsg = "Enregistrer les modifications dans "
        unNomFich = Mid(Caption, 8)
        unMsg = unMsg + unNomFich + " ?"
        If MsgBox(unMsg, vbYesNo + vbQuestion) = vbYes Then
            'Sauvegarde, sinon rien
            If uneModif2 Then
                'Sauvegarde après changement des données de calcul
                'et avant recalcul onde
                maCoherenceDataCalc = IncoherenceDonneeCalcul
            End If
            If Mid(unNomFich, 1, 8) = "Sans Nom" Then
                frmMain.RunSaveAs Me
            Else
                EcrireDansFichier unNomFich, Me
            End If
            'Mise en tête dans la liste des derniers fichiers ouverts
            'si on n'a pas fait annuler dans le choix du nouveau nom
            '==> monSite.Caption = "Site : Sans Nom unNum" ou lieu de
            '"Site : C:\ggg\etc .."
            If Mid(monSite.Caption, 8, 8) <> "Sans Nom" Then
                frmMain.MettreEnTeteRecents Mid(monSite.Caption, 8), False
            End If
        End If
    End If
    
    'Fermeture du fichier pour le délocker
    Close #monFichId
    
    'Réduction du nombres de fenêtres filles
    monNbFenFilles = monNbFenFilles - 1
    
    'libération de la mémoire des collections du site
    'A finir (à tester avec les events terminate des instances
    'contenu dans ces collections
    Set mesCarfReduitsSens2 = Nothing

    'On masque les boutons dans la toolbar permettant l'impression et
    'la sauvegarde s'il n'y a plus de fenêtre fille ouverte
    '==> Impression et sauvegarde impossible
    If monNbFenFilles = 0 Then
        frmMain.tbToolBar.Buttons("Print").Visible = False
        frmMain.tbToolBar.Buttons("Save").Visible = False
        frmMain.HelpContextID = 0 'retour au sommaire de l'aide
    End If
End Sub

Private Sub FrameVisuCarf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Désélection graphique
    Call Deselectionner(Me)
    'Mise à jour des contextes d'aide
    ChangerHelpID TabFeux.Tab
End Sub


Private Sub IconeArret_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection de l'arrêt Index et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Call MiseAJourSelection(Me, ArretSel, Index, IconeArret(Index))
End Sub

Private Sub IconeArret_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NomArret_MouseUp Index, Button, Shift, X, Y
End Sub

Private Sub IconeFeu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NumFeu_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub IconeFeu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NumFeu_MouseUp Index, Button, Shift, X, Y
End Sub





Private Sub NewArret_Click()
    CreerArretTC Me
End Sub

Private Sub NewTC_Click()
    'interdiction de créer des TC si moins de 2 carrefours
    If mesCarrefours.Count < 2 Then
        MsgBox "Pour créer un Transport collectif, au moins 2 carrefours doivent exister"
    Else
        Call NewOrRenameTC("new")
    End If
End Sub


Private Sub NomArret_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection de l'arrêt Index et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Call MiseAJourSelection(Me, ArretSel, Index, NomArret(Index), X)
End Sub

Private Sub NomArret_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'Affichage d'un menu contextuel
        frmMain.AfficherMenuContextuel "arrêt TC"
    End If
End Sub

Private Sub NomCarf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection du carrefour Index et de son premier feu et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    
    'Mise à jour sélection graphique et l'onglet Carrefour
    Call MiseAJourSelectionEtOngletCarrefour(Me, CarfSel, Fix(Val(NomCarf(Index).Tag)), 1)
End Sub

Private Sub NomCarf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'Affichage d'un menu contextuel
        frmMain.AfficherMenuContextuel "carrefour"
    End If
End Sub

Private Sub NumFeu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection du feu Index et de son carrefour et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unePos As Integer, unePosCarf As Long, unePosFeu As Long
    'Récupération du carrefour et du feu sélectionné
    'par décodage du tag de l'objet graphique NumFeu sélectionné
    unePos = InStr(1, NumFeu(Index).Tag, "-")
    unePosCarf = Val(Mid$(NumFeu(Index).Tag, 1, unePos - 1))
    unePosFeu = Val(Mid$(NumFeu(Index).Tag, unePos + 1))
    'Mise à jour sélection graphique et l'onglet Carrefour
    Call MiseAJourSelectionEtOngletCarrefour(Me, FeuSel, unePosCarf, unePosFeu)
End Sub

Private Sub NumFeu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'Affichage d'un menu contextuel
        frmMain.AfficherMenuContextuel "feu"
    End If
End Sub


Private Sub OptionOndeDouble_Click()
    monTypeOnde = OndeDouble
    'Affichage des poids
    FramePoids.Visible = OptionOndeDouble.Value
    'Effacement des paramètres TC
    FrameOndeTC.Visible = OptionTC.Value
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub


Private Sub OptionSensD_Click()
    monTypeOnde = OndeSensD
    'Effacement des poids
    FramePoids.Visible = OptionOndeDouble.Value
    'Effacement des paramètres TC
    FrameOndeTC.Visible = OptionTC.Value
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub

Private Sub OptionSensM_Click()
    monTypeOnde = OndeSensM
    'Effacement des poids
    FramePoids.Visible = OptionOndeDouble.Value
    'Effacement des paramètres TC
    FrameOndeTC.Visible = OptionTC.Value
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub

Private Sub OptionTC_Click()
    monTypeOnde = OndeTC
    'Effacement des poids
    FramePoids.Visible = OptionOndeDouble.Value
    'Affichage des paramètres TC au même endroit que les poids
    FrameOndeTC.Left = FramePoids.Left
    FrameOndeTC.Top = FramePoids.Top
    FrameOndeTC.Visible = OptionTC.Value
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub


Private Sub OptionVitConst_Click()
    monTypeVit = VitConst
    'On cache les colonnes 3 et 4 permettant la saisie des
    'vitesses montantes et descendantes de chaque carrefour
    TabInfoCalc.Col = 3
    TabInfoCalc.ColHidden = True
    TabInfoCalc.Col = 4
    TabInfoCalc.ColHidden = True
    'Désinhibition des vitesses constantes montantes et descendantes
    TextVitM.Enabled = True
    TextVitD.Enabled = True
    LabelVitSensM.Enabled = True
    LabelVitSensD.Enabled = True
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub

Private Sub OptionVitVar_Click()
    monTypeVit = VitVar
    'On rend visible les colonnes 3 et 4 permettant la saisie des
    'vitesses montantes et descendantes de chaque carrefour
    TabInfoCalc.Col = 3
    TabInfoCalc.ColHidden = False
    TabInfoCalc.ColWidth(3) = 1050 'Taille en twips fixée dans le spread designer (éviter un bug de retailage)
    TabInfoCalc.Col = 4
    TabInfoCalc.ColHidden = False
    TabInfoCalc.ColWidth(4) = 1050 'Taille en twips fixée dans le spread designer (éviter un bug de retailage)
    'Inhibition des vitesses constantes montantes et descendantes
    TextVitM.Enabled = False
    TextVitD.Enabled = False
    LabelVitSensM.Enabled = False
    LabelVitSensD.Enabled = False
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub


Private Sub RenameCarf_Click()
    RenommerCarrefour Me
End Sub

Private Sub RenameTC_Click()
    Call NewOrRenameTC("rename")
End Sub

Private Sub SupprCarf_Click()
    SupprimerCarrefour Me
End Sub

Private Sub SupprFeu_Click()
    SupprimerFeu Me
End Sub


Private Sub TabDecal_Change(ByVal Col As Long, ByVal Row As Long)
    Dim unSaveText As String
    
    If Col = 4 And TabDecal.Tag = "" Then
        'changement dans la colonne 4 et on ne vient pas de l'event
        'TabDecal_KeyUp, le même traitement y est fait
    
        'Positionnement sur la cellule active
        TabDecal.Row = Row
        TabDecal.Col = Col
        'Stockage de valeurs avant modif
        For i = 1 To mesCarrefours.Count
            maColDecSave.Add mesCarrefours(i).monDecCalcul
            maColDecSave.Add mesCarrefours(i).monDecModif
        Next i
        If TabDecal.Text = "Oui" Then
            unSaveText = "Non"
        ElseIf TabDecal.Text = "Non" Then
            unSaveText = "Oui"
        Else
            MsgBox "ERREUR de programmation dans frmDocument:TabDecal_Change", vbCritical
        End If
        'On fait un calcul à date imposé si on click dans la colonne
        '4 et si le type de décalage imposé du carrefour change
        RecalculerAvecDateImp mesCarrefours(Row), TabDecal.Text
        If maCoherenceDataCalc = CalculImpossible Then
            'Cas du calcul impossible on relance le calcul avec
            'l'autre valeur ("Oui" ou "Non") qui marchait avant
            'le changement avec les décalages avant modif
            For i = 1 To (maColDecSave.Count / 2)
                mesCarrefours(i).monDecCalcul = maColDecSave(2 * i - 1)
                mesCarrefours(i).monDecModif = maColDecSave(2 * i)
            Next i
            RecalculerAvecDateImp mesCarrefours(Row), unSaveText
        End If
        'On vide la collection de sauvegarde des décalages
        ViderCollection maColDecSave
    End If
End Sub

Private Sub TabDecal_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim unNewDecal As Integer
    Dim unOldDecal As Single
    
    'Positionnement sur la cellule active
    TabDecal.Row = Row
    TabDecal.Col = Col
        
    'Présence dans la colonne 4 on ne fait rien
    'cf KeyUp et Change events de TabDecal
    If Col = 4 Then Exit Sub
    
    If Mode = 0 Then
        'Cas où l'utilisateur a frappé Entréee, Return ou sort de la cellule
            
        'Recalcul des bandes passantes avec le nouveau décalage
        's'il a été modifié
        If ChangeMade Then
            'Cas où la valeur a changée
            
            'On ramène le nouveau décalage entre 0 et la durée du cycle
            unNewDecal = ModuloZeroCycle(TabDecal.Text, maDuréeDeCycle)
            'Affichage de cette nouvelle valeur
            TabDecal.Text = Format(unNewDecal)
            'Stockage du décalage et des bandes passantes avant modification
            unOldDecal = mesCarrefours(Row).monDecModif
            
            'Modif du décalage modifiable du carrefour choisi
            'On ajoute la différence entre la vrai en réelle et l'arrondi
            'en entier pour l'affichage pour ne pas perdre en précision
            'de calcul
            'Exemple : si le calcul trouve un décalage de 29.8 que l'on stocke
            'on affiche par contre 30, si l'utilisateur remet 30 il peut avoir
            'un résultat différent car le 30 qu'il voit, vaut en fait 29.8.
            'En ajoutant la différence du à l'arrondi on retrouve la même chose
            mesCarrefours(Row).monDecModif = mesCarrefours(Row).monDecModif - CIntCorrigé(mesCarrefours(Row).monDecModif) + unNewDecal
            
            If mesCarrefours(Row).monDecImp = 1 Then
                'Cas d'un carrefour à date imposé
                'Stockage de valeurs de décalages avant modif
                For i = 1 To mesCarrefours.Count
                    maColDecSave.Add mesCarrefours(i).monDecCalcul
                    maColDecSave.Add mesCarrefours(i).monDecModif
                Next i
                '==> Recalcul total de l'onde
                maModifDataOnde = True
                CalculerOndeVerte Me, True
                If maCoherenceDataCalc = CalculImpossible Then
                    'Cas où le recalcul a été impossible
                    'On réaffiche et remet la valeur du décalage avant modif
                    mesCarrefours(Row).monDecModif = unOldDecal
                    If CIntCorrigé(unOldDecal) = maDuréeDeCycle Then
                        'Une valeur valant durée du cycle s'affiche 0
                        TabDecal.Text = "0"
                    Else
                        TabDecal.Text = CIntCorrigé(unOldDecal)
                    End If
                    'On relance le calcul d'onde verte avec les décalages
                    'avant modif dues au calcul précédent
                    For i = 1 To (maColDecSave.Count / 2)
                        mesCarrefours(i).monDecCalcul = maColDecSave(2 * i - 1)
                        mesCarrefours(i).monDecModif = maColDecSave(2 * i)
                    Next i
                    '==> Recalcul total de l'onde
                    maModifDataOnde = True
                    CalculerOndeVerte Me, True
                End If
            Else
                'Cas d'un carrefour à décalage non imposé
                '==> Recalcul des bandes passantes uniquement
                If RecalculerBandesPassantes(Me) Then
                    'Cas où le recalcul a été possible
                    'Stockage d'une modification de valeurs dans les décalages
                    'Ceci permettra aussi de demander une sauvegarde à la fermeture
                    maModifDataDec = True
                    'Mise en grisée du menu Annuler dernière modif graphique car
                    'on fait une modif par saisie et pas par interaction graphique
                    frmMain.mnuGraphicOndeAnnul.Enabled = False
                Else
                    'Cas où le recalcul a été impossible
                    'On réaffiche et remet la valeur du décalage avant modif
                    mesCarrefours(Row).monDecModif = unOldDecal
                    If CIntCorrigé(unOldDecal) = maDuréeDeCycle Then
                        'Une valeur valant durée du cycle s'affiche 0
                        TabDecal.Text = "0"
                    Else
                        TabDecal.Text = CIntCorrigé(unOldDecal)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub TabDecal_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim unSaveText As String, uneToucheModifCombobox As Boolean
    
    If TabDecal.ActiveCol = 4 Then
        TabDecal.Row = TabDecal.ActiveRow
        TabDecal.Col = TabDecal.ActiveCol
        
        'liste des touches de pavés déplacement modifiant une combobox
        uneToucheModifCombobox = (KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft)
        
        If KeyCode = 78 Or KeyCode = 79 Or KeyCode = vbKeyEscape Or uneToucheModifCombobox Then
            'Touche pressée est o (code = 79) ou n (code = 78) ou Echap
            
            'Indication du passage dans cet event pour éviter le
            'traitement de l'event TabDecal_Change
            TabDecal.Tag = "Passée dans TabDecal_KeyUp"
            
            'Stockage de valeurs avant modif
            For i = 1 To mesCarrefours.Count
                maColDecSave.Add mesCarrefours(i).monDecCalcul
                maColDecSave.Add mesCarrefours(i).monDecModif
            Next i
            If TabDecal.Text = "Oui" Then
                unSaveText = "Non"
            ElseIf TabDecal.Text = "Non" Then
                unSaveText = "Oui"
            Else
                MsgBox "ERREUR de programmation dans frmDocument:TabDecal_KeyUp", vbCritical
            End If
        
            'On fait un calcul à date imposé si on click dans la colonne
            '4 et si le type de décalage imposé du carrefour change
            RecalculerAvecDateImp mesCarrefours(TabDecal.Row), TabDecal.Text
            
            If maCoherenceDataCalc = CalculImpossible Then
                'Cas du calcul impossible on relance le calcul avec
                'l'autre valeur ("Oui" ou "Non") qui marchait avant
                'le changement avec les décalages avant modif
                For i = 1 To (maColDecSave.Count / 2)
                    mesCarrefours(i).monDecCalcul = maColDecSave(2 * i - 1)
                    mesCarrefours(i).monDecModif = maColDecSave(2 * i)
                Next i
                RecalculerAvecDateImp mesCarrefours(TabDecal.Row), unSaveText
            End If
            'On vide la collection de sauvegarde des décalages
            ViderCollection maColDecSave
            'On vide Indication de passage dans cet event pour permettre
            'de nouveau le traitement de l'event TabDecal_Change
            TabDecal.Tag = ""
        End If
    End If
End Sub

Private Sub TabDecal_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Cas d'un changement de cellule active
    'Sélection du carrefour de la ligne active et de son premier feu
    'et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If NewRow > 0 And NewCol > 0 Then
        Set monCarrefourCourant = mesCarrefours(NewRow)
        MiseAJourSelectionParCellule Me, FeuSel, monCarrefourCourant.maPosition, 1
    End If
End Sub

Private Sub TabFeux_Click(PreviousTab As Integer)
    Dim unePosTC As Integer
    Dim unNomTC As String
    
    'Pointeur souris devient sablier pour prévenir attente
    MousePointer = vbHourglass
    
    'Utilisation de DoEvents pour vider la liste d'événements du
    'système pour corriger un bug bizarre : La variable maModifDataCarf
    'ne garde pas la valeur false après un calcul d'onde et le retour
    'dans l'onglet Cadrage
    '==> Recalcul alors qu'elle est bien mise à false après CalculerOndeVerte
    'd'où un problème de synchronisation résolu par le DoEvents,
    'l'utilisateur ne voit rien .
    DoEvents
    
    'Mise en grisé du menu Graphique onde verte
    'Dégrisée si onglet Graphique onde verte est actif (TabFeux.Tab = 4)
    frmMain.mnuGraphicOnde.Enabled = False
    
    'Mise à jour de la barre d'état
    frmMain.sbStatusBar.Panels.Item(1).Text = "OndeV version 1.0"
    
    If TabFeux.Tab = 0 Then
        'Cas de l'onglet Carrefours
        'Rafraichissement de TabPropCarf pour éviter l'apparition
        'd'un tableau TabPropCarf à moitié (Bug Spread)
        TabPropCarf.Refresh
        'Sélection graphique du carrefour courant et de son premier feu
        MiseAJourSelectionEtOngletCarrefour Me, CarfSel, monCarrefourCourant.maPosition, 1
    ElseIf TabFeux.Tab = 1 Then
        'Cas de l'onglet TC
        'Sélection graphique de l'arrêt TC correspondant
        'à la cellule active du spread TabYArret si FrameTC est visible
        If FrameTC.Visible = True Then
            'Rafraichir de FrameTC pour avoir tabYArret entier (bug spread)
            FrameTC.Refresh
            MiseAJourSelectionParCellule Me, ArretSel, ComboTC.ListIndex + 1, TabYArret.ActiveCol
            'Réaffichage des noms des carrefours de départ et d'arrivée pour
            'prendre en compte les changements d'indice lors d'une suppression
            unePosTC = ComboTC.ListIndex + 1
            ComboCarfDep.ListIndex = mesTC(unePosTC).monCarfDep.maPosition - 1
            ComboCarfArr.ListIndex = mesTC(unePosTC).monCarfArr.maPosition - 1
        End If
    ElseIf TabFeux.Tab = 2 Then
        'Cas de l'onglet Cadrage onde verte
        'Rafraichissement de TabInfoCalc pour éviter l'apparition
        'd'un tableau TabInfoCalc à moitié (Bug Spread)
        TabInfoCalc.Refresh
        'Affectation d'une couleur pour les cellules lockées
        TabInfoCalc.LockBackColor = LabelTrait.BackColor
        'Sélection graphique du carrefour courant et de son premier feu
        MiseAJourSelectionParCellule Me, CarfSel, monCarrefourCourant.maPosition, 1
        'On rend actif dans TabInfoCalc la ligne du carrefour courant
        TabInfoCalc.Row = monCarrefourCourant.maPosition
        TabInfoCalc.Col = 1
        TabInfoCalc.Action = SS_ACTION_ACTIVE_CELL
        
        'Mise à jour des bandes passantes TC montantes et descendantes
        TextBTCM.Text = Format(maBandeTCM)
        TextBTCD.Text = Format(maBandeTCD)
        'Mise à jour du TC cadrant l'onde verte en sens montant
        If monTCM = 0 Then
            ComboTCM.Text = "Aucun"
            'Masquage de la bande passante montante
            TextBTCM.Visible = False
            LabelSecM.Visible = False
        Else
            ComboTCM.Text = mesTC(monTCM).monNom
            'Afichage de la bande passante montante
            TextBTCM.Visible = True
            LabelSecM.Visible = True
        End If
        'Mise à jour du TC cadrant l'onde verte en sens descendant
        If monTCD = 0 Then
            ComboTCD.Text = "Aucun"
            'Masquage de la bande passante descendante
            TextBTCD.Visible = False
            LabelSecD.Visible = False
        Else
            ComboTCD.Text = mesTC(monTCD).monNom
            'Affichage de la bande passante descendante
            TextBTCD.Visible = True
            LabelSecD.Visible = True
        End If
    ElseIf TabFeux.Tab = 3 Then
        'Cas de l'onglet Résultat décalages
        'Affichage du calcul d'onde verte effectué
        If monTypeOnde = OndeDouble Then
            FrameTransDec.Caption = "Résultat du calcul d'onde verte à double sens"
        ElseIf monTypeOnde = OndeSensM Then
            FrameTransDec.Caption = "Résultat du calcul d'onde verte à sens privilégié montant"
        ElseIf monTypeOnde = OndeSensD Then
            FrameTransDec.Caption = "Résultat du calcul d'onde verte à sens privilégié descendant"
        ElseIf monTypeOnde = OndeTC Then
            FrameTransDec.Caption = "Résultat du calcul d'onde verte prenant en compte les TC"
        End If
        'Rafraichissement de TabBande et TabDecal pour éviter l'apparition
        'd'un tableau TabBande ou TabDecal à moitié (Bug Spread)
        TabBande.Refresh
        TabDecal.Refresh
        'Affectation d'une couleur pour les cellules lockées
        TabBande.LockBackColor = LabelTrait.BackColor
        'Affectation d'une couleur pour les cellules lockées
        TabDecal.LockBackColor = LabelTrait.BackColor
        'Sélection graphique du carrefour courant et de son premier feu
        MiseAJourSelectionParCellule Me, CarfSel, monCarrefourCourant.maPosition, 1
        'On rend actif dans TabInfoCalc la ligne du carrefour courant
        TabDecal.Row = monCarrefourCourant.maPosition
        TabDecal.Col = 1
        TabDecal.Action = SS_ACTION_ACTIVE_CELL
        'Calcul de l'onde verte si on ne vient pas des
        'onglets Graphique onde verte et Fiche Résultat
        unCalculOndeFait = True
        If PreviousTab <> 4 And PreviousTab <> 5 Then
            unCalculOndeFait = CalculerOndeVerte(Me)
        End If
        If unCalculOndeFait And maCoherenceDataCalc = OK Then
            'Remplir l'onglet Résultat décalages
            RemplirOngletTabDecalage Me
        Else
            TabFeux.Tab = PreviousTab
        End If
    ElseIf TabFeux.Tab = 4 Then
        'Mise en actif du menu permettant d'afficher les bandes
        'inter-carrefours voitures si onde cadrée par un TC montant
        'et/ou descendant sinon il est mis en inactif
        frmMain.mnuGraphicOndeInterCarfVP.Enabled = (monSite.monTypeOnde = OndeTC)
        'Activation du menu principal Graphique onde verte
        frmMain.mnuGraphicOnde.Enabled = True
        'Mise à jour Affichage de l'onglet Graphique onde verte
        AffichageOngletVisu
        'Calcul de l'onde verte si on ne vient pas des
        'onglets Résultat décalages et Fiche résultat
        unCalculOndeFait = True
        If PreviousTab <> 3 And PreviousTab <> 5 Then
            unCalculOndeFait = CalculerOndeVerte(Me)
        End If
        'Désélection si on clique dans un autre onglet
        Call Deselectionner(Me)
        
        If unCalculOndeFait And maCoherenceDataCalc = OK Then
            'Masquage des poignées si aucune n'a été sélectionnée
            PoigneeDroite.Visible = False
            PoigneeGauche.Visible = False
            'Dessiner le graphique de l'onde verte
            ZoneDessin.Cls
            unEspacement = 120 'même valeur que dans AffichageOngletVisu
            DessinerTout ZoneDessin, AxeTemps.X1, AxeTemps.Y1 - unEspacement / 4, AxeTemps.X2 - AxeTemps.X1, AxeOrdonnée.Y2 - AxeOrdonnée.Y1
            'le - unEsp/4 pour avoir l'origine de l'axe des temps au même
            'niveau que le min des Y
        Else
            TabFeux.Tab = PreviousTab
        End If
    ElseIf TabFeux.Tab = 5 Then
        'Cas de l'onglet Fiche résultat
        'Affichage du calcul d'onde verte effectué
        If monTypeOnde = OndeDouble Then
            LabelTypeOnde.Caption = "Résultat du calcul d'onde verte à double sens"
        ElseIf monTypeOnde = OndeSensM Then
            LabelTypeOnde.Caption = "Résultat du calcul d'onde verte à sens privilégié montant"
        ElseIf monTypeOnde = OndeSensD Then
            LabelTypeOnde.Caption = "Résultat du calcul d'onde verte à sens privilégié descendant"
        ElseIf monTypeOnde = OndeTC Then
            LabelTypeOnde.Caption = "Résultat du calcul d'onde verte prenant en compte les TC"
        End If
        
        'Rafraichissement de TabFicheOnde, TabFicheCarf et TabFicheTC pour
        'éviter l'apparition d'un de ces tableaux à moitié (Bug Spread)
        TabFicheOnde.Refresh
        TabFicheCarf.Refresh
        TabFicheTC.Refresh
        
        'Calcul de l'onde verte si on ne vient pas des
        'onglets Résultat décalages et Graphique onde verte
        unCalculOndeFait = True
        If PreviousTab <> 3 And PreviousTab <> 4 Then
            unCalculOndeFait = CalculerOndeVerte(Me)
        Else
            'Test si une modification manuelle des décalages a eu lieu
            If EstModifierManuel Then
                LabelTypeOnde.Caption = "Résultat du calcul d'onde verte avec décalages modifiés manuellement"
            End If
        End If
        
        If unCalculOndeFait And maCoherenceDataCalc = OK Then
            'Calcul des vitesses maximun
            CalculerVitMax Me
                        
            'Positionnement des TabFicheCarf et TabFicheTC
            TabFicheTC.Visible = True
            'Restauration des tailles initiales sous VB
            'avant retaillage automatique ==> Recalcul toujours
            'à partir des mêmes données ==> même résultat
            TabFicheCarf.Height = 1575
            TabFicheTC.Height = 1095
            TabFicheTC.Top = 4120
            LabelResTC.Top = 3900
            'Retaillage et positionnement
            unReste = TabFeux.Height - (TabFicheTC.Top + TabFicheTC.Height)
            If mesTCutil.Count = 0 Then
                LabelResTC.Top = TabFeux.Height - LabelResTC.Height - 200
                LabelResTC.Caption = "Aucun progression TC tracée ==> Pas de résultat pour les TC"
                TabFicheTC.Visible = False
                TabFicheCarf.Height = LabelResTC.Top - TabFicheCarf.Top - 100
            Else
                'Cas de TC utilisés
                unReste = unReste - 100
                LabelResTC.Caption = "Résultat par Transport Collectif"
                LabelResTC.Top = LabelResTC.Top + unReste * 2 / 3
                TabFicheCarf.Height = TabFicheCarf.Height + unReste * 2 / 3
                TabFicheTC.Top = TabFicheTC.Top + unReste * 2 / 3
                TabFicheTC.Height = TabFicheTC.Height + unReste / 3
            End If
            'Remplir l'onglet Fiche résultat
            RemplirOngletFicheResult Me
        Else
            'Cas où le calcul de l'onde verte est impossible
            TabFeux.Tab = PreviousTab
        End If
    End If
    
    'Aide contextuelle
    ChangerHelpID TabFeux.Tab
    
    'Pointeur souris redevient normal pour prévenir fin d'attente
    MousePointer = vbDefault
End Sub

Private Sub TabInfoCalc_Change(ByVal Col As Long, ByVal Row As Long)
    'Positionnement sur la cellule active
    TabInfoCalc.Row = TabInfoCalc.ActiveRow
    TabInfoCalc.Col = TabInfoCalc.ActiveCol

    If TabInfoCalc.ActiveCol = 2 Then
        'Cas d'une saisie de la prise en compte du carrefour dans les calculs
        If TabInfoCalc.Text = "Oui" Then
            monCarrefourCourant.monIsUtil = True
        Else
            monCarrefourCourant.monIsUtil = False
        End If
    End If
    'Stockage d'une modification de données pour le calcul de l'onde
    maModifDataOnde = True
End Sub

Private Sub TabInfoCalc_Click(ByVal Col As Long, ByVal Row As Long)
    'Cas d'un changement de cellule active
    'Sélection du carrefour de la ligne active et de son premier feu
    'et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If Row > 0 And Col > 0 Then
        Set monCarrefourCourant = mesCarrefours(Row)
        MiseAJourSelectionParCellule Me, FeuSel, monCarrefourCourant.maPosition, 1
    End If
End Sub

Private Sub TabInfoCalc_KeyUp(KeyCode As Integer, Shift As Integer)
    'Positionnement sur la cellule active
    TabInfoCalc.Row = TabInfoCalc.ActiveRow
    TabInfoCalc.Col = TabInfoCalc.ActiveCol

    'Stockage d'une modification de données pour le calcul de l'onde
    'si on n'a pas appuyer sur la touche Echappement (ou Escape)
    If KeyCode = vbKeyEscape Then
        maModifDataOnde = False
    Else
        maModifDataOnde = True
    End If
    
    If TabInfoCalc.ActiveCol = 2 Then
        'Cas d'une saisie de la prise en compte du carrefour dans les calculs
        If TabInfoCalc.Text = "Oui" Then
            monCarrefourCourant.monIsUtil = True
        Else
            monCarrefourCourant.monIsUtil = False
        End If
    ElseIf TabInfoCalc.ActiveCol = 3 Then
        'Cas d'une saisie d'une vitesse montante
        SaisieEntierPositifEntreMinMax KeyCode, TabInfoCalc, maVitSensM, 1, 99, "La vitesse du sens montant"
        monCarrefourCourant.maVitSensM = Val(TabInfoCalc.Text)
    ElseIf TabInfoCalc.ActiveCol = 4 Then
        'Cas d'une saisie d'une vitesse descendante
        SaisieEntierPositifEntreMinMax KeyCode, TabInfoCalc, maVitSensD, 1, 99, "La vitesse du sens descendant"
        monCarrefourCourant.maVitSensD = Val(TabInfoCalc.Text)
    End If
End Sub

Private Sub TabInfoCalc_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Cas d'un changement de cellule active
    'Sélection du carrefour de la ligne active et de son premier feu
    'et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If NewRow > 0 And NewCol > 0 Then
        Set monCarrefourCourant = mesCarrefours(NewRow)
        MiseAJourSelectionParCellule Me, FeuSel, monCarrefourCourant.maPosition, 1
    End If
End Sub

Private Sub TabPropCarf_Change(ByVal Col As Long, ByVal Row As Long)
    If TabPropCarf.ActiveCol = 1 Then
        'Cas d'une saisie d'un sens d'un feu par choix dans la combobox
        'Positionnement sur la cellule active
        TabPropCarf.Row = TabPropCarf.ActiveRow
        TabPropCarf.Col = TabPropCarf.ActiveCol
        If TabPropCarf.Text = "Montant" Then
            monCarrefourCourant.mesFeux(TabPropCarf.Row).monSensMontant = True
        Else
            monCarrefourCourant.mesFeux(TabPropCarf.Row).monSensMontant = False
        End If
        'Positionnement du feu (Numéro et icône Feu) à droite de l'axe des Y
        'pour un feu montant et à gauche pour un feu descendant
        PlacerFeuAxeY Me, monCarrefourCourant.maPosition, CInt(TabPropCarf.Row), monCarrefourCourant.mesFeuxGraphics(TabPropCarf.Row).Index
    End If
    
    'Stockage d'une modification de valeurs dans les carrefours
    maModifDataCarf = True
End Sub

Private Sub TabPropCarf_Click(ByVal Col As Long, ByVal Row As Long)
    'Cas d'un changement de cellule active
    'Sélection du feu et de son carrefour de la ligne active
    'et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If Row > 0 And Col > 0 Then
        MiseAJourSelectionParCellule Me, FeuSel, monCarrefourCourant.maPosition, Row
    End If
End Sub


Private Sub TabPropCarf_KeyUp(KeyCode As Integer, Shift As Integer)
    'Positionnement sur la cellule active
    TabPropCarf.Row = TabPropCarf.ActiveRow
    TabPropCarf.Col = TabPropCarf.ActiveCol

    'Stockage d'une modification de valeurs dans les carrefours
    'si on n'a pas appuyer sur la touche Echappement (ou Escape)
    If KeyCode = vbKeyEscape Then
        maModifDataCarf = False
    Else
        maModifDataCarf = True
    End If
    
    If TabPropCarf.ActiveCol = 1 Then
        'Cas d'une saisie d'un sens d'un feu
        If TabPropCarf.Text = "Montant" Then
            monCarrefourCourant.mesFeux(TabPropCarf.Row).monSensMontant = True
        Else
            monCarrefourCourant.mesFeux(TabPropCarf.Row).monSensMontant = False
        End If
        'Positionnement du feu (Numéro et icône Feu) à droite de l'axe des Y
        'pour un feu montant et à gauche pour un feu descendant
        PlacerFeuAxeY Me, monCarrefourCourant.maPosition, CInt(TabPropCarf.Row), monCarrefourCourant.mesFeuxGraphics(TabPropCarf.Row).Index
    ElseIf TabPropCarf.ActiveCol = 2 Then
        'Cas d'une saisie d'une ordonnée d'un feu
        ModifierYFeu Me, monCarrefourCourant, TabPropCarf.ActiveRow, TabPropCarf.Text
    ElseIf TabPropCarf.ActiveCol = 3 Then
        'Cas d'une saisie d'une durée de vert d'un feu
        Call VerifMinMaxDuréeVert
    ElseIf TabPropCarf.ActiveCol = 4 Then
        'Cas d'une saisie d'une position de référence d'un feu
        monCarrefourCourant.mesFeux(TabPropCarf.Row).maPositionPointRef = -Format(TabPropCarf.Text)
        '-PosRef car définition inverse entre dossier programmation et doc logiciel OndeV
    End If
End Sub



Private Sub TabPropCarf_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Cas d'un changement de cellule active
    'Sélection du feu et de son carrefour de la ligne active
    'et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If NewRow > 0 And NewCol > 0 Then
        MiseAJourSelectionParCellule Me, FeuSel, monCarrefourCourant.maPosition, NewRow
    End If
End Sub

Private Sub TabTrafCarf_KeyUp(KeyCode As Integer, Shift As Integer)
    'Positionnement sur la cellule active
    DoEvents
    TabTrafCarf.Row = TabTrafCarf.ActiveRow
    TabTrafCarf.Col = TabTrafCarf.ActiveCol
    DoEvents
    
    'Stockage d'une modification de valeurs dans les carrefours
    'si on n'a pas appuyer sur la touche Echappement (ou Escape)
    If KeyCode = vbKeyEscape Then
        maModifDataCarf = False
    Else
        maModifDataCarf = True
    End If
    
    If TabTrafCarf.Col = 1 Then
        'Cas d'une saisie d'une demande
        If TabTrafCarf.Row = 1 Then
            'Cas du sens montant
            monCarrefourCourant.maDemandeM = Format(TabTrafCarf.Text)
        Else
            'Cas du sens descendant
            monCarrefourCourant.maDemandeD = Format(TabTrafCarf.Text)
        End If
    Else
        'Cas d'une saisie d'un débit de saturation
        If TabTrafCarf.Row = 1 Then
            'Cas du sens montant
            monCarrefourCourant.monDebSatM = Format(TabTrafCarf.Text)
        Else
            'Cas du sens descendant
            monCarrefourCourant.monDebSatD = Format(TabTrafCarf.Text)
        End If
    End If
End Sub

Private Sub TabYArret_Click(ByVal Col As Long, ByVal Row As Long)
    'Cas d'un changement de cellule active
    'Sélection de l'arrêt de la colonne active et déselection de l'ancien objet sélectionné
    '==> suppression du gras de l'ancienne et mise en gras de la nouvelle sélection
    Dim unControl As Control
    
    If Col > 0 And Row > 0 Then
        MiseAJourSelectionParCellule Me, ArretSel, ComboTC.ListIndex + 1, Col
    End If

End Sub

Private Sub TabYArret_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim unInd As Integer
    Dim unY As Long, unAncienY As Long
    Dim uneListeArret As ColArretTC
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unArret As ArretTC
    
    TabYArret.Row = TabYArret.ActiveRow
    TabYArret.Col = TabYArret.ActiveCol
    
    'Indication d'un changement de données TC
    IndiquerModifTC

    'Récupération de l'arrêt sélectionné
    Set unArret = mesTC(ComboTC.ListIndex + 1).mesArrets(TabYArret.Col)
    
    If TabYArret.Row = 2 Then
        'Cas où l'on saisit une vitesse de marche
        'Modification de la vitesse de marche de l'arrêt
        unArret.maVitesseMarche = Val(TabYArret.Text)
        Exit Sub
    ElseIf TabYArret.Row = 3 Then
        'Cas où l'on saisit un temps d'arrêt
        'Modification du temps d'arrêt à l'arrêt sélectionné
        unArret.monTempsArret = Val(TabYArret.Text)
        Exit Sub
    End If
    
    'Cas où l'on saisit une ordonnée (1ère ligne, donc row = 1)
    Set uneListeArret = mesTC(ComboTC.ListIndex + 1).mesArrets
    'Test de l'existence d'un arrêt pour le TC courant en ce même Y = val(TabYArret.text)
    unY = Val(TabYArret.Text)
    If VerifierExistenceArret(unY, TabYArret, uneListeArret) = False Then
        'Cas où on sort par le bouton Annuler
        Exit Sub
    End If
    
    'Stockage de l'ancienne valeur du Y pour utilisation plus bas
    unAncienY = uneListeArret(TabYArret.Col).monOrdonnee
    'Modification de l'ordonnée de l'arrêt numéro TabYArret.Col du TC courant
    uneListeArret(TabYArret.Col).monOrdonnee = Val(TabYArret.Text)
    'Modification des objets graphiques de l'arrêt du TC courant
    unInd = mesTC(ComboTC.ListIndex + 1).mesObjGraphics(TabYArret.Col).Index
    unY = ConvertirReelEnEcran(monYMaxFeu - Val(TabYArret.Text), maLongueurAxeY, AxeOrdonnée.Y2 - AxeOrdonnée.Y1)
    NomArret(unInd).Top = unY + AxeOrdonnée.Y1 - NomArret(unInd).Height
    
    'Recherche des arrêts confondus en un Y valant Val(TabYArret.Text) pour
    'alimenter les listes d'arrêts et de TC trouvés
    unNb = RechercherArretConfondu(Val(TabYArret.Text), uneListeIndexTC, uneListeIndexArret)
    'Mise à jour des décalages des labels NomArrêt confondus en ce nouveau Y
    Call MiseAJourNomArret(Me, uneListeIndexTC, uneListeIndexArret)
    'Ajustement de la chaine de caractéres à l'axe des ordonnées
    NomArret(unInd).Width = AxeOrdonnée.X1 - NomArret(unInd).Left
    IconeArret(unInd).Top = unY + AxeOrdonnée.Y1 - IconeArret(unInd).Height
    
    'Recherche des arrêts confondus en un Y valant l'ancienne valeur de
    'l'ordonnée pour alimenter les listes d'arrêts et de TC trouvés
    ViderCollection uneListeIndexTC
    ViderCollection uneListeIndexArret
    unNb = RechercherArretConfondu(unAncienY, uneListeIndexTC, uneListeIndexArret)
    'Mise à jour des décalages des labels NomArrêt confondus en cet ancien Y
    Call MiseAJourNomArret(Me, uneListeIndexTC, uneListeIndexArret)
    
    'Libération de la mémoire
    ViderCollection uneListeIndexTC
    ViderCollection uneListeIndexArret
    Set uneListeIndexTC = Nothing
    Set uneListeIndexArret = Nothing
End Sub


Private Sub TabYArret_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Cas d'un changement de cellule active
    'Sélection de l'arrêt de la colonne active et déselection de l'ancien objet sélectionné
    '==> mise en gras de l'ancienne et de la nouvelle sélection
    Dim unControl As Control
    
    If NewCol > 0 And NewRow > 0 Then
        MiseAJourSelectionParCellule Me, ArretSel, ComboTC.ListIndex + 1, NewCol
    End If
End Sub

Private Sub TextArret_KeyUp(KeyCode As Integer, Shift As Integer)
    TabYArret.Col = TabYArret.ActiveCol
    mesTC(ComboTC.ListIndex + 1).mesArrets(TabYArret.Col).monLibelle = TextArret.Text
    'Indication d'un changement de données du site
    maModifDataSite = True
End Sub



Private Sub TextBTCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SaisieEntierPositifEntreMinMax(KeyCode, TextBTCD, 15, 1, maDuréeDeCycle, "La bande passante du sens descendant")
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
    maBandeTCD = Val(TextBTCD.Text)
End Sub

Private Sub TextBTCM_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SaisieEntierPositifEntreMinMax(KeyCode, TextBTCM, 15, 1, maDuréeDeCycle, "La bande passante du sens montant")
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
    maBandeTCM = Val(TextBTCM.Text)
End Sub

Private Sub TextDistAF_TC_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SaisieEntierPositifEntreMinMax(KeyCode, TextDistAF_TC, 30, 1, 100, "La distance d'accélération et de freinage du transport collectif")
    mesTC(ComboTC.ListIndex + 1).maDistAccFrein = Val(TextDistAF_TC.Text)
    IndiquerModifTC
End Sub

Private Sub TextDuréeAF_TC_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SaisieEntierPositifEntreMinMax(KeyCode, TextDuréeAF_TC, 8, 1, 20, "La durée d'accélération et de freinage du transport collectif")
    mesTC(ComboTC.ListIndex + 1).maDureeAccFrein = Val(TextDuréeAF_TC.Text)
    IndiquerModifTC
End Sub




Private Sub VerifMinMaxDuréeVert()
    'stockage de la cellule active
    uneRow = TabPropCarf.ActiveRow
    uneCol = TabPropCarf.ActiveCol
    'Positionnement sur la cellule active
    TabPropCarf.Col = uneCol
    TabPropCarf.Row = uneRow
    
    If Val(TabPropCarf.Text) < 1 Or Val(TabPropCarf.Text) >= Val(DuréeCycle.Text) Then
        'Test que la valeur de la durée de vert (colonne 3) est
        'entre 1 et la durée du cycle
        unMsg = "La durée de vert doit être >= 1 et < Durée du cycle, qui vaut " + DuréeCycle.Text
        unMsg = unMsg + Chr(13) + Chr(13) + "OndeV lui donnera comme valeur la moitié de la durée du cycle"
        MsgBox unMsg, vbCritical, "Message d'erreur de OndeV"
        'Positionnement sur la cellule initialement active
        TabPropCarf.Col = uneCol
        TabPropCarf.Row = uneRow
        TabPropCarf.Text = Format(Val(DuréeCycle.Text) / 2)
        'Positionnement sur la cellule initialement active
        TabPropCarf.Col = uneCol
        TabPropCarf.Row = uneRow
        TabPropCarf.Action = SS_ACTION_ACTIVE_CELL
    End If
    'Affectation d'une valeur valide dans l'instance
    monCarrefourCourant.mesFeux(uneRow).maDuréeDeVert = Val(TabPropCarf.Text)
End Sub












Private Sub TextPoidsD_Change()
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub

Private Sub TextPoidsD_KeyUp(KeyCode As Integer, Shift As Integer)
    SaisieEntierPositifEntreMinMax KeyCode, TextPoidsD, 1, 1, 10, "Le poids du sens descendant"
    monPoidsSensD = Val(TextPoidsD.Text)
End Sub


Private Sub TextPoidsM_Change()
    'Stockage d'une modif dans les données du calcul d'onde
    maModifDataOnde = True
End Sub

Private Sub TextPoidsM_KeyUp(KeyCode As Integer, Shift As Integer)
    SaisieEntierPositifEntreMinMax KeyCode, TextPoidsM, 1, 1, 10, "Le poids du sens montant"
    monPoidsSensM = Val(TextPoidsM.Text)
End Sub


Private Sub TextTDep_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SaisieEntierPositifEntreMinMax(KeyCode, TextTDep, 0, 0, maDuréeDeCycle, "L'instant de départ du transport collectif")
    mesTC(ComboTC.ListIndex + 1).monTDep = Val(TextTDep.Text)
    IndiquerModifTC
End Sub



Public Sub DessinerArretTC(unNumTC As Long, unYreel As Long)
    Dim unePos As Long
    
    'Conversion du Yréel en Y écran dans la FrameVisuCarf
    unePos = ConvertirReelEnEcran(monYMaxFeu - unYreel, maLongueurAxeY, AxeOrdonnée.Y2 - AxeOrdonnée.Y1)
    'Numérotation de l'arrêt TC
    unNumArret = mesTC(unNumTC).mesArrets.Count
    'Incrémentation du nombre d'objets graphiques TC créés
    monNbObjGraphicTC = monNbObjGraphicTC + 1
    i = monNbObjGraphicTC
    'Création du label pour le nom de l'arrêt TC
    Load NomArret(i)
    NomArret(i).ForeColor = mesTC(unNumTC).maCouleur
    Call ModifierCaptionLabel(mesTC(unNumTC).monNom, NomArret(i), unYreel)
    NomArret(i).Top = unePos + AxeOrdonnée.Y1 - NomArret(i).Height
    'Ajustement de la chaine de caractéres à l'axe des ordonnées
    NomArret(i).Width = AxeOrdonnée.X1 - NomArret(i).Left
    NomArret(i).Visible = True
    'Création de l'icone graphique STOP de l'arrêt TC
    Load IconeArret(i)
    IconeArret(i).Top = unePos + AxeOrdonnée.Y1 - IconeArret(i).Height
    IconeArret(i).Visible = True
    'Codage permettant de retrouver le TC et son arrêt à partir des objets graphiques
    'Tag = index dans la collection des TC plus un tiret et le numéro de l'arrêt
    NomArret(i).Tag = Format(unNumTC) + "-" + Format(unNumArret)
    IconeArret(i).Tag = NomArret(i).Tag
    'Stockage dans la liste des objets graphiques représentant les arrêts du TC
    mesTC(unNumTC).mesObjGraphics.Add NomArret(i)
    'Le nouvel arrêt créé est sélectionné
    MiseAJourSelectionParCellule Me, ArretSel, unNumTC, TabYArret.ActiveCol
End Sub



Private Sub TextTransDec_KeyUp(KeyCode As Integer, Shift As Integer)
    If TextTransDec.Text = "" Then
        'Cas où l'on supprime tous les caractères on remet à 0
        TextTransDec.Text = "0"
    End If
    'Suppression des 0 restants éventuellement en tête
    TextTransDec.Text = Format(Val(TextTransDec.Text))
    'Si un seul caractère on met le curseur à la fin
    If Len(TextTransDec.Text) = 1 Then TextTransDec.SelStart = 1
End Sub

Private Sub TextTransDec_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'Cas des touches Entrée ou Retour chariot
        'On applique la translation avec la valeur de TextTransDec
        TranslaterDecalages
    Else
        Call VerifSaisieEntier(KeyAscii, TextTransDec)
    End If
End Sub


Private Sub TextVitD_KeyUp(KeyCode As Integer, Shift As Integer)
    SaisieEntierPositifEntreMinMax KeyCode, TextVitD, 30, 1, 99, "La vitesse du sens descendant"
    If maVitSensD <> Val(TextVitD.Text) Then
        'Stockage d'une modif dans les données du calcul d'onde
        maModifDataOnde = True
        maVitSensD = Val(TextVitD.Text)
    End If
End Sub


Private Sub TextVitM_KeyUp(KeyCode As Integer, Shift As Integer)
    SaisieEntierPositifEntreMinMax KeyCode, TextVitM, 30, 1, 99, "La vitesse du sens montant"
    If maVitSensM <> Val(TextVitM.Text) Then
        'Stockage d'une modif dans les données du calcul d'onde
        maModifDataOnde = True
        maVitSensM = Val(TextVitM.Text)
    End If
End Sub


Private Sub TitreEtude_Change()
    monTitreEtude = TitreEtude.Text
    maModifDataSite = True
End Sub


Private Sub ModifierObjGraphicTC(unTypeModif As Integer, Optional unAncienNom As String)
    'Modification des objets graphiques d'un TC
    'Si unTypeModif : = ModifNomTC ==> Modification du nom dans les labels
    '                 = ModifColTC ==> Modification de la couleur
    '                 = SupprTC    ==> Suppression du TC
    Dim unObjGraphic As Control
    Dim unInd As Integer
    Dim unePos As Integer
    Dim uneStrAvantNomTC As String
    
    If unTypeModif = ModifColTC Then
        mesTC(ComboTC.ListIndex + 1).maCouleur = ColorTC.BackColor
    End If
    
    'Parcours de tous les objets graphiques du TC, donc les controls NomArret
    unInd = ComboTC.ListIndex + 1
    unNbArret = mesTC(unInd).mesObjGraphics.Count
    For i = 1 To unNbArret
        'Récupération du label NomArret rangé dans la collection mesObjgraphics
        Set unObjGraphic = mesTC(unInd).mesObjGraphics(i)
        If unTypeModif = ModifNomTC Then
            'Recherche du nom du TC dans la propriété Caption
            unePos = InStr(1, unObjGraphic.Caption, unAncienNom)
            'Extraction des caractéres se trouvant avant le nom du TC dans la propriété Caption
            If unePos = 1 Then
                uneStrAvantNomTC = ""
            Else
                uneStrAvantNomTC = Mid$(unObjGraphic.Caption, 1, unePos - 1)
            End If
            'Modification de la propriété caption
            unObjGraphic.Caption = uneStrAvantNomTC + mesTC(unInd).monNom + StringTrait
            'Ajustement de la chaine de caractéres à l'axe des ordonnées
            unObjGraphic.Width = AxeOrdonnée.X1 - unObjGraphic.Left
        ElseIf unTypeModif = ModifColTC Then
            unObjGraphic.ForeColor = ColorTC.BackColor
        ElseIf unTypeModif = SupprTC Then
            Unload IconeArret(unObjGraphic.Index)
            Unload NomArret(unObjGraphic.Index)
        End If
    Next i
End Sub


Public Function RechercherArretConfondu(unY As Long, uneListeIndexTC As Collection, uneListeIndexArret As Collection, Optional unNbTC As Integer = -1)
    'Recherche des arrêts de TC confondus à l'ordonnée unY, valeur réelle en
    'mètres, parmi les TC choisis grâce au paramètre unNbTC.
    'Si unNbTC = -1, paramètre par défaut ==> Recherche sur tous les TC du site
    'sinon ==> Recherche sur les TC du site compris entre le 1er et le numéro unNbTC
    'Retourne 0 si pas d'arrêt confondu, le nombre d'arrêts confondus sinon
    'uneListeTC et uneListeIndexArret contiendront la liste des TC avec les arrêts
    'confondus en Y
    Dim unTC As TC
    
    'Initialisation
    RechercherArretConfondu = 0
        
    'Choix des TC à tester
    If unNbTC = -1 Then
        'Cas où l'on parcourt tous les TC du site
        'C'est le choix par défaut
        unNbTC = mesTC.Count
    End If
    
    'Parcours de tous les TC choisis
    For i = 1 To unNbTC
        Set unTC = mesTC(i)
        'Parcours de tous les arrêts du TC
        For j = 1 To unTC.mesArrets.Count
            'Test d'égalité, ce sont des entiers
            If unTC.mesArrets(j).monOrdonnee = unY Then
                RechercherArretConfondu = RechercherArretConfondu + 1
                'Stockage des index des TC avec leurs arrêts trouvés
                uneListeIndexTC.Add i
                uneListeIndexArret.Add j
            End If
        Next j
    Next i
End Function

Private Sub ModifierCaptionLabel(unNomTC As String, unLabel As Label, unY As Long)
    Dim uneListeIndexTC As New Collection
    Dim uneListeIndexArret As New Collection
    Dim unNbArretsConfondus As Integer
    
    'Initialisation des variables locales
    unNbArretsConfondus = 0
    
    'Recherche des arrêts confondus pour affichage cote à cote
    unNbArretsConfondus = RechercherArretConfondu(unY, uneListeIndexTC, uneListeIndexArret)
    unLabel.Caption = ""
    For j = 2 To unNbArretsConfondus 'Début à 2 pour ne pas tenir compte que des arrêts différents de celui courant, celui d'ordonnée unY
        'On décale le nom d'autant de caractères qu'il y a de TC
        'sachant qu'on a droit à 5 caractères pour le nom d'un TC
        unLabel.Caption = unLabel.Caption + DonnerStringDecalage
    Next j
    'Modification du label contenant le nom TC
    unLabel.Caption = unLabel.Caption + unNomTC + StringTrait
End Sub

Public Sub NewOrRenameTC(uneAction As String)
    Dim unMsg As String, unTitre As String
    Dim unAncienNom As String, unNomTC As String
    Dim uneValeurDefaut As String
    Dim unYNew As Integer
    Dim unTC As TC, unArret As ArretTC
    Dim unCarfDep As Carrefour, unCarfArr As Carrefour
       
    If uneAction = "new" Then
        'Cas de la création d'un TC
        uneValeurDefaut = ""
        unMsg = "nom du nouveau"
        unTitre = "Création d'un " ' Définit le titre.
    ElseIf uneAction = "rename" Then
        'Cas du renommage d'un TC
        'recherche de la position dans la liste du TC à renommer
        unePos = ComboTC.ListIndex
        unAncienNom = mesTC(unePos + 1).monNom
        uneValeurDefaut = unAncienNom
        unMsg = "nouveau nom du"
        unTitre = "Changement du nom d'un " ' Définit le titre.
   End If
    
    ' Définit le message.
    unMsg = "Entrez le " + unMsg + " transport collectif (" + Format(NbCarMaxNomTC) + " caractères maximun)"
    unTitre = unTitre + "transport collectif" ' Définit le titre.
    
    ' Affiche le message, le titre et la valeur par défaut.
    Do
        unNomTC = InputBox(unMsg, unTitre, uneValeurDefaut)
        unNomTC = Trim(unNomTC) 'Suppression des blancs avant et après
        uneValeurDefaut = unNomTC
        If Len(unNomTC) > NbCarMaxNomTC Then
            unMsg1 = "Le nom du transport collectif est limité à " + Format(NbCarMaxNomTC) + " caractères"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        ElseIf Trim(unNomTC) = "" Then
            'Cas du click sur le bouton annuler ou sur OK sans rentrer de nom
            '==> Sortie sans rien faire comme un annuler
            uneSortie = True
        ElseIf PosInListe(unNomTC, ComboTC) <> -1 Then
            'Cas où le nom existe déjà
            unMsg1 = "Le transport collectif " + UCase(unNomTC) + " existe déjà"
            MsgBox unMsg1, vbCritical
            uneSortie = False
        Else
            uneSortie = True
            'Désélection de la sélection graphique précédente
            Call Deselectionner(Me)
            
            If uneAction = "new" Then
                'Cas de la création d'un TC
                'Désinhibition des boutons de TC
                RenameTC.Enabled = True
                DelTC.Enabled = True
                'Visualisation de la frame TC
                FrameTC.Visible = True
                FrameTC.Refresh
                'Calcul de la valeur par défaut du Y du premier arrêt
                'Positionner au quart de l'axe des Y quelque soit le zoom
                unYNew = (monYMaxFeu + monYMinFeu) / 4
                'Création de l'arrêt numéro 1
                TabYArret.MaxCols = 1
                TabYArret.Row = 1
                TabYArret.Col = 1
                TabYArret.Text = Format(unYNew)
                'Création d'une instance de TC
                Set unCarfDep = mesCarrefours(1) 'Premier carrefour
                Set unCarfArr = mesCarrefours(mesCarrefours.Count) 'Dernier carrefour
                Set unTC = mesTC.Add(unNomTC, 0, 30, 8, unCarfDep, unCarfArr, 255)
                Set unArret = unTC.mesArrets.Add(unYNew, 15, 30, "Arrêt 1 de " + unNomTC)
                'Dessin des objets graphiques de l'arrêt TC numéro 1
                DessinerArretTC ComboTC.ListCount + 1, CLng(unYNew)
                'Mise à jour de la combobox listant les TC
                ComboTC.AddItem unNomTC
                ComboTC.ListIndex = ComboTC.ListCount - 1
                
                'Mise à jour des combobox TC pour l'onde verte TC
                RemplirComboboxOndeTC Me, unTC
                
                'Indication d'une modification dans les données TC
                maModifDataTC = True
            ElseIf uneAction = "rename" Then
                'Cas du renommage d'un TC
                'Suppression de l'item correspondant à l'ancien nom
                ComboTC.RemoveItem unePos
                'Création en ajoutant le nouveau nom dans la liste à la même
                'position que l'ancien nom
                ComboTC.AddItem unNomTC, unePos
                ComboTC.ListIndex = unePos
                'Changement du nom de l'instance TC
                mesTC(unePos + 1).monNom = unNomTC
                'Changement des labels de tous les arrêts du TC courant
                Call ModifierObjGraphicTC(ModifNomTC, unAncienNom)
                'Mise à jour du nom dans les listes de TC montant et descendant
                If DonnerYCarrefour(mesTC(unePos + 1).monCarfDep) < DonnerYCarrefour(mesTC(unePos + 1).monCarfArr) Then
                    'Supression dans la liste des TC montant
                    i = -1
                    Do
                        i = i + 1
                    Loop Until unAncienNom = ComboTCM.List(i)
                    ComboTCM.RemoveItem i
                    ComboTCM.AddItem unNomTC, i
                Else
                    'Supression dans la liste des TC descendant
                    i = -1
                    Do
                        i = i + 1
                    Loop Until unAncienNom = ComboTCD.List(i)
                    ComboTCD.RemoveItem i
                    ComboTCD.AddItem unNomTC, i
                End If
                'Indication d'une modification dans les données du site et
                'pas TC car le changement de nom n'influence pas les calculs
                maModifDataSite = True
            End If
        End If
    Loop While uneSortie = False
End Sub

Private Sub AffichageOngletVisu()
    'Positionnement de la zone de dessin de l'onglet Graphique
    unEspacement = 120
    ZoneDessin.Top = TabFeux.TabHeight + unEspacement / 2
    ZoneDessin.Height = TabFeux.Height - TabFeux.TabHeight - unEspacement
    ZoneDessin.Left = unEspacement / 2
    ZoneDessin.Width = TabFeux.Width - unEspacement
    
    'Positionnement de l'axe des temps en face du bas de l'axe des ordonnées
    AxeTemps.Y1 = ZoneDessin.Height + unEspacement / 4 * 3 - (FrameVisuCarf.Height - AxeOrdonnée.Y2)
    'le unEp* 3/4 pour avoir l'axe des temps plus bas que le min des Y
    AxeTemps.Y2 = AxeTemps.Y1
    AxeTemps.X1 = unEspacement / 2
    AxeTemps.X2 = ZoneDessin.Width - unEspacement / 2
    LabelFleche.Left = AxeTemps.X2 - LabelFleche.Width / 2
    LabelFleche.Top = AxeTemps.Y1 - LabelFleche.Height / 2
End Sub


Private Sub UpDownSensD_Change()
    monPoidsSensD = Val(TextPoidsD.Text)
End Sub

Private Sub UpDownSensM_Change()
    monPoidsSensM = Val(TextPoidsM.Text)
End Sub

Public Sub InitIndiqModif()
    'Initialisation des variables indiquant des modifications de valeurs dans
    'les carrefours, TC, calculs d'onde, les décalages et la visu graphique
    maModifDataSite = False
    maModifDataCarf = False
    maModifDataTC = False
    maModifDataOndeTC = False
    maModifDataOnde = False
    maModifDataDec = False
    maModifDataDes = False
End Sub

Private Sub ZoneDessin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Sélection graphique si bouton gauche appuyé
        'Pas de sélection multiple
        SelectionGraphique ZoneDessin, X, Y
        'Indication que le bouton souris gauche enfoncé
        ZoneDessin.Tag = "DownBtnG"
    End If
End Sub

Private Sub ZoneDessin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Cas d'un MouseMove avec bouton gauche enfoncé ==> Modif Interactive
        ModifierSelection ZoneDessin, X, Y
    Else
        'Changement du pointeur souris en croix si on passe
        'sur les poignées de sélection si elles sont visibles
        ChangerPointeurSouris ZoneDessin, X, Y
        'Affichage dans la 1ère zone de la barre d'état lors du mouvement souris
        unMsg = "Sélection possible des plages de vert coupant les ondes vertes de même sens"
        frmMain.sbStatusBar.Panels.Item(1).Text = unMsg
    End If
End Sub

Private Sub ZoneDessin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        'Mise à jour après les modifications graphiques interactives
        MettreAJourSelection ZoneDessin, X
        'Indication que le bouton souris gauche n'est plus enfoncé
        ZoneDessin.Tag = ""
    ElseIf Button = 2 And ZoneDessin.Tag = "" Then
        'Si click bouton droit et bouton gauche pas enfoncé
        'Affichage du menu contextuel "Graphique onde verte"
        frmMain.PopupMenu frmMain.mnuGraphicOnde, vbPopupMenuRightButton
    End If
End Sub

