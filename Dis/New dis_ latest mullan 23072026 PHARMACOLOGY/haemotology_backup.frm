VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Differential Count in % (Manual)"
      Height          =   2895
      Left            =   0
      TabIndex        =   31
      Top             =   4320
      Width           =   11295
      Begin VB.TextBox txt_mneutrophils 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_meosinophils 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_mbasophills 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   50
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt_mmonocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt_alymphocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_promyelocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   47
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_myelocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   46
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_metamyelocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_prolymphocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txt_blasts 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   43
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_hairycells 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   42
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_plasmacells 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   41
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_Sezarycells 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   40
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_othercells 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   39
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox TXT_NORMOBLASTS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txt_mlymphocytes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_rbcmorphology 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   36
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txt_reticcount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   35
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt_malarialparasite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txt_filarialparasite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   33
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   32
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eosinophils"
         Height          =   195
         Left            =   3720
         TabIndex        =   73
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dmneutrophils 
         AutoSize        =   -1  'True
         Caption         =   "Neutrophils"
         Height          =   195
         Left            =   600
         TabIndex        =   72
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_mbasophills 
         AutoSize        =   -1  'True
         Caption         =   "Basophills"
         Height          =   195
         Left            =   6480
         TabIndex        =   71
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lbl_mmonocytes 
         AutoSize        =   -1  'True
         Caption         =   "Monocytes"
         Height          =   195
         Left            =   9000
         TabIndex        =   70
         Top             =   600
         Width           =   930
      End
      Begin VB.Label lbl_mlymphocytes 
         AutoSize        =   -1  'True
         Caption         =   "Lymphocytes"
         Height          =   195
         Left            =   360
         TabIndex        =   69
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lbl_matyplicallymphocytes 
         AutoSize        =   -1  'True
         Caption         =   "Atyplical Lymphocytes"
         Height          =   195
         Left            =   2760
         TabIndex        =   68
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl_promyelocytes 
         AutoSize        =   -1  'True
         Caption         =   "Promyelocytes"
         Height          =   195
         Left            =   6120
         TabIndex        =   67
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label lbl_myelocytes 
         AutoSize        =   -1  'True
         Caption         =   "Myelocytes"
         Height          =   195
         Left            =   9000
         TabIndex        =   66
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lbl_metamyelocytes 
         AutoSize        =   -1  'True
         Caption         =   "Metamyelocytes"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lbl_prolymphocytes 
         AutoSize        =   -1  'True
         Caption         =   "Prolymphocytes"
         Height          =   195
         Left            =   3240
         TabIndex        =   64
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl_blasts 
         AutoSize        =   -1  'True
         Caption         =   "Blasts"
         Height          =   195
         Left            =   6720
         TabIndex        =   63
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lbl_hairycells 
         Caption         =   "Hairy Cells"
         Height          =   255
         Left            =   9000
         TabIndex        =   62
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl_plasmacells 
         AutoSize        =   -1  'True
         Caption         =   "Plasma Cells"
         Height          =   195
         Left            =   480
         TabIndex        =   61
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label lbl_sezarycells 
         AutoSize        =   -1  'True
         Caption         =   "Sezary Cells"
         Height          =   195
         Left            =   3480
         TabIndex        =   60
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lbl_othercells 
         AutoSize        =   -1  'True
         Caption         =   "Other Cells"
         Height          =   195
         Left            =   6360
         TabIndex        =   59
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label lbl_normoblastsper100wbc 
         AutoSize        =   -1  'True
         Caption         =   "Normoblasts Per 100WBC"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   1800
         Width           =   2190
      End
      Begin VB.Label LBL_RBCMORPHOLOGY 
         AutoSize        =   -1  'True
         Caption         =   "RBC Morphology"
         Height          =   195
         Left            =   960
         TabIndex        =   57
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label lbl_reticcount 
         AutoSize        =   -1  'True
         Caption         =   "Retic Count [%]"
         Height          =   195
         Left            =   6000
         TabIndex        =   56
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lbl_malarialparasite 
         AutoSize        =   -1  'True
         Caption         =   "Malarial Parasite"
         Height          =   195
         Left            =   960
         TabIndex        =   55
         Top             =   2520
         Width           =   1425
      End
      Begin VB.Label lbl_filarialparasite 
         AutoSize        =   -1  'True
         Caption         =   "Filarial Parasite"
         Height          =   195
         Left            =   6000
         TabIndex        =   54
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "P. Morphology"
         Height          =   195
         Left            =   6120
         TabIndex        =   53
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      Begin VB.Frame Frame5 
         Caption         =   "Counts"
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5295
         Begin VB.CheckBox Check24 
            Caption         =   "Platelets"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "TLC"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Differential Count (counter)"
            Height          =   195
            Left            =   2040
            TabIndex        =   28
            Top             =   360
            Width           =   2655
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Hb"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "RBC Indicies"
            Height          =   195
            Left            =   2040
            TabIndex        =   26
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Manual"
         Height          =   1455
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   4935
         Begin VB.CheckBox Check6 
            Caption         =   "Platelets"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox Check7 
            Caption         =   "TC"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Differential Count (counter)"
            Height          =   255
            Left            =   2160
            TabIndex        =   22
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Hb"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Morphology"
            Height          =   255
            Left            =   2160
            TabIndex        =   20
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Haemogram"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1575
      End
      Begin VB.Frame Frame7 
         Caption         =   "Coagulation Profile"
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   6855
         Begin VB.CheckBox Check11 
            Caption         =   "BT"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check12 
            Caption         =   "CT"
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox Check13 
            Caption         =   "PT"
            Height          =   255
            Left            =   2520
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check14 
            Caption         =   "PTTK"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox Check15 
            Caption         =   "TT"
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   840
            Width           =   1455
         End
         Begin VB.Frame Frame8 
            Height          =   1215
            Left            =   4080
            TabIndex        =   10
            Top             =   120
            Width           =   2655
            Begin VB.CheckBox Check16 
               Caption         =   "FIBRINOGEN"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1575
            End
            Begin VB.CheckBox Check17 
               Caption         =   "FDP"
               Height          =   375
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   975
            End
         End
      End
      Begin VB.CheckBox Check19 
         Caption         =   "RBC Morphology"
         Height          =   195
         Left            =   4440
         TabIndex        =   8
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Others"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Frame Frame9 
         Height          =   1455
         Left            =   7080
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
         Begin VB.CheckBox Check18 
            Caption         =   "RBC Indicies(CH)"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Platelet  Morphology"
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CheckBox Check22 
         Caption         =   "ESR"
         Height          =   255
         Left            =   8520
         TabIndex        =   2
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox Check23 
         Caption         =   "PCV"
         Height          =   255
         Left            =   9600
         TabIndex        =   1
         Top             =   3360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
