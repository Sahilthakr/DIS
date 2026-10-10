VERSION 5.00
Begin VB.Form cardiology 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
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
   ScaleHeight     =   3195
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame cardiology 
      Caption         =   "Cardiology"
      ForeColor       =   &H000000C0&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cmb_ecprocedure 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txt_ecprocode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_ecprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Frame HLATYPING 
      Caption         =   "HLA Typing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3135
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txt_hladonar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_hlarelpatient 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txt_hlaage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txt_hlasex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4320
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_hlaabcdr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_hlaabc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt_hlab27 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7680
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_hladnadrb1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txt_hladnarb1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7680
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_hladnadqa1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_hladnadqb1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7680
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_hladr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txt_hlareqno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9600
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmb_hlapreinv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lbl_hlanameofdonar 
         AutoSize        =   -1  'True
         Caption         =   "Name of the donar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label lbl_hlarelpatient 
         AutoSize        =   -1  'True
         Caption         =   "Relationship with patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   32
         Top             =   480
         Width           =   2115
      End
      Begin VB.Label lbl_hlage 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2640
         TabIndex        =   31
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lbl_hlasex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lbl_hlaabcdr 
         AutoSize        =   -1  'True
         Caption         =   "HLA ABC/DR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   29
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label LBL_HLAABC 
         AutoSize        =   -1  'True
         Caption         =   "HLA ABC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6720
         TabIndex        =   28
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label LBL_HLAB27 
         Caption         =   "HLA B 27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6720
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   15
      End
      Begin VB.Label lbl_hladnadrb1 
         AutoSize        =   -1  'True
         Caption         =   "HLA DNA-DRB1(low resolution)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1800
         Width           =   2670
      End
      Begin VB.Label lbl_hladnarb1 
         AutoSize        =   -1  'True
         Caption         =   "HLA DNA-RB1(Subtyping)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5280
         TabIndex        =   24
         Top             =   1800
         Width           =   2205
      End
      Begin VB.Label lbl_hlsdnadqa1 
         AutoSize        =   -1  'True
         Caption         =   "HLA DNA-DQA1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   23
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label LBL_HLSDNADQB1 
         AutoSize        =   -1  'True
         Caption         =   "HLA DNA DQB1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   22
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label LBL_HLADR 
         AutoSize        =   -1  'True
         Caption         =   "HLA DR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2280
         TabIndex        =   21
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label LBL_HLAREQNO 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7680
         TabIndex        =   20
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lbl_hlsprevinv 
         AutoSize        =   -1  'True
         Caption         =   "Previous investigations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Width           =   1980
      End
   End
End
Attribute VB_Name = "cardiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
