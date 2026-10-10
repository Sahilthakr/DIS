VERSION 5.00
Begin VB.Form microbiology 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
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
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Microbiology 
      Caption         =   "Microbiology"
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
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmb_mbprocedure 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox cmb_mbprocode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox txt_mbonantib 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txt_mbspecimen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmb_mboutspecimen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7920
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmb_mboutsblock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox cmb_mboutslide 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7920
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_mbexamreq 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cmb_mbprevunv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txt_mbreqno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7920
         TabIndex        =   2
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txt_labidno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7920
         TabIndex        =   1
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lbl_mbprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lbl_mbprocode 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6360
         TabIndex        =   21
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lbl_mbongant 
         AutoSize        =   -1  'True
         Caption         =   "Ongoing Antibiotics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label lbl_mbspecimendesc 
         AutoSize        =   -1  'True
         Caption         =   "Specimen Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label lbl_mboutblock 
         AutoSize        =   -1  'True
         Caption         =   "Outside Block"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lbl_mboutslides 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6600
         TabIndex        =   17
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label lbl_mboutspeci 
         AutoSize        =   -1  'True
         Caption         =   "Outside Speciemen/Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   16
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label lbl_mbexamreq 
         AutoSize        =   -1  'True
         Caption         =   "Examination Required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   1860
      End
      Begin VB.Label lbl_mbpreinv 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label lbl_mbreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6480
         TabIndex        =   13
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label lbl_labno 
         AutoSize        =   -1  'True
         Caption         =   "Lab ID No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6840
         TabIndex        =   12
         Top             =   2640
         Width           =   945
      End
   End
End
Attribute VB_Name = "microbiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
