VERSION 5.00
Begin VB.Form cytologyreq 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CYTOLOGY 
      Caption         =   "Gynaec/NonGynaec"
      ForeColor       =   &H000000C0&
      Height          =   2775
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox cmb_cnsite 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox CMB_PROCEDURE 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   29
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_scopyfind 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cbo_proccode 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6600
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmb_gynaec 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txt_cnatmeterial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cmb_cnoslides 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   1560
         Width           =   1350
      End
      Begin VB.TextBox txt_cmetastasis 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6120
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmb_claterality 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmb_coutslides 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9600
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txt_trtdetail 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9720
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmb_cpreinv 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9600
         TabIndex        =   4
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox TXT_LMP 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cmb_cyoutspecimen 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6120
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_cyreqno1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9720
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Scopy Finding"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   780
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5040
         TabIndex        =   25
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of procedure"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label lbl_Gynaec 
         AutoSize        =   -1  'True
         Caption         =   "Gynaec/Non Gynaec"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label lbl_csiteofspec 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1650
         TabIndex        =   22
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lbl_cnatumaterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label lbl_cnoslides 
         AutoSize        =   -1  'True
         Caption         =   "No of Slides"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   945
         TabIndex        =   20
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label lbl_coutspecimen 
         AutoSize        =   -1  'True
         Caption         =   "Outside Specimen"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4410
         TabIndex        =   19
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label lbl_cmetastasis 
         Caption         =   "Metastasis"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4980
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl_claterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lbl_outslide 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8280
         TabIndex        =   16
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label lbl_ctrtdetail 
         AutoSize        =   -1  'True
         Caption         =   "Treatment Details"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8160
         TabIndex        =   15
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label lbl_cpreinvest 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7560
         TabIndex        =   14
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label lbl_lmp 
         AutoSize        =   -1  'True
         Caption         =   "LMP"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5565
         TabIndex        =   13
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label lbl_cyreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8400
         TabIndex        =   12
         Top             =   360
         Width           =   1320
      End
   End
End
Attribute VB_Name = "cytologyreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
