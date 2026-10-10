VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "PROJECT1.OCX"
Begin VB.Form cpl 
   Caption         =   "Clinical Pathology Laboratory"
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
   WindowState     =   2  'Maximized
   Begin Project1.UserControl1 UserControl11 
      Height          =   495
      Left            =   3600
      TabIndex        =   33
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   1680
      TabIndex        =   27
      Top             =   5520
      Width           =   8655
      Begin VB.CheckBox Check25 
         Caption         =   "Pregnancy test"
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check24 
         Caption         =   "Semen Analysis"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Malarial Parasite"
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Other body fluid cell count"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox Check21 
         Caption         =   "CSF Cell Count"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   21
      Top             =   7920
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6720
      TabIndex        =   19
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   8655
      Begin VB.CheckBox Check15 
         Caption         =   "PLATELETS"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox Check14 
         Caption         =   "FDP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check12 
         Caption         =   "FIBRINOGEN"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Check11 
         Caption         =   "TT"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check10 
         Caption         =   "PTTK"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Check9 
         Caption         =   "PT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox Check8 
         Caption         =   "CT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check7 
         Caption         =   "BT"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Caption         =   "COAGULATION PROFILE"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   8655
      Begin VB.CheckBox Check20 
         Caption         =   "MCHC"
         Height          =   255
         Left            =   7080
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check19 
         Caption         =   "MCH"
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check18 
         Caption         =   "MCV"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox Check17 
         Caption         =   "PCV"
         Height          =   195
         Left            =   5400
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check16 
         Caption         =   "ESR"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Differential Count"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Platelet"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "WBC"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "RBC"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hemoglobin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "COMPLETE BLOOD COUNT"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLINICAL LABORATORY SERVICES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3960
      TabIndex        =   34
      Top             =   600
      Width           =   5010
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6120
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LBL_CASENO 
      AutoSize        =   -1  'True
      Caption         =   "Case No."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "cpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
Unload Me
pathology.show
End Sub

