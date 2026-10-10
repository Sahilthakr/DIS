VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form micro_myco 
   Caption         =   "Form2"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form2"
   ScaleHeight     =   6885
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   1575
      Left            =   1800
      OleObjectBlob   =   "micro_myco.frx":0000
      TabIndex        =   7
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Test Performed"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "MICROBIOLOGY REPORT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "05/11/21"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "micro_myco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label6.Caption = micro_report.Label7

End Sub

Private Sub Form_Load()
'DBGrid1.Col = 3
DBGrid1.Text = "saaaaaaad"
'DBGrid1.Caption = "sasasa"
'DBGrid1.Columns = 3


End Sub
