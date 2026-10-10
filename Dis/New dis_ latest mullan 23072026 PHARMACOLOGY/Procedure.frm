VERSION 5.00
Begin VB.Form Procedure 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Procedure Selection"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   ControlBox      =   0   'False
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
   ScaleHeight     =   3315
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox list_procedure 
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmd_remove 
      Caption         =   "Remove <<"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "Add >>"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmb_procode 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   1350
   End
   Begin VB.ComboBox cmb_procedure 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lbl_procode 
      AutoSize        =   -1  'True
      Caption         =   "Procedure Code"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label lbl_procedure 
      AutoSize        =   -1  'True
      Caption         =   "Name of Procedure"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1650
   End
End
Attribute VB_Name = "Procedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()

End Sub

Private Sub cmd_ok_Click()
Unload Me
End Sub
