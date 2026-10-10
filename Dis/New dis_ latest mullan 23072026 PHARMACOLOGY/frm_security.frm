VERSION 5.00
Begin VB.Form frm_security 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security Check"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Ok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Txt_password 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "#"
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Password Below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frm_security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'added by amarjeet on 09-11-2002 suggested by Dr. ganesh for password protected biochem report so that
'no one can modify it unless authorised.
'this form uses public variable password defined in class module


Private Sub cmd_Ok_Click()
PASSWORD = Txt_password.Text
Unload Me
End Sub

Private Sub Form_Load()
PASSWORD = ""
End Sub

