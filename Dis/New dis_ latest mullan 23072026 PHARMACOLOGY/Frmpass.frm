VERSION 5.00
Begin VB.Form Frmpass 
   Caption         =   "Form2"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password : "
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Frmpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Trim(Txtpass.Text)) = 0 Then
      MsgBox "Password cannot be blank"
      Txtpass.SetFocus
      Exit Sub
  End If
 

  Frmpass.Txtpass.Text = Txtpass.Text
   If Frmpass.Txtpass.Text = "supervise" Then
      memoduplicate = False
   Else
      Frmpass.Txtpass.Text = ""
      memoduplicate = True
   End If
  Me.Hide
  'op_memo_d1.show
  
End Sub

