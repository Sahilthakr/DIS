VERSION 5.00
Begin VB.Form FRMCHGPWD 
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_MSG 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   6120
      Width           =   5895
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox TXTRENEW 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3000
      Width           =   1572
   End
   Begin VB.TextBox TXTNEW 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
      Width           =   1572
   End
   Begin VB.TextBox TXTOLD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   1572
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "CHANGE &PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      Caption         =   "Message :"
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
      Left            =   840
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   8160
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   9345
   End
   Begin VB.Label LBLRENEW 
      AutoSize        =   -1  'True
      Caption         =   "RE-TYPE  NEW PASSWORD"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label LBLNEW 
      AutoSize        =   -1  'True
      Caption         =   "NEW PASSWORD"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1770
   End
   Begin VB.Label LBLOLD 
      AutoSize        =   -1  'True
      Caption         =   "OLD PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   1464
   End
End
Attribute VB_Name = "FRMCHGPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALUSR As Boolean
Dim RS_USR As rdoResultset
Dim A_B As String
Dim STRSQL As String
Dim MYRS As rdoResultset
Dim RSPWD As rdoResultset
Public usrid, PASSWORD, usrgroup As String


Private Sub Back_Click()
Unload Me
'SELMNU.Show
'tmh_ris_mnu.show
End Sub

Private Sub cmd_exit_Click()
Unload Me
End
End Sub

Private Sub cmdCancel_Click()
 Unload Me
 tmh_ris_mnu.show
End Sub

Private Sub CMDOK_Click()
Dim MSG As String
On Error GoTo MSG_1:

If Len(TXTRENEW.Text) > 0 Then
TXT_MSG.Text = "WORKING..."
DoEvents

Screen.MousePointer = 11
usrid = UID

PASSWORD = Trim(TXTRENEW.Text)
usrgroup = Trim(USRGRP)

'STRSQL = "Call DISLIB.pwdenc11('" & usrid & "','" & PASSWORD & "','" & Trim(usrgroup) & "')"
STRSQL = "Call DISLIB.PWDRPGL('" & Trim(usrid) & "','" & Trim(PASSWORD) & "','" & Trim(usrgroup) & "','" & Trim(USRDEPT) & "')"
Set RSPWD = rdb.OpenResultset(STRSQL, rdOpenKeyset, 3)

'MsgBox "RECORD UPDATED"

MSG = MsgBox("Password has changed Successfully !!", vbInformation, "PASSWORD CHANGED")

TXT_MSG.Text = ""

Else

MsgBox "PASSWORD CANNOT BE EMPTY"
TXTNEW.SetFocus

End If
Screen.MousePointer = 0
Exit Sub

MSG_1: MsgBox Error(Err)
End Sub

Private Sub Form_Activate()
TXTOLD.SetFocus
End Sub

Private Sub Form_Load()

CMDOK.Enabled = False

TXTOLD.Enabled = True
TXTNEW.Enabled = False
TXTRENEW.Enabled = False
LBLNEW.Visible = False
LBLRENEW.Visible = False
TXTNEW.Visible = False
TXTRENEW.Visible = False
TXTOLD.Text = ""
TXTNEW.Text = ""
TXTRENEW.Text = ""
TXT_MSG.Text = ""
'Me.Show
'TXTOLD.SetFocus
End Sub

Private Sub Form_Paint()
'TXTOLD.SetFocus

End Sub

Private Sub TXTNEW_LostFocus()
TXTRENEW.SetFocus
End Sub

Private Sub TXTOLD_LostFocus()
If Len(Trim(TXTOLD.Text)) <> 0 Then
usrid = UID
'PASSWORD = ""

'****
STRSQL = "CALL DISLIB.SEC_exam1( '" & UID & "' , '" & PASSWD & "')"
Set MYRS = rdb.OpenResultset(STRSQL, rdOpenKeyset, rdConcurValues)

'MsgBox MYRS.RowCount
If MYRS.RowCount > 0 Then
        TXTNEW.Visible = True
        TXTNEW.Enabled = True
        TXTRENEW.Visible = True
        TXTRENEW.Enabled = True
        LBLNEW.Visible = True
        LBLRENEW.Visible = True
        TXTNEW.Text = ""
        TXTRENEW.Text = ""
        TXTOLD.Enabled = False
        TXTNEW.SetFocus
    Else
        'MsgBox "WRONG PASSWORD ENTERED"
        MsgBox "OLD PASSWORD IS NOT CORRECT", vbCritical
    End If
Else
    cmd_exit.SetFocus
End If
End Sub

Private Sub TXTRENEW_LostFocus()
Dim MSG As String
Static RETRIALS As Integer

If Trim(TXTNEW.Text) <> Trim(TXTRENEW.Text) Then
    If RETRIALS < 2 Then
         MSG = MsgBox("SORRY NEW PASSWORDS DOSEN'T MATCH,TRY AGAIN", vbCritical, "CHANGE PASSWORD")
         RETRIALS = RETRIALS + 1
        TXTNEW.SelStart = 0
        TXTNEW.SelLength = Len(TXTNEW.Text)
        TXTRENEW.Text = ""
        TXTNEW.SetFocus
        Exit Sub
    Else
        MSG = MsgBox("SORRY  CANNOT CHANGE THE PASSWORD")
        Exit Sub
    End If
Else
         CMDOK.Enabled = True
        CMDOK.SetFocus
End If
End Sub
