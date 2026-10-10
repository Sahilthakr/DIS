VERSION 5.00
Begin VB.Form frm_micro_SampleAdd 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Microbiology Sample Addition Form"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   ScaleHeight     =   4635
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txt_sample 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frm_micro_SampleAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_sample_QR As rdoQuery
Dim sel_sample_RS As rdoResultset

Dim ins_sample_qr As rdoQuery

Dim str1 As String
Dim str2 As String

Private Sub cmd_clear_Click()
    txt_sample.Text = ""
End Sub

Private Sub Form_Load()
str1 = "SELECT * FROM DISLIB.microsampl WHERE SAMPLENAME=?"
Set sel_sample_QR = rdb.CreateQuery("", str1)

str2 = "insert into DISLIB.microsampl (SAMPLENAME) values (?)"
Set ins_sample_qr = rdb.CreateQuery("", str2)
End Sub


Private Sub cmd_add_Click()
sel_sample_QR(0) = Trim(txt_sample.Text)
Set sel_sample_RS = sel_sample_QR.OpenResultset(1, 3)

If sel_sample_RS.EOF Then
    ins_sample_qr(0) = Trim(txt_sample.Text)
    ins_sample_qr.Execute
    MsgBox "Sample added successfully", vbInformation, "Microbiology Sample"
Else
    MsgBox "Sample already exists", vbCritical, "Microbiology Sample"
End If
'sample_RS.Close
End Sub

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_exit_Click()
End
End Sub
